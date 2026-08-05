from __future__ import annotations

import base64
import html
import json
import os
import re
import secrets
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable

from cryptography.exceptions import InvalidTag
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries


APP_NAME = "個人情報変換ちゃん"
APP_VERSION = "0.2.1"
ENCRYPTED_FORMAT = "excel-privacy-bridge-map-v1"
PLAIN_FORMAT = "excel-privacy-bridge-map-plain-v1"
KDF_ITERATIONS = 600_000
TOKEN_RE = re.compile(r"\bPII_[A-Z0-9]{6}_[0-9]{6}\b")


class PrivacyToolError(Exception):
    pass


class WrongPasswordError(PrivacyToolError):
    pass


@dataclass
class AnonymizeResult:
    output_path: str
    map_path: str
    replaced_cells: int
    unique_values: int
    skipped_blank_cells: int
    skipped_formula_cells: int


@dataclass
class RestoreResult:
    output_path: str
    replacements: int
    missing_expected_tokens: list[str]
    unknown_tokens: list[str]


def _derive_key(password: str, salt: bytes, iterations: int = KDF_ITERATIONS) -> bytes:
    if not password:
        raise PrivacyToolError("暗号化にはパスワードが必要です。")
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(), length=32, salt=salt, iterations=iterations
    )
    return kdf.derive(password.encode("utf-8"))


def save_mapping(mapping: dict, path: str | Path, password: str | None) -> None:
    destination = Path(path)
    destination.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(mapping, ensure_ascii=False, separators=(",", ":")).encode("utf-8")

    if password:
        salt = os.urandom(16)
        nonce = os.urandom(12)
        key = _derive_key(password, salt)
        encrypted = AESGCM(key).encrypt(nonce, payload, ENCRYPTED_FORMAT.encode("ascii"))
        wrapper = {
            "format": ENCRYPTED_FORMAT,
            "kdf": "PBKDF2-HMAC-SHA256",
            "iterations": KDF_ITERATIONS,
            "salt": base64.b64encode(salt).decode("ascii"),
            "nonce": base64.b64encode(nonce).decode("ascii"),
            "ciphertext": base64.b64encode(encrypted).decode("ascii"),
        }
    else:
        wrapper = {"format": PLAIN_FORMAT, "payload": mapping}

    destination.write_text(json.dumps(wrapper, ensure_ascii=False, indent=2), encoding="utf-8")


def load_mapping(path: str | Path, password: str | None = None) -> dict:
    try:
        wrapper = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise PrivacyToolError(f"対応表を読み込めません: {exc}") from exc

    format_name = wrapper.get("format")
    if format_name == PLAIN_FORMAT:
        return wrapper["payload"]
    if format_name != ENCRYPTED_FORMAT:
        raise PrivacyToolError("未対応の対応表形式です。")
    if not password:
        raise WrongPasswordError("この対応表は暗号化されています。パスワードを入力してください。")

    try:
        salt = base64.b64decode(wrapper["salt"])
        nonce = base64.b64decode(wrapper["nonce"])
        ciphertext = base64.b64decode(wrapper["ciphertext"])
        iterations = int(wrapper.get("iterations", KDF_ITERATIONS))
        key = _derive_key(password, salt, iterations)
        payload = AESGCM(key).decrypt(
            nonce, ciphertext, ENCRYPTED_FORMAT.encode("ascii")
        )
        return json.loads(payload.decode("utf-8"))
    except (InvalidTag, ValueError, KeyError, json.JSONDecodeError) as exc:
        raise WrongPasswordError("パスワードが違うか、対応表が破損しています。") from exc


def mapping_is_encrypted(path: str | Path) -> bool:
    try:
        wrapper = json.loads(Path(path).read_text(encoding="utf-8"))
        return wrapper.get("format") == ENCRYPTED_FORMAT
    except (OSError, json.JSONDecodeError):
        return False


def normalize_ranges(ranges: Iterable[str]) -> list[str]:
    normalized: list[str] = []
    for raw in ranges:
        for item in raw.split(","):
            value = item.strip().upper().replace("$", "")
            if not value:
                continue
            try:
                range_boundaries(value)
            except ValueError as exc:
                raise PrivacyToolError(f"セル範囲が正しくありません: {value}") from exc
            normalized.append(value)
    if not normalized:
        raise PrivacyToolError("1つ以上のセル範囲を指定してください。")
    return list(dict.fromkeys(normalized))


def anonymize_workbook(
    input_path: str | Path,
    output_path: str | Path,
    map_path: str | Path,
    sheet_name: str,
    ranges: Iterable[str],
    password: str | None = None,
) -> AnonymizeResult:
    source = Path(input_path)
    destination = Path(output_path)
    mapping_destination = Path(map_path)
    if source.resolve() == destination.resolve():
        raise PrivacyToolError("入力ファイルと出力ファイルは別名にしてください。")
    normalized_ranges = normalize_ranges(ranges)

    try:
        workbook = load_workbook(source)
    except Exception as exc:
        raise PrivacyToolError(f"Excelファイルを開けません: {exc}") from exc
    if sheet_name not in workbook.sheetnames:
        raise PrivacyToolError(f"シートが見つかりません: {sheet_name}")

    sheet = workbook[sheet_name]
    prefix = secrets.token_hex(3).upper()
    original_to_token: dict[str, str] = {}
    token_to_original: dict[str, str] = {}
    cell_records: list[dict[str, str]] = []
    processed_coordinates: set[str] = set()
    replaced = blank = formulas = 0

    for cell_range in normalized_ranges:
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        min_col = min_col or 1
        min_row = min_row or 1
        max_col = max_col or sheet.max_column
        max_row = max_row or sheet.max_row
        for row in sheet.iter_rows(
            min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col
        ):
            for cell in row:
                if cell.coordinate in processed_coordinates:
                    continue
                processed_coordinates.add(cell.coordinate)
                value = cell.value
                if value is None or (isinstance(value, str) and not value.strip()):
                    blank += 1
                    continue
                if cell.data_type == "f" or (isinstance(value, str) and value.startswith("=")):
                    formulas += 1
                    continue
                original = str(value)
                token = original_to_token.get(original)
                if token is None:
                    token = f"PII_{prefix}_{len(original_to_token) + 1:06d}"
                    original_to_token[original] = token
                    token_to_original[token] = original
                cell.value = token
                cell_records.append({"sheet": sheet_name, "cell": cell.coordinate, "token": token})
                replaced += 1

    if replaced == 0:
        raise PrivacyToolError("置換できるセルがありませんでした。範囲を確認してください。")

    mapping = {
        "schema": 1,
        "app": APP_NAME,
        "app_version": APP_VERSION,
        "created_at": datetime.now(timezone.utc).isoformat(),
        "source_filename": source.name,
        "anonymized_filename": destination.name,
        "sheet": sheet_name,
        "ranges": normalized_ranges,
        "token_prefix": f"PII_{prefix}_",
        "values": token_to_original,
        "cells": cell_records,
    }

    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(destination)
        save_mapping(mapping, mapping_destination, password)
    except Exception as exc:
        if destination.exists():
            destination.unlink(missing_ok=True)
        raise PrivacyToolError(f"ファイルを保存できません: {exc}") from exc

    return AnonymizeResult(
        output_path=str(destination),
        map_path=str(mapping_destination),
        replaced_cells=replaced,
        unique_values=len(token_to_original),
        skipped_blank_cells=blank,
        skipped_formula_cells=formulas,
    )


def _decode_html(raw: bytes) -> tuple[str, str, bytes]:
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw[3:].decode("utf-8"), "utf-8", b"\xef\xbb\xbf"
    if raw.startswith(b"\xff\xfe"):
        return raw[2:].decode("utf-16-le"), "utf-16-le", b"\xff\xfe"
    if raw.startswith(b"\xfe\xff"):
        return raw[2:].decode("utf-16-be"), "utf-16-be", b"\xfe\xff"
    try:
        return raw.decode("utf-8"), "utf-8", b""
    except UnicodeDecodeError:
        try:
            return raw.decode("cp932"), "cp932", b""
        except UnicodeDecodeError as exc:
            raise PrivacyToolError("HTMLの文字コードを判定できません。UTF-8で保存し直してください。") from exc


def restore_html(
    html_path: str | Path,
    map_path: str | Path,
    output_path: str | Path,
    password: str | None = None,
) -> RestoreResult:
    try:
        text, encoding, bom = _decode_html(Path(html_path).read_bytes())
    except OSError as exc:
        raise PrivacyToolError(f"HTMLを読み込めません: {exc}") from exc
    mapping = load_mapping(map_path, password)
    values: dict[str, str] = mapping.get("values", {})
    if not values:
        raise PrivacyToolError("対応表に復元データがありません。")

    present_before = {token for token in values if token in text}
    replacements = 0
    for token in sorted(values, key=len, reverse=True):
        count = text.count(token)
        if count:
            # HTML本文で安全に表示できる形へ戻す。通常の日本語氏名は変化しない。
            text = text.replace(token, html.escape(str(values[token]), quote=True))
            replacements += count

    expected = set(values)
    missing = sorted(expected - present_before)
    unknown = sorted(set(TOKEN_RE.findall(text)) - expected)
    destination = Path(output_path)
    try:
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(bom + text.encode(encoding))
    except OSError as exc:
        raise PrivacyToolError(f"復元HTMLを保存できません: {exc}") from exc

    return RestoreResult(
        output_path=str(destination),
        replacements=replacements,
        missing_expected_tokens=missing,
        unknown_tokens=unknown,
    )
