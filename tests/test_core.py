import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_privacy_tool.core import (
    PLAIN_FORMAT,
    WrongPasswordError,
    anonymize_workbook,
    load_mapping,
    restore_html,
)


class CoreWorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.source = self.root / "source.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "名簿"
        sheet.append(["氏名", "部署", "点数"])
        sheet.append(["山田 太郎", "営業", 80])
        sheet.append(["佐藤 花子", "開発", 90])
        sheet.append(["山田 太郎", "企画", 75])
        sheet["A5"] = None
        sheet["A6"] = "=B2"
        workbook.save(self.source)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_encrypted_round_trip(self):
        anonymized = self.root / "anon.xlsx"
        map_path = self.root / "data.epbmap"
        result = anonymize_workbook(
            self.source, anonymized, map_path, "名簿", ["A2:A6"], "correct horse"
        )
        self.assertEqual(result.replaced_cells, 3)
        self.assertEqual(result.unique_values, 2)
        self.assertEqual(result.skipped_blank_cells, 1)
        self.assertEqual(result.skipped_formula_cells, 1)

        workbook = load_workbook(anonymized, data_only=False)
        sheet = workbook["名簿"]
        self.assertRegex(sheet["A2"].value, r"^PII_[A-Z0-9]{6}_000001$")
        self.assertEqual(sheet["A2"].value, sheet["A4"].value)
        self.assertEqual(sheet["A6"].value, "=B2")

        mapping = load_mapping(map_path, "correct horse")
        token_yamada = sheet["A2"].value
        token_sato = sheet["A3"].value
        self.assertEqual(mapping["values"][token_yamada], "山田 太郎")
        with self.assertRaises(WrongPasswordError):
            load_mapping(map_path, "wrong")

        html_source = self.root / "analysis.html"
        html_source.write_text(
            f"<!doctype html><p>{token_yamada}: 高評価</p><p>{token_sato}</p>", encoding="utf-8"
        )
        restored = self.root / "restored.html"
        restore_result = restore_html(html_source, map_path, restored, "correct horse")
        restored_text = restored.read_text(encoding="utf-8")
        self.assertIn("山田 太郎", restored_text)
        self.assertIn("佐藤 花子", restored_text)
        self.assertEqual(restore_result.replacements, 2)
        self.assertFalse(restore_result.missing_expected_tokens)

    def test_plain_map_and_html_escaping(self):
        workbook = load_workbook(self.source)
        workbook["名簿"]["A2"] = "A&B <社長>"
        workbook.save(self.source)
        anonymized = self.root / "anon.xlsx"
        map_path = self.root / "plain.epbmap"
        anonymize_workbook(self.source, anonymized, map_path, "名簿", ["A2"], None)
        wrapper = json.loads(map_path.read_text(encoding="utf-8"))
        self.assertEqual(wrapper["format"], PLAIN_FORMAT)
        token = next(iter(wrapper["payload"]["values"]))
        html_source = self.root / "analysis.html"
        html_source.write_text(f"<p>{token}</p>", encoding="utf-8")
        restored = self.root / "restored.html"
        restore_html(html_source, map_path, restored)
        self.assertIn("A&amp;B &lt;社長&gt;", restored.read_text(encoding="utf-8"))

    def test_whole_column_range(self):
        anonymized = self.root / "column.xlsx"
        map_path = self.root / "column.epbmap"
        result = anonymize_workbook(self.source, anonymized, map_path, "名簿", ["A:A"], None)
        self.assertEqual(result.replaced_cells, 4)


if __name__ == "__main__":
    unittest.main()
