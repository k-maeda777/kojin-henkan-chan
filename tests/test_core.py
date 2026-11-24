import json
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from excel_privacy_tool.core import (
    PLAIN_FORMAT,
    PersonColumns,
    PersonConflictError,
    PrivacyToolError,
    WrongPasswordError,
    anonymize_workbook,
    anonymize_workbook_multi,
    load_mapping,
    restore_html,
    restore_workbook,
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

    def _multi_sheet_source(self):
        path = self.root / "multi.xlsx"
        workbook = Workbook()
        a = workbook.active
        a.title = "営業部"
        a.append(["社員番号", "氏名", "売上"])
        a.append([1001, "山田 太郎", 500])
        a.append([1002, "佐藤 花子", 300])
        b = workbook.create_sheet("開発部")
        b.append(["部署", "名前", "ID", "評価"])
        b.append(["開発", "山田　太郎", "1001", "A"])  # 全角スペース・社員番号が文字列
        b.append(["開発", "鈴木 一郎", "1003", "B"])
        c = workbook.create_sheet("異動履歴")
        c.append(["氏名", "旧所属"])
        c.append(["佐藤花子", "営業"])  # スペースなし表記
        workbook.save(path)
        return path

    def test_multi_sheet_same_person_same_token(self):
        source = self._multi_sheet_source()
        anonymized = self.root / "multi_anon.xlsx"
        map_path = self.root / "multi.epbmap"
        result = anonymize_workbook_multi(
            source, anonymized, map_path,
            [("営業部", ["A2:B3"]), ("開発部", ["B2:C3"]), ("異動履歴", ["A2"])],
            "pw",
        )
        # 山田太郎/1001/佐藤花子 は複数シートに出るので、異なる値は 山田・1001・佐藤・1002・鈴木・1003 の6個
        self.assertEqual(result.unique_values, 6)
        self.assertEqual(result.replaced_cells, 4 + 4 + 1)
        self.assertEqual(result.replaced_by_sheet, {"営業部": 4, "開発部": 4, "異動履歴": 1})

        wb = load_workbook(anonymized)
        sales, dev, moves = wb["営業部"], wb["開発部"], wb["異動履歴"]
        # 氏名: 空白・全角半角が違っても同一人物 → 同じID
        self.assertEqual(sales["B2"].value, dev["B2"].value)
        self.assertEqual(sales["B3"].value, moves["A2"].value)
        # 社員番号: 数値1001と文字列"1001" → 同じID
        self.assertEqual(sales["A2"].value, dev["C2"].value)
        # 範囲外の列は変更されない
        self.assertEqual(sales["C2"].value, 500)
        self.assertEqual(dev["A2"].value, "開発")
        self.assertEqual(dev["D2"].value, "A")
        # 見出しも変更されない
        self.assertEqual(sales["B1"].value, "氏名")

        mapping = load_mapping(map_path, "pw")
        self.assertEqual(mapping["values"][sales["B2"].value], "山田 太郎")
        self.assertEqual([t["sheet"] for t in mapping["targets"]], ["営業部", "開発部", "異動履歴"])

        html_source = self.root / "multi.html"
        html_source.write_text(f"<p>{dev['B2'].value}</p><p>{moves['A2'].value}</p>", encoding="utf-8")
        restored = self.root / "multi_restored.html"
        restore_html(html_source, map_path, restored, "pw")
        text = restored.read_text(encoding="utf-8")
        self.assertIn("山田 太郎", text)
        self.assertIn("佐藤 花子", text)

    def test_multi_sheet_unknown_sheet(self):
        source = self._multi_sheet_source()
        with self.assertRaises(PrivacyToolError):
            anonymize_workbook_multi(
                source, self.root / "x.xlsx", self.root / "x.epbmap",
                [("営業部", ["B2"]), ("存在しない", ["A2"])], None,
            )

    def _person_source(self, second_id="1002"):
        path = self.root / "person.xlsx"
        workbook = Workbook()
        a = workbook.active
        a.title = "正社員"
        a.append(["社員番号", "氏名", "部署"])
        a.append([1001, "山田 太郎", "営業"])
        a.append([1002, "佐藤 花子", "開発"])
        b = workbook.create_sheet("評価")  # 氏名のみ
        b.append(["氏名", "評価"])
        b.append(["山田　太郎", "A"])
        b.append(["佐藤花子", "B"])
        c = workbook.create_sheet("研修")  # 社員番号のみ（列位置も違う）
        c.append(["講座", "メモ", "ID"])
        c.append(["安全", "", 1001])
        c.append(["接遇", "", second_id])
        workbook.save(path)
        return path

    def _person_specs(self):
        return [
            PersonColumns("正社員", name_col="B", id_col="A"),
            PersonColumns("評価", name_col="A"),
            PersonColumns("研修", id_col="C"),
        ]

    def test_person_columns_one_id_per_person(self):
        source = self._person_source()
        out, map_path = self.root / "p.xlsx", self.root / "p.epbmap"
        result = anonymize_workbook_multi(source, out, map_path, person_columns=self._person_specs())
        self.assertEqual(result.person_count, 2)
        self.assertFalse(result.conflicts)
        wb = load_workbook(out)
        staff, review, training = wb["正社員"], wb["評価"], wb["研修"]
        base = staff["B2"].value
        self.assertRegex(base, r"^PII_[A-Z0-9]{6}_000001$")
        self.assertEqual(review["A2"].value, base)  # 氏名のみのシートも同じID
        self.assertEqual(staff["A2"].value, base + "_E1")  # 社員番号は基本ID+枝番
        self.assertEqual(training["C2"].value, base + "_E1")  # 社員番号のみのシートも同じ
        self.assertEqual(staff["C2"].value, "営業")
        self.assertEqual(training["A2"].value, "安全")

        mapping = load_mapping(map_path)
        self.assertEqual(mapping["values"][base], "山田 太郎")
        self.assertEqual(mapping["values"][base + "_E1"], "1001")

        html_source = self.root / "p.html"
        html_source.write_text(f"<p>{base}さんの社員番号は{base}_E1</p>", encoding="utf-8")
        restored = self.root / "p_restored.html"
        restore_html(html_source, map_path, restored)
        self.assertIn("山田 太郎さんの社員番号は1001", restored.read_text(encoding="utf-8"))

    def test_person_with_two_employee_numbers_raises_then_merges(self):
        # 1001 の山田さんが、研修シートでは 2001 でも登場する（登用で社員番号が変わった）→ 行をまたいで
        # 同一人物と判定できるよう、評価シートにも社員番号列を追加して2001と山田さんを結ぶ
        source = self._person_source()
        workbook = load_workbook(source)
        workbook["評価"]["C1"] = "社員番号"
        workbook["評価"]["C2"] = 2001
        workbook.save(source)
        specs = self._person_specs()
        specs[1] = PersonColumns("評価", name_col="A", id_col="C")
        out, map_path = self.root / "c.xlsx", self.root / "c.epbmap"
        with self.assertRaises(PersonConflictError) as ctx:
            anonymize_workbook_multi(source, out, map_path, person_columns=specs)
        self.assertFalse(out.exists())  # アラート時は何も保存しない
        conflict = ctx.exception.conflicts[0]
        self.assertEqual(conflict.kind, "multi_id")
        self.assertEqual(sorted(conflict.values), ["1001", "2001"])
        self.assertIn("山田 太郎", conflict.context)

        result = anonymize_workbook_multi(
            source, out, map_path, person_columns=specs, allow_conflicts=True
        )
        self.assertEqual(len(result.conflicts), 1)
        self.assertEqual(result.person_count, 2)
        wb = load_workbook(out)
        base = wb["正社員"]["B2"].value
        self.assertEqual(wb["正社員"]["A2"].value, base + "_E1")  # 1001
        self.assertEqual(wb["評価"]["C2"].value, base + "_E2")  # 2001
        mapping = load_mapping(map_path)
        self.assertEqual(mapping["values"][base + "_E2"], "2001")

    def test_same_employee_number_with_two_names_is_conflict(self):
        source = self._person_source()
        workbook = load_workbook(source)
        workbook["正社員"]["B3"] = "山田 次郎"
        workbook["正社員"]["A3"] = 1001
        workbook.save(source)
        with self.assertRaises(PersonConflictError) as ctx:
            anonymize_workbook_multi(
                source, self.root / "n.xlsx", self.root / "n.epbmap",
                person_columns=[PersonColumns("正社員", name_col="B", id_col="A")],
            )
        self.assertEqual(ctx.exception.conflicts[0].kind, "multi_name")

    def test_person_columns_validation(self):
        source = self._person_source()
        for spec in (
            PersonColumns("正社員"),
            PersonColumns("正社員", name_col="A", id_col="A"),
            PersonColumns("正社員", name_col="1"),
            PersonColumns("存在しない", name_col="A"),
        ):
            with self.assertRaises(PrivacyToolError):
                anonymize_workbook_multi(
                    source, self.root / "v.xlsx", self.root / "v.epbmap", person_columns=[spec]
                )

    # ---- Excelの復元 -------------------------------------------------------------

    def _anonymized_person_book(self):
        source = self._person_source()
        anon, map_path = self.root / "anon1.xlsx", self.root / "map1.epbmap"
        anonymize_workbook_multi(
            source, anon, map_path, password="pw", person_columns=self._person_specs()
        )
        return source, anon, map_path

    def test_restore_workbook_round_trip_keeps_types(self):
        source, anon, map_path = self._anonymized_person_book()
        restored = self.root / "restored.xlsx"
        result = restore_workbook(anon, map_path, restored, "pw")
        self.assertEqual(result.restored_cells, 8)  # 正社員4 + 評価2 + 研修2
        self.assertFalse(result.unknown_tokens)

        wb = load_workbook(restored)
        staff, review, training = wb["正社員"], wb["評価"], wb["研修"]
        self.assertEqual(staff["A2"].value, 1001)  # 数値は数値のまま戻る
        self.assertIsInstance(staff["A2"].value, int)
        self.assertEqual(staff["B2"].value, "山田 太郎")
        self.assertEqual(staff["C2"].value, "営業")
        self.assertEqual(review["A2"].value, "山田 太郎")  # 表記ゆれは最初の表記に統一される
        self.assertEqual(training["C2"].value, 1001)
        self.assertEqual(staff["A1"].value, "社員番号")

    def test_restore_workbook_dates_floats_and_errors(self):
        from datetime import datetime

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "S"
        sheet.append(["名前", "入社日", "係数"])
        sheet.append(["山田", datetime(2026, 4, 10), 3.5])
        source = self.root / "types.xlsx"
        workbook.save(source)
        anon, map_path = self.root / "types_anon.xlsx", self.root / "types.epbmap"
        anonymize_workbook_multi(source, anon, map_path, [("S", ["A2:C2"])], "pw")
        restored = self.root / "types_restored.xlsx"
        restore_workbook(anon, map_path, restored, "pw")
        row = [c.value for c in load_workbook(restored)["S"][2]]
        self.assertEqual(row, ["山田", datetime(2026, 4, 10), 3.5])

        with self.assertRaises(WrongPasswordError):
            restore_workbook(anon, map_path, self.root / "x.xlsx", "wrong")
        with self.assertRaises(PrivacyToolError):
            restore_workbook(anon, map_path, anon, "pw")  # 上書きは不可
        with self.assertRaises(PrivacyToolError):  # 仮IDのないExcelは対応表と合わない
            restore_workbook(source, map_path, self.root / "y.xlsx", "pw")

    def test_restore_workbook_flags_unknown_token_in_text(self):
        source, anon, map_path = self._anonymized_person_book()
        wb = load_workbook(anon)
        token = wb["正社員"]["B2"].value
        wb["正社員"]["C2"] = f"{token}さん担当 / {token}_E9"  # _E9 は対応表にないID
        wb.save(anon)
        restored = self.root / "r2.xlsx"
        result = restore_workbook(anon, map_path, restored, "pw")
        self.assertEqual(load_workbook(restored)["正社員"]["C2"].value, f"山田 太郎さん担当 / {token}_E9")
        self.assertEqual(result.unknown_tokens, [f"{token}_E9"])

    # ---- 匿名化済みExcelへのシート追加 ----------------------------------------------

    def _add_new_sheet(self, anon, rows, name="新規"):
        wb = load_workbook(anon)
        sheet = wb.create_sheet(name)
        sheet.append(["氏名", "社員番号"])
        for row in rows:
            sheet.append(list(row))
        plus = self.root / "plus.xlsx"
        wb.save(plus)
        return plus

    def _append(self, plus, base_map, allow=False, password="pw", **kwargs):
        out, map_path = self.root / "anon2.xlsx", self.root / "map2.epbmap"
        result = anonymize_workbook_multi(
            plus, out, map_path, password=password,
            person_columns=[PersonColumns("新規", name_col="A", id_col="B")],
            allow_conflicts=allow, base_map_path=base_map, base_map_password=password, **kwargs,
        )
        return result, out, map_path

    def test_append_sheet_shares_ids_with_existing(self):
        _, anon, map1 = self._anonymized_person_book()
        before = load_workbook(anon)
        base1 = before["正社員"]["B2"].value  # 山田
        plus = self._add_new_sheet(anon, [("山田　太郎", 1001), ("田中 次郎", 3001), (None, 1002)])
        result, out, map2 = self._append(plus, map1)

        wb = load_workbook(out)
        new, staff = wb["新規"], wb["正社員"]
        prefix = base1.rsplit("_", 1)[0]
        self.assertEqual(new["A2"].value, base1)  # 既存の人は同じID
        self.assertEqual(new["B2"].value, base1 + "_E1")
        self.assertEqual(new["A3"].value, f"{prefix}_000003")  # 新しい人は続き番号（既存は山田・佐藤の2人）
        self.assertEqual(new["B3"].value, f"{prefix}_000003_E1")
        self.assertEqual(new["B4"].value, before["正社員"]["A3"].value)  # 社員番号だけでも既存の人に結びつく
        self.assertEqual(new["A1"].value, "氏名")
        for row in range(1, 4):  # すでに匿名化したシートは1セルも変わらない
            for col in range(1, 4):
                self.assertEqual(staff.cell(row, col).value, before["正社員"].cell(row, col).value)
        self.assertEqual((result.new_persons, result.reused_ids, result.person_count), (1, 3, 3))
        self.assertEqual(result.replaced_by_sheet, {"新規": 5})

        # 更新後の対応表だけで、全シートが元に戻る。古い対応表は変わらない
        restored = self.root / "all_restored.xlsx"
        restore_workbook(out, map2, restored, "pw")
        rwb = load_workbook(restored)
        self.assertEqual(rwb["新規"]["A3"].value, "田中 次郎")
        self.assertEqual(rwb["新規"]["B3"].value, 3001)
        self.assertEqual(rwb["正社員"]["B2"].value, "山田 太郎")
        old = load_mapping(map1, "pw")["values"]
        new = load_mapping(map2, "pw")
        self.assertTrue(set(old) < set(new["values"]))
        self.assertEqual(len(new["updates"]), 1)
        restore_workbook(anon, map1, self.root / "old_restored.xlsx", "pw")

    def test_append_skips_already_anonymized_cells(self):
        _, anon, map1 = self._anonymized_person_book()
        before = load_workbook(anon)["正社員"]["B2"].value
        plus = self._add_new_sheet(anon, [("田中 次郎", 3001)])
        out, map2 = self.root / "s.xlsx", self.root / "s.epbmap"
        with self.assertRaises(PrivacyToolError):  # 匿名化済みシートだけを指定しても何も起きない
            anonymize_workbook_multi(
                plus, out, map2, [("正社員", ["A2:B3"])], "pw", base_map_path=map1, base_map_password="pw"
            )
        result = anonymize_workbook_multi(
            plus, out, map2, [("正社員", ["A2:B3"])], "pw",
            person_columns=[PersonColumns("新規", name_col="A", id_col="B")],
            base_map_path=map1, base_map_password="pw",
        )
        self.assertEqual(result.skipped_token_cells, 4)
        self.assertEqual(load_workbook(out)["正社員"]["B2"].value, before)

    def test_append_new_number_for_existing_person_is_alerted(self):
        _, anon, map1 = self._anonymized_person_book()
        base1 = load_workbook(anon)["正社員"]["B2"].value
        plus = self._add_new_sheet(anon, [("山田 太郎", 2001)])  # 登用で社員番号が変わった
        with self.assertRaises(PersonConflictError) as ctx:
            self._append(plus, map1)
        self.assertFalse((self.root / "anon2.xlsx").exists())
        conflict = ctx.exception.conflicts[0]
        self.assertEqual(conflict.kind, "multi_id")
        self.assertIn("2001", conflict.values)
        _, out, map2 = self._append(plus, map1, allow=True)
        self.assertEqual(load_workbook(out)["新規"]["B2"].value, base1 + "_E2")  # 1001がE1なので続きのE2
        self.assertEqual(load_mapping(map2, "pw")["values"][base1 + "_E2"], "2001")

    def test_append_linking_two_existing_people_is_alerted_not_merged(self):
        _, anon, map1 = self._anonymized_person_book()
        wb = load_workbook(anon)
        yamada, sato_no = wb["正社員"]["B2"].value, wb["正社員"]["A3"].value
        plus = self._add_new_sheet(anon, [("山田 太郎", 1002)])  # 山田さんと佐藤さんの番号が同じ行に
        with self.assertRaises(PersonConflictError) as ctx:
            self._append(plus, map1)
        self.assertIn("merge_existing", [c.kind for c in ctx.exception.conflicts])
        _, out, _ = self._append(plus, map1, allow=True)
        new = load_workbook(out)["新規"]
        self.assertEqual((new["A2"].value, new["B2"].value), (yamada, sato_no))  # 既存のIDはそのまま

    def test_append_with_legacy_map_without_roles_and_kinds(self):
        source = self._person_source()
        anon, map1 = self.root / "l_anon.xlsx", self.root / "l.epbmap"
        anonymize_workbook_multi(
            source, anon, map1, password=None, person_columns=self._person_specs()
        )
        wrapper = json.loads(map1.read_text(encoding="utf-8"))
        for key in ("token_roles", "kinds"):  # 0.3.0 までの対応表には無い項目
            wrapper["payload"].pop(key)
        map1.write_text(json.dumps(wrapper, ensure_ascii=False), encoding="utf-8")
        base1 = load_workbook(anon)["正社員"]["B2"].value
        plus = self._add_new_sheet(anon, [("山田 太郎", 1001), ("鈴木 三郎", 4001)])
        out, map2 = self.root / "l_out.xlsx", self.root / "l_out.epbmap"
        result = anonymize_workbook_multi(
            plus, out, map2, person_columns=[PersonColumns("新規", name_col="A", id_col="B")],
            base_map_path=map1,
        )
        self.assertEqual(load_workbook(out)["新規"]["A2"].value, base1)
        self.assertEqual(result.new_persons, 1)

    def test_append_rejects_same_map_path(self):
        _, anon, map1 = self._anonymized_person_book()
        plus = self._add_new_sheet(anon, [("田中 次郎", 3001)])
        with self.assertRaises(PrivacyToolError):
            anonymize_workbook_multi(
                plus, self.root / "z.xlsx", map1, password="pw",
                person_columns=[PersonColumns("新規", name_col="A", id_col="B")],
                base_map_path=map1, base_map_password="pw",
            )


if __name__ == "__main__":
    unittest.main()
