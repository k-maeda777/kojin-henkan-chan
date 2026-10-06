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


if __name__ == "__main__":
    unittest.main()
