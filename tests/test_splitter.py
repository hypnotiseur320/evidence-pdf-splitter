#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""evidence_pdf_splitter 自测用例。

运行：python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import csv
import hashlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import evidence_pdf_splitter as eps  # noqa: E402
from pypdf import PdfReader, PdfWriter  # noqa: E402

HEADERS = ["证据编号", "证据名称", "起始页", "结束页", "证据来源", "证明目的", "原件/复印件", "备注"]


def make_blank_pdf(path: Path, pages: int = 20) -> None:
    writer = PdfWriter()
    for _ in range(pages):
        writer.add_blank_page(width=595, height=842)
    with path.open("wb") as handle:
        writer.write(handle)


def write_csv(path: Path, rows: list[list[object]], encoding: str = "utf-8-sig") -> None:
    with path.open("w", encoding=encoding, newline="", errors="replace") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADERS)
        for row in rows:
            writer.writerow(["" if value is None else value for value in row])


def run_cli(argv: list[str]) -> tuple[int, str]:
    buffer = io.StringIO()
    with contextlib.redirect_stdout(buffer), contextlib.redirect_stderr(buffer):
        code = eps.main(argv)
    return code, buffer.getvalue()


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


class SplitterTestBase(unittest.TestCase):
    def setUp(self) -> None:
        self._temp = tempfile.TemporaryDirectory()
        self.root = Path(self._temp.name)
        self.pdf = self.root / "卷宗.pdf"
        make_blank_pdf(self.pdf, 20)

    def tearDown(self) -> None:
        self._temp.cleanup()

    def base_argv(self, list_path: Path, out: Path) -> list[str]:
        return ["-p", str(self.pdf), "-l", str(list_path), "-o", str(out)]


class TestNormalFlow(SplitterTestBase):
    def test_split_pages_and_offset(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(
            listing,
            [
                [1, "借条", 1, 1, "原告持有", "证明借贷合意", "原件", ""],
                [2, "施工合同", 2, 5, "双方签订", "证明合同关系", "复印件", ""],
                [3, "验收报告", 6, 6, "建设单位", "证明验收合格", "复印件", ""],
            ],
        )
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out) + ["--offset", "2"])
        self.assertEqual(code, 0, output)

        files = sorted((out / "证据文件").glob("*.pdf"))
        self.assertEqual(len(files), 3)
        self.assertEqual([len(PdfReader(str(f)).pages) for f in files], [1, 4, 1])
        self.assertIn("卷宗p3", files[0].name)

        total = sum(len(PdfReader(str(f)).pages) for f in files)
        self.assertEqual(total, 6)

        self.assertTrue((out / "证据目录.xlsx").exists())
        self.assertTrue((out / "证据目录.md").exists())
        self.assertTrue((out / "校验报告.txt").exists())

    def test_physical_mode_without_offset(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "证据一", 1, 2, "", "", "", ""]])
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out))
        self.assertEqual(code, 0, output)
        files = sorted((out / "证据文件").glob("*.pdf"))
        self.assertIn("PDFp1-2", files[0].name)

    def test_group_parent_not_split(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(
            listing,
            [
                [1, "函件组", None, None, "", "组证据", "", ""],
                ["1-1", "催告函", 1, 2, "", "", "", ""],
                ["1-2", "回函", 3, 3, "", "", "", ""],
            ],
        )
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out))
        self.assertEqual(code, 0, output)
        files = sorted((out / "证据文件").glob("*.pdf"))
        self.assertEqual(len(files), 2)
        self.assertTrue(files[0].name.startswith("01-1"))
        self.assertTrue(files[1].name.startswith("01-2"))

    def test_merged_bookmarks_follow_page_order(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(
            listing,
            [
                [1, "甲", 1, 1, "", "", "", ""],
                [2, "组", None, None, "", "", "", ""],
                ["2-1", "乙", 2, 2, "", "", "", ""],
                ["2-2", "丙", 3, 3, "", "", "", ""],
                [3, "丁", 4, 4, "", "", "", ""],
            ],
        )
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out) + ["--mode", "merged"])
        self.assertEqual(code, 0, output)
        reader = PdfReader(str(out / "合并卷宗_带书签.pdf"))
        self.assertEqual(len(reader.pages), 4)
        titles = [item.title for item in reader.outline if not isinstance(item, list)]
        self.assertEqual(titles, ["甲", "组", "丁"])

    def test_dry_run_writes_nothing(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "证据一", 1, 1, "", "", "", ""]])
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out) + ["--dry-run"])
        self.assertEqual(code, 0, output)
        self.assertFalse(out.exists())


class TestValidation(SplitterTestBase):
    def _expect_failure(self, rows: list[list[object]], code_name: str, extra: list[str] | None = None) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, rows)
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out) + (extra or []))
        self.assertEqual(code, 1, output)
        self.assertIn(code_name, output)
        self.assertFalse((out / "证据文件").exists(), "校验失败时不应写出证据文件")

    def test_overlap(self) -> None:
        self._expect_failure(
            [[1, "甲", 1, 3, "", "", "", ""], [2, "乙", 3, 5, "", "", "", ""]], "OVERLAP"
        )

    def test_overlap_allowed_by_flag(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "甲", 1, 3, "", "", "", ""], [2, "乙", 3, 5, "", "", "", ""]])
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out) + ["--allow-overlap"])
        self.assertEqual(code, 0, output)

    def test_out_of_range(self) -> None:
        self._expect_failure([[1, "甲", 1, 30, "", "", "", ""]], "OUT_OF_RANGE")

    def test_reversed(self) -> None:
        self._expect_failure([[1, "甲", 5, 3, "", "", "", ""]], "REVERSED")

    def test_duplicate_eid(self) -> None:
        self._expect_failure(
            [[1, "甲", 1, 1, "", "", "", ""], [1, "乙", 2, 2, "", "", "", ""]], "DUP_EID"
        )

    def test_missing_page_for_plain_evidence(self) -> None:
        self._expect_failure([[1, "甲", None, None, "", "", "", ""]], "NO_PAGE")

    def test_uncovered_pages_is_warning(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "甲", 1, 1, "", "", "", ""]])
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out))
        self.assertEqual(code, 0, output)
        self.assertIn("UNCOVERED", output)

    def test_strict_turns_warning_into_failure(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "甲", 1, 1, "", "", "", ""]])
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out) + ["--strict"])
        self.assertEqual(code, 1, output)


class TestEncodingAndPaths(SplitterTestBase):
    def test_gbk_csv(self) -> None:
        listing = self.root / "清单_gbk.csv"
        write_csv(listing, [[1, "借款合同", 1, 2, "原告", "证明借贷", "原件", "备注"]], encoding="gbk")
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out))
        self.assertEqual(code, 0, output)
        files = list((out / "证据文件").glob("*.pdf"))
        self.assertEqual(len(files), 1)
        self.assertIn("借款合同", files[0].name)

    def test_chinese_path_with_special_chars(self) -> None:
        target_dir = self.root / "案件材料 ♥ 测试"
        target_dir.mkdir()
        pdf = target_dir / "卷宗 副本.pdf"
        make_blank_pdf(pdf, 10)
        listing = target_dir / "清单.csv"
        write_csv(listing, [[1, "证据一", 1, 3, "", "", "", ""]])
        out = target_dir / "输出"
        code, output = run_cli(["-p", str(pdf), "-l", str(listing), "-o", str(out)])
        self.assertEqual(code, 0, output)
        self.assertEqual(len(list((out / "证据文件").glob("*.pdf"))), 1)

    def test_illegal_filename_chars_sanitized(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, '合同/协议:"原件"', 1, 1, "", "", "", ""]])
        out = self.root / "out"
        code, output = run_cli(self.base_argv(listing, out))
        self.assertEqual(code, 0, output)
        files = list((out / "证据文件").glob("*.pdf"))
        self.assertEqual(len(files), 1)
        self.assertNotIn("/", files[0].name)
        self.assertNotIn('"', files[0].name)


class TestIntegrity(SplitterTestBase):
    def test_source_file_unchanged(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "甲", 1, 2, "", "", "", ""]])
        before = sha256(self.pdf)
        code, _ = run_cli(self.base_argv(listing, self.root / "out"))
        self.assertEqual(code, 0)
        self.assertEqual(sha256(self.pdf), before)

    def test_idempotent(self) -> None:
        listing = self.root / "清单.csv"
        write_csv(listing, [[1, "甲", 1, 2, "", "", "", ""], [2, "乙", 3, 4, "", "", "", ""]])
        out = self.root / "out"
        run_cli(self.base_argv(listing, out))
        first = {p.name: len(PdfReader(str(p)).pages) for p in (out / "证据文件").glob("*.pdf")}
        run_cli(self.base_argv(listing, out))
        second = {p.name: len(PdfReader(str(p)).pages) for p in (out / "证据文件").glob("*.pdf")}
        self.assertEqual(first, second)
        self.assertEqual(len(second), 2)


class TestPageMapper(unittest.TestCase):
    def test_fixed_offset(self) -> None:
        mapper = eps.PageMapper(offset=2)
        self.assertEqual(mapper.to_physical(1), 3)
        self.assertEqual(mapper.to_physical(10), 12)

    def test_anchor_segments(self) -> None:
        # 卷宗第 10 页之前插入了 2 页未编号页，偏移由 2 变为 4
        mapper = eps.PageMapper(anchors=[(1, 3), (10, 14)])
        self.assertEqual(mapper.to_physical(1), 3)
        self.assertEqual(mapper.to_physical(9), 11)
        self.assertEqual(mapper.to_physical(10), 14)

    def test_overlapping_anchors_rejected(self) -> None:
        with self.assertRaises(eps.SplitterError):
            eps.PageMapper(anchors=[(1, 3), (10, 11)])

    def test_anchor_earlier_than_first(self) -> None:
        mapper = eps.PageMapper(anchors=[(5, 7)])
        with self.assertRaises(eps.SplitterError):
            mapper.to_physical(3)


class TestParsing(unittest.TestCase):
    def test_page_number_forms(self) -> None:
        self.assertEqual(eps.parse_page_number("12"), 12)
        self.assertEqual(eps.parse_page_number("第12页"), 12)
        self.assertEqual(eps.parse_page_number(12.0), 12)
        self.assertIsNone(eps.parse_page_number(""))
        self.assertIsNone(eps.parse_page_number("无"))

    def test_header_alias(self) -> None:
        rows = [["编号", "名称", "起页", "止页"], [1, "甲", 1, 2]]
        header_index, mapping = eps.locate_header(rows)
        self.assertEqual(header_index, 0)
        self.assertEqual(mapping["eid"], 0)
        self.assertEqual(mapping["start"], 2)

    def test_missing_required_header(self) -> None:
        rows = [["案号", "当事人"], ["(2026)浙0206民初1号", "张某"]]
        with self.assertRaises(eps.SplitterError):
            eps.locate_header(rows)

    def test_filename_sanitize(self) -> None:
        self.assertEqual(eps.sanitize_filename('合 同/协议:*?"<>|'), "合 同_协议")
        self.assertEqual(eps.sanitize_filename("  . "), "未命名")

    def test_prefix_format(self) -> None:
        self.assertEqual(eps.format_prefix("3", 3), "03")
        self.assertEqual(eps.format_prefix("5-1", 6), "05-1")
        self.assertEqual(eps.format_prefix("12", 1), "12")


if __name__ == "__main__":
    unittest.main(verbosity=2)
