#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""evidence_pdf_splitter - 案卷 PDF 结构化归档工具

按人工给定的证据清单，把一份完整案卷 PDF 切分为独立证据文件，
并生成可提交法院的证据目录表（xlsx + md）。

设计前提：脚本不判断证据边界。证据的起止范围由人工阅卷后以清单形式提供，
脚本只承担切分、命名、页码换算、目录生成与完整性校验。

依赖：pypdf、openpyxl
许可：MIT
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import re
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable, Sequence

__version__ = "1.0.0"

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from pypdf import PdfReader, PdfWriter
except ImportError as exc:  # pragma: no cover
    sys.stderr.write(
        "缺少依赖：%s\n请先运行：pip install -r requirements.txt\n" % exc
    )
    raise SystemExit(2)


# --------------------------------------------------------------------------
# 常量
# --------------------------------------------------------------------------

# 清单列名别名表。左侧为内部字段，右侧为可识别的中文/英文列名。
COLUMN_ALIASES: dict[str, tuple[str, ...]] = {
    "eid": ("证据编号", "编号", "证据序号", "证据号", "eid", "no", "id"),
    "name": ("证据名称", "名称", "证据", "name", "title"),
    "start": ("起始页", "起页", "开始页", "首页", "起始页码", "start", "startpage"),
    "end": ("结束页", "止页", "终止页", "末页", "结束页码", "end", "endpage"),
    "source": ("证据来源", "来源", "source"),
    "purpose": ("证明目的", "证明对象", "证明内容", "purpose"),
    "form": ("原件/复印件", "原件或复印件", "形式", "form"),
    "remark": ("备注", "说明", "remark", "note"),
}

REQUIRED_FIELDS = ("eid", "name", "start")

# 目录输出列（顺序即表格列顺序）
OUTPUT_COLUMNS: tuple[tuple[str, str, int], ...] = (
    ("eid", "证据编号", 12),
    ("name", "证据名称", 34),
    ("vol_range", "卷宗页码", 12),
    ("pdf_range", "PDF页码", 12),
    ("pages", "页数", 6),
    ("source", "证据来源", 18),
    ("purpose", "证明目的", 38),
    ("form", "原件/复印件", 11),
    ("remark", "备注", 20),
)

# 目录中需要居中显示的列（按 OUTPUT_COLUMNS 的位置，1 基）
CENTER_COLUMNS = (1, 3, 4, 5, 8)

ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|\r\n\t]')
MAX_NAME_LEN = 80

EID_PATTERN = re.compile(r"^(\d+)(?:-(\d+))?$")
PAGE_NUMBER_PATTERN = re.compile(r"-?\d+")


class SplitterError(Exception):
    """可预期的业务错误，main 捕获后以退出码 2 结束。"""


# --------------------------------------------------------------------------
# 数据结构
# --------------------------------------------------------------------------


@dataclass
class EvidenceRow:
    """清单中的一行证据。"""

    seq: int
    eid: str
    name: str
    vol_start: int | None = None
    vol_end: int | None = None
    source: str = ""
    purpose: str = ""
    form: str = ""
    remark: str = ""
    extra: dict[str, Any] = field(default_factory=dict)

    # 计算字段
    is_group: bool = False          # 是否为组证据的父级（有子级则父级不切分）
    physical_start: int | None = None
    physical_end: int | None = None
    out_name: str = ""
    merged_index: int | None = None

    @property
    def pages(self) -> int:
        if self.vol_start is None or self.vol_end is None:
            return 0
        return self.vol_end - self.vol_start + 1

    @property
    def vol_range_label(self) -> str:
        if self.vol_start is None:
            return ""
        if self.vol_end is None or self.vol_end == self.vol_start:
            return f"第{self.vol_start}页"
        return f"第{self.vol_start}-{self.vol_end}页"

    @property
    def pdf_range_label(self) -> str:
        if self.physical_start is None:
            return ""
        if self.physical_end is None or self.physical_end == self.physical_start:
            return f"第{self.physical_start}页"
        return f"第{self.physical_start}-{self.physical_end}页"

    @property
    def pdf_range_for_filename(self) -> str:
        if self.physical_start is None:
            return ""
        if self.physical_end is None or self.physical_end == self.physical_start:
            return f"p{self.physical_start}"
        return f"p{self.physical_start}-{self.physical_end}"


@dataclass
class Issue:
    """一条校验结果。level 取 error / warning。"""

    level: str
    code: str
    message: str

    def render(self) -> str:
        tag = "错误" if self.level == "error" else "警告"
        return f"[{tag}][{self.code}] {self.message}"


# --------------------------------------------------------------------------
# 基础工具
# --------------------------------------------------------------------------


def relax_stdout_encoding() -> None:
    """避免 Windows 控制台因个别字符无法编码而中断输出。"""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")  # type: ignore[attr-defined]
        except Exception:
            pass


def normalize_header(text: Any) -> str:
    """列名归一化：去空白、去全角空格、转小写、统一斜杠。"""
    s = str(text or "")
    s = s.replace("　", "").replace(" ", "").replace("／", "/")
    s = s.replace("（", "(").replace("）", ")")
    return s.strip().lower()


def parse_page_number(value: Any) -> int | None:
    """从单元格中提取页码整数，支持 '12'、'12.0'、'第12页' 等写法。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, int):
        return value
    if isinstance(value, float):
        return int(value) if value == int(value) else None
    text = str(value).strip()
    if not text:
        return None
    match = PAGE_NUMBER_PATTERN.search(text)
    if not match:
        return None
    try:
        return int(match.group())
    except ValueError:
        return None


def sanitize_filename(name: str, max_len: int = MAX_NAME_LEN) -> str:
    """清洗文件名片段：替换非法字符、去除首尾点与空格、限长。"""
    cleaned = ILLEGAL_FILENAME_CHARS.sub("_", str(name))
    cleaned = re.sub(r"_{2,}", "_", cleaned).strip(" ._")
    if len(cleaned) > max_len:
        cleaned = cleaned[:max_len].rstrip(" ._")
    return cleaned or "未命名"


def format_prefix(eid: str, fallback_seq: int) -> str:
    """生成文件名前缀：'3' -> '03'，'5-1' -> '05-1'，其他保留清洗值。"""
    match = EID_PATTERN.match(eid.strip())
    if match:
        major = match.group(1).zfill(2)
        minor = match.group(2)
        return f"{major}-{minor}" if minor else major
    cleaned = sanitize_filename(eid, 20)
    return cleaned or str(fallback_seq).zfill(2)


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


# --------------------------------------------------------------------------
# 清单读取与解析
# --------------------------------------------------------------------------


def load_table_rows(path: Path, sheet: str | None = None) -> list[list[Any]]:
    """读取 xlsx 或 csv，返回二维表（含表头行）。"""
    if not path.exists():
        raise SplitterError(f"清单文件不存在：{path}")
    suffix = path.suffix.lower()
    if suffix in (".xlsx", ".xlsm"):
        return _load_rows_from_xlsx(path, sheet)
    if suffix in (".csv", ".txt"):
        return _load_rows_from_csv(path)
    raise SplitterError(f"不支持的清单格式：{suffix}（仅支持 .xlsx/.xlsm/.csv）")


def _load_rows_from_xlsx(path: Path, sheet: str | None) -> list[list[Any]]:
    try:
        workbook = load_workbook(str(path), data_only=True, read_only=True)
    except Exception as exc:
        raise SplitterError(f"无法打开 Excel 清单：{exc}") from exc
    try:
        if sheet:
            if sheet not in workbook.sheetnames:
                raise SplitterError(
                    f"工作表 '{sheet}' 不存在，可用工作表：{', '.join(workbook.sheetnames)}"
                )
            worksheet = workbook[sheet]
        else:
            worksheet = workbook.worksheets[0]
        rows: list[list[Any]] = []
        for raw in worksheet.iter_rows(values_only=True):
            rows.append(["" if cell is None else cell for cell in raw])
        return rows
    finally:
        workbook.close()


def _load_rows_from_csv(path: Path) -> list[list[Any]]:
    raw = path.read_bytes()
    text = None
    for encoding in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise SplitterError("无法识别清单文件编码（已尝试 UTF-8 与 GBK 系列）。")

    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
    except csv.Error:
        dialect = csv.excel
    return [list(row) for row in csv.reader(io.StringIO(text), dialect)]


def locate_header(rows: Sequence[Sequence[Any]]) -> tuple[int, dict[str, int]]:
    """定位表头行并返回 字段 -> 列索引 映射。"""
    for row_index, row in enumerate(rows[:20]):
        if not any(str(cell).strip() for cell in row):
            continue
        normalized = [normalize_header(cell) for cell in row]
        mapping: dict[str, int] = {}
        for field_name, aliases in COLUMN_ALIASES.items():
            alias_set = {normalize_header(a) for a in aliases}
            for column_index, header in enumerate(normalized):
                if header and header in alias_set:
                    mapping[field_name] = column_index
                    break
        if all(name in mapping for name in ("eid", "name", "start")):
            return row_index, mapping
    actual = ""
    if rows:
        actual = "、".join(str(cell) for cell in rows[0] if str(cell).strip())
    raise SplitterError(
        "未能在清单前 20 行中找到必要的表头列：证据编号、证据名称、起始页。\n"
        f"检测到的首行内容为：{actual or '(空)'}\n"
        "请核对列名，或参考 README 的字段说明修改清单。"
    )


def parse_evidence_rows(rows: Sequence[Sequence[Any]]) -> tuple[list[EvidenceRow], list[Issue]]:
    """把二维表解析为证据行列表，同时收集逐行的字段级问题。"""
    header_index, mapping = locate_header(rows)
    issues: list[Issue] = []
    parsed: list[EvidenceRow] = []
    seq = 0

    for row_number, row in enumerate(rows[header_index + 1 :], start=header_index + 2):
        def cell(field_name: str) -> str:
            index = mapping.get(field_name)
            if index is None or index >= len(row):
                return ""
            return str(row[index]).strip()

        eid = cell("eid")
        name = cell("name")
        if not eid and not name:
            continue  # 空行

        seq += 1
        start = parse_page_number(cell("start"))
        end = parse_page_number(cell("end"))

        if not eid:
            issues.append(Issue("error", "MISSING_EID", f"第 {row_number} 行缺少证据编号。"))
        if not name:
            issues.append(Issue("error", "MISSING_NAME", f"第 {row_number} 行（编号 {eid or '?'}）缺少证据名称。"))
        # 起始页缺失不在解析阶段判定：组证据父级允许不填页码，
        # 是否构成错误取决于分组识别结果，统一在 validate() 中处理。
        if end is None:
            end = start

        parsed.append(
            EvidenceRow(
                seq=seq,
                eid=eid,
                name=name,
                vol_start=start,
                vol_end=end,
                source=cell("source"),
                purpose=cell("purpose"),
                form=cell("form"),
                remark=cell("remark"),
            )
        )
    return parsed, issues


def mark_group_parents(rows: list[EvidenceRow]) -> None:
    """识别组证据父级：若存在 '5-1'，则 '5' 标记为父级、不参与切分。"""
    child_majors = set()
    for row in rows:
        match = re.match(r"^(\d+)-", row.eid.strip())
        if match:
            child_majors.add(match.group(1))
    for row in rows:
        key = row.eid.strip()
        if key in child_majors and re.fullmatch(r"\d+", key):
            row.is_group = True


# --------------------------------------------------------------------------
# 页码映射
# --------------------------------------------------------------------------


class PageMapper:
    """把清单中的卷宗页码换算为 PDF 物理页序（1 基）。

    优先级：锚点表 > 固定偏移 > 恒等映射。
    锚点表每行为 (卷宗起始页, 物理起始页)，可表达卷宗中插入未编号页造成的分段偏移。
    """

    def __init__(self, offset: int = 0, anchors: Sequence[tuple[int, int]] | None = None) -> None:
        self.anchors = sorted(anchors or [])
        self.offset = offset
        self._validate_anchors()

    def _validate_anchors(self) -> None:
        """允许各段偏移不同（卷宗中间插入未编号页即为典型场景），
        但要求卷宗页与物理页均严格递增，且物理页增量不小于卷宗页增量，
        否则会出现两页映射到同一物理页的情况。
        """
        previous = None
        for vol_start, phys_start in self.anchors:
            if vol_start < 1 or phys_start < 1:
                raise SplitterError("锚点表中的页码必须为正整数。")
            if previous is not None:
                prev_vol, prev_phys = previous
                if vol_start <= prev_vol:
                    raise SplitterError(f"锚点表未按卷宗页码升序排列：{prev_vol} -> {vol_start}")
                if phys_start <= prev_phys:
                    raise SplitterError(
                        f"锚点表的物理页码未递增：{prev_phys} -> {phys_start}，请核对。"
                    )
                if phys_start - prev_phys < vol_start - prev_vol:
                    raise SplitterError(
                        "锚点表前后两段的物理页增量小于卷宗页增量，会导致页码重叠："
                        f"卷宗第 {prev_vol} 页 -> 物理第 {prev_phys} 页，"
                        f"卷宗第 {vol_start} 页 -> 物理第 {phys_start} 页。"
                    )
            previous = (vol_start, phys_start)

    @property
    def has_shift(self) -> bool:
        if self.anchors:
            return any(vol != phys for vol, phys in self.anchors)
        return self.offset != 0

    def to_physical(self, vol_page: int) -> int:
        if self.anchors:
            anchor = None
            for vol_start, phys_start in self.anchors:
                if vol_page >= vol_start:
                    anchor = (vol_start, phys_start)
                else:
                    break
            if anchor is None:
                raise SplitterError(
                    f"卷宗第 {vol_page} 页早于锚点表首个锚点（卷宗第 {self.anchors[0][0]} 页），无法换算。"
                )
            vol_start, phys_start = anchor
            return phys_start + (vol_page - vol_start)
        return vol_page + self.offset


def load_anchor_file(path: Path) -> list[tuple[int, int]]:
    """读取锚点表：每行 '卷宗页码,物理页码'，允许 # 注释。"""
    if not path.exists():
        raise SplitterError(f"锚点表文件不存在：{path}")
    anchors: list[tuple[int, int]] = []
    for line_number, line in enumerate(
        path.read_text(encoding="utf-8-sig", errors="replace").splitlines(), start=1
    ):
        text = line.split("#", 1)[0].strip()
        if not text:
            continue
        parts = re.split(r"[,;\t ]+", text)
        if len(parts) < 2:
            raise SplitterError(f"锚点表第 {line_number} 行格式错误：{line.strip()}")
        try:
            anchors.append((int(parts[0]), int(parts[1])))
        except ValueError as exc:
            raise SplitterError(f"锚点表第 {line_number} 行不是整数：{line.strip()}") from exc
    if not anchors:
        raise SplitterError("锚点表为空。")
    return anchors


# --------------------------------------------------------------------------
# 校验
# --------------------------------------------------------------------------


def validate(
    rows: Sequence[EvidenceRow],
    mapper: PageMapper,
    total_pages: int,
    allow_overlap: bool = False,
) -> tuple[list[Issue], list[Issue]]:
    """执行全部校验，返回 (错误, 警告)。"""
    errors: list[Issue] = []
    warnings: list[Issue] = []

    # 1. 证据编号重复
    seen: dict[str, int] = {}
    for row in rows:
        key = row.eid.strip()
        if key in seen:
            errors.append(Issue("error", "DUP_EID", f"证据编号重复：'{key}'（第 {seen[key]} 条与第 {row.seq} 条）。"))
        else:
            seen[key] = row.seq

    # 2. 页码倒序 / 缺失，并换算物理页
    splittable: list[EvidenceRow] = []
    for row in rows:
        if row.vol_start is None or row.vol_end is None:
            if not row.is_group:
                errors.append(Issue("error", "NO_PAGE", f"证据 {row.eid} '{row.name}' 缺少页码，无法切分。"))
            continue
        if row.vol_end < row.vol_start:
            errors.append(
                Issue("error", "REVERSED", f"证据 {row.eid} 页码倒序：起始 {row.vol_start} > 结束 {row.vol_end}。")
            )
            continue
        try:
            row.physical_start = mapper.to_physical(row.vol_start)
            row.physical_end = mapper.to_physical(row.vol_end)
        except SplitterError as exc:
            errors.append(Issue("error", "MAP_FAILED", f"证据 {row.eid} 页码换算失败：{exc}"))
            continue
        if row.physical_start < 1 or row.physical_end > total_pages:
            errors.append(
                Issue(
                    "error",
                    "OUT_OF_RANGE",
                    f"证据 {row.eid} 换算后为 PDF 第 {row.physical_start}-{row.physical_end} 页，"
                    f"超出源文件总页数 {total_pages}。",
                )
            )
            continue
        if row.is_group:
            warnings.append(
                Issue(
                    "warning",
                    "GROUP_PAGE_IGNORED",
                    f"证据 {row.eid} 为组证据父级，已按子证据切分，父级页码 {row.vol_range_label} 不单独出文件。",
                )
            )
            continue
        splittable.append(row)

    # 3. 页码区间重叠（按物理页排序后检测，与清单顺序无关）
    ordered = sorted(splittable, key=lambda r: (r.physical_start or 0, r.physical_end or 0))
    for previous, current in zip(ordered, ordered[1:]):
        if previous.physical_end is None or current.physical_start is None:
            continue
        if current.physical_start <= previous.physical_end:
            level = "warning" if allow_overlap else "error"
            message = (
                f"证据 {previous.eid}（{previous.pdf_range_label}）与证据 {current.eid}"
                f"（{current.pdf_range_label}）页码区间重叠。"
            )
            target = warnings if level == "warning" else errors
            target.append(Issue(level, "OVERLAP", message))

    # 4. 未覆盖页
    if splittable:
        covered: set[int] = set()
        for row in splittable:
            if row.physical_start and row.physical_end:
                covered.update(range(row.physical_start, row.physical_end + 1))
        first_physical = min(r.physical_start or 1 for r in splittable)
        uncovered = [p for p in range(first_physical, total_pages + 1) if p not in covered]
        if uncovered:
            preview = "、".join(str(p) for p in uncovered[:12])
            more = f" 等共 {len(uncovered)} 页" if len(uncovered) > 12 else ""
            warnings.append(
                Issue("warning", "UNCOVERED", f"以下 PDF 页未被任何证据覆盖：第 {preview} 页{more}，请确认是否为空白页或漏列。")
            )

    # 5. 子级缺少父级
    major_numbers = {r.eid.strip() for r in rows if re.fullmatch(r"\d+", r.eid.strip())}
    for row in rows:
        match = re.match(r"^(\d+)-", row.eid.strip())
        if match and match.group(1) not in major_numbers:
            warnings.append(
                Issue("warning", "ORPHAN_CHILD", f"证据 {row.eid} 为子证据，但清单中没有父级证据 {match.group(1)}。")
            )

    return errors, warnings


# --------------------------------------------------------------------------
# 输出
# --------------------------------------------------------------------------


def assign_output_names(rows: Iterable[EvidenceRow], page_label: str) -> None:
    for row in rows:
        prefix = format_prefix(row.eid, row.seq)
        name = sanitize_filename(row.name)
        suffix = f"_{page_label}{row.pdf_range_for_filename}" if row.pdf_range_for_filename else ""
        row.out_name = f"{prefix}_{name}{suffix}.pdf"


def split_pdfs(reader: PdfReader, rows: Sequence[EvidenceRow], out_dir: Path) -> list[Path]:
    written: list[Path] = []
    for row in rows:
        if row.is_group or row.physical_start is None or row.physical_end is None:
            continue
        writer = PdfWriter()
        for page_index in range(row.physical_start - 1, row.physical_end):
            writer.add_page(reader.pages[page_index])
        writer.add_metadata(
            {
                "/Producer": f"evidence_pdf_splitter {__version__}",
                "/Creator": f"evidence_pdf_splitter {__version__}",
                "/Title": row.name,
            }
        )
        target = out_dir / row.out_name
        with target.open("wb") as handle:
            writer.write(handle)
        written.append(target)
    return written


def build_merged_pdf(
    reader: PdfReader, rows: Sequence[EvidenceRow], target: Path
) -> list[EvidenceRow]:
    """按物理页序合并，并生成二级书签。返回参与合并的行（已填 merged_index）。"""
    splittable = [r for r in rows if not r.is_group and r.physical_start and r.physical_end]
    ordered = sorted(splittable, key=lambda r: (r.physical_start or 0, r.physical_end or 0))

    writer = PdfWriter()
    for row in ordered:
        row.merged_index = len(writer.pages)
        for page_index in range(row.physical_start - 1, row.physical_end):
            writer.add_page(reader.pages[page_index])

    # 书签：父级标题指向其首个（页序最前）子级
    groups = [r for r in rows if r.is_group]
    group_targets: dict[str, int] = {}
    for group in groups:
        key = group.eid.strip()
        children = [
            r for r in ordered if re.match(rf"^{re.escape(key)}-\d+$", r.eid.strip())
        ]
        if children:
            group_targets[key] = min(r.merged_index or 0 for r in children)

    # 书签分两阶段：先按页序创建全部顶层项（普通证据与组证据混排），
    # 再把子证据挂到对应组下面。pypdf 的书签按创建顺序追加，
    # 因此必须先排序再创建，否则父级会跑到最前面。
    group_keys = {g.eid.strip() for g in groups}
    top_items: list[tuple[int, str, Any]] = []
    for group in groups:
        key = group.eid.strip()
        if key in group_targets:
            top_items.append((group_targets[key], "group", group))
    for row in ordered:
        match = re.match(r"^(\d+)-", row.eid.strip())
        if match and match.group(1) in group_keys:
            continue  # 子证据不占顶层
        top_items.append((row.merged_index or 0, "row", row))
    top_items.sort(key=lambda item: item[0])

    parent_refs: dict[str, Any] = {}
    for target_index, kind, payload in top_items:
        if kind == "group":
            key = payload.eid.strip()
            parent_refs[key] = writer.add_outline_item(
                title=payload.name or f"证据{key}", page_number=target_index
            )
        else:
            writer.add_outline_item(
                title=payload.name or f"证据{payload.eid}", page_number=target_index
            )

    for row in ordered:
        match = re.match(r"^(\d+)-", row.eid.strip())
        if not match or match.group(1) not in parent_refs:
            continue
        writer.add_outline_item(
            title=row.name or f"证据{row.eid}",
            page_number=row.merged_index or 0,
            parent=parent_refs[match.group(1)],
        )

    writer.add_metadata(
        {
            "/Producer": f"evidence_pdf_splitter {__version__}",
            "/Creator": f"evidence_pdf_splitter {__version__}",
            "/Title": "合并卷宗（带书签）",
        }
    )
    with target.open("wb") as handle:
        writer.write(handle)
    return ordered


def _row_values(row: EvidenceRow) -> list[Any]:
    return [
        row.eid,
        row.name,
        row.vol_range_label,
        row.pdf_range_label,
        "—" if row.is_group else (row.pages if row.vol_start is not None else ""),
        row.source,
        row.purpose,
        row.form,
        row.remark,
    ]


def write_xlsx(rows: Sequence[EvidenceRow], target: Path, title: str | None = None) -> None:
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "证据目录"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(name="宋体", size=11, bold=True, color="FFFFFF")
    body_font = Font(name="宋体", size=10.5)

    start_row = 1
    if title:
        worksheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(OUTPUT_COLUMNS))
        cell = worksheet.cell(row=1, column=1, value=title)
        cell.font = Font(name="宋体", size=14, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        worksheet.row_dimensions[1].height = 32
        start_row = 2

    for column_index, (_, header, width) in enumerate(OUTPUT_COLUMNS, start=1):
        cell = worksheet.cell(row=start_row, column=column_index, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = border
        worksheet.column_dimensions[get_column_letter(column_index)].width = width
    worksheet.row_dimensions[start_row].height = 30

    for offset, row in enumerate(rows, start=1):
        for column_index, value in enumerate(_row_values(row), start=1):
            cell = worksheet.cell(row=start_row + offset, column=column_index, value=value)
            cell.font = body_font
            cell.border = border
            align = "center" if column_index in CENTER_COLUMNS else "left"
            cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        worksheet.row_dimensions[start_row + offset].height = 34

    worksheet.freeze_panes = worksheet.cell(row=start_row + 1, column=1)
    worksheet.page_setup.orientation = "landscape"
    worksheet.page_setup.paperSize = 9  # A4
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.print_title_rows = f"{start_row}:{start_row}"

    workbook.save(str(target))


def _md_escape(text: Any) -> str:
    return str(text or "").replace("|", "\\|").replace("\n", "<br>")


def write_markdown(
    rows: Sequence[EvidenceRow], target: Path, title: str | None, issues: Sequence[Issue]
) -> None:
    lines: list[str] = []
    lines.append(f"# {title or '证据目录'}")
    lines.append("")
    lines.append(f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append("")
    headers = [header for _, header, _ in OUTPUT_COLUMNS]
    lines.append("| " + " | ".join(headers) + " |")
    lines.append("|" + "|".join(["---"] * len(headers)) + "|")
    for row in rows:
        values = [_md_escape(value) for value in _row_values(row)]
        lines.append("| " + " | ".join(values) + " |")
    lines.append("")
    if issues:
        lines.append("## 校验提示")
        lines.append("")
        for issue in issues:
            lines.append(f"- {issue.render()}")
        lines.append("")
    lines.append("> 本目录由 evidence_pdf_splitter 依据人工提供的证据清单自动生成，")
    lines.append("> 页码与页数系机械换算结果，提交前请与卷宗原件逐条核对。")
    lines.append("")
    target.write_text("\n".join(lines), encoding="utf-8")


def write_report(
    issues: Sequence[Issue],
    target: Path,
    source_pdf: Path,
    source_hash: str,
    total_pages: int,
    split_count: int,
) -> None:
    lines = [
        "案卷归集校验报告",
        "=" * 32,
        f"生成时间：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"源文件：{source_pdf}",
        f"源文件 SHA256：{source_hash}",
        f"源文件总页数：{total_pages}",
        f"切分证据份数：{split_count}",
        "",
    ]
    errors = [i for i in issues if i.level == "error"]
    warnings = [i for i in issues if i.level == "warning"]
    lines.append(f"错误 {len(errors)} 项，警告 {len(warnings)} 项。")
    lines.append("")
    if issues:
        for issue in issues:
            lines.append(issue.render())
    else:
        lines.append("未发现校验问题。")
    lines.append("")
    target.write_text("\n".join(lines), encoding="utf-8")


# --------------------------------------------------------------------------
# 主流程
# --------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="evidence_pdf_splitter",
        description="按证据清单切分案卷 PDF，并生成证据目录（xlsx + md）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python evidence_pdf_splitter.py -p 卷宗.pdf -l 证据清单.xlsx -o 输出 --offset 2\n"
            "  python evidence_pdf_splitter.py -p 卷宗.pdf -l 清单.csv --map-file 锚点.csv --mode both\n"
        ),
    )
    parser.add_argument("-p", "--pdf", required=True, type=Path, help="源案卷 PDF 路径（只读）")
    parser.add_argument("-l", "--evidence-list", required=True, type=Path, help="证据清单（xlsx 或 csv）")
    parser.add_argument("-o", "--out", type=Path, default=Path("output"), help="输出目录，默认 ./output")
    parser.add_argument(
        "--mode",
        choices=("split", "merged", "both"),
        default="both",
        help="输出形态：split 仅切分，merged 仅合并加书签，both 两者都出（默认）",
    )
    parser.add_argument("--offset", type=int, default=0, help="固定页码偏移：物理页 = 卷宗页 + offset")
    parser.add_argument("--map-file", type=Path, help="锚点表 CSV（卷宗页码,物理页码），优先于 --offset")
    parser.add_argument("--sheet", help="Excel 工作表名，默认取第一个")
    parser.add_argument("--title", help="目录标题，如案件名称")
    parser.add_argument("--page-label", choices=("auto", "vol", "pdf"), default="auto", help="文件名页码前缀")
    parser.add_argument("--allow-overlap", action="store_true", help="页码重叠降级为警告")
    parser.add_argument("--strict", action="store_true", help="警告也视为失败")
    parser.add_argument("--dry-run", action="store_true", help="仅校验并打印清单，不写文件")
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    return parser


def run(args: argparse.Namespace) -> int:
    source_pdf: Path = args.pdf
    list_path: Path = args.evidence_list

    if not source_pdf.exists():
        raise SplitterError(f"源 PDF 不存在：{source_pdf}")
    if source_pdf.suffix.lower() != ".pdf":
        raise SplitterError(f"源文件不是 PDF：{source_pdf}")

    source_hash = file_sha256(source_pdf)

    try:
        reader = PdfReader(str(source_pdf))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception as exc:
                raise SplitterError("源 PDF 已加密，且无法用空密码打开，请先解密后再运行。") from exc
        total_pages = len(reader.pages)
    except SplitterError:
        raise
    except Exception as exc:
        raise SplitterError(f"无法读取源 PDF：{exc}") from exc

    if total_pages == 0:
        raise SplitterError("源 PDF 不含任何页面。")

    table = load_table_rows(list_path, args.sheet)
    rows, parse_issues = parse_evidence_rows(table)
    if not rows:
        raise SplitterError("清单中未解析到任何有效证据行。")

    mark_group_parents(rows)

    anchors = load_anchor_file(args.map_file) if args.map_file else None
    mapper = PageMapper(offset=args.offset, anchors=anchors)

    validate_errors, validate_warnings = validate(rows, mapper, total_pages, args.allow_overlap)
    all_issues = list(parse_issues) + validate_errors + validate_warnings

    # 页码前缀：有偏移时标"卷宗p"，否则标"PDFp"
    if args.page_label == "vol":
        page_label = "卷宗"
    elif args.page_label == "pdf":
        page_label = "PDF"
    else:
        page_label = "卷宗" if mapper.has_shift else "PDF"
    assign_output_names(rows, page_label)

    splittable = [r for r in rows if not r.is_group and r.physical_start and r.physical_end]

    print(f"源文件：{source_pdf}")
    print(f"总页数：{total_pages}    清单条数：{len(rows)}    待切分：{len(splittable)}")
    print(f"页码模式：{'锚点表' if anchors else ('固定偏移 ' + str(args.offset) if args.offset else '物理页序')}")
    print("")

    if all_issues:
        print("校验结果：")
        for issue in all_issues:
            print("  " + issue.render())
        print("")
    else:
        print("校验结果：未发现问题。")
        print("")

    if validate_errors or parse_issues:
        print(f"存在 {len(validate_errors) + len(parse_issues)} 项错误，已终止，未写入任何文件。")
        return 1

    if args.dry_run:
        print("预演模式（--dry-run）：不写文件。将生成以下证据文件：")
        for row in splittable:
            print(f"  {row.out_name}    卷宗{row.vol_range_label} / PDF{row.pdf_range_label}")
        return 0

    out_dir: Path = args.out
    if not out_dir.is_absolute():
        out_dir = Path.cwd() / out_dir
    evidence_dir = out_dir / "证据文件"
    out_dir.mkdir(parents=True, exist_ok=True)

    written: list[Path] = []
    merged_path: Path | None = None
    try:
        if args.mode in ("split", "both"):
            evidence_dir.mkdir(parents=True, exist_ok=True)
            written = split_pdfs(reader, splittable, evidence_dir)
        if args.mode in ("merged", "both"):
            merged_path = out_dir / "合并卷宗_带书签.pdf"
            build_merged_pdf(reader, rows, merged_path)
    except PermissionError as exc:
        raise SplitterError(f"写入被拒绝，请关闭已打开的同名 PDF 后重试：{exc}") from exc

    write_xlsx(rows, out_dir / "证据目录.xlsx", args.title)
    write_markdown(rows, out_dir / "证据目录.md", args.title, all_issues)
    write_report(all_issues, out_dir / "校验报告.txt", source_pdf, source_hash, total_pages, len(written))

    print("输出完成：")
    if written:
        print(f"  切分文件：{len(written)} 份 -> {evidence_dir}")
        total_split_pages = sum(r.pages for r in splittable)
        print(f"  切分总页数：{total_split_pages} 页")
    if merged_path:
        print(f"  合并卷宗：{merged_path}")
    print(f"  证据目录：{out_dir / '证据目录.xlsx'}")
    print(f"  目录副本：{out_dir / '证据目录.md'}")
    print(f"  校验报告：{out_dir / '校验报告.txt'}")

    after_hash = file_sha256(source_pdf)
    if after_hash != source_hash:
        print("警告：源 PDF 的哈希在运行后发生变化，请立即核查源文件。")
        return 1
    print(f"源文件未被修改（SHA256 {after_hash[:16]}...）。")

    if args.strict and validate_warnings:
        print(f"--strict 模式：存在 {len(validate_warnings)} 项警告，按失败处理。")
        return 1
    return 0


def main(argv: Sequence[str] | None = None) -> int:
    relax_stdout_encoding()
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return run(args)
    except SplitterError as exc:
        sys.stderr.write(f"错误：{exc}\n")
        return 2
    except KeyboardInterrupt:
        sys.stderr.write("已中断。\n")
        return 130


if __name__ == "__main__":
    raise SystemExit(main())
