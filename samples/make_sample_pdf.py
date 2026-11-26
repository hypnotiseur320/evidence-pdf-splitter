#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成虚构样例卷宗与配套证据清单。

本脚本仅用于生成演示数据，不属于工具的运行依赖。
所有当事人、金额、编号均为虚构，与任何真实案件无关。

用法：python samples/make_sample_pdf.py
输出：samples/样例卷宗_虚构.pdf、samples/样例证据清单.xlsx、samples/样例证据清单.csv
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

HERE = Path(__file__).resolve().parent

CASE_TITLE = "张某诉宁波某某建筑工程有限公司建设工程施工合同纠纷"
SAMPLE_NOTE = "（虚构样例，仅用于工具演示，与任何真实案件无关）"

FONT_CANDIDATES = [
    ("C:/Windows/Fonts/simsun.ttc", 0),
    ("C:/Windows/Fonts/msyh.ttc", 0),
    ("C:/Windows/Fonts/simhei.ttf", None),
]


def register_font() -> str:
    for path, index in FONT_CANDIDATES:
        if not Path(path).exists():
            continue
        try:
            if index is None:
                pdfmetrics.registerFont(TTFont("CNSample", path))
            else:
                pdfmetrics.registerFont(TTFont("CNSample", path, subfontIndex=index))
            return "CNSample"
        except Exception:
            continue
    raise SystemExit("未找到可用的中文字体，请修改 FONT_CANDIDATES。")


# 每页正文：标题 + 若干段落。第 1-2 页为封面与目录，正文从卷宗第 1 页开始。
BODY_PAGES: list[tuple[str, list[str]]] = [
    (
        "借　条",
        [
            "今借到张某人民币伍拾万元整（¥500,000.00），用于宁波某某建筑工程有限公司",
            "材料采购周转。借款期限自二〇二三年三月一日起至二〇二三年八月三十一日止，",
            "月利率百分之一，到期一次还本付息。",
            "",
            "借款人：宁波某某建筑工程有限公司（盖章）",
            "法定代表人：李某某（签字）",
            "二〇二三年三月一日",
        ],
    ),
    (
        "中国某某银行客户回单（第 1/3 页）",
        [
            "付款人名称：张某",
            "付款人账号：6222 **** **** 1234",
            "收款人名称：宁波某某建筑工程有限公司",
            "收款人账号：3301 **** **** 5678",
            "交易金额：人民币贰拾万元整（¥200,000.00）",
            "交易日期：二〇二三年三月二日",
            "摘要：借款（材料款）",
        ],
    ),
    (
        "中国某某银行客户回单（第 2/3 页）",
        [
            "付款人名称：张某",
            "收款人名称：宁波某某建筑工程有限公司",
            "交易金额：人民币贰拾万元整（¥200,000.00）",
            "交易日期：二〇二三年三月六日",
            "摘要：借款（材料款）",
        ],
    ),
    (
        "中国某某银行客户回单（第 3/3 页）",
        [
            "付款人名称：张某",
            "收款人名称：宁波某某建筑工程有限公司",
            "交易金额：人民币壹拾万元整（¥100,000.00）",
            "交易日期：二〇二三年三月十日",
            "摘要：借款（材料款）",
            "",
            "合计：人民币伍拾万元整（¥500,000.00）",
        ],
    ),
    (
        "建设工程施工合同（第 1/8 页）",
        [
            "发包人：宁波某某置业有限公司",
            "承包人：宁波某某建筑工程有限公司",
            "工程名称：某某科创园二期幕墙工程",
            "工程地点：宁波市某某区某某路 88 号",
            "签约合同价：人民币捌佰陆拾万元整（¥8,600,000.00）",
        ],
    ),
    (
        "建设工程施工合同（第 2/8 页）",
        [
            "第一条　工程概况",
            "本工程为幕墙工程，包括玻璃幕墙、石材幕墙及铝合金门窗的供应与安装，",
            "建筑面积约 21,400 平方米，幕墙面积约 12,800 平方米。",
        ],
    ),
    (
        "建设工程施工合同（第 3/8 页）",
        [
            "第二条　合同工期",
            "计划开工日期：二〇二三年四月一日；计划竣工日期：二〇二三年十二月三十一日，",
            "总日历天数 275 天。因发包人原因导致停工的，工期相应顺延。",
        ],
    ),
    (
        "建设工程施工合同（第 4/8 页）",
        [
            "第三条　合同价款",
            "本合同采用固定单价方式计价。单价详见附件一《工程量清单》。",
            "合同价款中包含材料费、人工费、机械费、管理费、利润、规费及税金。",
        ],
    ),
    (
        "建设工程施工合同（第 5/8 页）",
        [
            "第四条　工程款支付",
            "预付款为合同价的百分之十，于开工前七日内支付；",
            "进度款按月计量支付，支付比例为已完成工程量的百分之八十；",
            "竣工结算款于结算审定后三十日内支付至结算价的百分之九十七；",
            "剩余百分之三为质量保证金，保修期满后无息返还。",
        ],
    ),
    (
        "建设工程施工合同（第 6/8 页）",
        [
            "第五条　工程变更",
            "发包人要求变更的，应提前七日以书面形式通知承包人。",
            "因变更引起的价款调整，由双方在变更洽商记录中确认，作为结算依据。",
        ],
    ),
    (
        "建设工程施工合同（第 7/8 页）",
        [
            "第六条　竣工验收",
            "承包人完成全部工程内容并自检合格后，应向发包人提交竣工验收报告。",
            "发包人应在收到报告后十四日内组织验收，逾期未组织且未提出异议的，",
            "视为验收合格。",
        ],
    ),
    (
        "建设工程施工合同（第 8/8 页）",
        [
            "第七条　违约责任",
            "发包人逾期支付工程款的，应按同期全国银行间同业拆借中心公布的",
            "贷款市场报价利率（LPR）向承包人支付逾期付款利息。",
            "",
            "（以下无正文，为签署页）",
            "发包人（盖章）：　　　　　　承包人（盖章）：",
            "签订日期：二〇二三年三月二十五日",
        ],
    ),
    (
        "单位工程竣工验收报告（第 1/2 页）",
        [
            "工程名称：某某科创园二期幕墙工程",
            "建设单位：宁波某某置业有限公司",
            "施工单位：宁波某某建筑工程有限公司",
            "验收日期：二〇二三年十二月二十八日",
        ],
    ),
    (
        "单位工程竣工验收报告（第 2/2 页）",
        [
            "验收结论：",
            "经建设单位、施工单位、监理单位、设计单位共同验收，本工程质量符合",
            "设计文件及施工合同约定的标准，观感质量良好，资料基本齐全，",
            "一致同意通过竣工验收。",
            "",
            "建设单位（盖章）：　　　　施工单位（盖章）：",
            "监理单位（盖章）：　　　　设计单位（盖章）：",
        ],
    ),
    (
        "催告函（第 1/2 页）",
        [
            "致：宁波某某置业有限公司",
            "",
            "就贵我双方签订的《建设工程施工合同》（某某科创园二期幕墙工程），",
            "本工程已于二〇二三年十二月二十八日通过竣工验收，结算送审金额为",
            "人民币玖佰壹拾贰万元整（¥9,120,000.00）。",
            "截至本函发出之日，贵方尚欠工程款人民币贰佰叁拾万元整（¥2,300,000.00）。",
        ],
    ),
    (
        "催告函（第 2/2 页）",
        [
            "现函告贵方：请于收到本函之日起十五日内支付上述欠款。",
            "逾期未付，我方将依法主张工程价款及逾期付款利息，并保留追究",
            "其他法律责任的权利。",
            "",
            "宁波某某建筑工程有限公司（盖章）",
            "二〇二四年三月十八日",
        ],
    ),
    (
        "回函（第 1/2 页）",
        [
            "致：宁波某某建筑工程有限公司",
            "",
            "贵司二〇二四年三月十八日《催告函》收悉。",
            "经核对，我方确认尚欠贵司幕墙工程款人民币贰佰叁拾万元整",
            "（¥2,300,000.00），对结算送审金额中的石材品种调差部分存在异议，",
            "该部分金额为人民币肆拾壹万元整（¥410,000.00）。",
        ],
    ),
    (
        "回函（第 2/2 页）",
        [
            "我方拟于近期组织内部复核，并就调差部分与贵司协商。",
            "在协商完成前，请贵司暂缓采取进一步措施。",
            "",
            "宁波某某置业有限公司（盖章）",
            "二〇二四年四月二日",
        ],
    ),
]


def draw_common(c: canvas.Canvas, font: str, page_no: int | None) -> None:
    width, height = A4
    c.setFont(font, 8)
    c.setFillGray(0.35)
    c.drawString(20 * mm, height - 15 * mm, CASE_TITLE[:28])
    c.setStrokeGray(0.7)
    c.setLineWidth(0.5)
    c.line(20 * mm, height - 17 * mm, width - 20 * mm, height - 17 * mm)

    c.line(20 * mm, 20 * mm, width - 20 * mm, 20 * mm)
    c.setFont(font, 8)
    c.setFillGray(0.45)
    c.drawString(20 * mm, 15 * mm, SAMPLE_NOTE)
    if page_no is not None:
        c.setFont(font, 9)
        c.setFillGray(0.2)
        c.drawRightString(width - 20 * mm, 15 * mm, f"卷宗第 {page_no} 页")


def draw_cover(c: canvas.Canvas, font: str) -> None:
    width, height = A4
    c.setFont(font, 34)
    c.setFillGray(0.1)
    c.drawCentredString(width / 2, height - 95 * mm, "卷　　宗")
    c.setFont(font, 13)
    c.drawCentredString(width / 2, height - 125 * mm, CASE_TITLE[:20])
    c.drawCentredString(width / 2, height - 140 * mm, CASE_TITLE[20:] if len(CASE_TITLE) > 20 else "")
    c.setFont(font, 10)
    c.setFillGray(0.45)
    c.drawCentredString(width / 2, 60 * mm, SAMPLE_NOTE)
    c.setFont(font, 9)
    c.drawCentredString(width / 2, 50 * mm, "本册共 18 页（不含封面与目录）")


def draw_index(c: canvas.Canvas, font: str, entries: list[str]) -> None:
    width, height = A4
    c.setFont(font, 15)
    c.setFillGray(0.1)
    c.drawCentredString(width / 2, height - 40 * mm, "目　　录")
    c.setFont(font, 10.5)
    y = height - 60 * mm
    for line in entries:
        c.drawString(30 * mm, y, line)
        y -= 8 * mm


def draw_body(c: canvas.Canvas, font: str, title: str, paragraphs: list[str]) -> None:
    width, height = A4
    c.setFont(font, 13.5)
    c.setFillGray(0.08)
    c.drawCentredString(width / 2, height - 40 * mm, title)

    c.setFont(font, 10.5)
    c.setFillGray(0.12)
    y = height - 58 * mm
    for line in paragraphs:
        c.drawString(28 * mm, y, line)
        y -= 8.2 * mm


def build_pdf(target: Path, font: str) -> None:
    width, height = A4
    c = canvas.Canvas(str(target), pagesize=A4)
    c.setTitle("虚构样例卷宗")
    c.setAuthor("evidence_pdf_splitter sample")

    # 前置页 1：封面（无卷宗页码）
    draw_cover(c, font)
    c.showPage()

    # 前置页 2：目录（无卷宗页码）
    entries = [
        "证据一　借　条 ·································· 第 1 页",
        "证据二　银行转账凭证 ····························· 第 2-4 页",
        "证据三　建设工程施工合同 ························· 第 5-12 页",
        "证据四　竣工验收报告 ····························· 第 13-14 页",
        "证据五　催告与往来函件 ··························· 第 15-18 页",
        "　　5-1　催告函 ································· 第 15-16 页",
        "　　5-2　回函 ··································· 第 17-18 页",
    ]
    draw_index(c, font, entries)
    c.showPage()

    # 正文 18 页：卷宗第 1-18 页
    for vol_page, (title, paragraphs) in enumerate(BODY_PAGES, start=1):
        draw_body(c, font, title, paragraphs)
        draw_common(c, font, vol_page)
        c.showPage()

    c.save()


HEADERS = ["证据编号", "证据名称", "起始页", "结束页", "证据来源", "证明目的", "原件/复印件", "备注"]
ROWS = [
    [1, "借条", 1, 1, "原告持有", "证明双方存在借贷合意及款项交付事实", "原件", "被告盖章并法定代表人签字"],
    [2, "银行转账凭证", 2, 4, "某某银行调取", "证明原告已实际支付借款50万元", "复印件", "三笔转账合计50万元"],
    [3, "建设工程施工合同", 5, 12, "双方签订", "证明施工范围、合同价款及付款节点约定", "复印件", "含签署页"],
    [4, "竣工验收报告", 13, 14, "建设单位提供", "证明工程已于2023年12月28日验收合格", "复印件", "四方盖章"],
    [5, "催告与往来函件", None, None, "—", "组证据，不单独提交", "—", "由5-1、5-2组成"],
    ["5-1", "催告函", 15, 16, "原告发出", "证明原告已书面催告被告支付欠款", "复印件", "2024年3月18日发出"],
    ["5-2", "回函", 17, 18, "被告出具", "证明被告确认欠付工程款230万元", "复印件", "2024年4月2日出具"],
]


def build_xlsx(target: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "证据清单"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for column, header in enumerate(HEADERS, start=1):
        cell = sheet.cell(row=1, column=column, value=header)
        cell.font = Font(name="宋体", size=11, bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F3864")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border
    widths = [12, 24, 10, 10, 18, 40, 14, 28]
    for index, width in enumerate(widths, start=1):
        sheet.column_dimensions[chr(64 + index)].width = width

    for row_index, row in enumerate(ROWS, start=2):
        for column, value in enumerate(row, start=1):
            cell = sheet.cell(row=row_index, column=column, value=value)
            cell.font = Font(name="宋体", size=10.5)
            cell.border = border
            cell.alignment = Alignment(
                horizontal="center" if column in (1, 3, 4, 7) else "left",
                vertical="center",
                wrap_text=True,
            )
    sheet.freeze_panes = "A2"
    workbook.save(str(target))


def build_csv(target: Path) -> None:
    import csv

    with target.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle)
        writer.writerow(HEADERS)
        for row in ROWS:
            writer.writerow(["" if value is None else value for value in row])


def main() -> None:
    font = register_font()
    pdf_path = HERE / "样例卷宗_虚构.pdf"
    build_pdf(pdf_path, font)
    build_xlsx(HERE / "样例证据清单.xlsx")
    build_csv(HERE / "样例证据清单.csv")
    print(f"已生成：{pdf_path}")
    print(f"已生成：{HERE / '样例证据清单.xlsx'}")
    print(f"已生成：{HERE / '样例证据清单.csv'}")


if __name__ == "__main__":
    main()
