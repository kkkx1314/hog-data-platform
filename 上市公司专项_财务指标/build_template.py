# -*- coding: utf-8 -*-
"""生成线下 Excel 模板：全部 21 家公司的指标趋势、指标快照与五项资产。

用法：
    D:\\miniconda\\python.exe build_template.py

产出：
    output/上市公司财务指标_趋势与资产_模板.xlsx

模板不含宏，靠数据验证下拉 + SUMIFS 公式动态取数，WPS/Excel 打开自动重算。
数据更新后（build_data.py 跑完）再跑本脚本即可刷新模板。
"""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.data_source import StrRef
from openpyxl.chart.marker import Marker
from openpyxl.chart.text import Text
from openpyxl.chart.title import Title
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from build_data import CAT_MAP, FORMULAS, MULT_IND, PCT_IND

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
OUT = ROOT / "output" / "上市公司财务指标_趋势与资产_模板.xlsx"

ITEMS = ["投资性房地产", "固定资产", "在建工程", "生产性生物资产", "存货"]
ITEM_COLORS = ["#C0392B", "#2E5EAA", "#7FA8D9", "#2A9D8F", "#E9C46A"]
COMPANY_COLOR = {
    "牧原": "#2E5EAA", "温氏": "#D9822B", "新希望": "#2A9D8F", "正邦": "#8E44AD",
    "天邦": "#C0392B", "新五丰": "#16A085", "唐人神": "#E67E22", "巨星": "#2980B9",
    "天康": "#27AE60", "大北农": "#F39C12", "神农": "#D35400", "华统": "#7F8C8D",
    "立华": "#3498DB", "傲农": "#9B59B6", "京基智农": "#1ABC9C", "东瑞": "#E74C3C",
    "金新农": "#34495E", "罗牛山": "#B7950B", "克明": "#A04000", "益生": "#5DADE2",
    "正虹": "#AF7AC5",
}

TITLE_FILL = PatternFill("solid", fgColor="1F3864")
HEAD_FILL = PatternFill("solid", fgColor="2F5597")
SECTION_FILL = PatternFill("solid", fgColor="D9E2F3")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(bottom=THIN)


def period_key(p: str) -> int:
    y, q = int(p[:4]), p[4:]
    return y * 4 + {"Q1": 0, "H1": 1, "Q3": 2, "FY": 3}[q]


def ind_type(name: str) -> str:
    if name in PCT_IND:
        return "pct"
    if name in ("人均创收", "人均核心利润"):
        return "percap"
    return "mult"


def unit_of(typ: str) -> str:
    return {"pct": "%", "percap": "万元/人", "mult": "倍"}[typ]


def hx(color: str) -> str:
    """openpyxl 写 srgbClr 时不能带 #，否则 Excel 判定文件损坏。"""
    return color.lstrip("#")


def style_title(ws, text: str, last_col: int) -> None:
    c = ws.cell(1, 1, text)
    c.font = Font(bold=True, size=14, color="FFFFFF")
    c.fill = TITLE_FILL
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_col)
    ws.row_dimensions[1].height = 28


def style_header(ws, row: int, ncol: int) -> None:
    for j in range(1, ncol + 1):
        c = ws.cell(row, j)
        c.font = WHITE_BOLD
        c.fill = HEAD_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = 30


def section(ws, row: int, text: str, ncol: int) -> None:
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, color="1F3864")
    c.fill = SECTION_FILL
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncol)


def add_list_validation(ws, cell: str, formula: str) -> None:
    dv = DataValidation(type="list", formula1=formula, allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(cell)


def line_chart(data, cats, title, colors, height=20, width=32, numfmt="0.0000", title_ref=None):
    ch = LineChart()
    if title_ref:
        ch.title = Title(tx=Text(strRef=StrRef(title_ref)))
    else:
        ch.title = title
    ch.height = height
    ch.width = width
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.legend.position = "r"
    ch.legend.overlay = False
    ch.y_axis.numFmt = numfmt
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    for s, c in zip(ch.series, colors):
        s.graphicalProperties.line.solidFill = hx(c)
        s.graphicalProperties.line.width = 19000
        s.marker = Marker(symbol="circle", size=4)
        s.marker.graphicalProperties.solidFill = hx(c)
        s.marker.graphicalProperties.line.solidFill = hx(c)
    return ch


def stacked_bar(data, cats, title, colors, horizontal=True, height=12, width=30):
    ch = BarChart()
    ch.type = "bar" if horizontal else "col"
    ch.grouping = "stacked"
    ch.overlap = 100
    ch.gapWidth = 60
    if title:
        ch.title = title
    ch.height = height
    ch.width = width
    ch.add_data(data, titles_from_data=True)
    ch.set_categories(cats)
    ch.legend.position = "b"
    ch.legend.overlay = False
    ch.y_axis.numFmt = "0.0"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if horizontal:
        # 让图中公司顺序与上方表格一致（自上而下）
        ch.x_axis.scaling.orientation = "maxMin"
        ch.y_axis.crosses = "max"
    for s, c in zip(ch.series, colors):
        s.graphicalProperties.solidFill = hx(c)
    return ch


def main() -> None:
    q = pd.read_csv(DATA / "indicators_single.csv")
    c = pd.read_csv(DATA / "indicators_cum.csv")
    a = pd.read_csv(DATA / "assets_five.csv")

    companies = list(dict.fromkeys(q["公司"].tolist()))
    periods = sorted(q["期间"].unique(), key=period_key)
    inds = [x[1] for x in FORMULAS]

    # ------------------------------------------------------------ 长表
    q_long = q.melt(id_vars=["公司", "期间"], value_vars=inds,
                    var_name="指标", value_name="数值")
    q_long.insert(0, "口径", "单季")
    c_long = c.melt(id_vars=["公司", "期间"], value_vars=inds,
                    var_name="指标", value_name="数值")
    c_long.insert(0, "口径", "累计")
    ind_long = pd.concat([q_long, c_long], ignore_index=True)
    ind_long = ind_long[["口径", "指标", "公司", "期间", "数值"]]
    ind_long["指标"] = pd.Categorical(ind_long["指标"], categories=inds, ordered=True)
    ind_long["公司"] = pd.Categorical(ind_long["公司"], categories=companies, ordered=True)
    ind_long["期间"] = pd.Categorical(ind_long["期间"], categories=periods, ordered=True)
    ind_long = ind_long.sort_values(["口径", "指标", "公司", "期间"]).reset_index(drop=True)

    assets = a[["公司", "期间"] + ITEMS + ["五项合计", "资产总计"]].copy()
    assets["公司"] = pd.Categorical(assets["公司"], categories=companies, ordered=True)
    assets["期间"] = pd.Categorical(assets["期间"], categories=periods, ordered=True)
    assets = assets.sort_values(["公司", "期间"]).reset_index(drop=True)

    n_ind_rows = len(ind_long) + 1   # 含表头，公式用 $2:$n
    n_ast_rows = len(assets) + 1
    IND_RNG = f"'数据_指标'!$A$2:$A${n_ind_rows}"
    IND_RNG_B = f"'数据_指标'!$B$2:$B${n_ind_rows}"
    IND_RNG_C = f"'数据_指标'!$C$2:$C${n_ind_rows}"
    IND_RNG_D = f"'数据_指标'!$D$2:$D${n_ind_rows}"
    IND_RNG_E = f"'数据_指标'!$E$2:$E${n_ind_rows}"
    AST_RNG_A = f"'数据_资产'!$A$2:$A${n_ast_rows}"
    AST_RNG_B = f"'数据_资产'!$B$2:$B${n_ast_rows}"
    AST_RNG_ITEM = [f"'数据_资产'!${get_column_letter(3 + k)}$2:${get_column_letter(3 + k)}${n_ast_rows}"
                    for k in range(5)]
    AST_RNG_TOT = f"'数据_资产'!$I$2:$I${n_ast_rows}"

    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    ws0 = wb.active
    ws0.title = "00_说明"

    # ------------------------------------------------------------ 参数（隐藏）
    wp = wb.create_sheet("参数")
    wp["A1"], wp["B1"], wp["C1"], wp["D1"] = "口径", "指标", "公司", "期间"
    for i, v in enumerate(["单季", "累计"], start=2):
        wp.cell(i, 1, v)
    for i, v in enumerate(inds, start=2):
        wp.cell(i, 2, v)
    for i, v in enumerate(companies, start=2):
        wp.cell(i, 3, v)
    for i, v in enumerate(periods, start=2):
        wp.cell(i, 4, v)
    wp["F1"], wp["G1"], wp["H1"], wp["I1"], wp["J1"], wp["K1"] = "指标", "类型", "单位", "分类", "方向", "公式"
    fmap = {x[1]: x for x in FORMULAS}
    for i, name in enumerate(inds, start=2):
        cat, _nm, formula, _note = fmap[name]
        typ = ind_type(name)
        wp.cell(i, 6, name)
        wp.cell(i, 7, typ)
        wp.cell(i, 8, unit_of(typ))
        wp.cell(i, 9, cat)
        wp.cell(i, 10, "越高越好" if name not in ("期间费用率", "有息负债率", "重资产占比") else "越低越好")
        wp.cell(i, 11, formula)
    wp.sheet_state = "hidden"

    # ------------------------------------------------------------ 数据_指标
    ws1 = wb.create_sheet("数据_指标")
    ws1.append(["口径", "指标", "公司", "期间", "数值"])
    for row in ind_long.itertuples(index=False):
        ws1.append([row.口径, row.指标, row.公司, row.期间,
                    None if pd.isna(row.数值) else float(row.数值)])
    style_header(ws1, 1, 5)
    for col, w in zip("ABCDE", (8, 18, 10, 9, 14)):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = "A2"
    for cell in ws1["E"][1:]:
        cell.number_format = "0.00000000"
    ws1.auto_filter.ref = f"A1:E{n_ind_rows}"

    # ------------------------------------------------------------ 数据_资产
    ws2 = wb.create_sheet("数据_资产")
    ws2.append(["公司", "期间"] + ITEMS + ["五项合计", "资产总计"])
    for row in assets.itertuples(index=False):
        vals = [row.公司, row.期间]
        vals += [None if pd.isna(v) else float(v) for v in row[2:]]
        ws2.append(vals)
    style_header(ws2, 1, 9)
    for col, w in zip("ABCDEFGHI", (10, 9, 14, 14, 14, 16, 14, 14, 16)):
        ws2.column_dimensions[col].width = w
    ws2.freeze_panes = "A2"
    for row in ws2.iter_rows(min_row=2, min_col=3, max_col=9):
        for cell in row:
            cell.number_format = "#,##0"
    ws2.auto_filter.ref = f"A1:I{n_ast_rows}"

    # ------------------------------------------------------------ 01 指标趋势
    wt = wb.create_sheet("01_指标趋势")
    style_title(wt, "全部公司指标趋势（21 家 · 2020Q1–2026Q3）", len(companies) + 1)
    wt["A3"] = "口径"
    wt["B3"] = "单季"
    wt["A4"] = "指标"
    wt["B4"] = "核心利润率"
    wt["A5"] = "单位"
    wt["B5"] = ("=IFERROR(INDEX(参数!$H$2:$H$21,MATCH($B$4,参数!$F$2:$F$21,0)),\"\")")
    wt["A6"] = "说明"
    wt["B6"] = "切换口径/指标后图表自动更新；趋势图为全部 21 家公司。比率类数值为小数，0.1234 = 12.34%。"
    for r in (3, 4, 5, 6):
        wt.cell(r, 1).font = Font(bold=True)
    add_list_validation(wt, "B3", "=参数!$A$2:$A$3")
    add_list_validation(wt, "B4", "=参数!$B$2:$B$21")

    head_row = 8
    wt.cell(head_row, 1, "报告期")
    for j, comp in enumerate(companies, start=2):
        wt.cell(head_row, j, comp)
    style_header(wt, head_row, len(companies) + 1)
    wt.freeze_panes = "B9"

    first_data = head_row + 1
    last_data = first_data + len(periods) - 1
    for i, p in enumerate(periods):
        r = first_data + i
        wt.cell(r, 1, p).font = Font(bold=True)
        for j in range(2, len(companies) + 2):
            col = get_column_letter(j)
            wt.cell(r, j).value = (
                f'=IF(COUNTIFS({IND_RNG},$B$3,{IND_RNG_B},$B$4,{IND_RNG_C},{col}${head_row},'
                f'{IND_RNG_D},$A{r},{IND_RNG_E},"<>")=0,"",SUMIFS({IND_RNG_E},{IND_RNG},$B$3,{IND_RNG_B},$B$4,'
                f'{IND_RNG_C},{col}${head_row},{IND_RNG_D},$A{r}))'
            )
            wt.cell(r, j).number_format = "0.0000"
        wt.row_dimensions[r].height = 18

    wt.column_dimensions["A"].width = 9
    for j in range(2, len(companies) + 2):
        wt.column_dimensions[get_column_letter(j)].width = 10

    data = Reference(wt, min_col=2, max_col=len(companies) + 1,
                     min_row=head_row, max_row=last_data)
    cats = Reference(wt, min_col=1, min_row=first_data, max_row=last_data)
    colors = [COMPANY_COLOR.get(x, "#7F8C8D") for x in companies]
    wt.add_chart(line_chart(data, cats, "全部公司指标趋势", colors,
                            height=20, width=32, title_ref="'01_指标趋势'!$B$4"), "A38")

    # ------------------------------------------------------------ 02 五项资产
    wa = wb.create_sheet("02_五项资产")
    style_title(wa, "五项资产（全部 21 家公司）", 11)
    wa["A3"] = "对比报告期"
    wa["B3"] = periods[-1]
    wa["A4"] = "单家公司"
    wa["B4"] = companies[0]
    wa["A5"] = "说明"
    wa["B5"] = ("上方为全部公司同一报告期的结构对比；下方按所选公司展示 2020Q1 以来走势。"
                "金额单位：亿元；空白表示该公司该期尚未披露（如新希望截至 2026H1）。")
    for r in (3, 4, 5):
        wa.cell(r, 1).font = Font(bold=True)
    add_list_validation(wa, "B3", f"=参数!$D$2:$D${len(periods) + 1}")
    add_list_validation(wa, "B4", f"=参数!$C$2:$C${len(companies) + 1}")

    # 全部公司横向
    section(wa, 7, "全部公司 · 五项资产（对比报告期，亿元）", 8)
    head_all = 8
    for j, h in enumerate(["公司"] + ITEMS + ["五项合计", "资产总计"], start=1):
        wa.cell(head_all, j, h)
    style_header(wa, head_all, 8)
    first_all = head_all + 1
    for i, comp in enumerate(companies):
        r = first_all + i
        wa.cell(r, 1, comp)
        for k, rng in enumerate(AST_RNG_ITEM):
            col = get_column_letter(2 + k)
            wa.cell(r, 2 + k).value = (
                f'=IF(COUNTIFS({AST_RNG_A},$A{r},{AST_RNG_B},$B$3,{rng},"<>")=0,"",'
                f'SUMIFS({rng},{AST_RNG_A},$A{r},{AST_RNG_B},$B$3)/100000000)'
            )
            wa.cell(r, 2 + k).number_format = "0.0"
        wa.cell(r, 7).value = f'=IF(COUNT($B{r}:$F{r})=0,"",SUM($B{r}:$F{r}))'
        wa.cell(r, 8).value = (
            f'=IF(COUNTIFS({AST_RNG_A},$A{r},{AST_RNG_B},$B$3,{AST_RNG_TOT},"<>")=0,"",'
            f'SUMIFS({AST_RNG_TOT},{AST_RNG_A},$A{r},{AST_RNG_B},$B$3)/100000000)'
        )
        wa.cell(r, 7).number_format = "0.0"
        wa.cell(r, 8).number_format = "0.0"
    last_all = first_all + len(companies) - 1

    bar_data = Reference(wa, min_col=2, max_col=6, min_row=head_all, max_row=last_all)
    bar_cats = Reference(wa, min_col=1, min_row=first_all, max_row=last_all)
    wa.add_chart(stacked_bar(bar_data, bar_cats, None,
                             ITEM_COLORS, horizontal=True, height=13, width=30), "A31")

    # 单家公司走势
    sec2 = 60
    section(wa, sec2, "单家公司 · 五项资产走势（所选公司，亿元）", 11)
    head_one = sec2 + 1
    for j, h in enumerate(["期间"] + ITEMS + ["五项合计", "资产总计",
                                              "五项/总资产", "固定资产/总资产", "存货/总资产"], start=1):
        wa.cell(head_one, j, h)
    style_header(wa, head_one, 11)
    first_one = head_one + 1
    for i, p in enumerate(periods):
        r = first_one + i
        wa.cell(r, 1, p).font = Font(bold=True)
        for k, rng in enumerate(AST_RNG_ITEM):
            wa.cell(r, 2 + k).value = (
                f'=IF(COUNTIFS({AST_RNG_A},$B$4,{AST_RNG_B},$A{r},{rng},"<>")=0,"",'
                f'SUMIFS({rng},{AST_RNG_A},$B$4,{AST_RNG_B},$A{r})/100000000)'
            )
            wa.cell(r, 2 + k).number_format = "0.0"
        wa.cell(r, 7).value = f'=IF(COUNT($B{r}:$F{r})=0,"",SUM($B{r}:$F{r}))'
        wa.cell(r, 8).value = (
            f'=IF(COUNTIFS({AST_RNG_A},$B$4,{AST_RNG_B},$A{r},{AST_RNG_TOT},"<>")=0,"",'
            f'SUMIFS({AST_RNG_TOT},{AST_RNG_A},$B$4,{AST_RNG_B},$A{r})/100000000)'
        )
        wa.cell(r, 9).value = f'=IFERROR($G{r}/$H{r},"")'
        wa.cell(r, 10).value = f'=IFERROR($C{r}/$H{r},"")'
        wa.cell(r, 11).value = f'=IFERROR($F{r}/$H{r},"")'
        for cc in (7, 8):
            wa.cell(r, cc).number_format = "0.0"
        for cc in (9, 10, 11):
            wa.cell(r, cc).number_format = "0.0%"
    last_one = first_one + len(periods) - 1

    wa.column_dimensions["A"].width = 11
    for col, w in zip("BCDEFGHIJK", (14, 14, 14, 16, 14, 14, 16, 15, 15, 14)):
        wa.column_dimensions[col].width = w
    wa.freeze_panes = "B9"

    one_data = Reference(wa, min_col=2, max_col=6, min_row=head_one, max_row=last_one)
    one_cats = Reference(wa, min_col=1, min_row=first_one, max_row=last_one)
    wa.add_chart(stacked_bar(one_data, one_cats, "单家公司 · 五项资产结构（亿元）",
                             ITEM_COLORS, horizontal=False, height=13, width=16), "M31")
    share = LineChart()
    share.title = "占总资产比重"
    share.height = 12
    share.width = 16
    share.add_data(Reference(wa, min_col=9, max_col=11, min_row=head_one, max_row=last_one),
                   titles_from_data=True)
    share.set_categories(one_cats)
    share.legend.position = "b"
    share.legend.overlay = False
    share.y_axis.numFmt = "0%"
    share.y_axis.delete = False
    share.x_axis.delete = False
    for s, c in zip(share.series, ["#2E5EAA", "#7FA8D9", "#E9C46A"]):
        s.graphicalProperties.line.solidFill = hx(c)
        s.graphicalProperties.line.width = 22000
        s.marker = Marker(symbol="circle", size=4)
    wa.add_chart(share, "M60")

    # ------------------------------------------------------------ 03 指标快照
    wsn = wb.create_sheet("03_指标快照")
    style_title(wsn, "全部公司指标快照（20 项指标 × 21 家公司）", len(companies) + 2)
    wsn["A3"] = "口径"
    wsn["B3"] = "单季"
    wsn["A4"] = "报告期"
    wsn["B4"] = periods[-1]
    wsn["A5"] = "说明"
    wsn["B5"] = "切换口径/报告期后整表自动更新；比率类为小数，0.1234 = 12.34%。"
    for r in (3, 4, 5):
        wsn.cell(r, 1).font = Font(bold=True)
    add_list_validation(wsn, "B3", "=参数!$A$2:$A$3")
    add_list_validation(wsn, "B4", f"=参数!$D$2:$D${len(periods) + 1}")

    head_s = 7
    wsn.cell(head_s, 1, "分类")
    wsn.cell(head_s, 2, "指标")
    for j, comp in enumerate(companies, start=3):
        wsn.cell(head_s, j, comp)
    style_header(wsn, head_s, len(companies) + 2)
    for i, name in enumerate(inds):
        r = head_s + 1 + i
        typ = ind_type(name)
        wsn.cell(r, 1, CAT_MAP.get(name, "")).font = Font(size=9, color="666666")
        wsn.cell(r, 2, name).font = Font(bold=True)
        for j in range(3, len(companies) + 3):
            col = get_column_letter(j)
            wsn.cell(r, j).value = (
                f'=IF(COUNTIFS({IND_RNG},$B$3,{IND_RNG_B},$B{r},{IND_RNG_C},{col}${head_s},'
                f'{IND_RNG_D},$B$4,{IND_RNG_E},"<>")=0,"",SUMIFS({IND_RNG_E},{IND_RNG},$B$3,{IND_RNG_B},$B{r},'
                f'{IND_RNG_C},{col}${head_s},{IND_RNG_D},$B$4))'
            )
            wsn.cell(r, j).number_format = {"pct": "0.00%", "percap": "0.0", "mult": "0.00"}[typ]
        wsn.row_dimensions[r].height = 18
    wsn.column_dimensions["A"].width = 11
    wsn.column_dimensions["B"].width = 18
    for j in range(3, len(companies) + 3):
        wsn.column_dimensions[get_column_letter(j)].width = 10
    wsn.freeze_panes = "C8"

    # ------------------------------------------------------------ 09 口径说明
    wsm = wb.create_sheet("09_口径说明")
    style_title(wsm, "指标公式与口径", 6)
    heads = ["分类", "指标", "类型", "单位", "方向", "计算公式", "说明"]
    for j, h in enumerate(heads, start=1):
        wsm.cell(2, j, h)
    style_header(wsm, 2, len(heads))
    for i, (cat, name, formula, note) in enumerate(FORMULAS, start=3):
        typ = ind_type(name)
        wsm.cell(i, 1, cat)
        wsm.cell(i, 2, name)
        wsm.cell(i, 3, {"pct": "比率", "mult": "倍数", "percap": "人均"}[typ])
        wsm.cell(i, 4, unit_of(typ))
        wsm.cell(i, 5, "越高越好" if name not in ("期间费用率", "有息负债率", "重资产占比") else "越低越好")
        wsm.cell(i, 6, formula)
        wsm.cell(i, 7, note)
        for j in range(1, 8):
            wsm.cell(i, j).alignment = Alignment(vertical="top", wrap_text=True)
        wsm.row_dimensions[i].height = 42
    for col, w in zip("ABCDEFG", (10, 18, 8, 10, 10, 52, 62)):
        wsm.column_dimensions[col].width = w
    wsm.freeze_panes = "A3"

    # ------------------------------------------------------------ 00 说明
    style_title(ws0, "上市猪企财务指标 · 线下 Excel 模板（21 家）", 6)
    lines = [
        ("文件定位", "无需联网、无需 Streamlit 的线下模板，覆盖全部 21 家上市猪企、2020Q1–2026Q3 共 556 个公司-期。"),
        ("01_指标趋势", "下拉选择口径（单季/累计）与指标，一张图同时画出全部 21 家公司；附完整数据矩阵。"),
        ("02_五项资产", "上方为全部公司同一报告期的结构对比（横向堆叠柱），下方按所选公司看五项资产走势与占比。"),
        ("03_指标快照", "所选报告期的 20 项指标 × 21 家公司横向表，切换口径/报告期自动更新。"),
        ("数据表", "“数据_指标”为 20 指标 × 2 口径长表；\"数据_资产\"为五项资产明细。模板所有图表由 SUMIFS 公式取数。"),
        ("更新方法", "源数据更新后先跑 build_data.py，再跑 build_template.py 重新生成本文件。"),
        ("打开提示", "公式为动态取数，首次打开若显示空白，按 F9（WPS）或 Ctrl+Alt+F9（Excel）强制重算。"),
        ("口径约定", "单季=累计差分；比率类已年化；毛利率只扣营业成本，核心利润率扣营业总成本；详见 09_口径说明。"),
    ]
    for i, (k, v) in enumerate(lines, start=3):
        ws0.cell(i, 1, k).font = Font(bold=True)
        ws0.cell(i, 2, v).alignment = Alignment(wrap_text=True, vertical="top")
        ws0.merge_cells(start_row=i, start_column=2, end_row=i, end_column=6)
        ws0.row_dimensions[i].height = 44
    ws0.column_dimensions["A"].width = 14

    # tab 颜色
    for name, color in [("01_指标趋势", "2E5EAA"), ("02_五项资产", "2A9D8F"),
                        ("03_指标快照", "D9822B")]:
        wb[name].sheet_properties.tabColor = color

    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print("完成 ->", OUT)
    print(f"公司 {len(companies)} 家，期间 {len(periods)} 个，指标 {len(inds)} 项，"
          f"指标长表 {len(ind_long):,} 行，资产 {len(assets):,} 行")


if __name__ == "__main__":
    main()
