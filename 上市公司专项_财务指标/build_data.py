# -*- coding: utf-8 -*-
"""
上市公司专项（21 家上市猪企）财务指标底表构建

数据源：
  D:\\CC\\_shangshi_extract_all\\上市公司专项\\<公司>\\2.三大报表...\\*.xlsx
  D:\\CC\\_shangshi_extract_all\\上市公司专项\\<公司>\\3.原始公告...\\*.pdf

产出：
  data/panel_long.csv          面板长表（公司×期间×口径×指标）
  data/indicators_single.csv   单季指标宽表
  data/indicators_cum.csv      累计指标宽表
  data/industry_long.csv       行业合计指标（合并分子/合并分母）
  data/assets_five.csv         五项资产明细
  output/上市公司财务指标_底表.xlsx
"""
from __future__ import annotations

import os
import re
import glob
import json
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl
import fitz

warnings.filterwarnings("ignore")

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUT_DIR = ROOT / "output"
DATA_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

SRC = Path(r"D:\CC\_shangshi_extract_all\上市公司专项")

# 文件夹, 简称, 证券简称, 代码。顺序即底表/看板中的公司顺序（原压缩包编号）。
COMPANY_SPECS = [
    ("1.牧原", "牧原", "牧原股份", "002714.SZ"),
    ("2.温氏", "温氏", "温氏股份", "300498.SZ"),
    ("3.新希望", "新希望", "新希望", "000876.SZ"),
    ("5.正邦", "正邦", "正邦科技", "002157.SZ"),
    ("6.天邦", "天邦", "天邦食品", "002124.SZ"),
    ("8.新五丰", "新五丰", "新五丰", "600975.SH"),
    ("9.唐人神", "唐人神", "唐人神", "002567.SZ"),
    ("10.巨星", "巨星", "巨星农牧", "603477.SH"),
    ("11.天康", "天康", "天康生物", "002100.SZ"),
    ("12.大北农", "大北农", "大北农", "002385.SZ"),
    ("13.神农", "神农", "神农集团", "605296.SH"),
    ("14.华统", "华统", "华统股份", "002840.SZ"),
    ("15.立华", "立华", "立华股份", "300761.SZ"),
    ("16.傲农", "傲农", "傲农生物", "603363.SH"),
    ("17.京基智农", "京基智农", "京基智农", "000048.SZ"),
    ("18.东瑞", "东瑞", "东瑞股份", "001201.SZ"),
    ("19,金新农", "金新农", "金新农", "002548.SZ"),
    ("20.罗牛山", "罗牛山", "罗牛山", "000735.SZ"),
    ("21.克明", "克明", "克明食品", "002661.SZ"),
    ("22.益生", "益生", "益生股份", "002458.SZ"),
    ("23.正虹", "正虹", "正虹科技", "000702.SZ"),
]
COMPANY_FULL = {short: f"{name}({code})" for _, short, name, code in COMPANY_SPECS}
COMPANY_ORDER = [short for _, short, _, _ in COMPANY_SPECS]
COMPANY_DIR = {short: SRC / folder for folder, short, _, _ in COMPANY_SPECS}

# 源 xlsx 个别季度利润表列填的是单季值（与 PDF 报告的年初至今累计口径不符），需加回上半年累计。
# 巨星 2020Q3：源值 650,811,413.39 = 单季；PDF 三季报 1-9 月累计 811,412,779.92 = 2020H1 160,601,366.53 + 单季，
# 年报“分季度主要财务数据”亦为 Q1 77,918,476.05 / Q2 82,682,890.48 / Q3 650,811,413.39 / Q4 627,777,824.17。
SINGLE_Q_IS = {("巨星", "2020Q3")}

# 源 xlsx 个别期间利润行符号整体翻转；2023 年报原文：新五丰归母净利润 -1,203,418,006.16 元，
# 营业利润 -1,219,955,337.75、利润总额 -1,250,774,407.63、净利润 -1,255,307,733.70。
SIGN_FLIP = {
    ("新五丰", "2023FY"): (
        "三、营业利润（亏损以“—”号填列）",
        "四、利润总额（亏损总额以“—”号填列）",
        "五、净利润（净亏损以“—”号填列）",
        "1.持续经营净利润",
        "1.归属于母公司股东的净利润",
        "七、综合收益总额",
        "归属于母公司所有者的综合收益总额",
    ),
}


def _find_xlsx(folder: Path) -> Path | None:
    """三大报表目录（2.开头）下的 xlsx；兼容编号/目录名微调。"""
    for d in sorted(folder.glob("2.*")):
        if d.is_dir():
            xs = sorted(p for p in d.glob("*.xlsx") if not p.name.startswith("~$"))
            if xs:
                return xs[0]
    xs = sorted(p for p in folder.glob("*.xlsx") if "三大报表" in p.name)
    return xs[0] if xs else None


XLSX_PAT = {}
for _short in COMPANY_ORDER:
    _p = _find_xlsx(COMPANY_DIR[_short])
    if _p is not None:
        XLSX_PAT[_short] = _p

# ---------------------------------------------------------------- 报表科目
BS_ITEMS = {
    "货币资金": "货币资金",
    "存货": "存货",
    "投资性房地产": "投资性房地产",
    "固定资产": "固定资产",
    "在建工程": "在建工程",
    "生产性生物资产": "生产性生物资产",
    "资产总计": "资产总计",
    "流动负债合计": "流动负债合计",
    "负债合计": "负债合计",
    "短期借款": "短期借款",
    "一年内到期的非流动负债": "一年内到期的非流动负债",
    "长期借款": "长期借款",
    "应付债券": "应付债券",
    "租赁负债": "租赁负债",
    "归属于母公司所有者权益合计": "归属于母公司所有者权益合计",
    "所有者权益合计": "所有者权益合计",
}
IS_ITEMS = {
    "营业收入": "营业收入",
    "营业成本": "营业成本",
    "税金及附加": "税金及附加",
    "销售费用": "销售费用",
    "管理费用": "管理费用",
    "研发费用": "研发费用",
    "财务费用": "财务费用",
    "利息费用": "利息费用",
    "所得税费用": "所得税费用",
    "归母净利润": "归属于母公司股东的净利润",
    "营业利润": "营业利润",
}
CF_ITEMS = {
    "销售商品收到的现金": "销售商品、提供劳务收到的现金",
    "经营活动现金流入小计": "经营活动现金流入小计",
    "经营活动现金流量净额": "经营活动产生的现金流量净额",
    "购建长期资产支付的现金": "购建固定资产、无形资产和其他长期资产支付的现金",
    "投资活动现金流入小计": "投资活动现金流入小计",
    "筹资活动现金流入小计": "筹资活动现金流入小计",
}
FLOW_ITEMS = list(IS_ITEMS) + list(CF_ITEMS)


def norm_label(x) -> str:
    if x is None:
        return ""
    return re.sub(r"\s+", "", str(x)).replace("\u3000", "")


def clean_label(x) -> str:
    """\u6bd4\u8f83\u7528\uff1a\u53bb\u6389\u201c\uff08\u6216\u80a1\u4e1c\u6743\u76ca\uff09\u201d\u7b49\u62ec\u6ce8\uff0c\u517c\u5bb9\u4e0d\u540c\u5e74\u62a5\u7684\u79d1\u76ee\u5199\u6cd5\u3002"""
    s = norm_label(x)
    for junk in ("\uff08\u6216\u80a1\u4e1c\u6743\u76ca\uff09", "(\u6216\u80a1\u4e1c\u6743\u76ca)", "\uff08\u6216\u80a1\u4e1c\u6743\u76ca", "\uff08\u6216\u80a1\u672c\uff09"):
        s = s.replace(junk, "")
    return s


def to_num(v):
    if v is None:
        return np.nan
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "—", "－", "不适用", "nan", "None"):
        return np.nan
    try:
        return float(s)
    except ValueError:
        return np.nan


def read_sheet(path: Path, sheet: str):
    """read_only 模式下必须用 iter_rows 批量读，逐格 cell() 会慢到不可用。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet]
    it = ws.iter_rows(values_only=True)
    header = next(it)
    periods = [str(v).strip() for v in header[1:] if v is not None and str(v).strip()]
    rows = []  # (row_no, label, [values])
    for i, vals in enumerate(it, start=2):
        label = norm_label(vals[0] if vals else None)
        if not label:
            continue
        nums = [to_num(v) for v in vals[1:1 + len(periods)]]
        rows.append((i, label, nums))
    wb.close()
    return periods, rows


def pick_row(rows, name: str, exclude=()):
    """按包含关系取第一行；name 为规范名，exclude 为排除词（均在净化后的科目名上比较）。"""
    for r, label, vals in rows:
        cl = clean_label(label)
        if name in cl and not any(e in cl for e in exclude):
            return r, label, vals
    return None, None, None


def load_statements():
    """返回 df_all(long), periods_map。"""
    recs = []
    periods_map = {}
    missing = [c for c in COMPANY_ORDER if c not in XLSX_PAT]
    if missing:
        print("  [WARN] 缺少三大报表：", "、".join(missing))
    for comp, path in XLSX_PAT.items():
        if not path.exists():
            print(f"  [WARN] {comp} 三大报表不存在：{path}")
            continue
        specs = [
            ("资产负债表", "合并资产负债表", BS_ITEMS, ()),
            ("利润表", "合并利润表", IS_ITEMS, ()),
            ("现金流量表", "合并现金流量表", CF_ITEMS, ()),
        ]
        for sheet_kind, sheet_name, items, _ in specs:
            periods, rows = read_sheet(path, sheet_name)
            periods_map[(comp, sheet_kind)] = periods
            for canon, search in items.items():
                excl = ()
                if canon == "营业收入":
                    excl = ("总",)
                elif canon == "负债合计":
                    excl = ("流动",)  # 排除“流动负债合计”“非流动负债合计”
                elif canon == "所有者权益合计":
                    excl = ("归属于母公司",)  # 排除“归属于母公司所有者权益合计”
                _, label, vals = pick_row(rows, search, exclude=excl)
                if label is None:
                    print(f"  [WARN] {comp} {sheet_name} 未找到 {search}")
                    continue
                for p, v in zip(periods, vals):
                    recs.append({
                        "公司": comp, "报表": sheet_kind, "期间": p,
                        "项目": canon, "原始项目": label, "金额": v,
                    })
        # 原始全量报表
        for sheet_kind, sheet_name in [("资产负债表", "合并资产负债表"),
                                       ("利润表", "合并利润表"),
                                       ("现金流量表", "合并现金流量表")]:
            periods, rows = read_sheet(path, sheet_name)
            for r, label, vals in rows:
                for p, v in zip(periods, vals):
                    recs.append({
                        "公司": comp, "报表": sheet_kind, "期间": p,
                        "项目": label, "原始项目": label, "金额": v,
                        "行号": r, "全量": 1,
                    })
    df = pd.DataFrame(recs)
    return df


def period_key(p: str):
    m = re.match(r"(\d{4})(Q1|H1|Q3|FY)", str(p))
    if not m:
        return (9999, 9)
    y, k = int(m.group(1)), m.group(2)
    return (y, {"Q1": 0, "H1": 1, "Q3": 2, "FY": 3}[k])


def period_year(p):
    return int(re.match(r"(\d{4})", p).group(1))


def period_q(p):
    return {"Q1": 1, "H1": 2, "Q3": 3, "FY": 4}[re.search(r"(Q1|H1|Q3|FY)", p).group(1)]


def shift_period(p: str, years: int):
    return f"{period_year(p) - years}{p[-2:]}" if p.endswith(("Q1", "H1", "Q3", "FY")) else None


# ---------------------------------------------------------------- 扣非 & 员工
TOKEN_RE = re.compile(r"(-?\d{1,3}(?:,\d{3})*(?:\.\d+)?%)|(-?\d{1,3}(?:,\d{3})*\.\d{2})")
HEADER_RE = re.compile(
    r"(?:[\u4e00-\u9fa5A-Za-z]{2,30})?20\d{2}年(?:年度|半年度|第[一二三四]季度)报告(?:全文)?\d*(?:/\d+)?"
)


def pdf_period(fname: str):
    m = re.search(r"(20\d{2})年(.+?)报告", fname)
    if not m:
        return None
    y, kind = m.group(1), m.group(2)
    if "半年度" in kind:
        return y + "H1"
    if "一季度" in kind or "第一季度" in kind:
        return y + "Q1"
    if "三季度" in kind or "第三季度" in kind:
        return y + "Q3"
    if "年度" in kind:
        return y + "FY"
    return None


def flat_text(path: Path, max_pages: int | None = None, stop_after: str | None = None) -> str:
    """抽取 PDF 文本。stop_after 命中后提前结束（主要会计数据在前几页）。"""
    doc = fitz.open(path)
    n = doc.page_count if max_pages is None else min(doc.page_count, max_pages)
    parts = []
    for i in range(n):
        parts.append(re.sub(r"\s+", "", doc[i].get_text()))
        if stop_after and stop_after in parts[-1]:
            break
    doc.close()
    return "".join(parts)


def _tokens(seg):
    out = []
    for t in TOKEN_RE.finditer(seg):
        if t.group(1):
            out.append((float(t.group(1).rstrip("%").replace(",", "")), True))
        else:
            out.append((float(t.group(2).replace(",", "")), False))
    return out


def row_values(text, start_pat, end_pat):
    m = re.search(start_pat, text)
    if not m:
        return None
    e = re.search(end_pat, text[m.end():])
    if not e:
        return None
    return _tokens(text[m.end():m.end() + e.start()])


def parse_kf_from_text(t: str, cum_rev: float):
    t = HEADER_RE.sub("", t)
    i = t.find("主要会计数据和财务指标")
    if i < 0:
        i = t.find("主要会计数据")
    if i < 0:
        i = t.find("主要财务数据")
    if i < 0:
        return None
    t = t[i:i + 8000]
    # 沪深两版报表“营业收入”后有的带“（元）”有的不带；扣非也有“的/后的”两种写法
    rev = row_values(t, r"营业(?:总)?收入(?:（元）)?", r"归属于(?:上市公司|母公司)股东")
    kf_end = r"加权平均净资产收益率|经营活动产生的现金|基本每股收益|稀释每股收益"
    kf = row_values(t, r"扣除非经常性损益(?:后)?的净利润(?:（元）)?", kf_end)
    if kf is None:
        kf = row_values(t, r"扣除非经常性损益[^，。]{0,20}?净利润", kf_end)
    if not kf:
        # 部分沪市季报把“扣非”行标签拆在页码/金额两侧（如“扣除-66,549,509.74…非经常性损益的净利润”），
        # 用宽松模式把夹在标签中间的数字也取出来
        for pat in (r"扣除非经常性(.{0,80}?)润",
                    r"扣除(.{0,80}?)非经常性损益.{0,20}?润"):
            m = re.search(pat, t)
            if not m or not re.search(r"\d", m.group(1)):
                continue
            vals = _tokens(m.group(1))
            e = re.search(kf_end, t[m.end():])
            tail = _tokens(t[m.end():m.end() + e.start()]) if e else []
            kf = vals + tail
            break
    if not rev or not kf:
        return None
    # 上交所版季报的增减幅度不带 %（如 ...432.34-12.864,316...），
    # 用金额量级过滤掉百分比：A 股报表金额以元计，营收/扣非均远大于 1 万元。
    revs = [v for v, pct in rev if not pct and abs(v) >= 1e4]
    kfs = [v for v, pct in kf if not pct and abs(v) >= 1e4]
    if cum_rev is None or np.isnan(cum_rev) or not revs or not kfs:
        return None
    rel = [abs(v - cum_rev) / max(abs(cum_rev), 1.0) for v in revs]
    idx = int(np.argmin(rel))
    if rel[idx] > 0.01:
        return None
    if idx >= len(kfs):
        return None
    return kfs[idx], revs[idx]


def parse_emp(t: str):
    t = HEADER_RE.sub("", t)
    m = re.search(r"教育程度.{0,12}?类别数量（人）", t)
    if not m:
        return None
    seg = t[m.end():]
    k = seg.find("合计")
    if k < 0:
        return None
    body = seg[:k]
    # 名称允许“/、”等分隔符（如“大专/本科”“高中、中专及技校”）
    pairs = re.findall(r"([\u4e00-\u9fa5A-Za-z（）()/、·]+?)(\d[\d,]*)", body)
    cats = []
    for n, v in pairs:
        # 剔除夹在表格中的页眉页脚（公司名/年年度报告/页码）
        if re.search(r"(公司|报告|年度|有限|股份)", n) or n.isdigit():
            continue
        if not n.strip("/、·"):
            continue
        cats.append((n, int(v.replace(",", ""))))
    if not cats:
        return None
    m2 = re.match(r"合计(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d{3,6}(?:\.\d+)?)", seg[k:])
    total = int(m2.group(1).replace(",", "")) if m2 else sum(v for _, v in cats)
    ssum = sum(v for _, v in cats)
    if ssum and total != ssum and str(total).startswith(str(ssum)):
        total = ssum  # “合计6882、薪酬政策”：后面的节号被吞进合计
    if total:
        fixed = []
        for n, v in cats:
            if v > total and re.fullmatch(r"\d{4,7}", str(v)) and re.search(r"20\d{2}$", str(v)):
                v = int(str(v)[:-4])  # 去掉误粘的“2022”年份页眉
            fixed.append((n, v))
        cats = fixed
    return cats, total


def high_edu(cats):
    names = {n: v for n, v in cats}
    has_sep_bachelor = any(n in ("本科", "大学本科", "大学") for n in names)
    high = 0
    detail = []
    for n, v in cats:
        is_high = False
        if any(k in n for k in ("博士", "硕士", "研究生")):
            if n == "研究生" and any(("硕士" in x or "博士" in x) for x in names):
                continue
            is_high = True
        elif n in ("本科", "大学本科", "大学"):
            is_high = True
        elif ("本科及以上" in n or "本科以上" in n):
            # 若同时单列“本科”，此类别实为研究生及以上；两种情况都计入本科及以上
            is_high = True
        elif ("本科" in n and "大专" in n):
            # “本科及大专/大专/本科”为合并档，无法拆分：只计研究生，明细中注明
            detail.append(f"[{n}{v}合并未拆分]")
            continue
        if is_high:
            high += v
            detail.append(f"{n}{v}")
    return high, "; ".join(detail), has_sep_bachelor


def load_pdf_data(df_all):
    """解析全部 PDF 的扣非，及年报员工学历（带缓存）。"""
    cache = DATA_DIR / "pdf_parsed_all.pkl"
    if cache.exists():
        kf, emp, pdf_status = pd.read_pickle(cache)
        # 源目录/公司范围变化时丢弃旧缓存
        if set(pdf_status.get("公司", [])) >= set(XLSX_PAT):
            return kf, emp, pdf_status
    cum_rev = {}
    sub = df_all[(df_all["报表"] == "利润表") & (df_all["项目"] == "营业收入") & (df_all.get("全量") != 1)]
    for _, r in sub.iterrows():
        cum_rev[(r["公司"], r["期间"])] = r["金额"]

    kf_rows, emp_rows, pdf_rows = [], [], []
    for comp in XLSX_PAT:
        pdf_dir = None
        for d in sorted(COMPANY_DIR[comp].glob("3.*")):
            if d.is_dir():
                pdf_dir = d
                break
        if pdf_dir is None:
            continue
        pdfs = sorted(p for p in pdf_dir.iterdir() if p.suffix.lower() == ".pdf" and not p.name.startswith("~$"))
        print(f"      PDF {comp}: {len(pdfs)} 份")
        for f in pdfs:
            p = pdf_period(f.name)
            if p is None:
                continue
            # 季报只读前 12 页（主要会计数据在前部）；年报读全文以取员工学历
            t = flat_text(f, max_pages=None if p.endswith("FY") else 20)
            res = parse_kf_from_text(t, cum_rev.get((comp, p), np.nan))
            ok = res is not None
            if ok:
                kf_rows.append({"公司": comp, "期间": p, "扣非归母净利润_累计": res[0],
                                "营业收入_报告值": res[1]})
            pdf_rows.append({"公司": comp, "期间": p, "文件": f.name,
                             "扣非解析": "OK" if ok else "FAIL"})
            # 年报员工
            if p and p.endswith("FY"):
                r = parse_emp(t)
                if r:
                    cats, total = r
                    high, detail, _ = high_edu(cats)
                    emp_rows.append({
                        "公司": comp, "年度": period_year(p), "员工总数": total,
                        "本科及以上人数": high, "高学历密度": high / total if total else np.nan,
                        "学历结构": "; ".join(f"{n}:{v}" for n, v in cats),
                        "学历明细": detail,
                    })
    kf = pd.DataFrame(kf_rows)
    emp = pd.DataFrame(emp_rows).sort_values(["公司", "年度"])
    pdf_status = pd.DataFrame(pdf_rows)
    pd.to_pickle((kf, emp, pdf_status), cache)
    return kf, emp, pdf_status


# ---------------------------------------------------------------- 计算
def wide_from_long(df_all, report: str, items, full=False):
    d = df_all[(df_all["报表"] == report) & (df_all["项目"].isin(items))]
    if not full:
        d = d[d.get("全量").isna()] if "全量" in d else d
    w = d.pivot_table(index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")
    return w


def add_single_quarter(wide: pd.DataFrame, cols) -> pd.DataFrame:
    """把年初至今累计数差分出单季数。"""
    out = {}
    for comp, g in wide.groupby(level=0):
        d = g.droplevel(0)
        for col in cols:
            if col not in d.columns:
                continue
            s = d[col]
            for year in sorted({int(p[:4]) for p in d.index}):
                vals = {}
                for q, p in [(1, f"{year}Q1"), (2, f"{year}H1"), (3, f"{year}Q3"), (4, f"{year}FY")]:
                    if p not in s.index:
                        vals[q] = np.nan
                        continue
                    if q == 1:
                        vals[q] = s.loc[p]
                    else:
                        prev = {2: f"{year}Q1", 3: f"{year}H1", 4: f"{year}Q3"}[q]
                        vals[q] = s.loc[p] - s.loc[prev] if prev in s.index else np.nan
                for q, v in vals.items():
                    p = {1: f"{year}Q1", 2: f"{year}H1", 3: f"{year}Q3", 4: f"{year}FY"}[q]
                    out[(comp, p, col)] = v
    ser = pd.Series(out).rename("value")
    df = ser.reset_index()
    df.columns = ["公司", "期间", "项目", "value"]
    q = df.pivot_table(index=["公司", "期间"], columns="项目", values="value", aggfunc="first")
    return q


def safe_div(a, b):
    a = pd.to_numeric(a, errors="coerce")
    b = pd.to_numeric(b, errors="coerce")
    r = a / b.replace(0, np.nan)
    return r.replace([np.inf, -np.inf], np.nan)


def dict_to_series(d, index):
    """把 (公司,期间)->值 的字典按 MultiIndex 顺序对齐成 Series。"""
    return pd.Series([d.get(k, np.nan) for k in index], index=index, dtype=float)


def sort_group(g):
    keys = [period_key(p)[0] * 4 + period_key(p)[1] for p in g.index.get_level_values(1)]
    order = np.argsort(np.asarray(keys, dtype=int))
    return g.iloc[order]


def repair_single_q_is(df_all):
    """修正源 xlsx 的口径/符号错误：单季值还原为累计、利润行符号翻转。"""
    df_all = df_all.copy()
    canon_flip = ("营业利润", "归母净利润")  # 规范科目行（全量!=1）用简称
    for (comp, pstr), labels in SIGN_FLIP.items():
        m = ((df_all["公司"] == comp) & (df_all["报表"] == "利润表")
             & (df_all["期间"] == pstr)
             & (df_all["项目"].isin(labels) | df_all["项目"].isin(canon_flip)))
        df_all.loc[m, "金额"] = -df_all.loc[m, "金额"]
    for comp, pstr in SINGLE_Q_IS:
        h1 = f"{pstr[:4]}H1"
        base = df_all[(df_all["公司"] == comp) & (df_all["报表"] == "利润表") & (df_all["期间"] == h1)]
        h1map = {k: (0.0 if pd.isna(v) else v) for k, v in zip(base["项目"], base["金额"])}
        m = (df_all["公司"] == comp) & (df_all["报表"] == "利润表") & (df_all["期间"] == pstr)
        df_all.loc[m, "金额"] = df_all.loc[m, "金额"] + df_all.loc[m, "项目"].map(h1map)
    return df_all


def build():
    print("[1/5] 读取三大报表 ...")
    df_all = repair_single_q_is(load_statements())

    # 只取规范科目（不含全量重复）
    bones = df_all[df_all.get("全量").isna()].copy() if "全量" in df_all else df_all.copy()

    bs = bones[bones["报表"] == "资产负债表"].pivot_table(
        index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")
    is_cum = bones[bones["报表"] == "利润表"].pivot_table(
        index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")
    cf_cum = bones[bones["报表"] == "现金流量表"].pivot_table(
        index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")

    # 补齐列
    for c in BS_ITEMS:
        if c not in bs: bs[c] = np.nan
    for c in IS_ITEMS:
        if c not in is_cum: is_cum[c] = np.nan
    for c in CF_ITEMS:
        if c not in cf_cum: cf_cum[c] = np.nan

    print("[2/5] 解析 PDF 扣非 / 员工学历 ...")
    kf, emp, pdf_status = load_pdf_data(df_all)
    print(f"      扣非解析 {len(kf)} 条；员工年度 {len(emp)} 条")

    # 期间排序
    all_idx = sorted(set(bs.index), key=lambda x: (COMPANY_ORDER.index(x[0]), period_key(x[1])))
    mi = pd.MultiIndex.from_tuples(all_idx, names=["公司", "期间"])
    bs = bs.reindex(mi)
    is_cum = is_cum.reindex(mi)
    cf_cum = cf_cum.reindex(mi)

    print("[3/5] 差分单季、计算指标 ...")
    is_q = add_single_quarter(is_cum, list(IS_ITEMS))
    cf_q = add_single_quarter(cf_cum, list(CF_ITEMS))
    is_q = is_q.reindex(mi)
    cf_q = cf_q.reindex(mi)

    # 扣非并入
    kf_s = kf.set_index(["公司", "期间"])["扣非归母净利润_累计"] if len(kf) else pd.Series(dtype=float)
    is_cum["扣非归母净利润"] = kf_s.reindex(is_cum.index)
    # 扣非单季
    kf_q = add_single_quarter(is_cum[["扣非归母净利润"]], ["扣非归母净利润"]).reindex(mi)
    is_q["扣非归母净利润"] = kf_q["扣非归母净利润"]

    # 主面板
    p = pd.DataFrame(index=mi)
    for c in BS_ITEMS:
        p[f"bs_{c}"] = bs[c]
    for c in IS_ITEMS:
        p[f"is_{c}_cum"] = is_cum[c]
        p[f"is_{c}_q"] = is_q[c]
    for c in CF_ITEMS:
        p[f"cf_{c}_cum"] = cf_cum[c]
        p[f"cf_{c}_q"] = cf_q[c]
    p["kf_cum"] = is_cum["扣非归母净利润"]
    p["kf_q"] = is_q["扣非归母净利润"]

    # 期间属性
    p["年"] = [period_year(i[1]) for i in p.index]
    p["季度"] = [period_q(i[1]) for i in p.index]
    p["n_q"] = p["季度"]
    p["期间序号"] = [period_key(i[1])[0] * 4 + period_key(i[1])[1] for i in p.index]

    # 有息负债
    p["有息负债"] = (p["bs_短期借款"].fillna(0) + p["bs_一年内到期的非流动负债"].fillna(0)
                  + p["bs_长期借款"].fillna(0) + p["bs_应付债券"].fillna(0)
                  + p["bs_租赁负债"].fillna(0))

    # 平均资产/权益（单季：上一季；累计：上年末）
    def prev_period(pstr):
        y, q = period_year(pstr), pstr[-2:]
        if q == "Q1":
            return f"{y-1}FY"
        return {  # 上一季度
            "H1": f"{y}Q1", "Q3": f"{y}H1", "FY": f"{y}Q3"
        }[q]

    ta = p["bs_资产总计"]
    eq = p["bs_归属于母公司所有者权益合计"]
    avg_ta_q, avg_eq_q, avg_ta_ytd, avg_eq_ytd = {}, {}, {}, {}
    for comp, g in p.groupby(level=0):
        for pstr in g.index.get_level_values(1):
            prev = (comp, prev_period(pstr))
            cur = (comp, pstr)
            y = period_year(pstr)
            yend = (comp, f"{y-1}FY")
            # 2020 年缺 2019 年报，用 2020Q1 作为期初替代
            if yend not in p.index:
                fb = (comp, f"{y}Q1")
                if fb in p.index and fb != cur:
                    yend = fb
            avg_ta_q[cur] = np.nanmean([ta.get(prev, np.nan), ta.get(cur, np.nan)])
            avg_eq_q[cur] = np.nanmean([eq.get(prev, np.nan), eq.get(cur, np.nan)])
            avg_ta_ytd[cur] = np.nanmean([ta.get(yend, np.nan), ta.get(cur, np.nan)])
            avg_eq_ytd[cur] = np.nanmean([eq.get(yend, np.nan), eq.get(cur, np.nan)])
    p["avg_ta_q"] = dict_to_series(avg_ta_q, p.index)
    p["avg_eq_q"] = dict_to_series(avg_eq_q, p.index)
    p["avg_ta_ytd"] = dict_to_series(avg_ta_ytd, p.index)
    p["avg_eq_ytd"] = dict_to_series(avg_eq_ytd, p.index)

    # 员工映射：Q1-Q3 用上年末年报，Q4 用本年末年报
    emp_map = {(r["公司"], int(r["年度"])): r for _, r in emp.iterrows()}
    emp_n, high_n, high_d = [], [], []
    emp_years = {}
    for (c, y) in emp_map:
        emp_years.setdefault(c, []).append(y)
    for comp, pstr in p.index:
        y, q = period_year(pstr), period_q(pstr)
        fy = y if q == 4 else y - 1
        if fy < 2020:
            fy = 2020
        r = emp_map.get((comp, fy))
        # 2021 年上市的公司（神农/东瑞）缺上市前年报：Q1–Q3 退回最早一期年报
        ys = sorted(emp_years.get(comp, []))
        if r is None and ys and fy < ys[0]:
            r = emp_map.get((comp, ys[0]))
        emp_n.append(r["员工总数"] if r is not None else np.nan)
        high_n.append(r["本科及以上人数"] if r is not None else np.nan)
        high_d.append(r["高学历密度"] if r is not None else np.nan)
    p["员工总数"] = emp_n
    p["本科及以上人数"] = high_n
    p["高学历人数"] = high_n
    p["高学历人才密度"] = high_d

    # TTM 营收
    print("      DEBUG p.shape", p.shape, "dupcols", int(p.columns.duplicated().sum()))
    print("      DEBUG bad cols", [c for c in p.columns if p[c].dtype == object][:10])
    rev_q = p["is_营业收入_q"]
    ttm = {}
    for comp, g in p.groupby(level=0):
        g = sort_group(g)
        s = g["is_营业收入_q"]
        roll = s.rolling(4, min_periods=4).sum()
        for (c, pstr), v in roll.items():
            ttm[(c, pstr)] = v
    p["ttm_rev"] = dict_to_series(ttm, p.index)

    # 累计经营现金流（自2020Q1）
    cum_ocf = {}
    for comp, g in p.groupby(level=0):
        g = sort_group(g)
        s = g["cf_经营活动现金流量净额_q"].fillna(0).cumsum()
        for (c, pstr), v in s.items():
            cum_ocf[(c, pstr)] = v
    p["cum_ocf"] = dict_to_series(cum_ocf, p.index)

    # 区间平均总资产（自2020Q1起均值）
    avg_all_ta = {}
    for comp, g in p.groupby(level=0):
        g = sort_group(g)
        s = g["bs_资产总计"].expanding().mean()
        for (c, pstr), v in s.items():
            avg_all_ta[(c, pstr)] = v
    p["avg_all_ta"] = dict_to_series(avg_all_ta, p.index)

    # 3年复合增长率
    def cagr_series(col):
        out = {}
        for comp, g in p.groupby(level=0):
            g = sort_group(g)
            for (c, pstr), v in g[col].items():
                base_p = shift_period(pstr, 3)
                b = g[col].loc[(c, base_p)] if (c, base_p) in g.index else np.nan
                if pd.notna(v) and pd.notna(b) and b > 0:
                    out[(c, pstr)] = (v / b) ** (1 / 3) - 1
                else:
                    out[(c, pstr)] = np.nan
        return dict_to_series(out, p.index)
    p["资产复合增长率"] = cagr_series("bs_资产总计")
    p["营收复合增长率"] = cagr_series("ttm_rev")

    # ---------------- 指标计算
    def calc(period_type):
        """period_type: 'q' 单季 or 'cum' 累计"""
        sfx = "_q" if period_type == "q" else "_cum"
        ann = 4 if period_type == "q" else 4 / p["n_q"]
        if period_type == "q":
            avg_ta, avg_eq = p["avg_ta_q"], p["avg_eq_q"]
        else:
            avg_ta, avg_eq = p["avg_ta_ytd"], p["avg_eq_ytd"]

        rev = p[f"is_营业收入{sfx}"]
        cost = p[f"is_营业成本{sfx}"]
        tax = p[f"is_税金及附加{sfx}"].fillna(0)
        sell = p[f"is_销售费用{sfx}"].fillna(0)
        admin = p[f"is_管理费用{sfx}"].fillna(0)
        rd = p[f"is_研发费用{sfx}"].fillna(0)
        fin = p[f"is_财务费用{sfx}"].fillna(0)
        interest = p[f"is_利息费用{sfx}"]
        itax = p[f"is_所得税费用{sfx}"]
        ocf = p[f"cf_经营活动现金流量净额{sfx}"]
        capex = p[f"cf_购建长期资产支付的现金{sfx}"]
        cash_sales = p[f"cf_销售商品收到的现金{sfx}"]
        cfo_in = p[f"cf_经营活动现金流入小计{sfx}"]
        cfi_in = p[f"cf_投资活动现金流入小计{sfx}"].fillna(0)
        cff_in = p[f"cf_筹资活动现金流入小计{sfx}"].fillna(0)
        kf = p[f"kf_{period_type}"]

        core = rev - cost - tax - sell - admin - rd - fin
        d = pd.DataFrame(index=p.index)
        d["毛利率"] = safe_div(rev - cost, rev)
        d["期间费用率"] = safe_div(sell + admin + rd + fin, rev)
        d["核心利润率"] = safe_div(core, rev)
        d["归母扣非ROE"] = safe_div(kf * ann, avg_eq.where(avg_eq > 0))
        d["有息负债率"] = safe_div(p["有息负债"], p["bs_资产总计"])
        d["现金利息保障倍数"] = safe_div(ocf + interest + itax, interest.where(interest > 0))
        d["现金债务覆盖率"] = safe_div(ocf * ann, p["有息负债"].where(p["有息负债"] > 0))
        d["现金流量比率"] = safe_div(ocf * ann, p["bs_流动负债合计"].where(p["bs_流动负债合计"] > 0))
        d["重资产占比"] = safe_div(p["bs_投资性房地产"].fillna(0) + p["bs_固定资产"].fillna(0)
                                + p["bs_在建工程"].fillna(0) + p["bs_生产性生物资产"].fillna(0),
                                p["bs_资产总计"])
        d["总资产周转率"] = safe_div(rev * ann, avg_ta.where(avg_ta > 0))
        d["人均创收"] = safe_div(rev * ann, p["员工总数"]) / 1e4  # 万元/人
        d["人均核心利润"] = safe_div(core * ann, p["员工总数"]) / 1e4
        d["营收现金比率"] = safe_div(cash_sales, rev)
        d["自由现金流利润率"] = safe_div(ocf - capex, rev)
        d["资本开支覆盖倍数"] = safe_div(ocf, capex.where(capex > 0))
        d["经营现金流入占比"] = safe_div(cfo_in, cfo_in + cfi_in + cff_in)
        d["资产复合增长率"] = p["资产复合增长率"]
        d["营收复合增长率"] = p["营收复合增长率"]
        d["累计资产现金回报率"] = safe_div(p["cum_ocf"], p["avg_all_ta"].where(p["avg_all_ta"] > 0))
        d["高学历人才密度"] = p["高学历人才密度"]

        # 行业合计用“合并分子 / 合并分母”，不能把各公司的比率直接相加。
        # 复合增长率是水平值的增长率，单独按行业汇总后的总资产/TTM 营收计算。
        parts = {
            "毛利率": (rev - cost, rev),
            "期间费用率": (sell + admin + rd + fin, rev),
            "核心利润率": (core, rev),
            "归母扣非ROE": (kf * ann, avg_eq.where(avg_eq > 0)),
            "有息负债率": (p["有息负债"], p["bs_资产总计"]),
            "现金利息保障倍数": (ocf + interest + itax, interest.where(interest > 0)),
            "现金债务覆盖率": (ocf * ann, p["有息负债"].where(p["有息负债"] > 0)),
            "现金流量比率": (ocf * ann, p["bs_流动负债合计"].where(p["bs_流动负债合计"] > 0)),
            "重资产占比": (p["bs_投资性房地产"].fillna(0) + p["bs_固定资产"].fillna(0)
                         + p["bs_在建工程"].fillna(0) + p["bs_生产性生物资产"].fillna(0),
                         p["bs_资产总计"]),
            "总资产周转率": (rev * ann, avg_ta.where(avg_ta > 0)),
            "人均创收": (rev * ann / 1e4, p["员工总数"]),
            "人均核心利润": (core * ann / 1e4, p["员工总数"]),
            "营收现金比率": (cash_sales, rev),
            "自由现金流利润率": (ocf - capex, rev),
            "资本开支覆盖倍数": (ocf, capex.where(capex > 0)),
            "经营现金流入占比": (cfo_in, cfo_in + cfi_in + cff_in),
            "累计资产现金回报率": (p["cum_ocf"], p["avg_all_ta"].where(p["avg_all_ta"] > 0)),
            "高学历人才密度": (p["本科及以上人数"], p["员工总数"]),
        }
        return d, parts

    def industry_aggregate(parts):
        """按报告期汇总各公司分子/分母，得到行业口径指标与样本家数。"""
        periods = sorted(p.index.get_level_values(1).unique(), key=period_key)
        out = pd.DataFrame(index=periods)
        out.index.name = "期间"
        ns = pd.DataFrame(index=periods)
        ns.index.name = "期间"
        for name, (num, den) in parts.items():
            den_ok = den.where(den > 0)
            mask = num.notna() & den_ok.notna()
            nsum = num.where(mask).groupby(level=1).sum(min_count=1)
            dsum = den_ok.where(mask).groupby(level=1).sum(min_count=1)
            out[name] = (nsum / dsum).reindex(periods)
            ns[name] = mask.groupby(level=1).sum().reindex(periods).fillna(0).astype(int)

        def agg_cagr(series):
            # 只纳入“当前期与 3 年前同期都有值”的公司，避免样本变化造成增长率跳变。
            res, cnt = {}, {}
            for per in periods:
                base = shift_period(per, 3)
                try:
                    cur = series.xs(per, level=1)
                    old = series.xs(base, level=1)
                except KeyError:
                    res[per], cnt[per] = np.nan, 0
                    continue
                common = cur.index.intersection(old.index)
                cur, old = cur.reindex(common), old.reindex(common)
                ok = cur.notna() & old.notna() & (old > 0)
                if ok.any():
                    res[per] = (cur[ok].sum() / old[ok].sum()) ** (1 / 3) - 1
                    cnt[per] = int(ok.sum())
                else:
                    res[per], cnt[per] = np.nan, 0
            return (pd.Series(res).reindex(periods),
                    pd.Series(cnt).reindex(periods).fillna(0).astype(int))

        out["资产复合增长率"], ns["资产复合增长率"] = agg_cagr(p["bs_资产总计"])
        out["营收复合增长率"], ns["营收复合增长率"] = agg_cagr(p["ttm_rev"])
        return out, ns

    ind_q, parts_q = calc("q")
    ind_cum, parts_cum = calc("cum")
    ind_industry_q, ind_industry_n_q = industry_aggregate(parts_q)
    ind_industry_cum, ind_industry_n_cum = industry_aggregate(parts_cum)
    ind_industry_q = ind_industry_q[list(ind_q.columns)]
    ind_industry_cum = ind_industry_cum[list(ind_cum.columns)]
    ind_industry_n_q = ind_industry_n_q[list(ind_q.columns)]
    ind_industry_n_cum = ind_industry_n_cum[list(ind_cum.columns)]

    # ---------------- 五项资产
    assets = pd.DataFrame(index=p.index)
    for c in ["投资性房地产", "固定资产", "在建工程", "生产性生物资产", "存货"]:
        assets[c] = p[f"bs_{c}"]
    assets["五项合计"] = assets[["投资性房地产", "固定资产", "在建工程", "生产性生物资产", "存货"]].fillna(0).sum(axis=1)
    assets["资产总计"] = p["bs_资产总计"]
    assets["五项占总资产"] = safe_div(assets["五项合计"], assets["资产总计"])
    assets["固定资产占总资产"] = safe_div(assets["固定资产"], assets["资产总计"])
    assets["存货占总资产"] = safe_div(assets["存货"], assets["资产总计"])
    assets["在建工程/固定资产"] = safe_div(assets["在建工程"], assets["固定资产"])
    # 同比（同季度上年）
    yoy = {}
    for comp, g in assets.groupby(level=0):
        for (c, pstr), row in g.iterrows():
            base = (c, shift_period(pstr, 1))
            if base in assets.index and pd.notna(assets.loc[base, "五项合计"]) and assets.loc[base, "五项合计"] != 0:
                yoy[(c, pstr)] = row["五项合计"] / assets.loc[base, "五项合计"] - 1
            else:
                yoy[(c, pstr)] = np.nan
    assets["五项合计同比"] = dict_to_series(yoy, assets.index)
    assets = assets.reset_index()

    # ---------------- 长表
    def melt(ind, label):
        x = ind.copy()
        x["公司全称"] = [COMPANY_FULL[c] for c, _ in x.index]
        x = x.reset_index()
        x["口径"] = label
        long = x.melt(id_vars=["公司", "公司全称", "期间", "口径"],
                      var_name="指标", value_name="数值")
        return long

    long = pd.concat([melt(ind_q, "单季"), melt(ind_cum, "累计")], ignore_index=True)

    def melt_industry(df, ndf, label):
        d = df.reset_index().melt(id_vars="期间", var_name="指标", value_name="数值")
        n = ndf.reset_index().melt(id_vars="期间", var_name="指标", value_name="样本数")
        m = d.merge(n, on=["期间", "指标"], how="left")
        m["口径"] = label
        return m[["期间", "口径", "指标", "数值", "样本数"]]

    industry_long = pd.concat([
        melt_industry(ind_industry_q, ind_industry_n_q, "单季"),
        melt_industry(ind_industry_cum, ind_industry_n_cum, "累计"),
    ], ignore_index=True)
    ind_order = {name: i for i, name in enumerate(ind_q.columns)}
    industry_long["_p"] = industry_long["期间"].map(period_key)
    industry_long["_i"] = industry_long["指标"].map(ind_order)
    industry_long = (industry_long.sort_values(["_p", "_i"])
                      .drop(columns=["_p", "_i"]).reset_index(drop=True))

    # 排序
    def sort_df(df):
        df = df.copy()
        df["_c"] = df["公司"].map(lambda x: COMPANY_ORDER.index(x))
        df["_p"] = df["期间"].map(period_key)
        df = df.sort_values(["_c", "_p"]).drop(columns=["_c", "_p"])
        return df
    ind_q_out = sort_df(ind_q.reset_index())
    ind_cum_out = sort_df(ind_cum.reset_index())
    assets = sort_df(assets)
    long = sort_df(long)

    # 中文列顺序
    IND_ORDER = [
        "毛利率", "期间费用率", "核心利润率", "归母扣非ROE",
        "有息负债率", "现金利息保障倍数", "现金债务覆盖率", "现金流量比率",
        "重资产占比", "总资产周转率", "人均创收", "人均核心利润",
        "营收现金比率", "自由现金流利润率", "资本开支覆盖倍数", "经营现金流入占比",
        "资产复合增长率", "营收复合增长率", "累计资产现金回报率", "高学历人才密度",
    ]
    ind_q_out = ind_q_out[["公司", "期间"] + IND_ORDER]
    ind_cum_out = ind_cum_out[["公司", "期间"] + IND_ORDER]

    # 保存
    long.to_csv(DATA_DIR / "panel_long.csv", index=False, encoding="utf-8-sig")
    ind_q_out.to_csv(DATA_DIR / "indicators_single.csv", index=False, encoding="utf-8-sig")
    ind_cum_out.to_csv(DATA_DIR / "indicators_cum.csv", index=False, encoding="utf-8-sig")
    industry_long.to_csv(DATA_DIR / "industry_long.csv", index=False, encoding="utf-8-sig")
    assets.to_csv(DATA_DIR / "assets_five.csv", index=False, encoding="utf-8-sig")
    kf.to_csv(DATA_DIR / "kf_nonrecurring.csv", index=False, encoding="utf-8-sig")
    emp.to_csv(DATA_DIR / "employees.csv", index=False, encoding="utf-8-sig")
    pdf_status.to_csv(DATA_DIR / "pdf_status.csv", index=False, encoding="utf-8-sig")

    print("[4/5] 校验 ...")
    checks = build_checks(bones, kf, emp, pdf_status)

    print("[5/5] 写 Excel 底表 ...")
    write_excel(df_all, ind_q_out, ind_cum_out, assets, kf, emp, checks, pdf_status)
    print("完成 ->", OUT_DIR / "上市公司财务指标_底表.xlsx")

    return dict(long=long, ind_q=ind_q_out, ind_cum=ind_cum_out, assets=assets,
                kf=kf, emp=emp, checks=checks, pdf_status=pdf_status, df_all=df_all)


def build_checks(bones, kf, emp, pdf_status):
    rows = []
    # 资产负债表平衡
    bs = bones[bones["报表"] == "资产负债表"].pivot_table(
        index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")
    if all(c in bs for c in ["资产总计", "负债合计", "所有者权益合计"]):
        diff = bs["资产总计"] - bs["负债合计"] - bs["所有者权益合计"]
        rows.append({"校验项": "资产负债表平衡（资产=负债+权益）", "样本数": int(diff.notna().sum()),
                     "最大绝对差异(元)": float(diff.abs().max()),
                     "通过样本": int((diff.abs() < 1).sum()), "不通过样本": int((diff.abs() >= 1).sum())})
    # 经营现金流小计
    cf = bones[bones["报表"] == "现金流量表"].pivot_table(
        index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")
    if "经营活动现金流入小计" in cf and "经营活动现金流量净额" in cf:
        ok = int(cf["经营活动现金流入小计"].notna().sum())
        rows.append({"校验项": "经营现金流入小计覆盖率", "样本数": len(cf),
                     "最大绝对差异(元)": np.nan, "通过样本": ok, "不通过样本": len(cf) - ok})
    rows.append({"校验项": "PDF 扣非解析成功数", "样本数": len(pdf_status),
                 "最大绝对差异(元)": np.nan, "通过样本": int((pdf_status["扣非解析"] == "OK").sum()),
                 "不通过样本": int((pdf_status["扣非解析"] != "OK").sum())})
    rows.append({"校验项": "年报员工学历解析数", "样本数": len(emp),
                 "最大绝对差异(元)": np.nan, "通过样本": len(emp), "不通过样本": 0})
    return pd.DataFrame(rows)


# ---------------------------------------------------------------- Excel
CAT_MAP = {
    "毛利率": "盈利能力", "期间费用率": "盈利能力", "核心利润率": "盈利能力", "归母扣非ROE": "盈利能力",
    "有息负债率": "财务健康", "现金利息保障倍数": "财务健康", "现金债务覆盖率": "财务健康", "现金流量比率": "财务健康",
    "重资产占比": "运营效率", "总资产周转率": "运营效率", "人均创收": "运营效率", "人均核心利润": "运营效率",
    "营收现金比率": "现金韧性", "自由现金流利润率": "现金韧性", "资本开支覆盖倍数": "现金韧性", "经营现金流入占比": "现金韧性",
    "资产复合增长率": "发展质量", "营收复合增长率": "发展质量", "累计资产现金回报率": "发展质量", "高学历人才密度": "发展质量",
}
PCT_IND = {"毛利率", "期间费用率", "核心利润率", "归母扣非ROE", "有息负债率", "重资产占比",
           "营收现金比率", "自由现金流利润率", "经营现金流入占比", "资产复合增长率",
           "营收复合增长率", "累计资产现金回报率", "高学历人才密度"}
MULT_IND = {"现金利息保障倍数", "现金债务覆盖率", "现金流量比率", "资本开支覆盖倍数",
            "总资产周转率", "人均创收", "人均核心利润"}
UNIT_IND = {**{k: "万元/人" for k in ["人均创收", "人均核心利润"]}}

FORMULAS = [
    ("盈利能力", "毛利率", "(营业收入−营业成本)/营业收入", "只扣营业成本，不含税金及附加和期间费用；(营业总收入−营业总成本)/营业总收入 是核心利润率，两者不可混用。"),
    ("盈利能力", "期间费用率", "(销售费用+管理费用+研发费用+财务费用)/营业收入", "四费合计占收入比重；研发费用单列后计入。"),
    ("盈利能力", "核心利润率", "(营业收入−营业成本−税金及附加−销售费用−管理费用−研发费用−财务费用)/营业收入", "剔除投资收益、公允价值变动、政府补助等非经营性项目后的主业盈利能力（张新民口径）。"),
    ("盈利能力", "归母扣非ROE", "扣非归母净利润/平均归母净资产×100%", "单季×4、累计×4/n年化；平均净资产=(期初+期末)/2，期初单季取上季末、累计取上年末。"),
    ("财务健康", "有息负债率", "有息负债/总资产×100%", "有息负债=短期借款+一年内到期的非流动负债+长期借款+应付债券+租赁负债。"),
    ("财务健康", "现金利息保障倍数", "(经营活动现金流量净额+利息费用+所得税费用)/利息费用", "分母用利润表“其中：利息费用”，分子所得税为权责发生制近似。"),
    ("财务健康", "现金债务覆盖率", "经营活动现金流量净额(年化)/有息负债", "单季×4、累计×4/n；衡量经营现金对全部有息债务的覆盖。"),
    ("财务健康", "现金流量比率", "经营活动现金流量净额(年化)/流动负债", "单季×4、累计×4/n；短期偿债的现金保障。"),
    ("运营效率", "重资产占比", "(投资性房地产+固定资产+在建工程+生产性生物资产)/总资产×100%", "猪企生产性生物资产视作核心生产资产纳入。"),
    ("运营效率", "总资产周转率", "营业收入(年化)/平均总资产", "单季×4、累计×4/n；平均总资产=(期初+期末)/2。"),
    ("运营效率", "人均创收", "营业收入(年化)/员工总数", "员工数为最近一期已披露年报口径：Q1–Q3用上年末，Q4用本年末。"),
    ("运营效率", "人均核心利润", "核心利润(年化)/员工总数", "员工数口径同上。"),
    ("现金韧性", "营收现金比率", "销售商品、提供劳务收到的现金/营业收入", "收入转化为现金的程度，>1通常含税/预收因素。"),
    ("现金韧性", "自由现金流利润率", "(经营活动现金流量净额−购建长期资产支付的现金)/营业收入", "资本开支后剩余现金占收入比重。"),
    ("现金韧性", "资本开支覆盖倍数", "经营活动现金流量净额/购建长期资产支付的现金", "经营现金覆盖资本开支的倍数，<1说明需外部融资。"),
    ("现金韧性", "经营现金流入占比", "经营活动现金流入小计/(经营+投资+筹资活动现金流入小计)", "现金流入结构，衡量造血对总流入的贡献。"),
    ("发展质量", "资产复合增长率", "(期末总资产/3年前同期总资产)^(1/3)−1", "滚动3年复合增速；不足3年历史留空。"),
    ("发展质量", "营收复合增长率", "(TTM营业收入/3年前同期TTM营业收入)^(1/3)−1", "TTM=最近4个单季合计，剔除季节性；不足3年历史留空。"),
    ("发展质量", "累计资产现金回报率", "首个报告期至本期末累计经营现金流净额/区间平均总资产", "区间平均总资产=各季末总资产算术平均；2020 年上市的公司从 2020Q1 起算，神农/东瑞从 2021H1 起算。"),
    ("发展质量", "高学历人才密度", "本科及以上学历员工数/员工总数×100%", "本科及以上含博士、硕士/研究生、本科（含“大学”）；年报“教育程度”表口径。若“本科及大专”合并披露无法拆分，只计研究生及以上，为下限。"),
]


def write_sheet(wb, name, df, title=None, pct_cols=(), mult_cols=(), amount_cols=(), width_map=None):
    ws = wb.create_sheet(name)
    start = 0
    if title:
        ws.cell(1, 1, title).font = openpyxl.styles.Font(bold=True, size=14, color="FFFFFF")
        ws.cell(1, 1).fill = openpyxl.styles.PatternFill("solid", fgColor="1F4E79")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, len(df.columns)))
        ws.row_dimensions[1].height = 26
        start = 2
    header_row = start + 1
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(header_row, j, col)
        c.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
        c.fill = openpyxl.styles.PatternFill("solid", fgColor="2F5597")
        c.alignment = openpyxl.styles.Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = openpyxl.styles.Border(bottom=openpyxl.styles.Side("thin", color="FFFFFF"))
    for i, row in enumerate(df.itertuples(index=False), start=header_row + 1):
        for j, v in enumerate(row, start=1):
            if isinstance(v, (np.floating, float)) and (pd.isna(v) or np.isinf(v)):
                v = None
            c = ws.cell(i, j, v)
            col = df.columns[j - 1]
            if col in pct_cols:
                c.number_format = "0.00%"
            elif col in mult_cols:
                c.number_format = "0.00"
            elif col in amount_cols:
                c.number_format = "#,##0.00"
    ws.freeze_panes = ws.cell(header_row + 1, 3)
    ws.auto_filter.ref = f"A{header_row}:{openpyxl.utils.get_column_letter(len(df.columns))}{header_row + len(df)}"
    # 列宽
    for j, col in enumerate(df.columns, start=1):
        w = (width_map or {}).get(col)
        if w is None:
            w = min(34, max(10, len(str(col)) * 1.8 + 4))
        ws.column_dimensions[openpyxl.utils.get_column_letter(j)].width = w
    return ws


def write_excel(df_all, ind_q, ind_cum, assets, kf, emp, checks, pdf_status):
    out = OUT_DIR / "上市公司财务指标_底表.xlsx"
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # 00 说明
    comps = list(dict.fromkeys(ind_q["公司"]))
    _periods = sorted(ind_q["期间"].unique(), key=period_key)
    pmin, pmax = _periods[0], _periods[-1]
    n_pdf = len(pdf_status)
    ws = wb.create_sheet("00_说明")
    ws["A1"] = f"上市猪企财务指标底表（{len(comps)} 家 · {pmin}–{pmax}）"
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=16, color="1F4E79")
    lines = [
        ("数据范围", f"{len(comps)} 家上市猪企（{'、'.join(comps)}）；期间 {pmin}–{pmax}。"
                     "其中神农集团、东瑞股份 2021 年上市，自 2021H1 起；新希望已补充 2026Q3。"),
        ("数据来源", f"三大报表 xlsx（由年报/季报 PDF 解析汇总）+ {n_pdf} 份原始公告 PDF（扣非净利润、员工学历）。"),
        ("报表口径", "合并报表；资产负债表为期末时点；利润表、现金流量表原始为年初至今累计，本表另做单季差分。"),
        ("金额单位", "原始金额为人民币元；指标中人均创收/人均核心利润为万元/人。"),
        ("指标口径", "详见“09_口径说明”。"),
        ("重要提示", "季度数据未经审计；2020Q1温氏营业收入、扣非采用原季报值（与后续年报追溯值相差约0.03%）。"),
        ("重要提示", "巨星 2020Q3：源 xlsx 利润表该列为单季值，已按“H1 累计 + 单季”还原为年初至今累计"
                     "（PDF 三季报 9M 营业收入 811,412,779.92 元可验）；资产负债表、现金流量表原即为累计/时点口径。"),
        ("重要提示", "新五丰 2023FY：源 xlsx 营业利润/利润总额/净利润/归母净利润/综合收益等利润行为正数，"
                     "与 2023 年报原文（归母净利润 -1,203,418,006.16 元）符号相反，已按年报翻转；营业收入、成本、费用未动。"),
        ("重要提示", "人均指标与高学历密度使用最近一期年报员工数（Q1–Q3用上年末，Q4用本年末）；"
                     "2020Q1–Q3 及 2021 年新上市公司缺上市前年报时，暂用最早一期年报。"),
        ("重要提示", "归母扣非ROE在平均归母净资产≤0时留空（如天邦、傲农、金新农部分期间）。"),
        ("重要提示", "高学历人才密度=本科及以上/员工总数；立华部分年度“本科及大专”合并披露、巨星“大专/本科”合并，"
                     "无法拆分时只计研究生及以上，密度为下限（见 04 表学历明细）。"),
        ("生成日期", pd.Timestamp.today().strftime("%Y-%m-%d")),
    ]
    r = 3
    for k, v in lines:
        ws.cell(r, 1, k).font = openpyxl.styles.Font(bold=True)
        ws.cell(r, 2, v)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(r, 2).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30 if len(v) < 60 else 44
        r += 1
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 100

    # 01/02 指标
    for name, df, title in [
        ("01_指标_单季", ind_q, "单季指标（流量指标已做单季差分；存量指标为期末时点）"),
        ("02_指标_累计", ind_cum, "年初至今累计指标（流量指标为累计值，比率类已年化）"),
    ]:
        ws = write_sheet(wb, name, df, title,
                         pct_cols=PCT_IND, mult_cols=MULT_IND,
                         width_map={c: 13 for c in df.columns if c not in ("公司", "期间")})
        # 分类色带（第2行）
        ws.insert_rows(2)
        cat_colors = {"盈利能力": "DDEBF7", "财务健康": "E2EFDA", "运营效率": "FFF2CC",
                      "现金韧性": "FCE4D6", "发展质量": "E4DFEC"}
        j = 4
        ws.cell(2, 1, "").fill = openpyxl.styles.PatternFill("solid", fgColor="D9E1F2")
        ws.cell(2, 2, "").fill = openpyxl.styles.PatternFill("solid", fgColor="D9E1F2")
        ws.cell(2, 3, "")
        for cat in ["盈利能力", "财务健康", "运营效率", "现金韧性", "发展质量"]:
            n = sum(1 for c in df.columns if CAT_MAP.get(c) == cat)
            if n:
                ws.cell(2, j, cat).fill = openpyxl.styles.PatternFill("solid", fgColor=cat_colors[cat])
                ws.cell(2, j).font = openpyxl.styles.Font(bold=True)
                ws.merge_cells(start_row=2, start_column=j, end_row=2, end_column=j + n - 1)
                for jj in range(j, j + n):
                    ws.cell(2, jj).fill = openpyxl.styles.PatternFill("solid", fgColor=cat_colors[cat])
                j += n
        ws.freeze_panes = "C5"
        ws.auto_filter.ref = f"A4:{openpyxl.utils.get_column_letter(len(df.columns))}{4 + len(df)}"

    # 03 五项资产
    amt = ["投资性房地产", "固定资产", "在建工程", "生产性生物资产", "存货", "五项合计", "资产总计"]
    write_sheet(wb, "03_五项资产", assets,
                "五项资产明细（投资性房地产/固定资产/在建工程/生产性生物资产/存货）",
                pct_cols={"五项占总资产", "固定资产占总资产", "存货占总资产", "五项合计同比"},
                mult_cols={"在建工程/固定资产"}, amount_cols=amt,
                width_map={c: 15 for c in amt})

    # 04 扣非与员工
    ws = wb.create_sheet("04_扣非与员工")
    ws["A1"] = "扣非净利润与员工学历（PDF 解析）"
    ws["A1"].font = openpyxl.styles.Font(bold=True, size=14)
    r = 3
    ws.cell(r, 1, "扣非归母净利润（累计，元）").font = openpyxl.styles.Font(bold=True)
    r += 1
    if not kf.empty:
        for j, col in enumerate(kf.columns, 1):
            c = ws.cell(r, j, col); c.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            c.fill = openpyxl.styles.PatternFill("solid", fgColor="2F5597")
        for row in kf.itertuples(index=False):
            r += 1
            for j, v in enumerate(row, 1):
                c = ws.cell(r, j, v)
                if isinstance(v, float): c.number_format = "#,##0.00"
    r += 3
    ws.cell(r, 1, "员工学历（年报）").font = openpyxl.styles.Font(bold=True)
    r += 1
    if not emp.empty:
        for j, col in enumerate(emp.columns, 1):
            c = ws.cell(r, j, col); c.font = openpyxl.styles.Font(bold=True, color="FFFFFF")
            c.fill = openpyxl.styles.PatternFill("solid", fgColor="2F5597")
        for row in emp.itertuples(index=False):
            r += 1
            for j, v in enumerate(row, 1):
                c = ws.cell(r, j, v)
                if isinstance(v, float): c.number_format = "0.00%"
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 10
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 22

    # 05 原始三大报表（全量）
    raw = df_all[df_all.get("全量") == 1].copy()
    raw = raw[["公司", "报表", "期间", "项目", "金额", "行号"]].sort_values(
        ["公司", "报表", "期间", "行号"], key=lambda s: s.map(lambda x: COMPANY_ORDER.index(x)) if s.name == "公司" else s)
    write_sheet(wb, "05_原始三大报表", raw, "原始三大报表（全量科目，金额=元）",
                amount_cols=["金额"],
                width_map={"项目": 40, "金额": 18})

    # 06 单季利润与现金流
    qrows = []
    for comp in XLSX_PAT:
        for rep in ["利润表", "现金流量表"]:
            sub = df_all[(df_all["公司"] == comp) & (df_all["报表"] == rep) & (df_all.get("全量") != 1)]
            if sub.empty:
                continue
            w = sub.pivot_table(index=["公司", "期间"], columns="项目", values="金额", aggfunc="first")
            q = add_single_quarter(w, list(w.columns))
            for (_, pstr) in q.index:
                for item in q.columns:
                    qrows.append({"公司": comp, "报表": rep, "期间": pstr, "项目": item,
                                  "单季金额": q.loc[(comp, pstr), item]})
    qdf = pd.DataFrame(qrows)
    if not qdf.empty:
        qdf["_c"] = qdf["公司"].map(COMPANY_ORDER.index)
        qdf["_p"] = qdf["期间"].map(period_key)
        qdf = qdf.sort_values(["_c", "_p", "报表", "项目"]).drop(columns=["_c", "_p"])
    write_sheet(wb, "06_单季利润与现金流", qdf, "单季利润表与现金流量表（指标核心科目，差分值，元）",
                amount_cols=["单季金额"], width_map={"项目": 40})

    # 07 校验
    write_sheet(wb, "07_校验", checks, "数据校验")

    # 08 数据源
    write_sheet(wb, "08_数据源", pdf_status, "PDF 解析清单", width_map={"文件": 70})

    # 09 口径说明
    fdf = pd.DataFrame(FORMULAS, columns=["分类", "指标", "计算公式", "说明"])
    ws = write_sheet(wb, "09_口径说明", fdf, "指标计算公式与口径",
                     width_map={"分类": 10, "指标": 18, "计算公式": 62, "说明": 70})
    for i in range(4, 4 + len(fdf)):
        ws.row_dimensions[i].height = 34
        for j in range(1, 5):
            ws.cell(i, j).alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="top")

    wb.save(out)


if __name__ == "__main__":
    build()
