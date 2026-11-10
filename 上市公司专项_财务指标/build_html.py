# -*- coding: utf-8 -*-
"""生成线下单文件 HTML 看板：21 家公司指标趋势、五项资产与指标快照。

用法：
    D:\\miniconda\\python.exe build_html.py

产出：
    output/上市公司财务指标_趋势与资产_看板.html
    桌面副本：D:\\CC\\Desktop\\上市公司财务指标_趋势与资产_看板.html

数据来自 build_data.py 产出的 data/*.csv；ECharts 内嵌，单文件离线可用。
"""
from __future__ import annotations

import json
import math
import shutil
from pathlib import Path

import pandas as pd

from build_data import CAT_MAP, FORMULAS, MULT_IND, PCT_IND, UNIT_IND

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
TEMPLATE = ROOT / "html_template.html"
ECHARTS = ROOT.parent / "echarts.min.js"
OUT = ROOT / "output" / "上市公司财务指标_趋势与资产_看板.html"
DESKTOP = Path("D:/CC/Desktop/上市公司财务指标_趋势与资产_看板.html")

ASSET_ITEMS = ["投资性房地产", "固定资产", "在建工程", "生产性生物资产", "存货"]
LOWER_IS_BETTER = {"期间费用率", "有息负债率", "重资产占比"}


def period_key(p: str) -> int:
    y, q = int(p[:4]), p[4:]
    return y * 4 + {"Q1": 0, "H1": 1, "Q3": 2, "FY": 3}[q]


def build_indicators() -> list[dict]:
    out = []
    for cat, name, formula, note in FORMULAS:
        typ = "pct" if name in PCT_IND else "percap" if name in UNIT_IND else "mult"
        out.append({
            "name": name,
            "cat": cat,
            "type": typ,
            "dir": -1 if name in LOWER_IS_BETTER else 1,
            "formula": formula,
            "note": note,
            "unit": {"pct": "%", "percap": "万元/人", "mult": "倍"}[typ],
        })
    return out


def clean(v, digits: int):
    if v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v))):
        return None
    try:
        x = float(v)
    except (TypeError, ValueError):
        return None
    if math.isnan(x) or math.isinf(x):
        return None
    return round(x, digits)


def main() -> None:
    q = pd.read_csv(DATA / "indicators_single.csv")
    c = pd.read_csv(DATA / "indicators_cum.csv")
    a = pd.read_csv(DATA / "assets_five.csv")

    companies = list(dict.fromkeys(q["公司"].tolist()))
    periods = sorted(q["期间"].unique(), key=period_key)
    indicators = build_indicators()
    names = [x["name"] for x in indicators]

    assert len(companies) == 21, len(companies)
    assert len(periods) == 27, len(periods)
    assert len(indicators) == 20, len(indicators)
    for nm in names:
        assert nm in q.columns and nm in c.columns, nm
    for nm in ASSET_ITEMS + ["资产总计"]:
        assert nm in a.columns, nm

    series: dict[str, dict[str, dict[str, list]]] = {}
    for cal, df in [("单季", q), ("累计", c)]:
        series[cal] = {}
        for ind in indicators:
            nm = ind["name"]
            digits = 3 if ind["type"] == "percap" else 6
            series[cal][nm] = {}
            for comp in companies:
                sub = df[df["公司"] == comp].set_index("期间")
                arr = []
                for per in periods:
                    arr.append(clean(sub.at[per, nm], digits) if per in sub.index else None)
                series[cal][nm][comp] = arr

    assets: dict[str, list] = {}
    for comp in companies:
        sub = a[a["公司"] == comp].set_index("期间")
        rows = []
        for per in periods:
            if per not in sub.index:
                rows.append(None)
                continue
            rec = sub.loc[per]
            row = [clean(rec[item] / 1e8, 4) if not pd.isna(rec[item]) else None
                   for item in ASSET_ITEMS]
            row.append(clean(rec["资产总计"] / 1e8, 4) if not pd.isna(rec["资产总计"]) else None)
            rows.append(row)
        assets[comp] = rows

    payload = {
        "meta": {
            "title": "上市猪企财务质量面板",
            "companies": len(companies),
            "periods": len(periods),
            "indicators": len(indicators),
            "start": periods[0],
            "end": periods[-1],
        },
        "companies": companies,
        "periods": periods,
        "indicators": indicators,
        "assetItems": ASSET_ITEMS,
        "series": series,
        "assets": assets,
    }

    if not ECHARTS.exists():
        raise FileNotFoundError(f"缺少 ECharts：{ECHARTS}")
    ejs = ECHARTS.read_text(encoding="utf-8")
    if "</script" in ejs.lower():
        raise ValueError("echarts.min.js 含 </script，无法直接内嵌")

    json_text = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    html = TEMPLATE.read_text(encoding="utf-8")
    html = html.replace("__DATA_JSON__", json_text).replace("/*__ECHARTS__*/", ejs)
    html = html.replace("__BUILD_DATE__", "2026-11-10")

    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(html, encoding="utf-8")
    size_mb = OUT.stat().st_size / 1024 / 1024
    print(f"written: {OUT}  ({size_mb:.2f} MB)")
    print(f"companies={len(companies)} periods={len(periods)} indicators={len(indicators)}")

    if DESKTOP.parent.exists():
        shutil.copyfile(OUT, DESKTOP)
        print(f"desktop: {DESKTOP}")


if __name__ == "__main__":
    main()
