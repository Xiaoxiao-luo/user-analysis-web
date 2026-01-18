import base64
import io
from typing import Dict, List, Tuple

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib as mpl
from fastapi import FastAPI, File, Form, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

COL_REG = "注册时间"
COL_EXP = "体验金领取时间"
COL_FIRST = "首充时间"
COL_SECOND = "二充时间"
COL_PLUS = "升级PLUS时间"

def _set_cn_font():
    mpl.rcParams["font.sans-serif"] = ["Arial Unicode MS", "PingFang SC", "Heiti SC"]
    mpl.rcParams["axes.unicode_minus"] = False

def _annotate_bars(values):
    for i, v in enumerate(values):
        plt.text(i, v, str(int(v)), ha="center", va="bottom")

def _fig_to_base64_png() -> str:
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=160, bbox_inches="tight")
    plt.close()
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")

def _safe_to_datetime(df: pd.DataFrame, cols: List[str]) -> None:
    for c in cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

def _missing_cols(df: pd.DataFrame, cols: List[str]) -> List[str]:
    return [c for c in cols if c not in df.columns]

def analyze_module1(df: pd.DataFrame) -> Tuple[Dict, str, str, List[str], List[str]]:
    warnings, errors = [], []
    need = [COL_FIRST, COL_EXP]
    miss = _missing_cols(df, need)
    if miss:
        return {}, "", "", [f"缺少列：{', '.join(miss)}"], warnings

    _safe_to_datetime(df, [COL_REG, COL_EXP, COL_FIRST])

    base = df[df[COL_FIRST].notna()].copy()
    n_first = len(base)
    if n_first == 0:
        return {"完成首充用户数": 0}, "", "", [], ["首充时间全为空，无法生成分布图。"]

    base["delta_days"] = (base[COL_FIRST] - base[COL_EXP]).dt.total_seconds() / 86400.0

    def bucket(x):
        if pd.isna(x): return "未领取体验金(无法计算Δ)"
        if x < 0: return "先首充再领取体验金"
        if x < 1: return "同时领取体验金并首充人群"
        if x <= 3: return "1-3天"
        if x <= 6: return "4-6天"
        if x <= 10: return "7-10天"
        return "10天以上"

    base["bucket"] = base["delta_days"].apply(bucket)

    order = [
        "未领取体验金(无法计算Δ)",
        "先首充再领取体验金",
        "同时领取体验金并首充人群",
        "1-3天","4-6天","7-10天","10天以上"
    ]

    dist = base["bucket"].value_counts().reindex(order, fill_value=0)
    ratio = (dist / n_first).fillna(0)
    avg_days = base.loc[base["delta_days"].notna() & (base["delta_days"] >= 0), "delta_days"].mean()

    _set_cn_font()
    plt.figure()
    plt.bar(dist.index, dist.values)
    _annotate_bars(dist.values)
    plt.title("模块1：首充时间分布")
    plt.xlabel("时间区间")
    plt.ylabel("用户数")
    plt.xticks(rotation=15)
    bar_b64 = _fig_to_base64_png()

    _set_cn_font()
    plt.figure()
    plt.pie(dist.values, labels=dist.index, autopct=None)
    plt.title("模块1：首充分布（占比结构）")
    pie_b64 = _fig_to_base64_png()

    result = {
        "总首充人数(全表首充非空)": int(n_first),
        "平均耗时(天,仅Δ可算且>=0)": (None if pd.isna(avg_days) else round(float(avg_days), 2)),
        "分布(人数)": {k: int(dist[k]) for k in order},
        "分布(占比)": {k: round(float(ratio[k]), 4) for k in order},
        "分布加总校验": int(dist.sum())
    }
    return result, pie_b64, bar_b64, errors, warnings

def analyze_module2(df: pd.DataFrame) -> Tuple[Dict, str, str, List[str], List[str]]:
    warnings, errors = [], []
    need = [COL_FIRST, COL_SECOND]
    miss = _missing_cols(df, need)
    if miss:
        return {}, "", "", [f"缺少列：{', '.join(miss)}"], warnings

    _safe_to_datetime(df, [COL_FIRST, COL_SECOND])

    base = df[df[COL_FIRST].notna()].copy()
    base_n = len(base)
    if base_n == 0:
        return {"完成首充用户数(母体)": 0}, "", "", [], ["首充时间全为空，无法分析二充。"]

    completed = base[base[COL_SECOND].notna()].copy()
    n_second = len(completed)

    completed["delta_days"] = (completed[COL_SECOND] - completed[COL_FIRST]).dt.total_seconds() / 86400.0

    def bucket(x):
        if pd.isna(x): return "未知"
        if x < 0: return "时间倒流(二充早于首充)"
        if x <= 7: return "1-7天"
        if x <= 14: return "8-14天"
        if x <= 20: return "15-20天"
        return "20天以上"

    completed["bucket"] = completed["delta_days"].apply(bucket)

    order = ["1-7天","8-14天","15-20天","20天以上","时间倒流(二充早于首充)","尚未完成二充"]

    dist_dict = completed["bucket"].value_counts().to_dict()
    dist_dict["尚未完成二充"] = base_n - n_second

    dist = pd.Series(dist_dict).reindex(order, fill_value=0)
    ratio = (dist / base_n).fillna(0)

    _set_cn_font()
    plt.figure()
    plt.bar(dist.index, dist.values)
    _annotate_bars(dist.values)
    plt.title("模块2：二充时间分布")
    plt.xlabel("时间区间")
    plt.ylabel("用户数")
    plt.xticks(rotation=15)
    bar_b64 = _fig_to_base64_png()

    _set_cn_font()
    plt.figure()
    plt.pie(dist.values, labels=dist.index, autopct=None)
    plt.title("模块2：二充分布（占比结构）")
    pie_b64 = _fig_to_base64_png()

    result = {
        "完成首充用户数(母体)": int(base_n),
        "完成二充用户数": int(n_second),
        "二充转化率": round(float(n_second / base_n), 4) if base_n else None,
        "分布(人数)": {k: int(dist[k]) for k in order},
        "分布(占比)": {k: round(float(ratio[k]), 4) for k in order},
        "分布加总校验": int(dist.sum())
    }
    return result, pie_b64, bar_b64, errors, warnings

def analyze_module3(df: pd.DataFrame) -> Tuple[Dict, str, str, List[str], List[str]]:
    warnings, errors = [], []
    need = [COL_SECOND, COL_PLUS]
    miss = _missing_cols(df, need)
    if miss:
        return {}, "", "", [f"缺少列：{', '.join(miss)}"], warnings

    _safe_to_datetime(df, [COL_SECOND, COL_PLUS])

    plus_all = df[df[COL_PLUS].notna()].copy()
    plus_total = len(plus_all)
    plus_without_second = plus_all[plus_all[COL_SECOND].isna()].copy()
    n_plus_without_second = len(plus_without_second)

    base = df[df[COL_SECOND].notna()].copy()
    base_n = len(base)
    if base_n == 0:
        return {
            "全表PLUS总数": int(plus_total),
            "未二充直接PLUS": int(n_plus_without_second),
            "完成二充用户数(母体)": 0
        }, "", "", [], ["二充时间全为空：无法做“二充→PLUS”分布，但已返回全表PLUS来源。"]

    upgraded = base[base[COL_PLUS].notna()].copy()
    n_plus_after_second = len(upgraded)

    upgraded["delta_days"] = (upgraded[COL_PLUS] - upgraded[COL_SECOND]).dt.total_seconds() / 86400.0

    def bucket(x):
        if pd.isna(x): return "未知"
        if x < 0: return "时间倒流(PLUS早于二充)"
        if x <= 7: return "1-7天"
        if x <= 14: return "8-14天"
        if x <= 21: return "15-21天"
        if x <= 28: return "22-28天"
        return "28天以上"

    upgraded["bucket"] = upgraded["delta_days"].apply(bucket)

    order = ["1-7天","8-14天","15-21天","22-28天","28天以上","时间倒流(PLUS早于二充)","尚未升级PLUS"]

    dist_dict = upgraded["bucket"].value_counts().to_dict()
    dist_dict["尚未升级PLUS"] = base_n - n_plus_after_second

    dist = pd.Series(dist_dict).reindex(order, fill_value=0)
    ratio = (dist / base_n).fillna(0)

    _set_cn_font()
    plt.figure()
    plt.bar(dist.index, dist.values)
    _annotate_bars(dist.values)
    plt.title("模块3：PLUS时间分布（完成二充用户）")
    plt.xlabel("时间区间")
    plt.ylabel("用户数")
    plt.xticks(rotation=15)
    bar_b64 = _fig_to_base64_png()

    source_labels = ["完成二充后PLUS", "未二充直接PLUS"]
    source_values = [int(n_plus_after_second), int(n_plus_without_second)]
    _set_cn_font()
    plt.figure()
    plt.pie(source_values, labels=source_labels, autopct=None)
    plt.title("模块3：PLUS来源结构")
    pie_b64 = _fig_to_base64_png()

    result = {
        "全表PLUS总数": int(plus_total),
        "未二充直接PLUS": int(n_plus_without_second),
        "完成二充后PLUS": int(n_plus_after_second),
        "完成二充用户数(母体)": int(base_n),
        "PLUS转化率(母体=二充)": round(float(n_plus_after_second / base_n), 4) if base_n else None,
        "分布(人数,母体=二充)": {k: int(dist[k]) for k in order},
        "分布(占比,母体=二充)": {k: round(float(ratio[k]), 4) for k in order},
        "分布加总校验(母体=二充)": int(dist.sum())
    }
    return result, pie_b64, bar_b64, errors, warnings

app = FastAPI()
templates = Jinja2Templates(directory="app/templates")

@app.get("/", response_class=HTMLResponse)
def start(request: Request):
    return templates.TemplateResponse("start.html", {"request": request})

@app.get("/app", response_class=HTMLResponse)
def app_page(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/run")
async def run(module: str = Form(...), file: UploadFile = File(...)):
    if module not in {"1","2","3"}:
        return JSONResponse({"ok": False, "errors": ["模块必须是 1/2/3"], "warnings": []})

    if not (file.filename.endswith(".xlsx") or file.filename.endswith(".xls")):
        return JSONResponse({"ok": False, "errors": ["请上传 .xlsx/.xls 文件"], "warnings": []})

    try:
        content = await file.read()
        df = pd.read_excel(io.BytesIO(content))
    except Exception as e:
        return JSONResponse({"ok": False, "errors": [f"Excel读取失败：{str(e)}"], "warnings": []})

    try:
        if module == "1":
            result, pie_b64, bar_b64, errors, warnings = analyze_module1(df)
        elif module == "2":
            result, pie_b64, bar_b64, errors, warnings = analyze_module2(df)
        else:
            result, pie_b64, bar_b64, errors, warnings = analyze_module3(df)
    except Exception as e:
        return JSONResponse({"ok": False, "errors": [f"分析过程发生错误：{str(e)}"], "warnings": []})

    return JSONResponse({
        "ok": (len(errors) == 0),
        "module": module,
        "errors": errors,
        "warnings": warnings,
        "result": result,
        "pie_png_base64": pie_b64,
        "bar_png_base64": bar_b64,
    })

