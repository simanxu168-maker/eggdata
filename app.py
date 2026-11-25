"""
鸡蛋 · 平均淘汰日龄拐头跟踪（Streamlit 网页版）
--------------------------------------------------
运行方式（本地）：  streamlit run app.py
数据文件：          同目录下的「鸡蛋日龄拐头跟踪数据库.xlsx」
更新数据：          改 Excel → 保存 → 网页右上角 Rerun（或重启 app）
"""

from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from plotly.subplots import make_subplots

# ---------------- 阈值（想调就改这里） ----------------
PROFIT_LINE = 1.00   # 单斤利润临界点：跌破并传导 2–5 周
PROFIT_WIN_LO = 2
PROFIT_WIN_HI = 5
TAO_LINE = 5.00      # 淘鸡价「跌破并站稳」线
TAO_BREAK = 4.80     # 淘鸡价破位线
AGE_WARN = 550       # 日龄警戒线
CHAN_THR = 1.0       # 通道比值阈值

BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "鸡蛋日龄拐头跟踪数据库.xlsx"

# 表头（原表 A1 的列名是错位的，这里按列位置重新命名）
COLS = [
    "日期", "卓创日龄", "钢联日龄", "周出栏", "蛋价", "淘鸡价", "全价料", "单斤利润", "通道比值",
    "冻库容率", "大码占比", "库存生产", "库存流通", "五城销量", "日龄环比", "出栏环比",
    "状态", "三条件", "通道信号",
]

st.set_page_config(page_title="鸡蛋 · 淘汰日龄拐头跟踪", page_icon="🥚", layout="wide")


# ---------------- 读数据 ----------------
@st.cache_data(show_spinner=False)
def load_data(path: str, mtime: float) -> pd.DataFrame:
    """读取『每周数据』，清洗成一张干净的周度表。mtime 变化时缓存自动失效。"""
    df = pd.read_excel(path, sheet_name="每周数据", header=0)
    while df.shape[1] < len(COLS):
        df[df.shape[1]] = np.nan
    df = df.iloc[:, : len(COLS)].copy()
    df.columns = COLS

    df["日期"] = pd.to_datetime(df["日期"], errors="coerce")
    df = df.dropna(subset=["日期"]).sort_values("日期").reset_index(drop=True)

    for c in COLS[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    # 只保留有实际读数的周
    df = df.dropna(subset=["卓创日龄", "钢联日龄", "周出栏", "蛋价", "单斤利润"], how="all")
    df = df.reset_index(drop=True)

    # 环比缺列时自己算
    df["蛋价环比%"] = (df["蛋价"] / df["蛋价"].shift(1) - 1) * 100
    df["淘鸡环比%"] = (df["淘鸡价"] / df["淘鸡价"].shift(1) - 1) * 100
    df["日龄环比"] = df["日龄环比"].fillna(df["钢联日龄"].diff())
    df["出栏环比"] = df["出栏环比"].fillna((df["周出栏"] / df["周出栏"].shift(1) - 1) * 100)

    # 通道比值 = 淘鸡跌幅 ÷ 蛋价跌幅（表里这列大多是空的，这里自动补）
    calc = df["淘鸡环比%"] / df["蛋价环比%"]
    df["通道"] = df["通道比值"].where(df["通道比值"].notna(), calc)
    df["通道"] = df["通道"].replace([np.inf, -np.inf], np.nan)

    # 状态缺列时按矩阵补：出栏降→停滞；出栏增 + 日龄降/平/升 → 真去化/过渡/轮换
    def _state(row):
        if pd.isna(row["日龄环比"]) or pd.isna(row["出栏环比"]):
            return "—"
        if row["出栏环比"] <= 0:
            return "停滞"
        if row["日龄环比"] < 0:
            return "真去化"
        return "过渡" if row["日龄环比"] == 0 else "轮换"

    df["状态"] = df.apply(lambda r: r["状态"] if isinstance(r["状态"], str) and r["状态"].strip() else _state(r), axis=1)
    return df


def streak(cond: pd.Series) -> int:
    """从最新一周往前数，连续满足条件的周数。"""
    n = 0
    for v in cond.iloc[::-1].fillna(False).astype(bool):
        if v:
            n += 1
        else:
            break
    return n


if not DATA_FILE.exists():
    st.error(
        "没找到数据文件。\n\n"
        f"请把「鸡蛋日龄拐头跟踪数据库.xlsx」放在和 app.py 同一个文件夹里：\n\n`{BASE_DIR}`"
    )
    st.stop()

df = load_data(str(DATA_FILE), DATA_FILE.stat().st_mtime)
if len(df) < 2:
    st.error("『每周数据』里没有读到有效数据，请检查表里最新一周是否已填好。")
    st.stop()

latest, prev = df.iloc[-1], df.iloc[-2]

# ---------------- 侧边栏 ----------------
with st.sidebar:
    st.header("显示设置")
    weeks = st.slider("图表显示周数", min_value=13, max_value=104, value=52, step=13)
    table_weeks = st.slider("明细表周数", min_value=12, max_value=104, value=26, step=12)
    st.divider()
    st.caption(f"数据文件：{DATA_FILE.name}")
    st.caption(f"最后修改：{datetime.fromtimestamp(DATA_FILE.stat().st_mtime):%Y-%m-%d %H:%M}")
    st.caption(f"共 {len(df)} 周，最新 {latest['日期']:%Y-%m-%d}")
    if st.button("重新读取 Excel", width="stretch"):
        st.cache_data.clear()
        st.rerun()

d = df.tail(weeks).copy()

# ---------------- 标题与结论 ----------------
st.title("鸡蛋 · 平均淘汰日龄拐头跟踪")
st.caption(
    f"数据截至 **{latest['日期']:%Y-%m-%d}**（周一）　·　"
    f"来源：钢联 / 卓创 / iFinD　·　本页由 Excel 自动生成"
)

ps = streak(df["单斤利润"] < PROFIT_LINE)
up = streak(df["出栏环比"] > 0)
down = streak(df["日龄环比"] < 0)
chan_streak = streak(df["通道"] > CHAN_THR)
chan_now = latest["通道"]

c1 = (not pd.isna(latest["单斤利润"])) and ps >= PROFIT_WIN_LO
c2 = up >= 2 and down >= 1
c3 = chan_streak >= 2
ok1 = int(c1) + int(c2) + int(c3)

lead_ok = int(bool(chan_streak >= 2)) + int(bool(not pd.isna(latest["大码占比"]) and latest["大码占比"] < prev["大码占比"])) + int(bool(not pd.isna(latest["淘鸡价"]) and latest["淘鸡价"] < TAO_LINE))
gate_ok = int(bool(not pd.isna(latest["淘鸡价"]) and latest["淘鸡价"] < TAO_BREAK)) + int(bool(down >= 2)) + 0  # 第三关现金线缺全价料，暂不计

state = latest["状态"] if isinstance(latest["状态"], str) else "—"

if ok1 == 3:
    st.success(f"**拐头确认**：三前置条件全部成立（{ok1}/3）。")
elif ok1 >= 1:
    st.warning(f"**高位平台 · 尚未拐头**：三前置 {ok1}/3，日龄只是不再涨，不等于开始跌。")
else:
    st.info(f"**信号未启动**：三前置 0/3，日龄仍在高位僵持。")

# ---------------- KPI ----------------
row1 = st.columns(3)
row2 = st.columns(3)


def metric(col, label, value, delta=None):
    with col:
        st.metric(label, value, delta, delta_color="off")


def fmt(v, dp=0, suffix=""):
    return "—" if pd.isna(v) else f"{v:,.{dp}f}{suffix}"


def chg(v, dp=0, suffix=""):
    return "—" if pd.isna(v) else f"{v:+.{dp}f}{suffix}"


metric(row1[0], "钢联淘汰日龄", fmt(latest["钢联日龄"], 0, " 天"), chg(latest["日龄环比"], 0, " 天"))
metric(row1[1], "卓创淘汰日龄", fmt(latest["卓创日龄"], 0, " 天"), chg(latest["卓创日龄"] - prev["卓创日龄"], 0, " 天"))
metric(row1[2], "周度出栏量", fmt(latest["周出栏"], 0, " 万只"), chg(latest["出栏环比"], 2, "%"))
metric(row2[0], "蛋价周均价", fmt(latest["蛋价"], 2, " 元/斤"), chg((latest["蛋价"] / prev["蛋价"] - 1) * 100, 2, "%"))
metric(row2[1], "淘鸡周均价", fmt(latest["淘鸡价"], 2, " 元/斤"), chg((latest["淘鸡价"] / prev["淘鸡价"] - 1) * 100, 2, "%"))
metric(row2[2], "单斤利润", fmt(latest["单斤利润"], 2, " 元/斤"), chg(latest["单斤利润"] - prev["单斤利润"], 2, " 元"))

# ---------------- 主图：日龄 vs 出栏 ----------------
st.subheader("淘汰日龄 vs 周度出栏")
st.caption("「量增 + 龄降」才是真去化，缺一个都不算。")

fig = make_subplots(specs=[[{"secondary_y": True}]])
fig.add_trace(
    go.Bar(x=d["日期"], y=d["周出栏"], name="周出栏（万只）", marker_color="#ccdae7", opacity=0.9),
    secondary_y=True,
)
fig.add_trace(
    go.Scatter(x=d["日期"], y=d["钢联日龄"], name="钢联日龄", line=dict(color="#1d4e79", width=2)),
    secondary_y=False,
)
fig.add_trace(
    go.Scatter(x=d["日期"], y=d["卓创日龄"], name="卓创日龄", line=dict(color="#c08a1e", width=2, dash="dot")),
    secondary_y=False,
)
fig.add_trace(
    go.Scatter(
        x=[d["日期"].min(), d["日期"].max()], y=[AGE_WARN, AGE_WARN], mode="lines", name=f"{AGE_WARN} 天",
        line=dict(color="#c8453c", dash="dash", width=1.2), hoverinfo="skip",
    ),
    secondary_y=False,
)
fig.update_yaxes(title_text="日龄（天）", secondary_y=False)
fig.update_yaxes(title_text="出栏（万只）", secondary_y=True, rangemode="tozero")
fig.update_layout(
    height=420, hovermode="x unified",
    margin=dict(l=10, r=10, t=40, b=10),
    legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
)
st.plotly_chart(fig, width="stretch")

# ---------------- 两张小图 ----------------
c_left, c_right = st.columns(2)

with c_left:
    st.subheader("单斤利润")
    st.caption("跌破 1.00 并传导 2–5 周，是拐头的第一个前置条件。")
    fig2 = go.Figure()
    fig2.add_trace(go.Scatter(
        x=d["日期"], y=d["单斤利润"], name="单斤利润（元/斤）",
        fill="tozeroy", line=dict(color="#1f8a5b", width=1.8),
    ))
    fig2.add_hline(y=PROFIT_LINE, line_dash="dash", line_color="#c8453c",
                   annotation_text=f"{PROFIT_LINE:.2f} 元", annotation_position="top left")
    fig2.update_layout(height=260, margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified", showlegend=False)
    st.plotly_chart(fig2, width="stretch")

with c_right:
    st.subheader("蛋价 vs 淘鸡价")
    st.caption("淘鸡跌得比蛋价快（比值 > 1）＝ 屠宰端接不住。")
    fig3 = go.Figure()
    fig3.add_trace(go.Scatter(x=d["日期"], y=d["蛋价"], name="蛋价", line=dict(color="#1d4e79", width=1.8)))
    fig3.add_trace(go.Scatter(x=d["日期"], y=d["淘鸡价"], name="淘鸡价", line=dict(color="#b5603f", width=1.8)))
    fig3.add_hline(y=TAO_LINE, line_dash="dash", line_color="#c8453c", annotation_text=f"{TAO_LINE:.2f}")
    fig3.add_hline(y=TAO_BREAK, line_dash="dot", line_color="#9aa3b0", annotation_text=f"{TAO_BREAK:.2f}")
    fig3.update_layout(height=260, margin=dict(l=10, r=10, t=30, b=10), hovermode="x unified",
                       legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0))
    st.plotly_chart(fig3, width="stretch")

# ---------------- 信号面板 ----------------
st.subheader("信号面板")

ICON = {"ok": "🟢", "part": "🟡", "gray": "⚪"}

s1 = [
    ("① 利润传导", f"利润 {fmt(latest['单斤利润'], 2)} 元/斤，连续 {ps} 周低于 {PROFIT_LINE:.2f}",
     "已达成 · 传导期满" if ps >= PROFIT_WIN_HI else ("已达成 · 传导中" if ps >= PROFIT_WIN_LO else "未达"),
     "ok" if c1 else "gray"),
    ("② 池子释放", f"出栏 {fmt(latest['周出栏'], 0)} 万只，环比 {chg(latest['出栏环比'], 2, '%')}，连增 {up} 周；日龄环比 {chg(latest['日龄环比'], 1, ' 天')}",
     "已达成 · 量增龄降" if c2 else "未达", "ok" if c2 else "gray"),
    ("③ 屠宰通道", f"通道比值 {fmt(chan_now, 2)}（阈值 > {CHAN_THR:.1f}），连续 {chan_streak} 周",
     "已达成 · 连续两周" if c3 else "未达", "ok" if c3 else "gray"),
]
s2 = [
    ("① 通道比值", f"本周 {fmt(chan_now, 2)}，阈值 > {CHAN_THR:.1f}",
     "已亮" if chan_streak >= 2 else "未亮", "ok" if chan_streak >= 2 else "gray"),
    ("② 大码蛋占比", f"{fmt(latest['大码占比'], 2)}%，上周 {fmt(prev['大码占比'], 2)}%",
     "已亮 · 开始下降" if (not pd.isna(latest["大码占比"]) and latest["大码占比"] < prev["大码占比"]) else "未亮 · 大码还在堆",
     "ok" if (not pd.isna(latest["大码占比"]) and latest["大码占比"] < prev["大码占比"]) else "gray"),
    ("③ 淘鸡价水平", f"{fmt(latest['淘鸡价'], 2)} 元/斤，距 {TAO_LINE:.2f} {chg(latest['淘鸡价'] - TAO_LINE, 2, ' 元')}",
     "已亮 · 破位并站稳" if (not pd.isna(latest["淘鸡价"]) and latest["淘鸡价"] < TAO_LINE) else "未亮",
     "ok" if (not pd.isna(latest["淘鸡价"]) and latest["淘鸡价"] < TAO_LINE) else "gray"),
]
s3 = [
    ("第一关 · 现金线", "蛋价 < 550 天现金成本（全价料×2.45+0.45）",
     "无数据 · 全价料未填" if pd.isna(latest["全价料"]) else "待判定", "gray"),
    ("第二关 · 日龄线", f"钢联 {fmt(latest['钢联日龄'], 0)} 天，环比 {chg(latest['日龄环比'], 1, ' 天')}，连续下行 {down} 周",
     "已亮 · 连续两周下行" if down >= 2 else ("部分 · 下行 1 周" if down == 1 else "未亮"),
     "ok" if down >= 2 else ("part" if down == 1 else "gray")),
    ("第三关 · 淘鸡线", f"全国周均 {fmt(latest['淘鸡价'], 2)}，破位线 {TAO_BREAK:.2f}",
     "已亮 · 破 4.8" if (not pd.isna(latest["淘鸡价"]) and latest["淘鸡价"] < TAO_BREAK) else "未亮",
     "ok" if (not pd.isna(latest["淘鸡价"]) and latest["淘鸡价"] < TAO_BREAK) else "gray"),
]


def render_group(col, title, items, total):
    with col:
        st.markdown(f"**{title}　{total} / 3**")
        for name, read, judge, level in items:
            st.markdown(f"{ICON[level]} **{name}**  \n{judge}　·　{read}")


gc1, gc2, gc3 = st.columns(3)
render_group(gc1, "① 三前置条件", s1, ok1)
render_group(gc2, "② 先行征兆", s2, lead_ok)
render_group(gc3, "③ 三道关", s3, gate_ok)

st.caption(
    f"状态判别：**{state}**　·　"
    f"出栏 {'↑' if latest['出栏环比'] > 0 else '↓'} / 日龄 "
    f"{'↓' if latest['日龄环比'] < 0 else ('→' if latest['日龄环比'] == 0 else '↑')}"
    f"　·　只有「日龄降 + 出栏增」才是真去化"
)

# ---------------- 明细表 ----------------
st.subheader("周度数据")
tbl = df.tail(table_weeks)[::-1].copy()
tbl["日期"] = tbl["日期"].dt.strftime("%Y-%m-%d")
tbl = tbl.rename(columns={
    "周出栏": "周出栏(万只)", "蛋价": "蛋价(元/斤)", "淘鸡价": "淘鸡价(元/斤)",
    "单斤利润": "单斤利润(元)", "日龄环比": "日龄环比(天)", "出栏环比": "出栏环比(%)",
})
show_cols = ["日期", "卓创日龄", "钢联日龄", "周出栏(万只)", "蛋价(元/斤)", "淘鸡价(元/斤)",
             "单斤利润(元)", "日龄环比(天)", "出栏环比(%)", "通道", "状态"]
st.dataframe(
    tbl[show_cols].style.format({
        "卓创日龄": "{:.0f}", "钢联日龄": "{:.0f}", "周出栏(万只)": "{:.0f}",
        "蛋价(元/斤)": "{:.2f}", "淘鸡价(元/斤)": "{:.2f}", "单斤利润(元)": "{:.2f}",
        "日龄环比(天)": "{:+.0f}", "出栏环比(%)": "{:+.2f}", "通道": "{:.2f}",
    }, na_rep="—"),
    hide_index=True, width="stretch", height=430,
)

# ---------------- 说明 ----------------
with st.expander("数据口径与更新方法"):
    st.markdown(
        f"""
**数据来源**：钢联 Mysteel、卓创资讯、同花顺 iFinD，周度口径；单斤利润 = 蛋价 − 自算现金成本。

**阈值**（改 app.py 顶部的常量即可）：利润 {PROFIT_LINE:.2f} 元 · 传导 {PROFIT_WIN_LO}–{PROFIT_WIN_HI} 周 ·
淘鸡 {TAO_LINE:.2f} / 破位 {TAO_BREAK:.2f} · 日龄警戒 {AGE_WARN} 天 · 通道比值 {CHAN_THR:.1f}。

**更新数据**：替换同目录的 xlsx → 网页自动重新读（或点侧边栏「重新读取 Excel」）。
线上版本：在 GitHub 仓库里重新上传同名 xlsx，约 1 分钟后自动更新。

**缺口**：`全价料` 列为空时，「第一关 · 现金线」无法自动判定；`通道比值` 列大多为空，本页用
「淘鸡环比 ÷ 蛋价环比」自动补算。
"""
    )

st.divider()
st.caption("仅供研究参考 · 不构成投资建议")
