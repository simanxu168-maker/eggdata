"""
会议纪要 → 网页（Streamlit 版）
--------------------------------------------------
数据：同目录下最新的「会议纪要*.pptx」（自动取文件名最新的那份）
本地运行：streamlit run app_meeting.py
线上：GitHub 仓库里上传新的 pptx 覆盖/新增，约 1 分钟自动更新
"""

import html
import re
from datetime import datetime
from pathlib import Path

import streamlit as st
from pptx import Presentation

BASE_DIR = Path(__file__).resolve().parent

# 板块合并规则：同一个 Tab 里放哪些“板块”
MERGE = {
    "股指": "股指 · 贵金属 · 橡胶",
    "贵金属": "股指 · 贵金属 · 橡胶",
    "橡胶": "股指 · 贵金属 · 橡胶",
    "油脂油料": "油脂油料",
    "化工": "化工 · 氯碱",
    "氯碱": "化工 · 氯碱",
}

st.set_page_config(page_title="行情策略会纪要", page_icon="📄", layout="wide", initial_sidebar_state="collapsed")

st.markdown(
    """
<style>
@media (max-width: 640px){
  .block-container{padding:1.0rem .7rem 2rem !important;}
  h1{font-size:1.35rem !important;}
  .eg-card p{font-size:.82rem;}
}
.eg-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px;margin:6px 0 6px;}
.eg-card{background:#fff;border:1px solid #e4e8ee;border-radius:12px;padding:14px 16px;}
.eg-top{display:flex;align-items:center;justify-content:space-between;gap:10px;}
.eg-top h4{margin:0;font-size:15px;font-weight:600;color:#1c2430;}
.eg-stars{color:#b8892f;font-size:13px;letter-spacing:3px;}
.eg-card p{margin:8px 0 0;font-size:13px;line-height:1.75;color:#435062;}
.eg-grouptitle{font-size:12.5px;color:#435062;font-weight:600;letter-spacing:.06em;
  margin:18px 0 6px;padding-bottom:6px;border-bottom:1px solid #e4e8ee;}
.eg-table{width:100%;border-collapse:collapse;font-size:13px;background:#fff;
  border:1px solid #e4e8ee;border-radius:10px;overflow:hidden;}
.eg-table th{background:#fafbfc;color:#435062;font-weight:600;text-align:left;padding:10px 12px;border-bottom:1px solid #e4e8ee;}
.eg-table td{padding:11px 12px;border-bottom:1px solid #f0f2f5;vertical-align:top;color:#435062;}
.eg-table td b{color:#1c2430;}
.eg-table td.goal{color:#1d4e79;font-weight:500;}
.eg-table tr:last-child td{border-bottom:none;}
</style>
""",
    unsafe_allow_html=True,
)


def clean(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\r", " ").replace("\n", " ").replace("\v", " ")).strip()


@st.cache_data(show_spinner=False)
def find_source(root: str) -> tuple:
    """取目录下文件名最新的「会议纪要*.pptx」。"""
    files = sorted(Path(root).glob("会议纪要*.pptx"))
    files = [f for f in files if not f.name.startswith("~$")]
    if not files:
        return ("", 0.0)
    f = files[-1]
    return (str(f), f.stat().st_mtime)


@st.cache_data(show_spinner=False)
def parse_pptx(path: str, mtime: float) -> dict:
    prs = Presentation(path)
    groups, strategies = [], []

    for slide in prs.slides:
        title = ""
        for sh in slide.shapes:
            if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
                t = clean(sh.text_frame.text)
                if 0 < len(t) < 40:
                    title = t
                    break

        for sh in slide.shapes:
            if not getattr(sh, "has_table", False) or not sh.has_table:
                continue
            tbl = sh.table
            rows, cols = len(tbl.rows), len(tbl.columns)
            if rows < 2 or cols < 3:
                continue
            header = [clean(c.text) for c in tbl.rows[0].cells]
            hdr = "|".join(header)

            if "策略类型" in hdr:
                name = re.sub(r"^\s*交易策略\s*[·•\.]\s*", "", title) or "策略"
                out = []
                for r in range(1, rows):
                    cells = [clean(c.text) for c in tbl.rows[r].cells]
                    if not cells[0]:
                        continue
                    out.append({
                        "name": cells[0], "type": cells[1],
                        "act": cells[2], "goal": cells[3] if cols >= 4 else "",
                    })
                if out:
                    strategies.append({"title": name, "rows": out})

            elif ("板块" in hdr) or ("品种" in hdr):
                last_board, last_view = "", {}
                for r in range(1, rows):
                    cells = [clean(c.text) for c in tbl.rows[r].cells]
                    board = cells[0] or last_board
                    if cells[0]:
                        last_board = cells[0]
                    raw = cells[1]
                    if not raw:
                        continue
                    stars = raw.count("★")
                    nm = clean(raw.replace("★", ""))
                    view = cells[2] if cols >= 3 else ""
                    if view and view == last_view.get(board, ""):
                        view = "同上（原文为合并单元格，观点见上一行）"
                    else:
                        last_view[board] = view

                    grp = next((g for g in groups if g["board"] == board), None)
                    if grp is None:
                        grp = {"board": board, "items": []}
                        groups.append(grp)
                    grp["items"].append({"name": nm, "stars": stars, "text": view})

    return {"groups": groups, "strategies": strategies}


def tab_name(board: str) -> str:
    return MERGE.get(board, board)


def stars(n: int) -> str:
    return "★" * max(0, min(3, n or 0))


path, mtime = find_source(str(BASE_DIR))
if not path:
    st.error("没找到「会议纪要*.pptx」。请把 PPT 放在和 app_meeting.py 同一个目录里。")
    st.stop()

with st.spinner("正在解析 PPT…"):
    data = parse_pptx(path, mtime)

groups, strategies = data["groups"], data["strategies"]

# 组装 Tab
order, by_tab = [], {}
for g in groups:
    t = tab_name(g["board"])
    if t not in by_tab:
        by_tab[t] = {"boards": [], "items": []}
        order.append(t)
    by_tab[t]["boards"].append(g["board"])
    for it in g["items"]:
        by_tab[t]["items"].append({**it, "board": g["board"]})
if strategies:
    order.append("交易策略")

# 头部
src_name = Path(path).name
m = re.search(r"(\d{4})(\d{2})(\d{2})", Path(path).stem)
date = f"{m.group(1)}-{m.group(2)}-{m.group(3)}" if m else datetime.fromtimestamp(mtime).strftime("%Y-%m-%d")

n_items = sum(len(g["items"]) for g in groups)
n_rows = sum(len(s["rows"]) for s in strategies)

st.title("品种观点与交易策略汇编")
st.caption(f"会议日期 **{date}**　·　源文件 {src_name}　·　{len(order)} 个板块 / {n_items} 条观点 / {n_rows} 条策略")

with st.sidebar:
    st.header("显示设置")
    st.caption(f"源文件：{src_name}")
    st.caption(f"修改时间：{datetime.fromtimestamp(mtime):%Y-%m-%d %H:%M}")
    st.caption("换新 PPT：上传新的「会议纪要YYYYMMDD.pptx」覆盖或新增，本页自动取最新一份。")
    if st.button("重新读取 PPT", width="stretch"):
        st.cache_data.clear()
        st.rerun()

labels = [f"{i+1} {t}" for i, t in enumerate(order)]
tabs = st.tabs(labels)

for i, title in enumerate(order):
    with tabs[i]:
        if title == "交易策略":
            st.caption("股指 CTA 仅交易股指期货；时序量价取单品种自身价量趋势；套利遵循“多低估、空高估”。")
            for s in strategies:
                st.markdown(f"**{html.escape(s['title'])}　·　{len(s['rows'])} 条**")
                rows_html = "".join(
                    "<tr>"
                    f"<td><b>{html.escape(r['name'])}</b></td>"
                    f"<td>{html.escape(r['type'] or '—')}</td>"
                    f"<td>{html.escape(r['act'] or '—')}</td>"
                    f"<td class='goal'>{html.escape(r['goal'] or '—')}</td>"
                    "</tr>"
                    for r in s["rows"]
                )
                st.markdown(
                    "<table class='eg-table'><thead><tr><th>品种</th><th>策略类型</th>"
                    "<th>操作要点</th><th>目标 / 风险</th></tr></thead><tbody>"
                    f"{rows_html}</tbody></table>",
                    unsafe_allow_html=True,
                )
                st.write("")
        else:
            g = by_tab[title]
            boards = list(dict.fromkeys(g["boards"]))
            st.caption("覆盖 " + "、".join(boards) + "；星级数量代表观点强度（原文标注）。")
            for b in boards:
                items = [it for it in g["items"] if it["board"] == b]
                st.markdown(f"<div class='eg-grouptitle'>{html.escape(b)}</div>", unsafe_allow_html=True)
                cards = "".join(
                    "<div class='eg-card'>"
                    f"<div class='eg-top'><h4>{html.escape(it['name'])}</h4>"
                    f"<span class='eg-stars'>{stars(it['stars'])}</span></div>"
                    f"<p>{html.escape(it['text'] or '—')}</p>"
                    "</div>"
                    for it in items
                )
                st.markdown(f"<div class='eg-grid'>{cards}</div>", unsafe_allow_html=True)

st.divider()
st.caption(
    "口径说明：观点与星级为原报告标注，未做改动；★ 数量代表观点强度。"
    f"　·　由 {src_name} 自动生成，仅供研究参考，不构成投资建议。"
)
