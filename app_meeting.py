"""
会议纪要 → 网页（Streamlit 版 · 左侧目录）
--------------------------------------------------
左侧：周次 + 板块 + 品种 三级目录
右侧：所选周次、所选板块的内容
要求：标题与原文不改，只做排版；每周上传新的「会议纪要*.pptx」自动进入目录。
"""

import re
import html
from pathlib import Path

import streamlit as st
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

st.set_page_config(page_title="研究中心内部行情策略会", layout="wide")

CSS = """
<style>
:root{--side:#20222a;--side2:#2a2d37;--teal:#009688;--top:#393d49;--bg:#f2f3f5;--card:#fff;
--ink:#333;--ink2:#666;--ink3:#9aa0ab;--line:#e6e8ec;--red:#bf242a;--green:#0a7d63;--gold:#ff5722;}
.stApp{background:var(--bg);overflow-x:hidden;}
html,body{overflow-x:hidden;}
html,body,[class*="css"]{font-family:"Helvetica Neue",Helvetica,"PingFang SC","Microsoft YaHei",Tahoma,Arial,sans-serif;}
header[data-testid="stHeader"]{background:var(--top) !important;display:none;}
.block-container{max-width:100%;padding:14px 18px 40px;}
/* 左侧深色栏 */
section[data-testid="stSidebar"]{background:var(--side);border-right:0;width:236px !important;min-width:236px !important;max-width:236px !important;}
section[data-testid="stSidebar"] > div{padding-top:.9rem;}
section[data-testid="stSidebar"] .block-container{padding:.9rem .9rem 2rem;max-width:100%;}
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p,
section[data-testid="stSidebar"] [data-testid="stCaptionContainer"] p,
[data-testid="stSidebar"] small{color:rgba(255,255,255,.55) !important;}
section[data-testid="stSidebar"] p{color:rgba(255,255,255,.8) !important;}
section[data-testid="stSidebar"] [role="radiogroup"] label p{color:rgba(255,255,255,.8) !important;font-size:13.5px;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"] > div > div:first-child{
background:rgba(255,255,255,.12) !important;border-color:transparent !important;}
section[data-testid="stSidebar"] label[data-testid="stRadioOption"][data-selected="true"] > div > div:first-child{
background:var(--teal) !important;}
section[data-testid="stSidebar"] [data-baseweb="select"] > div{background:#2a2d36 !important;border-color:#3a3e49 !important;color:#fff !important;}
section[data-testid="stSidebar"] [data-baseweb="select"] *{color:#fff !important;}
section[data-testid="stSidebar"] hr{border-color:rgba(255,255,255,.08);}
.sb-brand{display:flex;align-items:center;gap:9px;padding:0 0 10px;margin-bottom:2px;}
.sb-brand .mark{width:26px;height:26px;border-radius:6px;background:linear-gradient(135deg,#00b39f,#00796b);
display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:#fff;}
.sb-brand b{font-size:14px;font-weight:600;color:#fff;line-height:1.35;letter-spacing:.01em;}
.sb-brand span{font-size:11px;color:rgba(255,255,255,.5);display:block;}
.sb-line{height:1px;background:rgba(255,255,255,.08);margin:10px 0 12px;}
/* 顶部深色条 */
/* 顶部分段标签 */
[data-testid="stButtonGroup"]{margin:0 0 14px;background:#fff;border:1px solid #e3e8ec;border-radius:8px;
padding:4px;display:inline-flex;box-shadow:0 1px 2px rgba(16,24,40,.05);}
[data-testid="stButtonGroup"] button{border-radius:6px !important;border:0 !important;
background:transparent !important;color:#4a5560 !important;font-size:13.5px !important;
min-height:38px !important;padding:0 20px !important;}
[data-testid="stButtonGroup"] button[data-selected="true"],
[data-testid="stButtonGroup"] button[aria-checked="true"]{background:var(--teal) !important;
border-color:var(--teal) !important;color:#fff !important;}
[data-testid="stButtonGroup"] button:hover{background:#f1f7f6 !important;color:var(--teal) !important;}
.mobrand{display:none;}
.topbar{display:flex;align-items:center;gap:10px;background:var(--top);color:rgba(255,255,255,.9);
border-radius:0;height:52px;padding:0 20px;font-size:14px;margin:-14px -18px 12px;}
.topbar .sep{color:rgba(255,255,255,.4);}
.topbar b{color:#fff;font-weight:500;}
.topbar .r{margin-left:auto;font-size:12px;color:rgba(255,255,255,.62);border:1px solid rgba(255,255,255,.16);
border-radius:3px;padding:2px 9px;}
/* 彩色分类条 */
.strip{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px;}
.chipcat{background:#fff;border:1px solid var(--line);border-radius:6px;overflow:hidden;
box-shadow:0 1px 2px rgba(0,0,0,.03);}
.chipcat .lab{padding:4px 4px;color:#fff;font-size:12px;font-weight:600;text-align:center;letter-spacing:.04em;}
.chipcat .val{padding:11px 4px 12px;text-align:center;font-size:22px;font-weight:700;color:#222;line-height:1.15;}
.chipcat .val small{display:block;font-size:11px;font-weight:400;color:var(--ink3);margin-top:2px;}
.chipcat.c1 .lab{background:#132b6b;} .chipcat.c2 .lab{background:var(--teal);}
.chipcat.c3 .lab{background:#6e0000;} .chipcat.c4 .lab{background:#00401c;}
/* 白色卡片 */
.card{background:#fff;border:1px solid var(--line);border-radius:6px;margin-bottom:14px;
box-shadow:0 1px 2px rgba(0,0,0,.03);}
.card-h{display:flex;align-items:baseline;gap:10px;padding:13px 18px;border-bottom:1px solid #f2f3f5;}
.card-h h2{margin:0;font-size:15px;font-weight:600;color:#222;line-height:1.25;}
.card-h .sub{font-size:12px;color:var(--ink3);}
.card-b{padding:14px 18px 16px;}
.bt{margin:0 0 10px;font-size:15px;font-weight:600;color:#222;}
.bt em{font-style:normal;font-size:12px;color:var(--ink3);margin-left:10px;font-weight:400;}
.cover .meta{display:flex;flex-wrap:wrap;gap:6px 22px;font-size:12.5px;color:var(--ink2);margin-bottom:10px;}
.cover .note{margin:0;padding-top:10px;border-top:1px solid #f4f4f4;font-size:11.5px;line-height:1.9;color:var(--ink3);}
.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.gsep{grid-column:1/-1;display:flex;align-items:center;gap:8px;font-size:12px;color:var(--ink2);
margin:2px 0 0;}
.gsep::before{content:"";width:3px;height:11px;background:var(--teal);}
.pcard{border:1px solid #edeff2;border-radius:5px;padding:11px 13px;background:#fff;}
.phead{display:flex;align-items:baseline;gap:8px;margin-bottom:4px;}
.pname{font-size:13.5px;font-weight:600;color:#222;}
.stars{font-size:11px;color:var(--gold);letter-spacing:1.5px;white-space:nowrap;}
.pview{margin:0;font-size:12.5px;line-height:1.75;color:var(--ink2);}
.tip{margin:12px 0 0;font-size:11.5px;color:var(--ink3);}
.pools{display:grid;grid-template-columns:1fr 1fr;gap:12px;}
.pool{background:#fff;border:1px solid #edeff2;border-radius:5px;overflow:hidden;}
.ph{display:flex;align-items:center;justify-content:space-between;padding:9px 12px;
font-size:13px;font-weight:600;border-bottom:1px solid #f2f2f2;}
.ph .dot{width:6px;height:6px;border-radius:50%;display:inline-block;margin-right:7px;vertical-align:1px;}
.pb{padding:11px 12px 12px;display:flex;flex-wrap:wrap;gap:7px;}
.chip{padding:4px 9px;border-radius:4px;font-size:12px;background:#fff;border:1px solid #e6e8ec;}
.chip i{font-style:normal;font-size:10.5px;letter-spacing:1.5px;margin-left:3px;}
.bull .ph{color:var(--red);background:#fdf4f3;border-bottom-color:#f7dfdc;} .bull .dot{background:var(--red);}
.bull .chip i{color:var(--red);} .bull .chip{border-color:#f0d5d1;}
.bear .ph{color:var(--green);background:#f2faf7;border-bottom-color:#d9ece5;} .bear .dot{background:var(--green);}
.bear .chip i{color:var(--green);} .bear .chip{border-color:#d5eae2;}
.desc{font-size:12px;color:var(--ink3);margin:-4px 0 10px;}
.scard{border:1px solid #edeff2;border-radius:5px;background:#fff;padding:11px 13px;margin-bottom:9px;}
.shead{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:5px;}
.sname{font-size:13.5px;font-weight:600;color:#222;}
.tag{font-size:11.5px;color:var(--teal);background:#eff9f7;border:1px solid #cfe9e3;
padding:2px 8px;border-radius:3px;white-space:nowrap;}
.srow{display:grid;grid-template-columns:64px 1fr;gap:8px;font-size:12.5px;color:var(--ink2);line-height:1.8;}
.srow b{font-weight:400;color:var(--ink3);font-size:11.5px;}
.srow.goal span{color:var(--green);font-weight:600;}
.skel{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;}
.sk{border:1px dashed #dcdcdc;border-radius:4px;height:110px;display:flex;align-items:center;
justify-content:center;color:#bbb;font-size:13px;background:#fcfcfc;}
.tl{display:grid;gap:10px;margin-bottom:14px;}
.tlc{background:#fff;border:1px solid var(--line);border-left:3px solid var(--teal);
border-radius:4px;padding:13px 16px;}
.tld{display:flex;align-items:center;gap:10px;font-size:12px;color:var(--ink3);margin-bottom:5px;}
.tld b{color:#374151;font-weight:600;font-size:12.5px;}
.tlv{margin:0;font-size:13px;line-height:1.85;color:var(--ink2);}
@media(max-width:860px){
header[data-testid="stHeader"]{display:flex !important;align-items:center;height:52px;background:var(--top) !important;
left:0 !important;right:0 !important;width:100% !important;padding:0 8px;}
header[data-testid="stHeader"]::after{content:none !important;}
header[data-testid="stHeader"] [data-testid="stToolbarActions"],
header[data-testid="stHeader"] [data-testid="stMainMenu"]{display:none !important;}
.mobrand{display:flex;position:fixed;left:72px;right:8px;top:0;height:52px;align-items:center;
background:var(--top);color:#fff;font-size:12px;font-weight:600;z-index:1000000;pointer-events:none;
letter-spacing:0;}
header[data-testid="stHeader"] [data-testid="stSidebarCollapseButton"] *{color:#fff !important;fill:#fff !important;}
header[data-testid="stHeader"] button{color:#fff !important;background:rgba(0,150,136,.95) !important;
border-radius:6px !important;min-height:32px !important;padding:0 10px !important;}
header[data-testid="stHeader"] svg{fill:#fff !important;color:#fff !important;}
header[data-testid="stHeader"] [data-testid="stMainMenu"]{display:none !important;}
header[data-testid="stHeader"] [data-testid="stAppDeployButton"]{display:none !important;}
.block-container{padding-top:64px;}
.topbar{display:none;}
[data-testid="stButtonGroup"]{display:flex !important;width:100%;flex-wrap:wrap !important;justify-content:center;}
[data-testid="stButtonGroup"]{justify-content:space-evenly;}
[data-testid="stButtonGroup"] button{flex:0 0 auto !important;padding:0 9px !important;font-size:12px !important;min-height:34px !important;}
.strip{grid-template-columns:1fr 1fr;gap:8px;} .chipcat .val{font-size:13.5px;}
.grid,.pools{grid-template-columns:1fr;} .pcard{padding:12px 13px;}
.pview{font-size:12.8px;} .srow{grid-template-columns:66px 1fr;}
.skel{grid-template-columns:1fr 1fr;} .sk{height:80px;}
.cover .meta{display:grid;gap:5px;font-size:12px;}
.block-container{padding-left:.85rem;padding-right:.85rem;}
.topbar{padding:0 12px;font-size:12px;} .topbar .r{display:none;}
}
</style>
"""


def clean(t):
    return re.sub(r"\s+", " ", str(t or "")).strip()


def split_stars(text):
    m = re.search(r"(★+)$", text)
    if m:
        return text[: m.start()].strip(), m.group(1)
    return text.strip(), ""


def esc(s):
    return html.escape(str(s or ""))


@st.cache_data(show_spinner=False)
def parse_week(path_str, mtime):
    prs = Presentation(path_str)
    slides = list(prs.slides)

    def shapes_text(shapes):
        out = []
        for sh in shapes:
            if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
                out.extend(shapes_text(sh.shapes))
            elif getattr(sh, "has_text_frame", False):
                for p in sh.text_frame.paragraphs:
                    t = clean("".join(r.text for r in p.runs))
                    if t:
                        out.append((t, sh.left or 0, sh.top or 0))
        return out

    def title_of(slide):
        best = None
        for sh in slide.shapes:
            if not getattr(sh, "has_text_frame", False):
                continue
            t = clean(sh.text_frame.text)
            if not t or len(t) > 40:
                continue
            size = None
            for p in sh.text_frame.paragraphs:
                for r in p.runs:
                    if r.font.size:
                        size = r.font.size.pt
                        break
                if size:
                    break
            key = (size or 0, -(sh.top or 0))
            if best is None or key > best[0]:
                best = (key, t)
        return best[1] if best else ""

    def tables_of(slide):
        res = []
        for sh in slide.shapes:
            if getattr(sh, "has_table", False) and sh.has_table:
                t = sh.table
                rows = []
                for r in t.rows:
                    cells, last = [], None
                    for c in r.cells:
                        v = clean(c.text)
                        if v and last is not None and v == last:
                            v = ""
                        if v:
                            last = v
                        cells.append(v)
                    if any(cells):
                        rows.append(cells)
                if rows:
                    res.append(rows)
        return res

    def notes_of(slide, title):
        out = []
        for t, _, _ in shapes_text(slide.shapes):
            if t == title or t in ("多头品种池", "空头品种池"):
                continue
            if t.startswith("★") or len(t) < 8:
                continue
            out.append(t)
        return out

    cover = {"title": "", "qual": "", "maker": "", "disclaimer": ""}
    fund, pools, strat, empty = [], {}, [], False

    for idx, slide in enumerate(slides):
        title = title_of(slide)
        if idx == 0:
            cover["title"] = title
            for t, _, _ in shapes_text(slide.shapes):
                if t == title:
                    continue
                if t.startswith("投资咨询业务资格"):
                    cover["qual"] = t
                elif t.startswith("证监许可"):
                    cover["qual"] = (cover["qual"] + " " + t).strip()
                elif t.startswith("报告制作人"):
                    cover["maker"] = t
                elif t.startswith("从业资格号") or t.startswith("投资咨询号"):
                    cover["maker"] = (cover["maker"] + "  " + t).strip()
                elif t.startswith("重要提示"):
                    cover["disclaimer"] = t
            continue

        if title.startswith("基本面"):
            tb = tables_of(slide)
            rows, board, last_view = [], "", ""
            for r in (tb[0][1:] if tb else []):
                if len(r) < 3:
                    continue
                if r[0]:
                    board = r[0]
                if r[2]:
                    last_view = r[2]
                if r[1]:
                    rows.append([board, r[1], r[2] or last_view])
            fund.append({
                "title": title,
                "rows": rows,
                "note": " ".join(notes_of(slide, title)),
            })
        elif "多空池" in title:
            pools = {"bull": {"label": "多头品种池", "items": []},
                     "bear": {"label": "空头品种池", "items": []}}
            txt_pools = {"bull": [], "bear": []}
            for t, _, _ in shapes_text(slide.shapes):
                if "多头品种池" in t or "空头品种池" in t:
                    body = t.split("：", 1)[-1] if "：" in t else ""
                    key = "bull" if "多头" in t else "bear"
                    for seg in re.split("）", body):
                        if "（" in seg:
                            nm = seg.split("（")[0].strip()
                            st_ = seg.split("（", 1)[1].strip()
                            if nm:
                                txt_pools[key].append((nm, st_))
            if txt_pools["bull"] or txt_pools["bear"]:
                for k in ("bull", "bear"):
                    pools[k]["items"] = txt_pools[k]
            else:
                names, stars = [], []
                for t, left, top in shapes_text(slide.shapes):
                    if t in ("多头品种池", "空头品种池", title):
                        continue
                    if t.startswith("★") and len(t) <= 4:
                        stars.append((t, left, top))
                    elif len(t) <= 12:
                        names.append((t, left, top))
                for name, x, y in names:
                    best, bestd = "", None
                    for s, sx, sy in stars:
                        dy = sy - y
                        if dy < -100000 or dy > 900000:
                            continue
                        d = abs(sx - x) + dy
                        if bestd is None or d < bestd:
                            best, bestd = s, d
                    key = "bull" if x < 5800000 else "bear"
                    pools[key]["items"].append((name, best, y))
                for k in ("bull", "bear"):
                    pools[k]["items"].sort(key=lambda p: p[2])
                    pools[k]["items"] = [(n, s) for n, s, _ in pools[k]["items"]]
            pools["note"] = " ".join(notes_of(slide, title))
        elif title.startswith("交易策略"):
            tb = tables_of(slide)
            strat.append({
                "title": title,
                "desc": " ".join(notes_of(slide, title)),
                "rows": tb[0][1:] if tb else [],
            })
        elif "账户" in title:
            empty = True

    return {"cover": cover, "fund": fund, "pools": pools, "strat": strat, "empty": empty}


def week_date(name):
    m = re.search(r"(\d{8})", name)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return ""


def load_weeks(data_dir):
    files = sorted(
        Path(data_dir).glob("会议纪要*.pptx"),
        key=lambda p: (week_date(p.name), p.name),
        reverse=True,
    )
    weeks = []
    for f in files:
        data = parse_week(str(f), f.stat().st_mtime)
        data["date"] = week_date(f.name) or f.stem
        data["file"] = f.name
        weeks.append(data)
    return weeks


def records(week):
    """把一周的基本面压平成 品种记录。"""
    out = []
    for b in week["fund"]:
        board = ""
        for r in b["rows"]:
            if len(r) < 2:
                continue
            if r[0]:
                board = r[0]
            name, stars = split_stars(r[1])
            view = " ".join(c for c in r[2:] if c)
            out.append({
                "block": b["title"], "board": board,
                "name": name, "stars": stars, "view": view,
            })
    return out


def fund_html(block, only=None):
    parts = ['<div class="grid">']
    board, last = "", None
    for r in block["rows"]:
        if len(r) < 2:
            continue
        if r[0]:
            board = r[0]
        name, stars = split_stars(r[1])
        if only and name != only:
            continue
        if board and board != last:
            last = board
            parts.append(f'<div class="gsep">{esc(board)}</div>')
        view = " ".join(c for c in r[2:] if c)
        parts.append(
            f'<article class="pcard"><div class="phead">'
            f'<span class="pname">{esc(name)}</span>'
            f'<span class="stars">{esc(stars)}</span></div>'
            f'<p class="pview">{esc(view)}</p></article>'
        )
    parts.append("</div>")
    return "".join(parts)


def pool_html(key, p):
    chips = "".join(
        f'<span class="chip">{esc(n)}<i>{esc(s)}</i></span>' for n, s in p["items"]
    )
    return (
        f'<div class="pool {key}"><div class="ph">'
        f'<span><span class="dot"></span>{esc(p["label"])}</span>'
        f"<span>{len(p['items'])}</span></div>"
        f'<div class="pb">{chips}</div></div>'
    )


def strategy_html(s):
    parts = []
    for r in s["rows"]:
        cells = list(r) + [""] * (4 - len(r))
        parts.append(
            f'<div class="scard"><div class="shead">'
            f'<span class="sname">{esc(cells[0])}</span>'
            f'<span class="tag">{esc(cells[1])}</span></div>'
            f'<div class="srow"><b>操作要点</b><span>{esc(cells[2])}</span></div>'
            f'<div class="srow goal"><b>目标 / 风险</b><span>{esc(cells[3])}</span></div>'
            f"</div>"
        )
    return "".join(parts)


def topbar(week, section):
    st.html(
        f'<div class="mobrand">研究中心内部行情策略会</div>'
        f'<div class="topbar"><b>研究中心内部行情策略会</b>'
        f'<span class="r">{esc(week["date"])} · 内部资料</span></div>'
    )


def strip(week, section):
    names = sorted({r["name"] for r in records(week)})
    bull = len(week["pools"].get("bull", {}).get("items", [])) if week["pools"] else 0
    bear = len(week["pools"].get("bear", {}).get("items", [])) if week["pools"] else 0
    srows = sum(len(s["rows"]) for s in week["strat"])
    cells = [
        ("c1", "覆盖品种", f"{len(names)} 个品种", "本周跟踪"),
        ("c2", "多头池", f"{bull} 个", "看多品种"),
        ("c3", "空头池", f"{bear} 个", "看空品种"),
        ("c4", "策略超市", f"{srows} 条", "在跟策略"),
    ]
    html_parts = "".join(
        f'<div class="chipcat {c}"><div class="lab">{esc(t)}</div>'
        f'<div class="val">{esc(v)}<small>{esc(s)}</small></div></div>'
        for c, t, v, s in cells
    )
    st.markdown(f'<div class="strip">{html_parts}</div>', unsafe_allow_html=True)


SECTIONS = ["策略超市", "品种观点", "品种多空池", "账户运行情况"]


def _jump_to_opinion():
    name = st.session_state.get("jmp")
    if name:
        st.session_state.sec = "品种观点"
        st.session_state.vf = name
        st.session_state.jmp = None


def main():
    data_dir = Path(__file__).parent
    st.html(CSS)
    weeks = load_weeks(data_dir)

    if weeks:
        topbar(weeks[0], "")
    if st.session_state.get("sec") not in SECTIONS:
        st.session_state.sec = "品种观点"
    section = st.pills(
        "板块", SECTIONS, key="sec", label_visibility="collapsed",
    ) or "品种观点"

    # ---------------- 左侧目录 ----------------
    with st.sidebar:
        st.html(
            '<div class="sb-brand"><div class="mark">研</div>'
            '<b>研究中心内部行情策略会</b></div>'
        )

        if not weeks:
            st.info("目录里还没有「会议纪要*.pptx」，上传后会自动出现。")
            return

        labels = [f"{w['date']}" + ("　最新" if i == 0 else "") for i, w in enumerate(weeks)]
        widx = st.selectbox("周次", range(len(weeks)), format_func=lambda i: labels[i])


        week = weeks[widx]
        names = sorted({r["name"] for r in records(week)})
        pick = None
        if section == "品种观点":
            st.html('<div class="sb-line"></div>')
            opts = ["全部品种"] + names
            cur = st.session_state.get("vf")
            if cur and cur != "全部品种" and cur not in opts:
                opts.insert(1, cur)
            if cur not in opts:
                st.session_state.vf = "全部品种"
            sel = st.selectbox("品种", opts, key="vf")
            pick = None if sel == "全部品种" else sel
            if pick and len(weeks) > 1:
                st.caption("选择单个品种时，右侧会按周列出它的观点变化。")

        st.html('<div class="sb-line"></div>')
        st.caption(f"共 {len(weeks)} 期 · 上传新 PPT 自动更新")

    # ---------------- 右侧内容 ----------------

    if section == "品种观点":
        if pick and len(weeks) > 1:
            st.markdown(
                f'<div class="bt">{esc(pick)} · 按周追踪<em>共 {len(weeks)} 期</em></div>',
                unsafe_allow_html=True,
            )
            items = []
            for w in weeks:
                for rec in records(w):
                    if rec["name"] == pick:
                        items.append((w["date"], rec))
                        break
            cards = "".join(
                f'<div class="tlc"><div class="tld"><b>{esc(d)}</b>'
                f'<span>{esc(r["board"])}</span>'
                f'<span class="stars">{esc(r["stars"])}</span></div>'
                f'<p class="tlv">{esc(r["view"])}</p></div>'
                for d, r in items
            )
            st.markdown(f'<div class="tl">{cards}</div>', unsafe_allow_html=True)
        else:
            for b in week["fund"]:
                shown = b["rows"]
                if pick:
                    shown = [r for r in b["rows"]
                             if len(r) > 1 and split_stars(r[1])[0] == pick]
                    if not shown:
                        continue
                note = f'<p class="tip">{esc(b["note"])}</p>' if b["note"] else ""
                st.html(
                    f'<div class="card"><div class="card-h"><h2>{esc(b["title"])}</h2>'
                    f'<span class="sub">{len(shown)} 个品种</span></div>'
                    f'<div class="card-b">{fund_html(b, only=pick)}{note}</div></div>'
                )

    elif section == "品种多空池":
        if week["pools"]:
            inner = (
                '<div class="pools">'
                + pool_html("bull", week["pools"]["bull"])
                + pool_html("bear", week["pools"]["bear"])
                + "</div>"
            )
            tip = (f'<p class="tip">{esc(week["pools"]["note"])}</p>'
                   if week["pools"].get("note") else "")
            st.html(
                f'<div class="card"><div class="card-h"><h2>品种多空池</h2>'
                f'<span class="sub">本周多空品种</span></div>'
                f'<div class="card-b">{inner}{tip}</div></div>'
            )
            pool_names = sorted(
                {x[0] for x in week["pools"]["bull"]["items"] + week["pools"]["bear"]["items"]}
            )
            if pool_names:
                st.caption("点击任一品种，直接跳到它在「品种观点」里的原文")
                st.pills("快速跳转品种观点", pool_names, key="jmp",
                         on_change=_jump_to_opinion, label_visibility="collapsed")
        else:
            st.markdown('<p class="tip">本期品种多空池暂无内容。</p>', unsafe_allow_html=True)

    elif section == "策略超市":
        if not week["strat"]:
            st.markdown('<p class="tip">本期策略超市暂无内容。</p>', unsafe_allow_html=True)
        for s in week["strat"]:
            desc = f'<p class="desc">{esc(s["desc"])}</p>' if s["desc"] else ""
            st.html(
                f'<div class="card"><div class="card-h"><h2>{esc(s["title"])}</h2></div>'
                f'<div class="card-b">{desc}{strategy_html(s)}</div></div>'
            )

    else:
        st.html(
            '<div class="card"><div class="card-h"><h2>账户运行情况</h2>'
            '<span class="sub">本期待补充</span></div><div class="card-b"><div class="skel">'
            + "".join('<div class="sk">—</div>' for _ in range(4))
            + "</div></div></div>"
        )


main()





