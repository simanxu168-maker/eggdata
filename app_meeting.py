"""
会议纪要 → 网页（Streamlit 版 · 忠实版式）
--------------------------------------------------
4 个板块：基本面 / 多空池 / 交易策略 / 账户运行情况
PPT 原始标题与文字全部原样呈现，只做排版。
自动读取目录中最新的「会议纪要*.pptx」，每周把新 PPT 传上来即自动更新。
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
:root{--ink:#16202c;--ink2:#3d4a5a;--ink3:#7a8698;--line:#dfe5ec;--line2:#eef1f5;
--navy:#12365e;--navy2:#1d4e79;--gold:#a9822f;}
.block-container{max-width:1120px;padding-top:1.6rem;padding-bottom:3rem;}
h1,h2,h3,h4,p,td,th{font-family:"Segoe UI","Microsoft YaHei","PingFang SC",sans-serif;}
.sheet{background:#fff;border:1px solid #dfe4ea;box-shadow:0 10px 30px rgba(18,32,48,.06);}
.cover{padding:38px 44px 26px;border-bottom:1px solid var(--line);}
.cover h1{font-size:30px;line-height:1.3;margin:0 0 22px;font-weight:600;color:var(--navy);}
.cover-grid{display:flex;justify-content:space-between;gap:30px;flex-wrap:wrap;}
.qual{font-size:13px;color:var(--ink2);white-space:pre-line;}
.maker{font-size:13px;color:var(--ink2);white-space:pre-line;text-align:right;}
.disclaimer{margin-top:22px;padding-top:14px;border-top:1px solid var(--line);
font-size:12px;line-height:1.85;color:var(--ink3);}
.bt{display:flex;align-items:center;gap:10px;margin:26px 0 10px;font-size:16px;
font-weight:600;color:var(--navy2);letter-spacing:.02em;}
.bt::before{content:"";width:3px;height:15px;background:var(--navy2);display:inline-block;}
.tbwrap{overflow-x:auto;}
table.mt{width:100%;border-collapse:collapse;table-layout:fixed;font-size:13.5px;}
table.mt th{background:#f3f6f9;color:var(--ink2);font-weight:600;text-align:left;
padding:9px 12px;border-top:1px solid var(--line);border-bottom:1px solid var(--line);
font-size:12.5px;letter-spacing:.04em;}
table.mt td{padding:11px 12px;border-bottom:1px solid var(--line2);vertical-align:top;color:var(--ink2);}
table.mt td.board{color:var(--navy2);font-weight:600;background:#fafcfe;width:11%;}
table.mt td.nm{color:var(--ink);font-weight:600;width:18%;}
table.mt td.nm .st{color:var(--gold);font-weight:400;letter-spacing:2px;font-size:12px;}
table.mt td.vw{line-height:1.8;}
table.mt td.goal{color:var(--navy2);font-weight:500;}
.note{margin:9px 0 0;font-size:12px;color:var(--ink3);}
.pools{display:flex;gap:20px;flex-wrap:wrap;margin-top:8px;}
.pool{flex:1 1 320px;border:1px solid var(--line);}
.pool h4{margin:0;padding:11px 16px;font-size:14px;font-weight:600;color:var(--navy2);
background:#f3f6f9;border-bottom:1px solid var(--line);}
.chips{padding:14px 16px 16px;display:flex;flex-wrap:wrap;gap:8px;min-height:90px;}
.chip{border:1px solid var(--line);padding:5px 10px;font-size:13px;color:var(--ink);background:#fff;white-space:nowrap;}
.chip .st{color:var(--gold);letter-spacing:2px;margin-left:3px;font-size:12px;}
.empty{border:1px dashed #cfd8e3;min-height:280px;display:flex;align-items:center;
justify-content:center;color:var(--ink3);font-size:14px;}
.foot{padding:14px 0 0;border-top:1px solid var(--line);color:var(--ink3);font-size:11.5px;margin-top:30px;}
@media (max-width:820px){
 .cover{padding:26px 18px 20px;} .cover h1{font-size:22px;}
 .maker{text-align:left;} .maker br{display:none;}
 .bt{font-size:15px;} .pools{flex-direction:column;}
 table.mt{min-width:0;table-layout:auto;font-size:12px;}
 table.mt td.board{width:1%;white-space:nowrap;}
 table.mt td.nm{width:1%;white-space:nowrap;}
 table.mt th{white-space:nowrap;}
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


@st.cache_data(show_spinner=False)
def parse_ppt(path_str, mtime):
    prs = Presentation(path_str)
    W = prs.slide_width
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
            if t == title:
                continue
            if t in ("多头品种池", "空头品种池"):
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
                    cover["qual"] = (cover["qual"] + "\n" + t).strip()
                elif t.startswith("报告制作人"):
                    cover["maker"] = t
                elif t.startswith("从业资格号") or t.startswith("投资咨询号"):
                    cover["maker"] = (cover["maker"] + "\n" + t).strip()
                elif t.startswith("重要提示"):
                    cover["disclaimer"] = t
            continue

        if title.startswith("基本面"):
            tb = tables_of(slide)
            fund.append({
                "title": title,
                "header": tb[0][0] if tb else [],
                "rows": tb[0][1:] if tb else [],
                "note": " ".join(notes_of(slide, title)),
            })
        elif "多空池" in title:
            pools = {"bull": {"label": "多头品种池", "items": []},
                     "bear": {"label": "空头品种池", "items": []}}
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
                pools[k]["items"].sort(key=lambda p: (p[2],))
                pools[k]["items"] = [(n, s) for n, s, _ in pools[k]["items"]]
            pools["note"] = " ".join(notes_of(slide, title))
        elif title.startswith("交易策略"):
            tb = tables_of(slide)
            strat.append({
                "title": title,
                "desc": " ".join(notes_of(slide, title)),
                "header": tb[0][0] if tb else [],
                "rows": tb[0][1:] if tb else [],
            })
        elif "账户" in title:
            empty = True

    return cover, fund, pools, strat, empty


def find_ppt(data_dir):
    files = sorted(
        Path(data_dir).glob("会议纪要*.pptx"),
        key=lambda p: (p.stat().st_mtime, p.name),
        reverse=True,
    )
    return files[0] if files else None


def esc(s):
    return html.escape(s or "")


def render_table(header, rows, col_widths):
    if not header or not rows:
        return ""
    cols = "".join(f'<col style="width:{w}">' for w in col_widths)
    th = "".join(f"<th>{esc(h)}</th>" for h in header)
    body = []
    n = len(header)
    if n >= 3 and "板块" in header[0]:
        # 第一列：空单元格视为上一板块的续行，合并成一条
        groups, cur = [], None
        for r in rows:
            if r[0]:
                cur = [r[0], [r]]
                groups.append(cur)
            elif cur is not None:
                cur[1].append(r)
            else:
                cur = ["", [r]]
                groups.append(cur)
        for board, grp in groups:
            span = len(grp)
            for k, r in enumerate(grp):
                name, stars = split_stars(r[1] if len(r) > 1 else "")
                view = " ".join(c for c in r[2:] if c)
                if k == 0:
                    body.append(
                        f'<tr><td class="board" rowspan="{span}">{esc(board)}</td>'
                        f'<td class="nm">{esc(name)}<span class="st">{esc(stars)}</span></td>'
                        f'<td class="vw">{esc(view)}</td></tr>'
                    )
                else:
                    body.append(
                        f'<tr><td class="nm">{esc(name)}<span class="st">{esc(stars)}</span></td>'
                        f'<td class="vw">{esc(view)}</td></tr>'
                    )
    elif "目标" in header[-1] or "风险" in header[-1]:
        for r in rows:
            cells = list(r) + [""] * (n - len(r))
            body.append(
                "<tr>"
                f'<td class="nm">{esc(cells[0])}</td>'
                + "".join(f"<td>{esc(c)}</td>" for c in cells[1:-1])
                + f'<td class="goal">{esc(cells[-1])}</td></tr>'
            )
    else:
        for r in rows:
            cells = list(r) + [""] * (n - len(r))
            body.append("<tr>" + "".join(f"<td>{esc(c)}</td>" for c in cells) + "</tr>")
    return (
        f'<div class="tbwrap"><table class="mt"><colgroup>{cols}</colgroup>'
        f"<thead><tr>{th}</tr></thead><tbody>{''.join(body)}</tbody></table></div>"
    )


def main():
    data_dir = Path(__file__).parent
    ppt = find_ppt(data_dir)

    st.markdown(CSS, unsafe_allow_html=True)
    if ppt is None:
        st.warning("目录里没有找到「会议纪要*.pptx」，把 PPT 上传到仓库根目录即可。")
        return

    with st.spinner("正在读取会议纪要…"):
        cover, fund, pools, strat, empty = parse_ppt(str(ppt), ppt.stat().st_mtime)

    # ---------- 封面 ----------
    st.markdown(
        f"""<div class="sheet cover">
<h1>{esc(cover['title'] or '会议纪要')}</h1>
<div class="cover-grid">
<div class="qual">{esc(cover['qual'])}</div>
<div class="maker">{esc(cover['maker'])}</div>
</div>
<div class="disclaimer">{esc(cover['disclaimer'])}</div>
</div>""",
        unsafe_allow_html=True,
    )

    tabs = st.tabs(["基本面", "多空池", "交易策略", "账户运行情况"])

    with tabs[0]:
        if not fund:
            st.markdown('<p class="note">本期基本面暂无内容。</p>', unsafe_allow_html=True)
        for b in fund:
            st.markdown(f'<div class="bt">{esc(b["title"])}</div>', unsafe_allow_html=True)
            st.markdown(
                render_table(b["header"], b["rows"], ["11%", "18%", "71%"]),
                unsafe_allow_html=True,
            )
            if b["note"]:
                st.markdown(f'<p class="note">{esc(b["note"])}</p>', unsafe_allow_html=True)

    with tabs[1]:
        if pools:
            st.markdown('<div class="bt">多空池</div>', unsafe_allow_html=True)
            cols = st.columns(2)
            for col, key in zip(cols, ["bull", "bear"]):
                p = pools.get(key)
                if not p:
                    continue
                chips = "".join(
                    f'<span class="chip">{esc(n)}<span class="st">{esc(s)}</span></span>'
                    for n, s in p["items"]
                )
                with col:
                    st.markdown(
                        f'<div class="pool"><h4>{esc(p["label"])}</h4>'
                        f'<div class="chips">{chips}</div></div>',
                        unsafe_allow_html=True,
                    )
            if pools.get("note"):
                st.markdown(
                    f'<p class="note">{esc(pools["note"])}</p>', unsafe_allow_html=True
                )
        else:
            st.markdown('<p class="note">本期多空池暂无内容。</p>', unsafe_allow_html=True)

    with tabs[2]:
        if not strat:
            st.markdown('<p class="note">本期交易策略暂无内容。</p>', unsafe_allow_html=True)
        for s in strat:
            st.markdown(f'<div class="bt">{esc(s["title"])}</div>', unsafe_allow_html=True)
            if s["desc"]:
                st.markdown(
                    f'<p class="note" style="margin:0 0 8px">{esc(s["desc"])}</p>',
                    unsafe_allow_html=True,
                )
            st.markdown(
                render_table(s["header"], s["rows"], ["19%", "18%", "43%", "20%"]),
                unsafe_allow_html=True,
            )

    with tabs[3]:
        st.markdown('<div class="bt">账户运行情况</div>', unsafe_allow_html=True)
        st.markdown('<div class="empty">—</div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="foot">本页由会议 PPT 排版生成，文字与观点未做改动。</div>',
        unsafe_allow_html=True,
    )


main()
