# -*- coding: utf-8 -*-
"""
生成纪要网页.py —— 把目录里所有「会议纪要*.pptx」排版成一个网页（meeting.html）

用法：
    1) 把新的 PPT 放进本文件夹（命名形如 会议纪要20261130.pptx）
    2) 双击运行本脚本（或命令行 python 生成纪要网页.py）
    3) 生成 meeting.html，双击即可打开；周次自动进左侧目录

说明：只做排版，不改标题、不改正文文字；解析不到的字段留空，不编造。
"""

import json
import re
import sys
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

HERE = Path(__file__).resolve().parent
CSS_FILE = HERE / "纪要样式.css"
OUT = HERE / "meeting.html"


def clean(t):
    return re.sub(r"\s+", " ", str(t or "")).strip()


def split_stars(text):
    m = re.search(r"(★+)$", text)
    if m:
        return text[: m.start()].strip(), m.group(1)
    return text.strip(), ""


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


def walk(shapes, out):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            walk(sh.shapes, out)
        elif getattr(sh, "has_text_frame", False):
            t = clean(sh.text_frame.text)
            if t:
                out.append((t, sh.left or 0, sh.top or 0))


def tables_of(slide):
    res = []
    for sh in slide.shapes:
        if getattr(sh, "has_table", False) and sh.has_table:
            t = sh.table
            rows = []
            for r in t.rows:
                rows.append([clean(c.text) for c in r.cells])
            if rows:
                res.append(rows)
    return res


def pool_items_from_text(text):
    body = text.split("：", 1)[-1] if "：" in text else text
    out = []
    for seg in re.split("）", body):
        if "（" in seg:
            name = seg.split("（")[0].strip()
            stars = seg.split("（", 1)[1].strip()
            if name:
                out.append([name, stars])
    return out


def parse_week(path):
    prs = Presentation(str(path))
    cover = {"title": "", "qual": "", "maker": "", "disc": ""}
    fund, strat, pools, empty = [], [], {}, False
    pool_text_items = {"bull": [], "bear": []}
    notes = {}

    for idx, slide in enumerate(prs.slides):
        title = title_of(slide)
        texts = []
        walk(slide.shapes, texts)

        if idx == 0:
            cover["title"] = title
            for t, _, _ in texts:
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
                    cover["disc"] = t
            continue

        if "多空池" in title:
            for t, _, _ in texts:
                if "多头品种池" in t:
                    pool_text_items["bull"] += pool_items_from_text(t)
                elif "空头品种池" in t:
                    pool_text_items["bear"] += pool_items_from_text(t)
            pools["_slide"] = slide
            pools["_title"] = title
            continue

        if title.startswith("基本面"):
            tb = tables_of(slide)
            rows, board, last_view = [], "", ""
            for r in tb[0][1:] if tb else []:
                if len(r) < 3:
                    continue
                if r[0]:
                    board = r[0]
                name, stars = split_stars(r[1]) if r[1] else ("", "")
                view = r[2] or last_view
                if r[2]:
                    last_view = r[2]
                if name:
                    rows.append([board, name, stars, view])
            note = ""
            for t, _, _ in texts:
                if "★即为看空" in t:
                    note = t
            fund.append({"title": title, "rows": rows, "note": note})
            continue

        if title.startswith("交易策略"):
            tb = tables_of(slide)
            desc = ""
            for t, _, _ in texts:
                if t == title:
                    continue
                if ("CTA：" in t) or ("跨期：" in t):
                    desc = t
            rows = []
            for r in (tb[0][1:] if tb else []):
                cells = list(r) + [""] * (4 - len(r))
                if not any(cells):
                    continue
                rows.append({"name": cells[0], "tag": cells[1],
                             "op": cells[2], "goal": cells[3]})
            strat.append({"title": title, "desc": desc, "rows": rows})
            continue

        if "账户" in title:
            empty = True

    # 多空池：优先用文本框（本周格式），否则用坐标分组（旧格式）
    bull, bear = pool_text_items["bull"], pool_text_items["bear"]
    if not bull and not bear and pools.get("_slide") is not None:
        slide = pools["_slide"]
        names, stars = [], []
        for t, left, top in texts_of_slide(slide):
            if t in ("多头品种池", "空头品种池", pools.get("_title", "")):
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
            (bull if x < 5800000 else bear).append([name, best])

    pools_out = {}
    if bull or bear:
        pools_out = {"bull": bull, "bear": bear}

    return {"cover": cover, "fund": fund, "strat": strat,
            "pools": pools_out, "empty": empty}


def texts_of_slide(slide):
    out = []
    walk(slide.shapes, out)
    return out


def week_date(name):
    m = re.search(r"(\d{8})", name)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"
    return ""


def main():
    files = sorted(HERE.glob("会议纪要*.pptx"),
                   key=lambda p: (week_date(p.name), p.name), reverse=True)
    if not files:
        print("目录里没有「会议纪要*.pptx」")
        return 1

    weeks = []
    for f in files:
        w = parse_week(f)
        w["date"] = week_date(f.name) or f.stem
        weeks.append(w)
        print(f"{f.name} → {w['date']}  基本面 {len(w['fund'])} 块 / "
              f"策略 {sum(len(s['rows']) for s in w['strat'])} 条 / "
              f"多空 {len(w['pools'].get('bull', []))}+{len(w['pools'].get('bear', []))}")

    css = CSS_FILE.read_text(encoding="utf-8") if CSS_FILE.exists() else ""
    html = SHELL.replace("__CSS__", css).replace(
        "__DATA__", json.dumps(weeks, ensure_ascii=False))
    OUT.write_text(html, encoding="utf-8")
    print(f"已生成 {OUT.name}（{len(weeks)} 期）")
    return 0


SHELL = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>研究中心内部行情策略会</title>
<style>
__CSS__
</style>
</head>
<body>

<div class="topbar">
  <div class="burger" onclick="document.querySelector('.side').classList.toggle('open')"><span></span><span></span><span></span></div>
  <h1>研究中心内部行情策略会<small id="hdate">每周纪要</small></h1>
  <div class="search"><input id="q" placeholder="搜索品种，如：菜粕" oninput="filterByName(this.value)"></div>
  <div class="topdate" id="topdate"></div>
</div>

<aside class="side">
  <div class="brand"><i></i><b>研究中心内部行情策略会</b></div>
  <div class="grp">
    <div class="grp-t">目 录</div>
    <button class="item" data-p="p3"><span class="ic">◎</span>策略超市<span class="cnt" id="c3"></span></button>
    <button class="item on" data-p="p1"><span class="ic">◈</span>品种观点<span class="cnt" id="c1"></span></button>
    <button class="item" data-p="p2"><span class="ic">◑</span>品种多空池<span class="cnt" id="c2"></span></button>
    <button class="item" data-p="p4"><span class="ic">▤</span>账户运行情况<span class="cnt">—</span></button>
  </div>
  <div class="grp">
    <div class="grp-t">周 次</div>
    <div id="weeks"></div>
  </div>
  <div class="side-foot">渤海期货 · 研究中心<br>每周更新，上传 PPT 自动生成</div>
</aside>

<button class="fab" onclick="document.querySelector('.side').classList.toggle('open')">☰ 目录</button>

<main class="main">
  <div class="col-main">
    <div class="strip" id="strip"></div>
    <div class="tabs">
      <button class="tab" data-p="p3"><span class="ti">◎</span>策略超市<em class="tn" id="t3"></em></button>
      <button class="tab on" data-p="p1"><span class="ti">◈</span>品种观点<em class="tn" id="t1"></em></button>
      <button class="tab" data-p="p2"><span class="ti">◑</span>品种多空池<em class="tn" id="t2"></em></button>
      <button class="tab" data-p="p4"><span class="ti">▤</span>账户运行情况</button>
    </div>
    <div class="pillbar"><span class="plab">品种</span><div class="pills" id="pills"></div></div>
    <section class="panel on" id="p1"></section>
    <section class="panel" id="p2"></section>
    <section class="panel" id="p3"></section>
    <section class="panel" id="p4"></section>
  </div>
  <aside class="rail">
    <div class="card">
      <div class="card-h"><h2 style="font-size:13.5px">本周多空池</h2><span class="right" id="railcnt">共 0</span></div>
      <div class="card-b" id="railList"></div>
    </div>
  </aside>
</main>

<script id="DATA" type="application/json">__DATA__</script>
<script>
var DATA = JSON.parse(document.getElementById("DATA").textContent);
var W = 0, SEC = "p1", FILTER = "全部品种";

function esc(s){ return String(s == null ? "" : s).replace(/[&<>"]/g, function(c){
  return {"&":"&amp;","<":"&lt;",">":"&gt;","\\"":"&quot;"}[c]; }); }

function namesOf(w){
  var n = {};
  w.fund.forEach(function(b){ b.rows.forEach(function(r){ if (r[1]) n[r[1]] = 1; }); });
  return Object.keys(n).sort();
}

function renderWeeks(){
  document.getElementById("weeks").innerHTML = DATA.map(function(w,i){
    return '<button class="item wk' + (i === W ? " on" : "") + '" data-i="' + i + '"'
      + ' onclick="pickWeek(' + i + ')">'
      + esc(w.date) + (i === 0 ? '　最新' : '') + '</button>';
  }).join("");
  document.getElementById("topdate").textContent = DATA[W] ? DATA[W].date : "";
}

function pickWeek(i){ W = i; FILTER = "全部品种"; renderWeeks(); renderAll(); }

function renderTiles(w){
  var n = namesOf(w).length;
  var b = (w.pools.bull || []).length, s = (w.pools.bear || []).length;
  var st = w.strat.reduce(function(a,x){ return a + x.rows.length; }, 0);
  var t = [
    ["t1","覆盖品种",n,"本周跟踪",""],
    ["t2","多头池",b,"看多品种","#009688"],
    ["t3","空头池",s,"看空品种","#bf242a"],
    ["t4","策略超市",st,"在跟策略","#0a7d63"]
  ];
  document.getElementById("strip").innerHTML = t.map(function(x){
    return '<div class="tile ' + x[0] + '"><div class="h">' + x[1] + '</div><div class="b">'
      + '<div class="v"' + (x[4] ? ' style="color:' + x[4] + '"' : '') + '>' + x[2] + '</div>'
      + '<div class="s">' + x[3] + '</div></div></div>';
  }).join("");
  var set = function(id, v){ var e = document.getElementById(id); if (e) e.textContent = v; };
  set("c1", n); set("c2", b + s); set("c3", st);
  set("t1", n); set("t2", b + s); set("t3", st);
  set("hdate", w.date);
}


function renderPills(w){
  var names = namesOf(w).slice();
  if (FILTER !== "全部品种" && names.indexOf(FILTER) < 0) names.push(FILTER);
  document.getElementById("pills").innerHTML =
    '<button class="pill' + (FILTER === "全部品种" ? " on" : "") + '" data-name="全部品种">全部品种</button>'
    + names.map(function(n){
        return '<button class="pill' + (FILTER === n ? " on" : "") + '" data-name="' + esc(n) + '">' + esc(n) + '</button>';
      }).join("");
}

function setFilter(name){
  FILTER = name; renderPills(DATA[W]); renderSection(); scrollPillIntoView();
}

function scrollPillIntoView(){
  var bar = document.getElementById("pills");
  var on = bar && bar.querySelector(".pill.on");
  if (on) bar.scrollLeft = Math.max(0, on.offsetLeft - 80);
}

function jumpOpinion(name){
  FILTER = name;
  renderPills(DATA[W]);
  show("p1");
  renderSection();
  scrollPillIntoView();
  var c = document.querySelector(".col-main");
  if (c) c.scrollTop = 0;
  window.scrollTo({top:0, behavior:"smooth"});
}

function renderFund(w){
  var out = "";
  w.fund.forEach(function(b){
    var rows = b.rows.filter(function(r){ return FILTER === "全部品种" || r[1] === FILTER; });
    if (!rows.length) return;
    var cards = rows.map(function(r){
      return '<div class="kcard" data-name="' + esc(r[1]) + '"><div class="khead">'
        + '<span class="kname">' + esc(r[1]) + '</span>'
        + '<span class="stars">' + esc(r[2]) + '</span>'
        + (r[0] ? '<span class="kboard">' + esc(r[0]) + '</span>' : '')
        + '</div>' + (r[3] ? '<p class="kview">' + esc(r[3]) + '</p>' : '') + '</div>';
    }).join("");
    var note = (b.note && FILTER === "全部品种") ? '<p class="note">' + esc(b.note) + '</p>' : "";
    out += '<div class="card"><div class="card-h"><h2>' + esc(b.title) + '</h2><span class="sub">'
      + rows.length + ' 个品种</span></div><div class="card-b"><div class="grid3">'
      + cards + '</div>' + note + '</div></div>';
  });
  return out || '<div class="card"><div class="card-b"><p class="note">该品种本周暂无观点记录。</p></div></div>';
}

function renderPools(w){
  function pool(key, label, items){
    return '<div class="pool ' + key + '"><div class="ph"><span>' + label + '</span><span>'
      + items.length + '</span></div><div class="pb">'
      + items.map(function(it){ return '<button class="chip" data-jump="' + esc(it[0]) + '" title="查看该品种观点">'
        + esc(it[0]) + '<i>' + esc(it[1]) + '</i></button>'; }).join("")
      + '</div></div>';
  }
  var bull = w.pools.bull || [], bear = w.pools.bear || [];
  if (!bull.length && !bear.length) return '<div class="card"><div class="card-b"><p class="note">本期多空池暂无内容。</p></div></div>';
  return '<div class="card"><div class="card-h"><h2>多空池</h2><span class="sub">多头 ' + bull.length
    + ' · 空头 ' + bear.length + ' · 点击品种可查看观点</span></div><div class="card-b"><div class="pools">'
    + pool("bull", "多头品种池", bull) + pool("bear", "空头品种池", bear)
    + '</div></div></div>';
}

function renderStrat(w){
  if (!w.strat.length) return '<div class="card"><div class="card-b"><p class="note">本期交易策略暂无内容。</p></div></div>';
  return w.strat.map(function(s){
    var cards = s.rows.map(function(r){
      return '<div class="sc"><div class="sc-top"><span class="sc-name">' + esc(r.name) + '</span>'
        + (r.tag ? '<span class="sc-tag">' + esc(r.tag) + '</span>' : '') + '</div>'
        + '<div class="sc-row"><b>操作要点</b><span>' + esc(r.op) + '</span></div>'
        + '<div class="sc-row goal"><b>目标 / 风险</b><span>' + esc(r.goal) + '</span></div></div>';
    }).join("");
    return '<div class="card"><div class="card-h"><h2>' + esc(s.title) + '</h2></div><div class="card-b">'
      + (s.desc ? '<p class="desc">' + esc(s.desc) + '</p>' : "") + cards + '</div></div>';
  }).join("");
}

function renderAccount(){
  return '<div class="card"><div class="card-h"><h2>账户运行情况</h2><span class="sub">本期待补充</span></div>'
    + '<div class="card-b"><div class="skel">'
    + '<div class="sk">—</div><div class="sk">—</div><div class="sk">—</div><div class="sk">—</div>'
    + '</div></div></div>';
}

function renderRail(w){
  var bull = w.pools.bull || [], bear = w.pools.bear || [];
  var all = bull.map(function(x){ return [x[0], x[1], "b"]; })
    .concat(bear.map(function(x){ return [x[0], x[1], "s"]; }));
  document.getElementById("railcnt").textContent = "共 " + all.length;
  document.getElementById("railList").innerHTML = all.map(function(p){
    return '<div class="rj ' + p[2] + '" data-jump="' + esc(p[0]) + '" title="查看该品种观点"><b>'
      + esc(p[0]) + '</b><i>' + esc(p[1]) + '</i></div>';
  }).join("") || '<p class="note">本期暂无</p>';
}

function renderSection(){
  var w = DATA[W];
  document.getElementById("p1").innerHTML = renderFund(w);
  document.getElementById("p2").innerHTML = renderPools(w);
  document.getElementById("p3").innerHTML = renderStrat(w);
  document.getElementById("p4").innerHTML = renderAccount();
}

function renderAll(){
  var w = DATA[W];
  if (!w) return;
  renderTiles(w); renderPills(w); renderSection(); renderRail(w);
}

document.querySelectorAll(".tab").forEach(function(t){
  t.addEventListener("click", function(){ show(t.dataset.p); });
});
document.querySelectorAll(".item[data-p]").forEach(function(t){
  t.addEventListener("click", function(){ show(t.dataset.p); });
});
document.getElementById("weeks").addEventListener("click", function(e){
  var b = e.target.closest(".wk"); if (b) pickWeek(Number(b.dataset.i));
});
document.getElementById("pills").addEventListener("click", function(e){
  var b = e.target.closest(".pill"); if (b) setFilter(b.dataset.name);
});
document.addEventListener("click", function(e){
  var j = e.target.closest("[data-jump]");
  if (j) { jumpOpinion(j.dataset.jump); }
});
function show(id){
  SEC = id;
  document.querySelectorAll(".panel").forEach(function(p){ p.classList.toggle("on", p.id === id); });
  document.querySelectorAll(".tab").forEach(function(t){ t.classList.toggle("on", t.dataset.p === id); });
  document.querySelectorAll(".item[data-p]").forEach(function(t){ t.classList.toggle("on", t.dataset.p === id); });
  document.querySelector(".side").classList.remove("open");
}

function filterByName(q){
  q = (q || "").trim();
  if (q) show("p1");
  if (!q) { setFilter("全部品种"); return; }
  var w = DATA[W], hit = namesOf(w).filter(function(n){ return n.indexOf(q) >= 0; })[0];
  if (hit) setFilter(hit);
}

(function(){ renderWeeks(); renderAll(); show("p1"); })();
</script>
</body>
</html>
"""

if __name__ == "__main__":
    sys.exit(main())
