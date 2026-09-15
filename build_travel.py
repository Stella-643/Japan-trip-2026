# -*- coding: utf-8 -*-
"""生成旅行安排 HTML v4（旅行手帐 · 暖色调）:
- 统一单色线性 SVG 图标（景点=鸟居 / 交通=列车 / 住宿=床 / 餐饮=刀叉 / 出发=飞机）
- Maps / 备注 / 备选餐厅 也用统一 SVG；地图标记改为纯色圆点（不堆图标）
- 倒计时实时计算，放进「今日重点」
- 路线连线改蓝色实线
- 新增 总览 tab（全程地图按天着色 + 每日关键）
- 新增 酒店·机票 tab（可填信息 + 传截图，存浏览器本地）
- 内联 Leaflet，免 CDN
"""
import urllib.parse, json, re, datetime
import openpyxl

LEAFLET_CSS = open("/workspace/assets/leaflet.css", encoding="utf-8").read()
LEAFLET_JS = open("/workspace/assets/leaflet.js", encoding="utf-8").read()

COLOR = {"morning": "#8A79C6", "afternoon": "#6D8FC9", "night": "#3F3566"}
DAY_COLOR = {"d1": "#59478C", "d2": "#7E6BC0", "d3": "#6D8FC9", "d4": "#8A79C6",
             "d5": "#3F3566", "d6": "#9C6FC0", "d7": "#5C86C9", "d8": "#7A5EA0", "d9": "#4E6FB0"}

ICONS = {
    "stay": '<path d="M3 18v-6a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v6"/><path d="M3 14h18"/><path d="M3 18v2"/><path d="M21 18v2"/><path d="M7 10V8a1 1 0 0 1 1-1h3a1 1 0 0 1 1 1v2"/>',
    "transit": '<rect x="5" y="3" width="14" height="13" rx="3"/><path d="M5 11h14"/><circle cx="9" cy="14" r="1"/><circle cx="15" cy="14" r="1"/><path d="M8 16l-2 4"/><path d="M16 16l2 4"/>',
    "sight": '<path d="M4 6h16"/><path d="M5 6v2"/><path d="M19 6v2"/><path d="M6.5 9v11"/><path d="M17.5 9v11"/><path d="M10 9h4"/>',
    "food": '<path d="M7 3v7a2 2 0 0 0 4 0V3"/><path d="M9 3v18"/><path d="M17 3c-1.6 0-2.5 2-2.5 5s.9 4 2.5 4v9"/>',
    "flight": '<path d="M21 16v-2l-8-5V3.5a1.5 1.5 0 0 0-3 0V9l-8 5v2l8-2.5V19l-2 1.5V22l3.5-1 3.5 1v-1.5L13 19v-5.5z"/>',
    "maps": '<path d="M12 21s-7-7.5-7-12a7 7 0 0 1 14 0c0 4.5-7 12-7 12z"/><circle cx="12" cy="9" r="2.5"/>',
    "note": '<path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/>',
    "fork": '<path d="M7 3v7a2 2 0 0 0 4 0V3"/><path d="M9 3v18"/><path d="M17 3c-1.6 0-2.5 2-2.5 5s.9 4 2.5 4v9"/>',
    "clock": '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/>',
    "globe": '<circle cx="12" cy="12" r="9.5"/><path d="M2.5 12h19"/><path d="M12 2.5c2.6 2.6 4 6 4 9.5s-1.4 6.9-4 9.5c-2.6-2.6-4-6-4-9.5s1.4-6.9 4-9.5z"/>',
    "calendar": '<rect x="3" y="4.5" width="18" height="16" rx="3"/><path d="M3 9.5h18"/><path d="M8 2.5v4"/><path d="M16 2.5v4"/>',
    "route": '<path d="M9 4 3 6.2v13.6l6-2.2 6 2.2 6-2.2V4l-6 2.2z"/><path d="M9 4v13.6"/><path d="M15 6.2v13.6"/>',
    "spark": '<path d="M12 3.4l2.3 5.7 5.7 2.3-5.7 2.3L12 19.4l-2.3-5.7L4 11.4l5.7-2.3z"/>',
    "check": '<rect x="3" y="3" width="18" height="18" rx="4.5"/><path d="M8 12.3l2.6 2.7L16 9"/>',
    "camera": '<path d="M4 8h3l2-2.5h6L17 8h3a1 1 0 0 1 1 1v9a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1V9a1 1 0 0 1 1-1z"/><circle cx="12" cy="13" r="3.1"/>',
    "ticket": '<path d="M3.5 8.5h17v2.2a1.8 1.8 0 0 0 0 3.6v2.2h-17v-2.2a1.8 1.8 0 0 0 0-3.6z"/><path d="M14 8.5v8"/>',
    "bag": '<rect x="5" y="7" width="14" height="13.5" rx="3"/><path d="M9 7V5a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2"/><path d="M9.5 11.5v5"/><path d="M14.5 11.5v5"/>',
    "shop": '<path d="M4 9h16l1.5-4.4L18 4H6L2.5 4.6z"/><path d="M4 9v10h16V9"/><path d="M9.5 19v-6h5v6"/>',
}


def svg(key):
    return (f'<svg class="isc" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
            f'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">{ICONS[key]}</svg>')


def maps_url(place):
    return "https://www.google.com/maps/search/?api=1&query=" + urllib.parse.quote(place)


def render_item(it, period):
    icon = it.get("icon", "sight")
    cls = "entry" + (" key" if it.get("key") else "")
    time = f'<div class="time">{it["time"]}</div>' if it.get("time") else ""
    maps = ""
    if it.get("place"):
        maps = (f'<a class="maps" href="{maps_url(it["place"])}" target="_blank" rel="noopener">'
                f'{svg("maps")}<span>Maps</span></a>')
    title = (f'<div class="title"><span class="ic" style="color:{COLOR[period]}">{svg(icon)}</span>'
             f'{it["title"]} {maps}</div>')
    desc = f'<div class="desc">{it["desc"]}</div>' if it.get("desc") else ""
    note = f'<div class="note">{svg("note")}<span>{it["note"]}</span></div>' if it.get("note") else ""
    tag = (f'<div class="tagrow"><span class="tag">{it["tag"]}</span></div>' if it.get("tag") else "")
    backup = ""
    if it.get("backup"):
        chips = ""
        for b in it["backup"]:
            if isinstance(b, dict):
                name, price = b.get("name", ""), b.get("price")
            else:
                name, price = b, None
            pr = f'<span class="bk-price">{price}</span>' if price else ""
            chips += (f'<a class="bk" href="{maps_url(name)}" target="_blank" rel="noopener">'
                      f'{name}{svg("maps")}{pr}</a>')
        backup = (f'<div class="backup"><span class="bk-label">{svg("fork")}<span>备选餐厅</span></span>'
                  f'<div class="bk-list">{chips}</div></div>')
    dot = f'<span class="edot" style="--c:{COLOR[period]}"></span>'
    return (f'<div class="{cls}">{dot}<div class="ebody">{time}{title}{desc}'
            f'{note}{tag}{backup}</div></div>')


def day_markers(sections_data):
    out, seen = [], set()
    for period, items in sections_data:
        for it in items:
            if not it.get("coord"):
                continue
            key = (round(it["coord"][0], 5), round(it["coord"][1], 5))
            if key in seen:
                continue
            seen.add(key)
            out.append({"lat": it["coord"][0], "lng": it["coord"][1],
                        "name": it["title"], "url": maps_url(it["place"]),
                        "color": COLOR[period]})
    return out


def render_lines(sections_data):
    parts = []
    for period, items in sections_data:
        label = "早上" if period == "morning" else ("下午" if period == "afternoon" else "晚上")
        parts.append(f'<div class="phead {period}"><span class="pdot" style="--c:{COLOR[period]}"></span>{label}</div>')
        for it in items:
            parts.append(render_item(it, period))
    return f'<div class="lines">{"".join(parts)}</div>'


def countdown_html(iso, time, name, fallback):
    """time 有值 → 倒计时到「最近的具体行程」（精确到分）；
    无 time（占位日）→ 退化为倒计时到当天 00:00。"""
    if time:
        return (f'<div class="cd">{svg("clock")}'
                f'<span class="cd-text" data-datetime="{iso}T{time}" data-name="{name}">计算中…</span></div>')
    return (f'<div class="cd">{svg("clock")}'
            f'<span class="cd-text" data-date="{iso}" data-label="{fallback}">计算中…</span></div>')


def render_focus(focus, cd_html):
    if not focus and not cd_html:
        return ""
    items = "".join(
        f'<div class="focus-item"><span class="fd"></span><span class="ftime">{t}</span><span class="fname">{n}</span></div>'
        for t, n in focus
    )
    return f'<div class="card focus">{cd_html}<div class="focus-title">{svg("spark")}<span>今日重点</span></div>{items}</div>'


def render_panel(pid, info):
    head = (f'<div class="dayhead"><div class="dh-label">{info["label"]}</div>'
            f'<div class="dh-date">{svg("calendar")}<span>{info["date"]} · {info["wd"]}</span></div></div>')
    cd_html = info.get("cd_html", "")
    if not info.get("sections"):
        inner = (f'<div class="card placeholder">{cd_html}'
                 f'<div class="ph-emoji">{svg("bag")}</div>'
                 f'<div class="ph-title">{info["label"]} · 行程待补充</div>'
                 f'<div class="ph-sub">日期：{info["date"]} {info["wd"]}<br>参照 Excel 模板填写后即可扩展。</div></div>')
        return f'<section class="panel" id="{pid}">{head}{inner}</section>'
    mapcard = ""
    if pid in map_data:
        mapcard = (f'<div class="card mapwrap"><div class="mapcard-title">{svg("route")}<span>今日路线</span></div>'
                   f'<div class="map" id="map-{pid}"><div class="mfb">地图加载中…<br>若长时间空白，请用系统自带浏览器（Safari / Chrome）打开本页</div></div></div>')
    return f'<section class="panel" id="{pid}">{head}{mapcard}{render_focus(info.get("focus"), cd_html)}{render_lines(info["sections"])}</section>'


# ======================= 数据（从 Excel 模板读取 D1–D9） =======================
EXCEL_PATH = "/workspace/Japan_Trip_D1-D9_Template.xlsx"

TYPE_ICON = {"✈️": "flight", "🏨": "stay", "🍽️": "food", "🏛️": "sight",
             "🚅": "transit", "☕": "food", "🛍️": "shop", "🎟️": "ticket",
             "🚶": "sight", "🌃": "sight"}
PERIOD = {"☀️早上": "morning", "🌤️下午": "afternoon", "🌙晚上": "night"}
WD = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]
THEME = {"d1": "大阪到达日", "d2": "京都一日游", "d3": "大阪·梅田", "d4": "USJ 环球影城",
         "d5": "大阪→东京转移", "d6": "东京·浅草银座", "d7": "富士山一日游",
         "d8": "东京·上野", "d9": "返程回沪"}
# 地点关键词 → 经纬度（命中关键词即打点；无则不显示该点）
COORDS = {
    "关西机场": (34.4347, 135.2440), "北滨": (34.6920, 135.5060),
    "心斋桥": (34.6687, 135.5017), "道顿堀": (34.6682, 135.5012),
    "伏见稻荷": (34.9671, 135.7727), "清水寺": (34.9949, 135.7850),
    "三年坂": (34.9965, 135.7825), "高台寺": (34.9995, 135.7815),
    "うなぎ": (35.0037, 135.7755), "八坂神社": (35.0037, 135.7784),
    "祇园": (35.0038, 135.7755), "鸭川": (35.0085, 135.7716),
    "祇园四条": (35.0038, 135.7720), "梅田": (34.7053, 135.4983),
    "JOJO": (34.7053, 135.4983), "Fountain": (34.7053, 135.4983),
    "Universal": (34.6654, 135.4322), "USJ": (34.6654, 135.4322),
    "新大阪": (34.7335, 135.5002), "茅场町": (35.6810, 139.7780),
    "涩谷": (35.6595, 139.6995), "浅草寺": (35.7148, 139.7967),
    "秋叶原": (35.6995, 139.7717), "银座": (35.6717, 139.7650),
    "富士山": (35.3606, 138.7274), "西洋美术馆": (35.7171, 139.7731),
    "city walk": (35.7171, 139.7731), "东京": (35.6762, 139.6503),
}
# 当天无任何坐标时的城市兜底点（避免地图空白）
CITY_FALLBACK = {"d1": (34.6937, 135.5023), "d2": (34.9949, 135.7850), "d3": (34.7053, 135.4983),
                 "d4": (34.6654, 135.4322), "d5": (35.6580, 139.6963), "d6": (35.7148, 139.7967),
                 "d7": (35.3606, 138.7274), "d8": (35.7171, 139.7731), "d9": (35.6762, 139.6503)}


def coord_for(title):
    for kw, c in COORDS.items():
        if kw in title:
            return list(c)
    return None


def parse_backup(cell):
    """备选餐厅：支持 `店名 ¥人均` 形式，用 / 或 、 分隔多家。"""
    if not cell or not str(cell).strip():
        return []
    out = []
    for p in re.split(r"[/、]", str(cell)):
        p = p.strip()
        if not p:
            continue
        m = re.search(r"¥\s*([\d,]+)", p)
        price = "¥" + m.group(1) if m else None
        name = re.sub(r"¥\s*[\d,]+", "", p).strip()
        if name:
            out.append({"name": name, "price": price})
    return out


def fmt_time(v):
    if v is None:
        return None
    if isinstance(v, datetime.time):
        return v.strftime("%H:%M")
    s = str(v).strip()
    m = re.search(r"(\d{1,2})[:：](\d{2})", s)
    return f"{int(m.group(1)):02d}:{m.group(2)}" if m else (s or None)


def build_days(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["D1-D9行程"]
    out = {}
    for r in ws.iter_rows(values_only=True):
        if not isinstance(r[0], datetime.datetime):
            continue  # 跳过表头
        if not any(c is not None and str(c).strip() != "" for c in r[:5]):
            continue
        date_v, dayid, period_v, type_v, activity = r[0], r[1], r[2], r[3], r[4]
        ktime, percap, note, gmaps, key_v, cd_v, backup = r[5], r[6], r[7], r[8], r[9], r[10], r[11]
        pid = str(dayid).strip().lower()
        if pid not in out:
            iso = date_v.strftime("%Y-%m-%d")
            wd = WD[date_v.weekday()]
            out[pid] = {"label": f"{dayid}｜{THEME.get(pid, '')}", "date": iso.replace("-", "."),
                        "wd": wd, "iso": iso,
                        "cdlabel": f"{dayid}（{iso[5:].replace('-', '/')} {wd[:2]}）",
                        "sections": {"morning": [], "afternoon": [], "night": []}, "focus": []}
        period = PERIOD.get(str(period_v).strip(), "morning")
        icon = TYPE_ICON.get(str(type_v).strip(), "sight")
        t = fmt_time(ktime)
        item = {"time": t, "icon": icon, "title": str(activity).strip(),
                "place": str(activity).strip(), "key": str(key_v).strip() == "是"}
        coord = coord_for(str(activity))
        if coord:
            item["coord"] = coord
        if note and str(note).strip():
            item["note"] = str(note).strip()
        if percap is not None and str(percap).strip() not in ("", "None"):
            try:
                item["tag"] = f"¥{int(float(percap)):,}"
            except (ValueError, TypeError):
                pass
        bk = parse_backup(backup)
        if bk:
            item["backup"] = bk
        out[pid]["sections"][period].append(item)
        if item["key"] and t:
            out[pid]["focus"].append((t, str(activity).strip()))
    for d in out.values():
        d["sections"] = [(p, d["sections"][p]) for p in ("morning", "afternoon", "night") if d["sections"][p]]
    return out

# 各天元信息（由 Excel 模板驱动生成）
days = build_days(EXCEL_PATH)


def first_event(info):
    """当天最早、且带时间的具体行程，用于「最近行程倒计时」。"""
    best = None
    for period, items in info.get("sections") or []:
        for it in (items or []):
            if it.get("time") and (best is None or it["time"] < best[0]):
                best = (it["time"], it["title"])
    return best


for pid, info in days.items():
    fe = first_event(info)
    info["cd_html"] = countdown_html(info["iso"], fe[0], fe[1], info["cdlabel"]) if fe \
        else countdown_html(info["iso"], None, None, info["cdlabel"])

# ======================= 总览 =======================
def overview_markers():
    out = []
    for pid, info in days.items():
        for period, items in info.get("sections") or []:
            for it in items:
                if not it.get("coord"):
                    continue
                out.append({"lat": it["coord"][0], "lng": it["coord"][1],
                            "name": it["title"], "url": maps_url(it["place"]),
                            "color": DAY_COLOR[pid]})
    return out


ov_items = ""
for pid, info in days.items():
    dot = f'<span class="ov-dot" style="background:{DAY_COLOR[pid]}"></span>'
    ov_items += (f'<div class="ov-item">{dot}'
                 f'<span class="ov-date">{info["date"]} {info["wd"][:2]}</span>'
                 f'<span class="ov-label">{info["label"]}</span></div>')

overview_panel = (
    f'<section class="panel" id="overview">'
    f'<div class="dayhead"><div class="dh-label">{svg("globe")}<span>全程总览</span></div>'
    f'<div class="dh-date">{svg("calendar")}<span>2026.10.31 – 11.08 · 9天8晚</span></div></div>'
    f'<div class="card mapwrap"><div class="mapcard-title">{svg("route")}<span>全程路线（按天着色）</span></div>'
    f'<div class="map" id="map-overview"><div class="mfb">地图加载中…<br>若长时间空白，请用系统自带浏览器（Safari / Chrome）打开本页</div></div></div>'
    f'<div class="card"><div class="focus-title">{svg("calendar")}<span>每日关键</span></div><div class="ov-list">{ov_items}</div></div>'
    f'</section>'
)

# ======================= 地图数据（各天 + 总览） =======================
map_data = {"overview": {"markers": overview_markers()}}
for _pid, _info in days.items():
    _mk = day_markers(_info["sections"])
    if not _mk and _pid in CITY_FALLBACK:
        _mk = [{"lat": CITY_FALLBACK[_pid][0], "lng": CITY_FALLBACK[_pid][1],
                "name": THEME.get(_pid, _pid), "url": maps_url(THEME.get(_pid, _pid)),
                "color": COLOR["morning"]}]
    if _mk:
        map_data[_pid] = {"markers": _mk}

# ======================= 酒店·机票 =======================
hotel_panel = (
    f'<section class="panel" id="hotel">'
    f'  <div class="dayhead"><div class="dh-label">{svg("stay")}<span>酒店 · 机票</span></div>'
    f'  <div class="dh-date">{svg("calendar")}<span>信息仅保存在本机浏览器</span></div></div>'
    f'  <div class="card">'
    f'    <div class="doc-title">{svg("stay")}<span>大阪酒店</span></div>'
    f'    <div class="field"><label>名称</label><input data-store="osaka_name" placeholder="如 相铁FRESA INN 北滨"></div>'
    f'    <div class="field"><label>房号</label><input data-store="osaka_room" placeholder="如 1208"></div>'
    f'    <div class="field"><label>地址</label><input data-store="osaka_addr" placeholder="日文/英文地址"></div>'
    f'    <div class="field"><label>电话</label><input data-store="osaka_tel" placeholder="如 +81-6-xxxx-xxxx"></div>'
    f'    <div class="field"><label>入住</label><input data-store="osaka_in" placeholder="如 10/31"></div>'
    f'    <div class="field"><label>退房</label><input data-store="osaka_out" placeholder="如 11/02"></div>'
    f'    <div class="upload"><label class="up-btn">{svg("camera")}<span>上传酒店截图</span><input type="file" accept="image/*" id="osakaImg" hidden></label><div class="imgprev" id="osakaImgPrev"></div></div>'
    f'  </div>'
    f'  <div class="card">'
    f'    <div class="doc-title">{svg("stay")}<span>东京酒店</span></div>'
    f'    <div class="field"><label>名称</label><input data-store="tokyo_name" placeholder="如 东京某酒店"></div>'
    f'    <div class="field"><label>房号</label><input data-store="tokyo_room" placeholder="如 0930"></div>'
    f'    <div class="field"><label>地址</label><input data-store="tokyo_addr" placeholder="日文/英文地址"></div>'
    f'    <div class="field"><label>电话</label><input data-store="tokyo_tel" placeholder="如 +81-3-xxxx-xxxx"></div>'
    f'    <div class="field"><label>入住</label><input data-store="tokyo_in" placeholder="如 11/02"></div>'
    f'    <div class="field"><label>退房</label><input data-store="tokyo_out" placeholder="如 11/08"></div>'
    f'    <div class="upload"><label class="up-btn">{svg("camera")}<span>上传酒店截图</span><input type="file" accept="image/*" id="tokyoImg" hidden></label><div class="imgprev" id="tokyoImgPrev"></div></div>'
    f'  </div>'
    f'  <div class="card">'
    f'    <div class="doc-title">{svg("flight")}<span>去程机票</span></div>'
    f'    <div class="field"><label>航班号</label><input data-store="go_no" placeholder="如 NH 956"></div>'
    f'    <div class="field"><label>出发</label><input data-store="go_from" placeholder="如 上海 PVG"></div>'
    f'    <div class="field"><label>到达</label><input data-store="go_to" placeholder="如 大阪 KIX"></div>'
    f'    <div class="field"><label>日期</label><input data-store="go_date" placeholder="如 10/31"></div>'
    f'    <div class="field"><label>时间</label><input data-store="go_time" placeholder="如 09:20"></div>'
    f'    <div class="upload"><label class="up-btn">{svg("camera")}<span>上传机票截图</span><input type="file" accept="image/*" id="goImg" hidden></label><div class="imgprev" id="goImgPrev"></div></div>'
    f'  </div>'
    f'  <div class="card">'
    f'    <div class="doc-title">{svg("flight")}<span>回程机票</span></div>'
    f'    <div class="field"><label>航班号</label><input data-store="back_no" placeholder="如 NH 955"></div>'
    f'    <div class="field"><label>出发</label><input data-store="back_from" placeholder="如 东京 HND"></div>'
    f'    <div class="field"><label>到达</label><input data-store="back_to" placeholder="如 上海 PVG"></div>'
    f'    <div class="field"><label>日期</label><input data-store="back_date" placeholder="如 11/08"></div>'
    f'    <div class="field"><label>时间</label><input data-store="back_time" placeholder="如 18:05"></div>'
    f'    <div class="upload"><label class="up-btn">{svg("camera")}<span>上传机票截图</span><input type="file" accept="image/*" id="backImg" hidden></label><div class="imgprev" id="backImgPrev"></div></div>'
    f'  </div>'
    f'  <div class="doc-tip">提示：内容仅存于本机浏览器（不上传），清缓存会丢失；建议另外截图备份关键证件。</div>'
    f'</section>'
)

# ======================= Checklist（Excel 真实数据） =======================
checklist_groups_raw = [
    ("bag", "出发前", ["护照", "签证", "信用卡", "Suica", "漫游/eSIM", "日元现金", "境外保险", "身份证",
                  "一只黑笔", "转换插头", "信用卡 visa或者master", "湿厕纸", "拖鞋/洞洞鞋", "纸巾",
                  "小垃圾袋", "晕车药", "盖章册子", "一次性雨衣", "充电宝"]),
    ("ticket", "预约项目", ["うなぎ四代目菊川", "USJ", "富士山一日游", "Shibuya Sky"]),
]
chk_html = ""
chk_index = 0
for gicon, g, items in checklist_groups_raw:
    lis = ""
    for i in items:
        key = "c" + str(chk_index)
        chk_index += 1
        lis += f'<li><label><input type="checkbox" data-key="{key}"><span>{i}</span></label></li>'
    chk_html += (f'<div class="card chk-group"><div class="chk-gtitle">{svg(gicon)}<span>{g}</span></div>'
                 f'<ul class="check">{lis}</ul></div>')
chk_html = (f'<div class="chk-head"><div class="chk-count" id="chkCount">已完成 0 / {chk_index}</div>'
            f'<button class="chk-reset" id="chkReset">清除全部</button></div>' + chk_html)
checklist_panel = (f'<section class="panel" id="checklist">'
                   f'<div class="dayhead"><div class="dh-label">{svg("check")}<span>出行 Checklist</span></div>'
                   f'<div class="dh-date">{svg("calendar")}<span>出发前逐项核对</span></div></div>{chk_html}</section>')

# ============ 组装 ============
tab_defs = [("overview", "总览", "globe"), ("d1", "D1", None), ("d2", "D2", None), ("d3", "D3", None),
            ("d4", "D4", None), ("d5", "D5", None), ("d6", "D6", None), ("d7", "D7", None),
            ("d8", "D8", None), ("d9", "D9", None),
            ("hotel", "酒店·机票", "stay"), ("checklist", "Checklist", "check")]
tabs = "".join(
    f'<button class="tab{" active" if pid == "overview" else ""}" data-day="{pid}">'
    f'{(svg(ic) if ic else "")}<span>{name}</span></button>'
    for pid, name, ic in tab_defs
)

panels = [overview_panel]
for pid, info in days.items():
    panels.append(render_panel(pid, info))
panels.append(hotel_panel)
panels.append(checklist_panel)
panels_html = "".join(panels).replace(
    '<section class="panel" id="overview">', '<section class="panel active" id="overview">', 1)

# map_data 已在上方「地图数据」段统一构建
map_json = json.dumps(map_data, ensure_ascii=False)

CSS = """
:root{--bg:#F7F6FB;--card:#FFFFFF;--text:#33304A;--sub:#8A8795;--accent:#59478C;--accent-soft:#EAE5F5;--morning:#8A79C6;--afternoon:#6D8FC9;--night:#3F3566;--blue:#7E6BC0;--hair:#E9E6F2}
*{box-sizing:border-box}
html,body{overflow-x:hidden}
body{margin:0;background:var(--bg);color:var(--text);font-family:"PingFang SC",-apple-system,"Segoe UI",sans-serif;line-height:1.6}
.container{max-width:880px;margin:0 auto;padding:0 18px 60px}
.topbar{padding:26px 0 14px;text-align:center}
.topbar h1{margin:0;font-size:25px;letter-spacing:.5px}
.trip-sub{margin-top:6px;color:var(--sub);font-size:13px}
.tabs{position:sticky;top:0;z-index:20;display:flex;flex-wrap:wrap;gap:8px;padding:12px 0;background:var(--bg);border-bottom:1px solid var(--hair);margin-bottom:8px}
.tab{display:inline-flex;align-items:center;gap:6px;flex:0 0 auto;border:1px solid var(--hair);background:#fff;color:var(--text);font-size:15px;font-weight:600;padding:8px 15px;border-radius:999px;cursor:pointer;box-shadow:0 2px 6px rgba(0,0,0,.05);white-space:nowrap}
.tab .isc{width:16px;height:16px}
.tab.active{background:var(--accent);color:#fff;border-color:var(--accent);box-shadow:0 4px 12px rgba(89,71,140,.3)}
.panel{display:none;padding-top:8px}
.panel.active{display:block}
.card{background:var(--card);border:1px solid var(--hair);border-radius:22px;padding:20px;margin-bottom:18px;box-shadow:0 6px 18px rgba(90,70,40,.05)}
.dayhead{margin:14px 0 14px}
.dh-label{display:flex;align-items:center;gap:9px;font-size:22px;font-weight:800;letter-spacing:.3px}
.dh-label .isc{width:23px;height:23px;color:var(--accent)}
.dh-date{margin-top:6px;color:var(--sub);font-size:14px;display:inline-flex;align-items:center;gap:6px}
.dh-date .isc{width:15px;height:15px}
.focus-title{display:flex;align-items:center;gap:8px;font-size:18px;font-weight:800;margin-bottom:10px}
.focus-title .isc{width:19px;height:19px;color:var(--accent)}
.focus{background:linear-gradient(135deg,#EFEAF8,#fff)}
.focus-item{display:flex;align-items:center;gap:12px;padding:7px 0}
.fd{width:9px;height:9px;border-radius:50%;background:var(--accent);flex:0 0 auto}
.ftime{color:var(--accent);font-weight:800;font-size:14px;min-width:54px}
.fname{flex:1;font-size:15px}
.cd{display:inline-flex;align-items:center;gap:6px;background:#EFEAF8;color:var(--blue);font-weight:700;font-size:13px;padding:5px 12px;border-radius:999px;margin-bottom:12px}
.lines{position:relative;padding-left:34px;margin-top:6px}
.lines:before{content:'';position:absolute;left:7px;top:10px;bottom:14px;width:2px;background:#ECE7F5}
.phead{position:relative;margin:26px 0 12px;font-weight:800;font-size:17px}
.phead:first-child{margin-top:4px}
.pdot{position:absolute;left:-31px;top:6px;width:12px;height:12px;border-radius:50%;background:var(--c);border:2.5px solid #fff;box-shadow:0 0 0 1.5px var(--c)}
.entry{position:relative;padding-bottom:20px}
.edot{position:absolute;left:-31px;top:7px;width:11px;height:11px;border-radius:50%;background:var(--c);border:2px solid #fff;box-shadow:0 0 0 1.5px rgba(0,0,0,.06)}
.entry.key .edot{width:13px;height:13px;left:-32px;top:6px;box-shadow:0 0 0 3px rgba(89,71,140,.22)}
.ebody{background:#fff;border:1px solid var(--hair);border-radius:16px;padding:12px 16px;box-shadow:0 3px 10px rgba(90,70,40,.04)}
.time{color:var(--accent);font-weight:800;font-size:14px}
.ic{display:inline-flex;vertical-align:-4px;margin-right:3px}
.title{font-size:19px;font-weight:700;line-height:1.5;margin-top:2px}
.isc{width:18px;height:18px;flex:0 0 auto}
.title .isc{width:20px;height:20px}
.desc{color:var(--sub);margin-top:6px;font-size:14px;line-height:1.7}
.note{margin-top:8px;font-size:13.5px;color:#6E6A80;line-height:1.6;display:flex;align-items:flex-start;gap:5px}
.note .isc{width:14px;height:14px;margin-top:2px}
.tagrow{margin-top:8px}
.tag{display:inline-block;background:var(--accent-soft);color:#4A3A78;padding:4px 12px;border-radius:999px;font-size:13px;font-weight:600}
.backup{margin-top:10px;background:#F1EDFA;border-radius:12px;padding:9px 12px}
.bk-label{font-weight:700;color:#56478A;display:inline-flex;align-items:center;gap:5px}
.bk-label .isc{width:15px;height:15px}
.bk-list{margin-top:6px;line-height:2}
.bk{display:inline-flex;align-items:center;gap:4px;background:#fff;border:1px solid #DED5F0;border-radius:999px;padding:3px 10px;margin:2px 8px 2px 0;color:#56478A;text-decoration:none;font-size:13px}
.bk:hover{background:#F6F3FC}
.bk .isc{width:13px;height:13px}
.bk-price{color:#7E6BC0;font-weight:700;font-size:12px;margin-left:3px}
.maps{display:inline-flex;align-items:center;gap:3px;margin-left:8px;font-size:12.5px;font-weight:700;color:var(--blue);background:#EFEAF8;padding:2px 10px;border-radius:999px;text-decoration:none;vertical-align:middle;white-space:nowrap}
.maps:hover{background:#E4DCF5}
.maps .isc{width:13px;height:13px}
.mapcard-title{display:flex;align-items:center;gap:8px;font-size:15px;font-weight:800;color:var(--sub);margin-bottom:10px}
.mapcard-title .isc{width:17px;height:17px;color:var(--accent)}
.map{position:relative;height:360px;border-radius:18px;overflow:hidden;background:#EDE9F4;z-index:0}
.map .leaflet-container{border-radius:18px;font:inherit}
.mfb{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;text-align:center;color:#9A94AC;font-size:13px;line-height:1.9;padding:24px}
.map-tip{position:absolute;left:10px;right:10px;bottom:10px;z-index:600;background:rgba(63,53,102,.92);color:#fff;font-size:12.5px;line-height:1.6;padding:9px 13px;border-radius:11px;box-shadow:0 4px 14px rgba(0,0,0,.2)}
.map-tip a{color:#D9CCFF;font-weight:700}
.njs{max-width:840px;margin:0 0 14px;background:#FFF4E5;border:1px solid #F0D8B0;color:#8A6A2F;font-size:13.5px;line-height:1.7;padding:12px 16px;border-radius:14px}
.njs b{color:#7A5A1F}
.placeholder{text-align:center;padding:46px 20px}
.ph-emoji{color:#C6BFDF;display:flex;justify-content:center}
.ph-emoji .isc{width:50px;height:50px;stroke-width:1.4}
.ph-title{font-size:20px;font-weight:800;margin-top:12px}
.ph-sub{margin-top:10px;color:var(--sub);line-height:1.8;font-size:14px}
.ov-list{padding-top:4px}
.ov-item{display:flex;align-items:center;gap:10px;padding:8px 0}
.ov-dot{width:11px;height:11px;border-radius:50%;flex:0 0 auto}
.ov-date{color:var(--sub);font-size:13px;font-weight:700;min-width:80px}
.ov-label{font-size:15px;flex:1}
.ov-focus{margin:0 0 6px 90px;color:#6E6A80;font-size:13px;line-height:1.7}
.doc-title{font-size:17px;font-weight:800;margin-bottom:10px;display:flex;align-items:center;gap:8px}
.doc-title .isc{width:19px;height:19px;color:var(--accent)}
.field{display:flex;align-items:center;gap:10px;padding:7px 0;border-bottom:1px dashed var(--hair)}
.field label{width:56px;color:var(--sub);font-size:14px;flex:0 0 auto}
.field input{flex:1;border:none;background:#F5F3FA;border-radius:10px;padding:9px 12px;font-size:15px;color:var(--text);font-family:inherit;min-width:0}
.field input:focus{outline:2px solid var(--accent-soft)}
.upload{margin-top:12px}
.up-btn{display:inline-flex;align-items:center;gap:7px;background:var(--accent-soft);color:#4A3A78;font-weight:700;font-size:14px;padding:8px 16px;border-radius:999px;cursor:pointer}
.up-btn .isc{width:16px;height:16px}
.up-btn:hover{background:#DED5F0}
.imgprev{margin-top:10px}
.imgprev img{max-width:100%;border-radius:14px;border:1px solid var(--hair);display:block}
.doc-tip{color:var(--sub);font-size:13px;line-height:1.7;margin-top:6px}
.chk-head{display:flex;justify-content:space-between;align-items:center;margin:14px 0 16px}
.chk-count{font-size:14px;color:var(--accent);font-weight:700}
.chk-reset{font-size:13px;color:var(--sub);background:none;border:1px solid var(--hair);border-radius:999px;padding:5px 12px;cursor:pointer}
.chk-reset:hover{color:var(--text)}
.chk-group{margin-bottom:14px}
.chk-gtitle{display:flex;align-items:center;gap:8px;font-size:16px;font-weight:800;margin-bottom:8px}
.chk-gtitle .isc{width:18px;height:18px;color:var(--accent)}
.check{list-style:none;padding:0;margin:0}
.check li{padding:4px 0}
.check label{display:flex;align-items:center;gap:12px;cursor:pointer;font-size:16px;padding:6px 4px;border-radius:10px}
.check label:hover{background:#F1EDFA}
.check input{width:20px;height:20px;accent-color:var(--accent);flex:0 0 auto}
.check label.done{color:#9A96A8;text-decoration:line-through}
.pin2{width:16px;height:16px;border-radius:50%;background:var(--c,#59478C);border:3px solid #fff;box-shadow:0 1px 4px rgba(0,0,0,.45)}
.leaflet-popup-content{font-family:"PingFang SC",-apple-system,sans-serif;font-size:13px}
.leaflet-popup-content a{color:var(--blue);font-weight:700;text-decoration:none;display:inline-flex;align-items:center;gap:4px}
.leaflet-popup-content a svg{width:14px;height:14px}
.leaflet-control-layers-toggle{background-image:url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='20' height='20' viewBox='0 0 24 24' fill='none' stroke='%2359478C' stroke-width='2' stroke-linejoin='round'%3E%3Cpolygon points='12 2 2 7 12 12 22 7 12 2'/%3E%3Cpolyline points='2 17 12 22 22 17'/%3E%3Cpolyline points='2 12 12 17 22 12'/%3E%3C/svg%3E");background-size:20px 20px;background-position:center;background-repeat:no-repeat}
.leaflet-control-zoom a{color:#59478C;font-weight:700}
"""

JS = """
var MAPS = __MAPDATA__;
var LOC = '<svg viewBox="0 0 24 24" fill="none" stroke="#7E6BC0" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M12 21s-7-7.5-7-12a7 7 0 0 1 14 0c0 4.5-7 12-7 12z"/><circle cx="12" cy="9" r="2.5"/></svg>';
var _inited = {};

/* WGS-84 → GCJ-02（高德/国内瓦片纠偏，标记才不会偏移） */
function wgs2gcj(lat, lng){
  function outOfChina(la,ln){ return (ln<72.004||ln>137.8347||la<0.8293||la>55.8271); }
  function tLat(x,y){var r=-100+2*x+3*y+0.2*y*y+0.1*x*y+0.2*Math.sqrt(Math.abs(x));
    r+=(20*Math.sin(6*x*Math.PI)+20*Math.sin(2*x*Math.PI))*2/3;
    r+=(20*Math.sin(y*Math.PI)+40*Math.sin(y/3*Math.PI))*2/3;
    r+=(160*Math.sin(y/12*Math.PI)+320*Math.sin(y*Math.PI/30))*2/3;return r;}
  function tLng(x,y){var r=300+x+2*y+0.1*x*x+0.1*x*y+0.1*Math.sqrt(Math.abs(x));
    r+=(20*Math.sin(6*x*Math.PI)+20*Math.sin(2*x*Math.PI))*2/3;
    r+=(20*Math.sin(x*Math.PI)+40*Math.sin(x/3*Math.PI))*2/3;
    r+=(150*Math.sin(x/12*Math.PI)+300*Math.sin(x/30*Math.PI))*2/3;return r;}
  if(outOfChina(lat,lng)) return [lat,lng];
  var a=6378245.0, ee=0.00669342162296594323;
  var dLat=tLat(lng-105,lat-35), dLng=tLng(lng-105,lat-35);
  var rad=lat/180*Math.PI, m=Math.sin(rad); m=1-ee*m*m; var sm=Math.sqrt(m);
  dLat=(dLat*180)/((a*(1-ee))/(m*sm)*Math.PI);
  dLng=(dLng*180)/(a/sm*Math.cos(rad)*Math.PI);
  return [lat+dLat, lng+dLng];
}

/* 底图源：第 1 个为默认；加载失败时按顺序自动切换 */
var BASES = [
  {label:'街道图', gcj:false, url:'https://server.arcgisonline.com/ArcGIS/rest/services/World_Street_Map/MapServer/tile/{z}/{y}/{x}', maxZoom:19, attr:'Tiles &copy; Esri'},
  {label:'高德(国内)', gcj:true, url:'https://webrd0{s}.is.autonavi.com/appmaptile?lang=zh_cn&size=1&scale=1&style=8&x={x}&y={y}&z={z}', subdomains:'1234', maxZoom:18, attr:'&copy; 高德地图'},
  {label:'OpenStreetMap', gcj:false, url:'https://tile.openstreetmap.de/{z}/{x}/{y}.png', maxZoom:19, attr:'&copy; OpenStreetMap'},
  {label:'影像图', gcj:false, url:'https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', maxZoom:19, attr:'Tiles &copy; Esri'}
];

function initMap(id){
  var cfg = MAPS[id];
  if(!cfg){ return; }
  if(_inited[id]){ _inited[id].invalidateSize(); return; }
  var el = document.getElementById('map-'+id);
  if(!el || el._leaflet_id){ return; }
  el.innerHTML = '';

  var map = L.map(el,{scrollWheelZoom:false});
  var layerByName = {}, defByName = {};
  BASES.forEach(function(d){
    var opts = {maxZoom:d.maxZoom, attribution:d.attr};
    if(d.subdomains) opts.subdomains = d.subdomains;
    layerByName[d.label] = L.tileLayer(d.url, opts);
    defByName[d.label] = d;
  });
  var curLabel = BASES[0].label;
  layerByName[curLabel].addTo(map);

  var ov = L.layerGroup().addTo(map);
  function drawOverlay(){
    ov.clearLayers();
    var gcj = defByName[curLabel].gcj, pts = [];
    cfg.markers.forEach(function(m){
      var la = m.lat, ln = m.lng;
      if(gcj){ var c = wgs2gcj(la, ln); la = c[0]; ln = c[1]; }
      var icon = L.divIcon({className:'', html:'<div class="pin2" style="--c:'+m.color+'"></div>',
        iconSize:[16,16], iconAnchor:[8,8], popupAnchor:[0,-12]});
      L.marker([la,ln],{icon:icon}).addTo(ov)
        .bindPopup('<b>'+m.name+'</b><br><a href="'+m.url+'" target="_blank" rel="noopener">'+LOC+' Maps</a>');
      pts.push([la,ln]);
    });
    if(pts.length){
      L.polyline(pts,{color:'#7E6BC0',weight:4,opacity:.9,lineCap:'round',lineJoin:'round'}).addTo(ov);
      map.fitBounds(pts,{padding:[46,46],maxZoom:16});
    }
  }
  drawOverlay();

  L.control.layers(layerByName, null, {position:'topright',collapsed:true}).addTo(map);
  map.on('baselayerchange', function(e){ curLabel = e.name; drawOverlay(); });

  function showTip(html){
    var tip = el.querySelector('.map-tip');
    if(!tip){ tip = document.createElement('div'); tip.className = 'map-tip'; el.appendChild(tip); }
    tip.innerHTML = html;
  }
  var order = BASES.map(function(d){ return d.label; });
  var failed = {}, errCount = 0, done = false;
  function onBaseErr(){
    if(done) return;
    errCount++;
    if(errCount < 6) return;
    var i = order.indexOf(curLabel), nxt = null;
    for(var k=1;k<order.length;k++){ var c = order[(i+k)%order.length]; if(!failed[c]){ nxt = c; break; } }
    if(nxt){
      failed[curLabel] = true;
      map.removeLayer(layerByName[curLabel]);
      layerByName[nxt].addTo(map);
      curLabel = nxt; errCount = 0; drawOverlay();
      showTip('当前网络加载底图失败，已自动切换到「' + nxt + '」');
      setTimeout(function(){ var t = el.querySelector('.map-tip'); if(t) t.remove(); }, 6000);
    } else {
      done = true;
      showTip('地图瓦片加载失败（可能是网络限制）。请改用系统自带浏览器（Safari / Chrome）打开本页，或点右上角图层图标手动切换底图。');
    }
  }
  order.forEach(function(name){
    layerByName[name].on('tileerror', function(){ if(name === curLabel) onBaseErr(); });
  });

  _inited[id] = map;
}
document.querySelectorAll('.tab').forEach(function(t){
  t.addEventListener('click', function(){
    document.querySelectorAll('.tab').forEach(function(x){x.classList.remove('active');});
    document.querySelectorAll('.panel').forEach(function(x){x.classList.remove('active');});
    t.classList.add('active');
    var d = t.getAttribute('data-day');
    var p = document.getElementById(d);
    if(p) p.classList.add('active');
    window.scrollTo({top:0,behavior:'smooth'});
    initMap(d);
    if(_inited[d]){ setTimeout(function(){ _inited[d].invalidateSize(); },60); }
  });
});
function paintCountdowns(){
  var now = new Date();
  document.querySelectorAll('.cd-text').forEach(function(el){
    var dt = el.getAttribute('data-datetime');
    if(dt){
      var dp = dt.split('T');
      var d0 = dp[0].split('-'), t0 = dp[1].split(':');
      var d = new Date(+d0[0], +d0[1]-1, +d0[2], +t0[0], +t0[1], 0);
      var name = el.getAttribute('data-name') || '';
      var ms = d - now;
      if(ms <= 0){ el.textContent = name + ' 已开始'; return; }
      var days = Math.floor(ms/86400000);
      var hours = Math.floor((ms%86400000)/3600000);
      var mins = Math.floor((ms%3600000)/60000);
      var txt = '距 ' + name + ' 还有 ';
      if(days > 0) txt += days + ' 天 ' + hours + ' 小时';
      else if(hours > 0) txt += hours + ' 小时 ' + mins + ' 分';
      else txt += mins + ' 分钟';
      el.textContent = txt;
      return;
    }
    var date = el.getAttribute('data-date');
    var q = date.split('-');
    var dd = new Date(+q[0], +q[1]-1, +q[2]); dd.setHours(0,0,0,0);
    var n2 = new Date(); n2.setHours(0,0,0,0);
    var diff = Math.round((dd - n2)/86400000);
    var label = el.getAttribute('data-label') || '';
    var t;
    if(diff > 0) t = '距 ' + label + ' 还有 ' + diff + ' 天';
    else if(diff === 0) t = label + ' 就是今天';
    else t = label + ' 已过去 ' + (-diff) + ' 天';
    el.textContent = t;
  });
}
window.addEventListener('load', function(){ initMap('overview'); });
initMap('overview');
paintCountdowns();
setInterval(paintCountdowns, 60000);

(function(){
  var KEY='jp_checklist_v1';
  var saved={};
  try{ saved=JSON.parse(localStorage.getItem(KEY)||'{}'); }catch(e){}
  var boxes=document.querySelectorAll('.check input[type=checkbox]');
  function refresh(){
    var done=0; boxes.forEach(function(b){ if(b.checked) done++; });
    var c=document.getElementById('chkCount'); if(c){ c.textContent='已完成 '+done+' / '+boxes.length; }
  }
  boxes.forEach(function(b){
    var k=b.getAttribute('data-key');
    if(saved[k]) b.checked=true;
    if(b.checked) b.closest('label').classList.add('done');
    b.addEventListener('change', function(){
      saved[k]=b.checked;
      b.closest('label').classList.toggle('done', b.checked);
      try{ localStorage.setItem(KEY, JSON.stringify(saved)); }catch(e){}
      refresh();
    });
  });
  refresh();
  var rst=document.getElementById('chkReset');
  if(rst){ rst.addEventListener('click', function(){
    boxes.forEach(function(b){ b.checked=false; b.closest('label').classList.remove('done'); saved[b.getAttribute('data-key')]=false; });
    try{ localStorage.setItem(KEY, JSON.stringify(saved)); }catch(e){}
    refresh();
  }); }
})();

(function(){
  var KEY='jp_trip_docs_v1';
  var data={}; try{ data=JSON.parse(localStorage.getItem(KEY)||'{}'); }catch(e){}
  function save(){ try{ localStorage.setItem(KEY, JSON.stringify(data)); }catch(e){} }
  document.querySelectorAll('[data-store]').forEach(function(el){
    var k=el.getAttribute('data-store');
    if(data[k]!==undefined && data[k]!==null) el.value=data[k];
    el.addEventListener('input', function(){ data[k]=el.value; save(); });
  });
  function bindUpload(inpId, prevId, key){
    var inp=document.getElementById(inpId), prev=document.getElementById(prevId);
    if(!inp||!prev) return;
    if(data[key]){ prev.innerHTML='<img src="'+data[key]+'" alt="截图">'; }
    inp.addEventListener('change', function(){
      var f=inp.files[0]; if(!f) return;
      var r=new FileReader();
      r.onload=function(){ data[key]=r.result; save(); prev.innerHTML='<img src="'+data[key]+'" alt="截图">'; };
      r.readAsDataURL(f);
    });
  }
  bindUpload('osakaImg','osakaImgPrev','osaka_img');
  bindUpload('tokyoImg','tokyoImgPrev','tokyo_img');
  bindUpload('goImg','goImgPrev','go_img');
  bindUpload('backImg','backImgPrev','back_img');
})();
"""

DOC = """<!DOCTYPE html><html lang="zh-CN"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>大阪东京9天8晚｜旅行安排</title>
<style>__LEAFLET_CSS__</style>
<style>__CSS__</style>
</head><body><div class="container"><div class="topbar"><h1>大阪东京 9天8晚</h1><div class="trip-sub">2026.10.31 – 11.08 · D1–D9</div></div><noscript><div class="njs"><b>提示：</b>本页需要开启 JavaScript 才能显示地图、倒计时、清单和上传。如果你在 App 的内嵌预览里看到地图空白，请点右上角「分享 / 用浏览器打开」，改用手机自带浏览器（Safari / Chrome）打开本文件。</div></noscript><div class="tabs">__TABS__</div>__PANELS__</div>
<script>__LEAFLET_JS__</script>
<script>__JS__</script>
</body></html>"""

html = (DOC
        .replace("__LEAFLET_CSS__", LEAFLET_CSS)
        .replace("__LEAFLET_JS__", LEAFLET_JS)
        .replace("__CSS__", CSS)
        .replace("__TABS__", tabs)
        .replace("__PANELS__", panels_html)
        .replace("__JS__", JS.replace("__MAPDATA__", map_json)))

with open("/workspace/travel_plan.html", "w", encoding="utf-8") as f:
    f.write(html)
# 部署入口：Cloudflare Pages 默认用根目录的 index.html 当首页
with open("/workspace/index.html", "w", encoding="utf-8") as f:
    f.write(html)
print("OK bytes=", len(html))
print("tabs=", len(tab_defs))
print("d2 markers=", len(map_data["d2"]["markers"]))
print("overview markers=", len(map_data["overview"]["markers"]))
print("backup blocks=", html.count('class="bk-label"'))
print("svg icons=", html.count('class="isc"'))
print("countdown lines=", html.count('class="cd-text"'))
