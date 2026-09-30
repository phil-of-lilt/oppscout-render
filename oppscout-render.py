#!/usr/bin/env python3
"""
OppScout briefing renderer  v1.1.1  (30 Sep 2026)

v1.1.1: every listed item must carry at least one link, at every score, and
every also_screened entry must carry a url; a linkless item is refused rather
than rendered.  One-line items (Worth a look, Watch) now show every link, one
per line, instead of only the first.

v1.1.0: subject prefix "Lilt OppScout" (the Apps Script watches for it), warning
glyph moved after the prefix, persistent item identifiers (18Sep26-A) instead of
integers, optional "logged" block echoing how the previous reply was read,
optional compact list of 1-2 items under "Also screened" for the wide-net period,
plain-language reply footer, no GitHub log link.

Reads a briefing JSON (written by the screening agent) and produces:
  - subject line        (stdout line 1, or --subject)
  - HTML email body     (--html PATH)
  - plain-text body     (--text PATH)

Standard library only. Fails loudly on any schema problem: a bad briefing must
never render into a plausible-looking email.

Usage:
  python3 oppscout-render.py briefing.json --html out.html --text out.txt --subject
"""
import json, sys, html, argparse, datetime as dt

# ---------- palette (Lilt brand) ----------
NAVY   = "#0B1A28"
TEAL   = "#36BBA1"
ORANGE = "#FAA433"
INK    = "#1A2430"
MUTED  = "#6B7683"
HAIR   = "#E3E7EB"
SOFT   = "#F4F6F8"
WHITE  = "#FFFFFF"
GREY   = "#C9CFD6"
TINT_O = "#FFF6E5"
RED    = "#D64545"
GREEN  = "#2E9E6E"
AMBER  = "#E0A100"

FONT = "'DM Sans','Helvetica Neue',Helvetica,Arial,sans-serif"

# ---------- schema ----------
REQUIRED = ["date", "screened", "verdict", "health", "needs_you", "items",
            "also_screened", "radar", "run_window", "run_time_pt"]
HEALTH_STATES = {"healthy", "degraded", "failed"}

def fail(msg):
    sys.stderr.write("BRIEFING INVALID: " + msg + "\n")
    sys.exit(2)

def validate(b):
    for k in REQUIRED:
        if k not in b:
            fail(f"missing field '{k}'")
    try:
        dt.date.fromisoformat(b["date"])
    except Exception:
        fail("date must be YYYY-MM-DD")
    if not isinstance(b["screened"], int) or b["screened"] < 0:
        fail("screened must be a non-negative integer")
    if not isinstance(b["verdict"], str) or not (1 <= len(b["verdict"]) <= 110):
        fail("verdict must be a string of 1-110 characters")
    h = b["health"]
    if h.get("status") not in HEALTH_STATES:
        fail("health.status must be healthy | degraded | failed")
    if not isinstance(h.get("lines", []), list) or len(h.get("lines", [])) > 3:
        fail("health.lines must be a list of at most 3 strings")
    if not isinstance(b["needs_you"], list) or len(b["needs_you"]) > 3:
        fail("needs_you must be a list of at most 3 strings")
    for s in b["needs_you"]:
        if not isinstance(s, str) or len(s) > 200:
            fail("each needs_you entry must be a string of at most 200 characters")
    seen = set()
    for it in b["items"]:
        for k in ["n", "score", "title", "agency", "due"]:
            if k not in it:
                fail(f"item missing '{k}': {it}")
        if not (1 <= it["score"] <= 10):
            fail(f"item {it['n']} score out of range")
        if not isinstance(it["n"], str) or not it["n"].strip():
            fail(f"item id 'n' must be a non-empty string such as 18Sep26-A: {it}")
        if it["n"] in seen:
            fail(f"duplicate item id {it['n']}")
        seen.add(it["n"])
        lk = it.get("links")
        if not isinstance(lk, list) or not lk:
            fail(f"item {it['n']} must carry at least one link (Notice first)")
        for l in lk:
            if not isinstance(l, dict) or not l.get("label") or not str(l.get("url", "")).startswith("https://"):
                fail(f"item {it['n']} has a link without a label or an https url: {l}")
        if it["score"] >= 7:
            for k in ["fit", "do", "links"]:
                if k not in it or not it[k]:
                    fail(f"item {it['n']} scores {it['score']} and must carry '{k}'")
        if it["due"] is not None:
            try:
                dt.date.fromisoformat(it["due"])
            except Exception:
                fail(f"item {it['n']} due must be YYYY-MM-DD or null")
        if len(it["title"]) > 90:
            fail(f"item {it['n']} title exceeds 90 characters")
    a = b["also_screened"]
    if not isinstance(a.get("count"), int):
        fail("also_screened.count must be an integer")
    if any(i["score"] <= 2 for i in b["items"]):
        fail("items scoring 1-2 must not be listed; put them in also_screened")
    extra = a.get("items", [])
    if not isinstance(extra, list) or len(extra) > 15:
        fail("also_screened.items must be a list of at most 15 entries")
    for e in extra:
        for k in ["n", "title"]:
            if k not in e or not e[k]:
                fail(f"also_screened item missing '{k}': {e}")
        if not str(e.get("url", "")).startswith("https://"):
            fail(f"also_screened item {e.get('n')} must carry an https url")
    lg = b.get("logged", [])
    if not isinstance(lg, list) or len(lg) > 6 or any((not isinstance(x, str)) or len(x) > 240 for x in lg):
        fail("logged must be a list of at most 6 strings of at most 240 characters")

# ---------- helpers ----------
def esc(s):
    return html.escape(str(s), quote=True)

def fmt_date(d):
    """Fri 18 Sep 2026"""
    return d.strftime("%a %d %b %Y").replace(" 0", " ")

def fmt_short(d):
    """Fri 18 Sep"""
    return d.strftime("%a %d %b").replace(" 0", " ")

def days_left(run, due):
    return (due - run).days

def due_text(run, due_iso):
    if not due_iso:
        return ("No deadline stated", False)
    due = dt.date.fromisoformat(due_iso)
    n = days_left(run, due)
    if n < 0:
        return (f"Closed {fmt_short(due)}", False)
    if n == 0:
        return (f"Due today, {fmt_short(due)}", True)
    if n == 1:
        return (f"Due tomorrow, {fmt_short(due)}", True)
    return (f"Due {fmt_short(due)}, {n} days", n <= 10)

def band(items, lo, hi):
    return sorted([i for i in items if lo <= i["score"] <= hi], key=lambda i: (-i["score"], i["n"]))

def badge_colors(score):
    if score >= 7:
        return ORANGE, NAVY
    if score >= 5:
        return TEAL, WHITE
    return GREY, INK

def health_dot(status):
    return {"healthy": GREEN, "degraded": AMBER, "failed": RED}[status]

def fmt_int(n):
    return f"{n:,}"

# ---------- subject ----------
def subject_line(b, act, look):
    run = dt.date.fromisoformat(b["date"])
    core = f"Lilt OppScout · {fmt_short(run)} · "
    if b["needs_you"] or b["health"]["status"] != "healthy":
        core += "\u26a0 "
    if act:
        top = act[0]
        extra = len(act) + len(look) - 1
        tail = f" (+{extra})" if extra > 0 else ""
        title = top["title"]
        if len(title) > 48:
            title = title[:45].rstrip() + "..."
        core += f"ACT: {top['score']} {title}{tail}"
    else:
        core += f"all clear · {fmt_int(b['screened'])} screened"
        if look:
            core += f" · {len(look)} worth a look"
    return core

# ---------- HTML pieces ----------
def section_head(label, tag=None):
    tag_html = f'<span style="color:{MUTED};font-weight:400;font-size:12px;padding-left:8px;">{esc(tag)}</span>' if tag else ""
    return f'''
<tr><td style="padding:26px 28px 8px 28px;">
  <div style="font-family:{FONT};font-size:13px;font-weight:700;color:{NAVY};letter-spacing:0.2px;">{esc(label)}{tag_html}</div>
  <div style="height:1px;background:{HAIR};margin-top:8px;"></div>
</td></tr>'''

def empty_line(text):
    return f'''
<tr><td style="padding:0 28px;">
  <div style="font-family:{FONT};font-size:13px;color:{MUTED};padding:6px 0 2px 0;">{esc(text)}</div>
</td></tr>'''

def badge(score, size=36, font=16):
    bg, fg = badge_colors(score)
    return (f'<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>'
            f'<td width="{size}" height="{size}" align="center" valign="middle" bgcolor="{bg}" '
            f'style="width:{size}px;height:{size}px;background:{bg};border-radius:6px;'
            f'font-family:{FONT};font-size:{font}px;font-weight:700;color:{fg};line-height:{size}px;">{score}</td></tr></table>')

def links_html(links):
    parts = []
    for l in links or []:
        parts.append(f'<a href="{esc(l["url"])}" style="color:{TEAL};text-decoration:underline;font-weight:600;">{esc(l["label"])}</a>')
    sep = f'&nbsp;&nbsp;<span style="color:{GREY};">|</span>&nbsp;&nbsp;'
    return sep.join(parts)

def item_card(run, it):
    due_txt, urgent = due_text(run, it.get("due"))
    due_style = f"color:{ORANGE};font-weight:700;" if urgent else f"color:{INK};"
    meta = " · ".join([x for x in [it.get("agency"), it.get("type")] if x])
    updated = ""
    if it.get("updated"):
        updated = f'<div style="font-family:{FONT};font-size:12px;color:{MUTED};padding-top:6px;">Updated: {esc(it["updated"])}</div>'
    return f'''
<tr><td style="padding:10px 28px 0 28px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border:1px solid {HAIR};border-radius:8px;">
<tr>
  <td width="52" valign="top" style="padding:16px 0 16px 14px;">{badge(it["score"])}</td>
  <td valign="top" style="padding:14px 16px 16px 4px;">
    <div style="font-family:{FONT};font-size:16px;font-weight:700;color:{NAVY};line-height:22px;">
      <span style="color:{MUTED};font-weight:400;">{esc(it["n"])}</span>&nbsp; {esc(it["title"])}</div>
    <div style="font-family:{FONT};font-size:13px;color:{MUTED};padding-top:3px;">{esc(meta)}</div>
    <div style="font-family:{FONT};font-size:13px;{due_style}padding-top:6px;">{esc(due_txt)}</div>
    <div style="font-family:{FONT};font-size:14px;color:{INK};line-height:20px;padding-top:10px;">
      <span style="color:{MUTED};">Fit</span>&nbsp;&nbsp;{esc(it["fit"])}</div>
    <div style="font-family:{FONT};font-size:14px;color:{INK};line-height:20px;padding-top:4px;">
      <span style="color:{MUTED};">Do</span>&nbsp;&nbsp;&nbsp;{esc(it["do"])}</div>
    {updated}
    <div style="font-family:{FONT};font-size:13px;padding-top:10px;">{links_html(it.get("links"))}</div>
  </td>
</tr></table>
</td></tr>'''

def item_row(run, it):
    due_txt, urgent = due_text(run, it.get("due"))
    due_style = f"color:{ORANGE};font-weight:700;" if urgent else f"color:{MUTED};"
    link = "<br>".join(
        f'<a href="{esc(l["url"])}" style="color:{TEAL};text-decoration:underline;font-weight:600;">{esc(l["label"])}</a>'
        for l in (it.get("links") or []))
    reason = f'<div style="font-family:{FONT};font-size:12px;color:{MUTED};padding-top:2px;">{esc(it["fit"])}</div>' if it.get("fit") else ""
    return f'''
<tr><td style="padding:8px 28px 0 28px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
<tr>
  <td width="34" valign="top" style="padding-top:1px;">{badge(it["score"], 26, 13)}</td>
  <td valign="top" style="padding-left:6px;">
    <div style="font-family:{FONT};font-size:14px;font-weight:600;color:{NAVY};line-height:19px;">
      <span style="color:{MUTED};font-weight:400;">{esc(it["n"])}</span>&nbsp; {esc(it["title"])}
      <span style="color:{MUTED};font-weight:400;font-size:13px;">&nbsp;&nbsp;{esc(it.get("agency",""))}</span></div>
    {reason}
  </td>
  <td width="150" valign="top" align="right" style="font-family:{FONT};font-size:12px;{due_style}padding-top:2px;white-space:nowrap;">
    {esc(due_txt)}<br>{link}
  </td>
</tr></table>
</td></tr>'''

def tile(n, label, color):
    return f'''<td width="25%" align="left" style="padding:0 8px 0 0;">
  <div style="font-family:{FONT};font-size:26px;font-weight:700;color:{color};line-height:30px;">{esc(n)}</div>
  <div style="font-family:{FONT};font-size:12px;color:{MUTED};padding-top:2px;">{esc(label)}</div>
</td>'''

# ---------- render ----------
def render_html(b):
    run = dt.date.fromisoformat(b["date"])
    items = b["items"]
    act, look, watch = band(items, 7, 10), band(items, 5, 6), band(items, 3, 4)

    pill_bg, pill_text = (ORANGE, "Act today") if act else (TEAL, "All clear")
    pill_fg = NAVY if act else WHITE

    out = [f'''<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="color-scheme" content="light"><meta name="supported-color-schemes" content="light">
<title>OppScout {esc(fmt_date(run))}</title>
</head>
<body style="margin:0;padding:0;background:{SOFT};" bgcolor="{SOFT}">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="{SOFT}" style="background:{SOFT};">
<tr><td align="center" style="padding:24px 12px;">
<table role="presentation" width="600" cellpadding="0" cellspacing="0" border="0" bgcolor="{WHITE}" style="width:600px;max-width:600px;background:{WHITE};border-radius:10px;overflow:hidden;">

<!-- header -->
<tr><td bgcolor="{NAVY}" style="background:{NAVY};padding:16px 28px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
  <td style="font-family:{FONT};font-size:15px;font-weight:700;color:{WHITE};letter-spacing:0.3px;">
    <span style="color:{TEAL};">LILT</span>&nbsp;OppScout</td>
  <td align="right" style="font-family:{FONT};font-size:13px;color:#B7C2CC;">{esc(fmt_date(run))}</td>
</tr></table>
</td></tr>

<!-- verdict -->
<tr><td style="padding:26px 28px 0 28px;">
  <table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>
    <td bgcolor="{pill_bg}" style="background:{pill_bg};border-radius:12px;padding:3px 10px;font-family:{FONT};font-size:11px;font-weight:700;color:{pill_fg};letter-spacing:0.4px;">{pill_text}</td>
  </tr></table>
  <div style="font-family:{FONT};font-size:22px;font-weight:700;color:{NAVY};line-height:28px;padding-top:12px;">{esc(b["verdict"])}</div>
</td></tr>

<!-- tiles -->
<tr><td style="padding:20px 28px 0 28px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"><tr>
{tile(len(act), "Act on", ORANGE if act else NAVY)}
{tile(len(look), "Worth a look", TEAL if look else NAVY)}
{tile(len(watch), "Watch", NAVY)}
{tile(fmt_int(b["screened"]), "Screened", NAVY)}
</tr></table>
</td></tr>''']

    # needs you
    if b["needs_you"]:
        rows = "".join(
            f'<div style="font-family:{FONT};font-size:14px;color:{INK};line-height:20px;padding:4px 0;">'
            f'<span style="color:{ORANGE};font-weight:700;">&#9656;</span>&nbsp; {esc(s)}</div>'
            for s in b["needs_you"])
        out.append(f'''
<tr><td style="padding:22px 28px 0 28px;">
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="{TINT_O}" style="background:{TINT_O};border-left:4px solid {ORANGE};border-radius:0 6px 6px 0;">
<tr><td style="padding:12px 16px 12px 14px;">
  <div style="font-family:{FONT};font-size:13px;font-weight:700;color:{NAVY};padding-bottom:4px;">Needs you ({len(b["needs_you"])})</div>
  {rows}
</td></tr></table>
</td></tr>''')

    # act on
    out.append(section_head("Act on", "scores 7 to 10"))
    if act:
        out.extend(item_card(run, it) for it in act)
    else:
        out.append(empty_line("None today."))

    # worth a look
    out.append(section_head("Worth a look", "scores 5 to 6"))
    if look:
        out.extend(item_row(run, it) for it in look)
    else:
        out.append(empty_line("None today."))

    # watch
    out.append(section_head("Watch", "scores 3 to 4"))
    if watch:
        out.extend(item_row(run, it) for it in watch)
    else:
        out.append(empty_line("None today."))

    # also screened
    a = b["also_screened"]
    also = f'{a["count"]} notice{"s" if a["count"] != 1 else ""} scored 1 to 2'
    if a.get("note"):
        also += f' · {a["note"]}'
    out.append(section_head("Also screened", "scores 1 to 2"))
    out.append(empty_line(also))
    for e in a.get("items", []):
        meta = " · ".join([x for x in [e.get("agency")] if x])
        link = f' <a href="{esc(e["url"])}" style="color:{TEAL};text-decoration:underline;">notice</a>' if e.get("url") else ""
        out.append(f'''
<tr><td style="padding:0 28px;">
  <div style="font-family:{FONT};font-size:12px;color:{MUTED};line-height:17px;padding:2px 0;">
    <span style="color:{GREY};">{esc(e["n"])}</span>&nbsp; {esc(e["title"])}{(" · " + esc(meta)) if meta else ""}{link}</div>
</td></tr>''')

    # radar
    out.append(section_head("Recompete radar and open calls"))
    out.append(empty_line(b["radar"]))

    # logged (how the previous reply was read)
    if b.get("logged"):
        rows = "".join(
            f'<div style="font-family:{FONT};font-size:13px;color:{INK};line-height:19px;padding:2px 0;">{esc(x)}</div>'
            for x in b["logged"])
        out.append(section_head("Logged from your reply"))
        out.append(f'''
<tr><td style="padding:0 28px;">{rows}</td></tr>''')

    # system
    h = b["health"]
    lines = "".join(f'<div style="font-family:{FONT};font-size:13px;color:{INK};line-height:19px;padding:2px 0;">{esc(l)}</div>' for l in h.get("lines", []))
    out.append(f'''
<tr><td style="padding:28px 28px 0 28px;">
  <div style="height:1px;background:{HAIR};"></div>
  <table role="presentation" cellpadding="0" cellspacing="0" border="0" style="margin-top:16px;"><tr>
    <td width="10" valign="middle" style="width:10px;"><div style="width:10px;height:10px;border-radius:5px;background:{health_dot(h["status"])};font-size:0;line-height:0;"></div></td>
    <td style="padding-left:8px;font-family:{FONT};font-size:13px;font-weight:700;color:{NAVY};">System: {esc(h["status"].capitalize())}</td>
  </tr></table>
  <div style="padding-top:6px;">{lines}</div>
</td></tr>''')

    # footer
    log = f' &nbsp;·&nbsp; <a href="{esc(b["log_url"])}" style="color:{TEAL};text-decoration:underline;font-weight:600;">Full log</a>' if b.get("log_url") else ""
    out.append(f'''
<tr><td style="padding:22px 28px 26px 28px;">
  <div style="font-family:{FONT};font-size:12px;color:{MUTED};line-height:18px;">
    Run {esc(b["run_time_pt"])} PT &nbsp;·&nbsp; window {esc(b["run_window"])}{log}</div>
  <div style="font-family:{FONT};font-size:12px;color:{MUTED};line-height:18px;padding-top:4px;">
    Reply in plain language. Refer to items by their ID, for example "18Sep26-A yes" or "18Sep26-C no, staffing".</div>
</td></tr>

</table>
</td></tr></table>
</body></html>''')
    return "\n".join(out)

def render_text(b):
    run = dt.date.fromisoformat(b["date"])
    items = b["items"]
    act, look, watch = band(items, 7, 10), band(items, 5, 6), band(items, 3, 4)
    L = []
    L.append(f"LILT OppScout  {fmt_date(run)}")
    L.append("")
    L.append(("ACT TODAY" if act else "ALL CLEAR") + "  " + b["verdict"])
    L.append(f"Act on {len(act)}  Worth a look {len(look)}  Watch {len(watch)}  Screened {fmt_int(b['screened'])}")
    if b["needs_you"]:
        L += ["", f"Needs you ({len(b['needs_you'])})"] + [f"  > {s}" for s in b["needs_you"]]
    def block(title, lst, full):
        L.append(""); L.append(title)
        if not lst:
            L.append("  None today."); return
        for it in lst:
            due_txt, _ = due_text(run, it.get("due"))
            L.append(f"  [{it['score']}] {it['n']}  {it['title']}  ({it.get('agency','')})  {due_txt}")
            if full:
                L.append(f"      Fit: {it['fit']}")
                L.append(f"      Do:  {it['do']}")
            elif it.get("fit"):
                L.append(f"      {it['fit']}")
            for l in it.get("links") or []:
                L.append(f"      {l['label']}: {l['url']}")
    block("Act on (7-10)", act, True)
    block("Worth a look (5-6)", look, False)
    block("Watch (3-4)", watch, False)
    a = b["also_screened"]
    L += ["", f"Also screened (1-2): {a['count']}" + (f" · {a['note']}" if a.get("note") else "")]
    for e in a.get("items", []):
        L.append(f"  {e['n']}  {e['title']}" + (f"  ({e['agency']})" if e.get("agency") else "") + (f"  {e['url']}" if e.get("url") else ""))
    if b.get("logged"):
        L += ["", "Logged from your reply"] + [f"  {x}" for x in b["logged"]]
    L += ["", "Recompete radar and open calls: " + b["radar"]]
    L += ["", f"System: {b['health']['status'].capitalize()}"] + [f"  {l}" for l in b["health"].get("lines", [])]
    L += ["", f"Run {b['run_time_pt']} PT · window {b['run_window']}" + (f" · Full log: {b['log_url']}" if b.get("log_url") else "")]
    L += ['Reply in plain language. Refer to items by their ID, for example "18Sep26-A yes" or "18Sep26-C no, staffing".']
    return "\n".join(L)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("briefing")
    ap.add_argument("--html")
    ap.add_argument("--text")
    ap.add_argument("--subject", action="store_true")
    a = ap.parse_args()
    with open(a.briefing, encoding="utf-8") as f:
        b = json.load(f)
    validate(b)
    items = b["items"]
    act, look = band(items, 7, 10), band(items, 5, 6)
    subj = subject_line(b, act, look)
    if a.html:
        with open(a.html, "w", encoding="utf-8") as f:
            f.write(render_html(b))
    if a.text:
        with open(a.text, "w", encoding="utf-8") as f:
            f.write(render_text(b))
    if a.subject or not (a.html or a.text):
        print(subj)

if __name__ == "__main__":
    main()
