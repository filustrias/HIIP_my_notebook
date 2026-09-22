"""
Generates `phc-investment-actions.qmd` from
PHC_investment_actions_narrative_structure_EN.docx.

The document is highly regular: every domain repeats the same blocks
(narrative, possible actions, actions to avoid, synergy, cost, horizon), and
the evidence for all domains comes at the end. The script parses that
structure and lays each block out to use the width of the screen: options
side by side, the horizon as a timeline, the synergy as a lever grid.

Usage:
    python scripts/generate_investment_actions.py

The synergy diagrams in the .docx are images. Which levers each action
*activates* is also in the text (the line under each diagram) and is parsed
from there. Which lever the action *belongs to* appears only in the images,
so it is recorded by hand in OWN_LEVER below: check it when a domain is
added or renumbered.
"""

import html
import re
import shutil
import subprocess
import tempfile
from pathlib import Path

import docx
from docx.text.paragraph import Paragraph

ROOT = Path(__file__).resolve().parent.parent
DOCX = ROOT / "PHC_investment_actions_narrative_structure_EN.docx"
OUT = ROOT / "phc-investment-actions.qmd"

# Read from the synergy diagrams (dark tile = the action's own lever).
OWN_LEVER = {1: 3, 2: 5, 3: 12, 4: 6, 6: 11}

# WHO/UNICEF Operational Framework for PHC, names as on the diagrams.
LEVERS = {
    1: "Political commitment", 2: "Governance", 3: "Funding", 4: "Communities",
    5: "Models of care", 6: "PHC workforce", 7: "Physical infrastructure",
    8: "Medicines", 9: "Private sector", 10: "Purchasing & payment",
    11: "Digital technologies", 12: "Quality", 13: "Research", 14: "Monitoring",
}

MARKERS = ("Situation", "Complication", "Reason for sequencing", "Answer")
SECTION_LABELS = (
    "Possible actions", "Actions to avoid", "Synergy with the other levers",
    "Order of magnitude of cost", "What we expect to see",
)
NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


# ------------------------------------------------------------------- read
def open_doc(path):
    """Open the .docx even while Word holds an exclusive lock on it."""
    try:
        return docx.Document(path)
    except PermissionError:
        pass
    tmp = Path(tempfile.gettempdir()) / f"_hiip_{path.name}"
    try:
        shutil.copy2(path, tmp)
    except PermissionError:
        subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             f"Copy-Item -LiteralPath '{path}' -Destination '{tmp}' -Force"],
            check=True, capture_output=True,
        )
    print("  (document open in another program; read a copy)")
    return docx.Document(tmp)


def segments(p):
    """Runs merged into (text, bold, italic) segments."""
    out = []
    for r in p.runs:
        if not r.text:
            continue
        f = (bool(r.bold), bool(r.italic))
        if out and out[-1][1:] == f:
            out[-1] = (out[-1][0] + r.text, *f)
        else:
            out.append((r.text, *f))
    return out


def read_blocks(d):
    """Flatten the body into dicts: kind, text, segs, style, image."""
    blocks = []
    for ch in d.element.body.iterchildren():
        if ch.tag.split("}")[-1] != "p":
            continue
        p = Paragraph(ch, d)
        has_img = bool(ch.findall(f".//{NS_A}blip"))
        text = p.text.strip()
        if not text and not has_img:
            continue
        segs = segments(p)
        visible = [s for s in segs if s[0].strip()]
        blocks.append({
            "text": text,
            "segs": segs,
            "style": p.style.name if p.style is not None else "",
            "bold": bool(visible) and all(s[1] for s in visible),
            "italic": bool(visible) and all(s[2] for s in visible),
            "image": has_img,
        })
    return blocks


# ----------------------------------------------------------------- render
def inline(segs):
    """Segments -> HTML, with markers and {gaps} styled."""
    h = "".join(
        f"<strong>{html.escape(t)}</strong>" if b else html.escape(t)
        for t, b, _ in segs
    )
    h = re.sub(
        r"<strong>\s*\[(" + "|".join(MARKERS) + r")\]\s*</strong>",
        r' <span class="pia-mark">\1</span>', h)
    h = re.sub(r"\{([^}]*)\}", r' <span class="pia-gap">\1</span>', h)
    return re.sub(r"\s+", " ", h).strip()


def gaps(text):
    return re.sub(r"\{([^}]*)\}",
                  lambda m: f'<span class="pia-gap">{html.escape(m.group(1))}</span>',
                  html.escape(text))


def split_label(segs):
    """'**Label — **rest' or '**Label. **rest' -> (label, rest segments)."""
    if segs and segs[0][1]:
        label = segs[0][0].strip().rstrip("—").rstrip(".").strip()
        return label, segs[1:]
    return "", segs


def raw(s):
    return "```{=html}\n" + s + "\n```\n"


# ----------------------------------------------------------------- parse
def parse(blocks):
    doc = {"title": "", "subtitle": [], "intro": [], "blocks9": [],
           "intro_after": [], "domains": [], "evidence_intro": "",
           "evidence": {}, "studies": [], "note": []}
    i, n = 0, len(blocks)

    doc["title"] = blocks[0]["text"]
    doc["subtitle"] = [s.strip() for s in blocks[1]["text"].split("\n") if s.strip()]
    i = 2

    mode, dom, sec, ev_dom = "intro", None, None, None
    while i < n:
        b = blocks[i]
        t = b["text"]
        m = re.match(r"Domain (\d+) — (.+)$", t) if b["bold"] else None

        if b["bold"] and t == "How these texts are organised":
            mode = "intro"
        elif b["bold"] and t == "Evidence":
            mode, dom = "evidence", None
        elif b["bold"] and t == "The diagnostic studies":
            mode = "studies"
        elif b["bold"] and t.startswith("Note on the nature"):
            mode = "note"
        elif m and mode == "evidence":
            ev_dom = int(m.group(1))
            doc["evidence"][ev_dom] = {"title": m.group(2), "basis": "", "studies": []}
        elif m:
            mode = "domain"
            dom = {"n": int(m.group(1)), "title": m.group(2), "narrative": None,
                   "options": [], "prior": "", "avoid": [], "synergy": [],
                   "cost": None, "horizon": []}
            doc["domains"].append(dom)
            sec = "narrative"
        elif mode == "domain" and b["bold"] and t in SECTION_LABELS:
            sec = t
        elif mode == "intro":
            if b["style"] == "List Bullet":
                doc["blocks9"].append(split_label(b["segs"]))
            elif doc["blocks9"]:
                doc["intro_after"].append(b["segs"])
            else:
                doc["intro"].append(b["segs"])
        elif mode == "domain":
            if sec == "narrative":
                dom["narrative"] = b["segs"]
            elif sec == "Possible actions":
                if b["style"] == "List Bullet":
                    dom["options"].append(split_label(b["segs"]))
                else:
                    dom["prior"] = t
            elif sec == "Actions to avoid":
                dom["avoid"].append(t)
            elif sec == "Synergy with the other levers" and not b["image"]:
                for part in t.split("·"):
                    mm = re.match(r"\s*(\d+)\s+(.+?)\s*$", part)
                    if mm:
                        dom["synergy"].append((int(mm.group(1)), mm.group(2)))
            elif sec == "Order of magnitude of cost":
                dom["cost"] = b["segs"]
            elif sec == "What we expect to see":
                dom["horizon"].append(split_label(b["segs"]))
        elif mode == "evidence":
            if ev_dom is None:
                doc["evidence_intro"] = t
            elif t.startswith("Diagnostic studies needed"):
                segs = b["segs"]
                # drop the leading bold "Diagnostic studies needed."
                segs = [s for s in segs]
                if segs and segs[0][0].strip().startswith("Diagnostic studies needed"):
                    head = segs[0][0].split("Diagnostic studies needed.", 1)[1]
                    segs = ([(head, True, False)] if head.strip() else []) + segs[1:]
                name = None
                for txt, bold, _ in segs:
                    if bold and txt.strip():
                        name = txt.strip()
                    elif name:
                        desc = txt.strip().lstrip("—").strip().rstrip(";").rstrip(".").strip()
                        doc["evidence"][ev_dom]["studies"].append((name, desc))
                        name = None
            else:
                doc["evidence"][ev_dom]["basis"] = t
        elif mode == "studies":
            doc["studies"].append(t)
        elif mode == "note":
            doc["note"].append(t)
        i += 1
    return doc


# ---------------------------------------------------------------- output
def lever_grid(dom):
    own = OWN_LEVER.get(dom["n"])
    act = dict(dom["synergy"])
    rows = [("Strategic", [1, 2, 3, 4]), ("Operational", [5, 6, 7, 8, 9]),
            ("", [10, 11, 12, 13, 14])]
    h = ['<div class="lever-legend">'
         '<span><i class="sw own"></i>Lever of the action</span>'
         '<span><i class="sw act"></i>Levers activated</span>'
         '<span><i class="sw none"></i>No direct effect</span></div>',
         '<div class="lever-grid">']
    for r, (label, nums) in enumerate(rows, start=1):
        if label:  # "Operational" spans the two operational rows, as on the diagram
            span = 2 if label == "Operational" else 1
            h.append(f'<div class="lever-rowlabel" style="grid-row:{r}/span {span}">{label}</div>')
        for k, num in enumerate(nums):
            start = 2 + k * 2 + (1 if len(nums) == 4 else 0)
            kind = "own" if num == own else ("act" if num in act else "none")
            body = ""
            if kind == "own":
                body = '<span class="lever-note">This action</span>'
            elif kind == "act":
                body = f'<span class="lever-note">{html.escape(act[num])}</span>'
            h.append(
                f'<div class="lever {kind}" style="grid-row:{r};grid-column:{start}/span 2">'
                f'<span class="lever-n">{num}</span>'
                f'<span class="lever-name">{html.escape(LEVERS[num])}</span>{body}</div>')
    h.append("</div>")
    return "\n".join(h)


def domain_section(dom):
    n = dom["n"]
    own = OWN_LEVER.get(n)
    plain = "".join(s[0] for s in dom["narrative"])
    ans = re.search(r"\[Reason for sequencing\](.*?)\[Answer\]", plain, re.S)
    answer = ans.group(1).strip().lstrip(".;, ").rstrip(".") + "." if ans else ""

    o = [f"## Domain {n} — {dom['title']} {{#domain-{n}}}\n"]

    o.append(raw(
        '<div class="pia-lead">'
        f'<div class="pia-narrative"><p>{inline(dom["narrative"])}</p></div>'
        '<div class="pia-answer">'
        + (f'<div class="pia-lever-tag">Lever {own} · {html.escape(LEVERS[own])}</div>' if own else "")
        + '<div class="pia-kicker">The investment action</div>'
        f'<p>{gaps(answer)}</p>'
        f'<a class="pia-evlink" href="#evidence-{n}">Evidence and studies needed ↓</a>'
        '</div></div>'))

    cards = "".join(
        f'<div class="pia-card"><div class="pia-card-n">{k}</div>'
        f'<div class="pia-card-title">{html.escape(lbl)}</div>'
        f'<p>{inline(rest)}</p></div>'
        for k, (lbl, rest) in enumerate(dom["options"], start=1))
    prior = (f'<div class="pia-prior"><span>Prior step</span>{gaps(dom["prior"])}</div>'
             if dom["prior"] else "")
    o.append(raw(
        '<div class="pia-block"><div class="pia-label">Possible actions '
        '<em>orientations, not a recommendation · in order of ambition →</em></div>'
        f'<div class="pia-options">{cards}</div>{prior}</div>'))

    avoid = "".join(f"<li>{gaps(a)}</li>" for a in dom["avoid"])
    o.append(raw(
        '<div class="pia-two">'
        '<div class="pia-block"><div class="pia-label">Actions to avoid</div>'
        f'<ul class="pia-avoid">{avoid}</ul></div>'
        '<div class="pia-block"><div class="pia-label">Order of magnitude of cost</div>'
        f'<p class="pia-cost">{inline(dom["cost"])}</p></div></div>'))

    o.append(raw(
        '<div class="pia-block"><div class="pia-label">Synergy with the other levers</div>'
        + lever_grid(dom) + "</div>"))

    steps = "".join(
        f'<div class="pia-step"><div class="pia-when">{html.escape(lbl)}</div>'
        f'<p>{inline(rest)}</p></div>'
        for lbl, rest in dom["horizon"])
    o.append(raw(
        '<div class="pia-block"><div class="pia-label">What we expect to see '
        '<em>a horizon, not monitoring indicators</em></div>'
        f'<div class="pia-timeline">{steps}</div></div>'))
    return "\n".join(o)


def evidence_section(n, ev):
    items = re.findall(r"\s*(.+?)\s*\{([^}]*)\}\.?", ev["basis"])
    rows = "".join(
        f'<li><span class="pia-ev-item">{html.escape(it)}</span>'
        f'<span class="pia-status{" read" if st.startswith("read") else ""}">'
        f'{html.escape(st)}</span></li>'
        for it, st in items)
    studies = "".join(
        f'<div class="pia-study"><div class="pia-study-name">{html.escape(nm)}</div>'
        f'<p>{html.escape(ds)}</p></div>'
        for nm, ds in ev["studies"])
    return (f"### Domain {n} — {ev['title']} {{#evidence-{n}}}\n\n" + raw(
        '<div class="pia-two pia-ev">'
        '<div class="pia-block"><div class="pia-label">Basis</div>'
        f'<ul class="pia-ev-list">{rows}</ul></div>'
        '<div class="pia-block"><div class="pia-label">Diagnostic studies needed</div>'
        f'<div class="pia-studies">{studies}</div></div></div>')
        + f'\n[↑ Back to Domain {n}](#domain-{n})\n')


def main():
    d = open_doc(DOCX)
    blocks = read_blocks(d)
    n_img = sum(1 for b in blocks if b["image"])
    doc = parse(blocks)

    missing = [x["n"] for x in doc["domains"] if x["n"] not in OWN_LEVER]
    if missing:
        print(f"  WARNING: no OWN_LEVER for domains {missing}; check the diagrams")
    if n_img != len(doc["domains"]):
        print(f"  WARNING: {n_img} images for {len(doc['domains'])} domains")

    o = ["---", f'title: "{doc["title"]}"']
    if doc["subtitle"]:
        o.append(f'subtitle: "{" · ".join(doc["subtitle"])}"')
    o += ["format:", "  html:", "    grid:", "      body-width: 1150px", "---", ""]

    o.append("## How these texts are organised\n")
    for segs in doc["intro"]:
        o.append(raw(f"<p>{inline(segs)}</p>"))
    nine = "".join(
        f'<div class="pia-nine-item{" pia-nine-para" if k <= 4 else ""}">'
        f'<div class="pia-nine-n">{k}</div><div><strong>{html.escape(lbl)}</strong>'
        f'<p>{inline(rest)}</p></div></div>'
        for k, (lbl, rest) in enumerate(doc["blocks9"], start=1))
    o.append(raw(f'<div class="pia-nine">{nine}</div>'
                 '<p class="pia-nine-key"><span class="pia-nine-sw"></span>'
                 'The first four form one continuous paragraph on each page.</p>'))
    for segs in doc["intro_after"]:
        o.append(raw(f"<p>{inline(segs)}</p>"))

    for dom in doc["domains"]:
        o.append(domain_section(dom))

    o.append("## Evidence {#evidence}\n")
    if doc["evidence_intro"]:
        o.append(raw(f"<p>{gaps(doc['evidence_intro'])}</p>"))
    for n, ev in doc["evidence"].items():
        o.append(evidence_section(n, ev))

    if doc["studies"]:
        o.append("## The diagnostic studies\n")
        for t in doc["studies"]:
            o.append(t + "\n")

    if doc["note"]:
        o.append("## Note on the nature of the missing information\n")
        text = " ".join(doc["note"])
        sents = re.split(r"(?<=\.)\s+(?=[A-Z])", text)
        if len(sents) >= 4:
            o.append(sents[0] + "\n")
            kinds = "".join(
                f'<div class="pia-card"><div class="pia-card-n">{k}</div><p>{gaps(s)}</p></div>'
                for k, s in enumerate(sents[1:], start=1))
            o.append(raw(f'<div class="pia-options">{kinds}</div>'))
        else:
            o.append(text + "\n")

    OUT.write_text("\n".join(o), encoding="utf-8")
    print(f"Written: {OUT}")
    print(f"  {len(doc['domains'])} domains "
          f"({', '.join(str(x['n']) for x in doc['domains'])}), "
          f"{sum(len(x['options']) for x in doc['domains'])} options, "
          f"{sum(len(x['avoid']) for x in doc['domains'])} actions to avoid, "
          f"{sum(len(x['synergy']) for x in doc['domains'])} synergies, "
          f"{sum(len(e['studies']) for e in doc['evidence'].values())} study mentions")


if __name__ == "__main__":
    main()
