"""
Generates `planning-and-roadmap.qmd` from HIIP_full_plan_EN.xlsx.

The single source of truth is the "2. Master" sheet (workstream > output >
task). "0. Key" supplies the conventions, "1. Workstreams" the external
validations, "4. Sprints" the calendar and "6. Founding document" the
contractual calendar.

Usage:
    python scripts/generate_roadmap.py

Re-running after editing the workbook rewrites the chapter in full. The
authored commentary on each workstream lives in the NOTES dict below so that
it survives regeneration: edit it here, never in the .qmd. Keep counts and
dates out of NOTES; they go stale. Anything numeric is computed.

The workbook colours outputs and validations as Delayed or Critical through
conditional formatting, which is not stored as a value. The script derives
the same flags from the dates, using the rules in the Key, as of the day it
runs.
"""

import datetime
import re
import shutil
import subprocess
import tempfile
from collections import Counter, OrderedDict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
XLSX = ROOT / "HIIP_full_plan_EN.xlsx"
OUT = ROOT / "planning-and-roadmap.qmd"
TODAY = datetime.date.today()

# Key: "Delayed: the date is less than two weeks away and the work has not
# started."
DELAYED_WINDOW = 14

# --------------------------------------------------------------------------
# Authored commentary per workstream. Survives regeneration. {n_twg} is
# replaced by the number of other workstreams validated at a TWG-PHC session.
# --------------------------------------------------------------------------
NOTES = {
    "TWG-PHC": (
        "The technical working group for primary health care is the body the "
        "rest of the plan answers to: {n_twg} of the other workstreams have "
        "their external validation at a TWG-PHC session. Reconstituting it is "
        "not one workstream among others but the precondition that governs "
        "the calendar. It is also where the plan is currently held up from "
        "outside: its work cannot move until the Ministry does."
    ),
    "Background document - PHC": (
        "The diagnosis. It inherits work by earlier consultants, which is "
        "audited before being reused, and combines a quantitative backbone "
        "with pilot field visits. It is the direct input to the strategy "
        "document: no strategic prioritisation can be defended without it."
    ),
    "Strategy document - PHC": (
        "Turns the diagnosis into choices through an explicit chain: map of "
        "barriers and gaps, theory of change, prioritised areas. Validation "
        "requires the Ministry to confirm that the strategy serves as the "
        "enabling policy document; without that, the investment plan and the "
        "essential package have no normative basis."
    ),
    "HHFA and situational diagnosis": (
        "Revised in September 2026. The inter-agency agreement with UNICEF "
        "and UNFPA was abandoned and replaced by an approach to GEPE, the "
        "Ministry's planning and statistics office. The HHFA becomes a light "
        "version built into the Angolan assessment tool already in use, in "
        "support of the Health Map (*Mapa Sanitário*), on a December 2026 "
        "horizon. Its legacy is a proposal to extend that tool with the HHFA "
        "items, so that assessment does not end with the project. The "
        "instrument itself is covered in the HHFA part of this notebook."
    ),
    "PFM": (
        "Public financial management. It runs with its own working group, "
        "the TWG-PFM, separate from the TWG-PHC and with its own validation, "
        "and shares the December 2026 horizon with the HHFA. The method that "
        "structures it is covered in the FinHealth 2.0 part of this notebook."
    ),
    "Investment case - PHC": (
        "The only product with **two** external validations: a political "
        "dialogue with the Minister on the first version, and a technical "
        "validation of the second, with HHFA data, by the TWG-PHC. The split "
        "recognises that persuading and proving are distinct moments, with "
        "different audiences and different criteria."
    ),
    "PHC investment plan (costed)": (
        "Turns the strategy into figures. The order matters: consensus on "
        "the investment areas comes before any costing, because costing "
        "before deciding where to invest produces a number nobody owns."
    ),
    "Essential package - PHC": (
        "Defines what primary care should actually offer, by filtering the "
        "UHC Compendium down to the primary care level and fitting it to the "
        "country's epidemiological profile."
    ),
    "Care model and professional profiles - PHC": (
        "The operational counterpart of the essential package: where the "
        "package says *what*, this workstream says *by whom and how* — team "
        "composition, the functions actually performed, professional "
        "profiles."
    ),
    "Project scoping - PHC": (
        "Maps who already finances primary care, where and for what, "
        "identifies overlaps and gaps, and turns the result into project "
        "profiles that can be put to financiers."
    ),
    "Financing mobilisation - PHC": (
        "The end of the arc: engaging the European Investment Bank, other "
        "donors and the EU delegation, and a World Bank loan. Its validation "
        "is the furthest away in time and depends on almost every other "
        "workstream."
    ),
    "HIIP contractual reporting": (
        "Contractual reporting obligations. The only workstream with no "
        "external validation, because the test is not acceptance by a "
        "technical client but meeting the contractual deadline."
    ),
}

STATUS_CLASS = {
    "Done": "st-done",
    "In progress": "st-prog",
    "To do": "st-todo",
    "Blocked": "st-block",
    "Backlog": "st-backlog",
    "Not started": "st-none",
    "Abandoned": "st-aband",
    "Delayed": "st-delayed",
    "Critical": "st-critical",
}


# ------------------------------------------------------------------ helpers
def open_workbook(path):
    """Open the workbook even while Excel holds an exclusive lock on it."""
    try:
        return openpyxl.load_workbook(path, data_only=True)
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
    print(f"  (workbook open in another program; read a copy at {tmp})")
    return openpyxl.load_workbook(tmp, data_only=True)


def txt(v):
    if v is None:
        return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def as_date(iso):
    try:
        return datetime.date.fromisoformat(iso) if iso else None
    except ValueError:
        return None


def readable(iso):
    """2026-10-23 -> 23 October 2026"""
    d = as_date(iso)
    return f"{d.day} {d.strftime('%B')} {d.year}" if d else (iso or "")


def cell(s):
    """Escape what would break a pipe table."""
    return txt(s).replace("|", "\\|").replace("\n", " ")


def badge(status, tip=None):
    if not status:
        return ""
    attr = f' title="{tip}"' if tip else ""
    return f"[{status}]{{.{STATUS_CLASS.get(status, 'st-none')}{attr}}}"


def callout_title(s):
    return txt(s).replace('"', "'")


def flag(status, iso):
    """Delayed / Critical as the workbook's conditional formatting shows them."""
    if status in ("Done", "Abandoned"):
        return None
    d = as_date(iso)
    if not d:
        return None
    if d < TODAY:
        return "Critical", f"Due {readable(iso)}, which has passed"
    if status in ("Not started", "") and (d - TODAY).days <= DELAYED_WINDOW:
        return "Delayed", f"Due {readable(iso)} and not started"
    return None


def status_badges(status, iso):
    out = badge(status)
    f = flag(status, iso)
    if f:
        out += " " + badge(f[0], f[1])
    return out


def sprint_key(s):
    m = re.fullmatch(r"S(\d+)", s or "")
    return int(m.group(1)) if m else None


def sprint_span(tasks):
    nums = [n for n in (sprint_key(t["Sprint"]) for t in tasks) if n is not None]
    if not nums:
        return ""
    lo, hi = min(nums), max(nums)
    return f"S{lo}" if lo == hi else f"S{lo}–S{hi}"


def slug(s):
    return "ws-" + re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def plural(n, one, many=None):
    return f"{n} {one if n == 1 else (many or one + 's')}"


# --------------------------------------------------------------------- read
def read_rows(ws, header_row):
    hdr = [txt(c) for c in next(ws.iter_rows(min_row=header_row,
                                             max_row=header_row,
                                             values_only=True))]
    rows = []
    for r in ws.iter_rows(min_row=header_row + 1, values_only=True):
        d = {hdr[i]: txt(v) for i, v in enumerate(r) if i < len(hdr) and hdr[i]}
        rows.append(d)
    return hdr, rows


def read_key(ws):
    """Sections are rows with a label only; items are label/text pairs."""
    key, section = OrderedDict(), None
    for row in ws.iter_rows(min_row=2, values_only=True):
        a = txt(row[0])
        b = txt(row[1]) if len(row) > 1 else ""
        if a and not b:
            section = a
            key[section] = []
        elif b and section:
            key[section].append((a, b))
    return key


def main():
    wb = open_workbook(XLSX)

    _, master = read_rows(wb["2. Master"], 3)
    master = [r for r in master if r.get("Task") or r.get("Output")]
    key = read_key(wb["0. Key"])
    _, sprints = read_rows(wb["4. Sprints"], 3)
    sprints = [s for s in sprints if s.get("Sprint")]
    _, founding = read_rows(wb["6. Founding document"], 7)
    founding = [f for f in founding if f.get("Activity")]
    founding_notes = []
    source_line = ""
    for row in wb["6. Founding document"].iter_rows(values_only=True):
        t = " ".join(txt(x) for x in row if txt(x))
        if t.startswith("Source:"):
            source_line = t
    notes_started = False
    for row in wb["6. Founding document"].iter_rows(values_only=True):
        t = " ".join(txt(x) for x in row if txt(x))
        if t == "Notes":
            notes_started = True
        elif notes_started and t:
            founding_notes.append(t)

    # group: workstream > output > tasks
    streams = OrderedDict()
    for r in master:
        w = r["Workstream (product)"]
        streams.setdefault(w, {"meta": r, "outputs": OrderedDict()})
        streams[w]["outputs"].setdefault(r["Output"], []).append(r)

    n_tasks = len(master)
    st = Counter(r["Task status"] for r in master)
    n_outputs = sum(len([o for o in s["outputs"] if o]) for s in streams.values())
    owners = sorted({r["Owner"] for r in master if r["Owner"]})
    n_twg = sum(1 for w, s in streams.items()
                if w != "TWG-PHC" and "TWG-PHC" in s["meta"]["External validation"])

    out = []
    A = out.append

    # ============================================================ header
    A("# Planning and roadmap\n")
    A(
        "This chapter is the readable version of the HIIP work plan. The "
        "single source of truth remains the workbook `HIIP_full_plan_EN.xlsx`, "
        "on its *Master* tab; what appears here is generated from it and is "
        "meant for reading and discussion, not for editing.\n"
    )
    summary = (
        f"The plan covers **{plural(len(streams), 'workstream')}**, "
        f"**{plural(n_outputs, 'output')}** and **{plural(n_tasks, 'task')}**, "
        f"spread over fortnightly sprints from "
        f"{readable(sprints[0]['Start'])} to "
        f"{readable([s for s in sprints if s['End']][-1]['End'])}. "
        f"**{st['Done']}** tasks are done ({st['Done'] / n_tasks:.0%}) and "
        f"{st['In progress']} are in progress"
    )
    if st["Abandoned"]:
        summary += f"; {st['Abandoned']} have been abandoned and are kept for the record"
    summary += "."
    if len(owners) == 1:
        summary += f" All tasks are assigned to the {owners[0].lower()}."
    A(summary + "\n")
    A(
        f"Delayed and Critical flags are computed from the dates as of "
        f"**{readable(TODAY.isoformat())}**, when this chapter was generated.\n"
    )

    # ------------------------------------------------ how to read
    A("## How to read this plan\n")
    A(
        "The plan has three levels, and the distinction between them is what "
        "makes it usable. A **workstream** produces a product and is the only "
        "level with external validation. An **output** is an intermediate "
        "delivery that closes within one or two sprints, and is the only "
        "level with a definition of done. A **task** is the concrete work: it "
        "has no criteria of its own, being specific enough that once carried "
        "out it is done.\n"
    )
    A(
        "The separation that matters most is between the **definition of "
        "done** and **external validation**. The definition of done is a set "
        "of internal criteria under the project manager's control: the output "
        "closes when they are met. External validation is acceptance by the "
        "client — the TWG-PHC, the TWG-PFM, the Minister, the financiers — "
        "with its own date and **outside the project manager's control**. A "
        "product can meet every internal criterion and still not be "
        "validated.\n"
    )
    A(
        "The plan also records **deadlocks**: local constraints that stop an "
        "output from being achieved even when the work is ready to proceed. "
        "A deadlock is not a delay in the work; it is the reason the work "
        "cannot move.\n"
    )
    A(
        "Some tasks have no output. That is not an omission: they serve the "
        "workstream's external validation directly, without passing through "
        "an intermediate delivery, and are grouped separately under each "
        "workstream.\n"
    )

    conv = [s for s in ("Statuses", "Other columns") if s in key]
    if conv:
        A('::: {.callout-note collapse="true" title="Statuses and conventions"}\n')
        if "Statuses" in key:
            A("| Level or flag | Meaning |")
            A("|:--|:--|")
            for k, v in key["Statuses"]:
                A(f"| {cell(k)} | {cell(v)} |")
            A("\n: {tbl-colwidths=\"[28,72]\"}\n")
        if "Other columns" in key:
            A("| Column | Meaning |")
            A("|:--|:--|")
            for k, v in key["Other columns"]:
                A(f"| {cell(k)} | {cell(v)} |")
            A("\n: {tbl-colwidths=\"[18,82]\"}\n")
        A(":::\n")

    rule = next((v for v in (key.get("Escalation rule") or []) if v[1]), None)
    if rule:
        A('::: {.callout-tip title="Escalation rule"}')
        A(rule[1])
        A(":::\n")

    # ------------------------------------------------ overview
    A("## Workstreams at a glance\n")
    A(
        "Each workstream has its own section below, with its external "
        "validation, outputs and tasks. The validation dates are what "
        "structure the calendar: they do not move for internal convenience.\n"
    )
    A("| Workstream | Validation | Outputs | Tasks | Done | Deadlocked | Sprints |")
    A("|:---|:---|---:|---:|---:|---:|:---|")
    for name, s in streams.items():
        tasks = [t for lst in s["outputs"].values() for t in lst]
        dv = s["meta"]["Validation date"]
        fl = flag(s["meta"]["Validation status"], dv)
        when = readable(dv) if dv else "—"
        if fl:
            when += " " + badge(fl[0], fl[1])
        dl = sum(1 for t in tasks if t["Deadlock"])
        A(
            f"| [{cell(name)}](#{slug(name)}) | {when} "
            f"| {len([o for o in s['outputs'] if o])} | {len(tasks)} "
            f"| {sum(1 for t in tasks if t['Task status'] == 'Done')} "
            f"| {dl if dl else '—'} | {sprint_span(tasks) or '—'} |"
        )
    A("")

    deadlocks = OrderedDict()
    for r in master:
        if r["Deadlock"]:
            deadlocks.setdefault(r["Deadlock"], Counter())[r["Workstream (product)"]] += 1
    if deadlocks:
        A('::: {.callout-important title="Current deadlocks"}\n')
        for text, by_ws in deadlocks.items():
            where = "; ".join(f"{w} ({plural(n, 'task')})" for w, n in by_ws.items())
            A(f"- [Deadlock]{{.st-deadlock}} **{text}** — {where}.")
        A("\n:::\n")

    A('::: {.callout-note collapse="true" title="Sprint calendar"}\n')
    A("| Sprint | Start | End | Tasks | Days |")
    A("|:---|:---|:---|---:|---:|")
    for s in sprints:
        A(f"| {cell(s['Sprint'])} | {readable(s.get('Start')) or '—'} "
          f"| {readable(s.get('End')) or '—'} | {cell(s.get('Tasks planned'))} "
          f"| {cell(s.get('Days planned'))} |")
    A("\n:::\n")

    # ------------------------------------------------ contractual calendar
    if founding:
        A("## The contractual calendar\n")
        A(
            "The plan answers to a contract. The [Proposal for Action]"
            "(hiip-proposal-for-action.qmd) recognises three contractual "
            "deliverables and two reports, scheduled in project months; every "
            "other product in the plan is intermediate work supporting them. "
            "Where the plan's dates and the contract's diverge, **the contract "
            "prevails**.\n"
        )
        A("| Contractual deliverable | Activity | Months | Period | Workstream in the plan |")
        A("|:---|:---|:--|:--|:---|")
        for f in founding:
            dlv = f.get("Contractual deliverable", "")
            m = f["Start month"] if f["Start month"] == f["End month"] else f"{f['Start month']}–{f['End month']}"
            per = f"{readable(f['Start date'])} – {readable(f['End date'])}"
            A(f"| {'**' + cell(dlv) + '**' if dlv else ''} | {cell(f['Activity'])} "
              f"| {m} | {per} | {cell(f['Workstream in the plan'])} |")
        A('\n: {tbl-colwidths="[22,34,8,20,16]"}\n')

        # the Key asks the reader to compare the two calendars; do it here
        contract_end = {}
        for f in founding:
            w, e = f["Workstream in the plan"], as_date(f["End date"])
            if w and e:
                contract_end[w] = max(contract_end.get(w, e), e)
        cmp_rows = []
        for w, e in contract_end.items():
            v = as_date(streams.get(w, {}).get("meta", {}).get("Validation date", ""))
            if v:
                cmp_rows.append((w, e, v, (v - e).days))
        if cmp_rows:
            A("### Plan against contract\n")
            A(
                "For each workstream the contract names, the latest contractual "
                "end date against the workstream's validation date in the plan. "
                "A positive gap means the plan runs past the contract.\n"
            )
            A("| Workstream | Contract ends | Plan validates | Gap |")
            A("|:---|:---|:---|---:|")
            for w, e, v, gap in sorted(cmp_rows, key=lambda x: -x[3]):
                g = f"+{gap} days" if gap > 0 else f"{gap} days"
                if gap > 0:
                    g = f"**{g}** [Past contract]{{.st-delayed title=\"Plan validates after the contract ends\"}}"
                A(f"| [{cell(w)}](#{slug(w)}) | {readable(e.isoformat())} "
                  f"| {readable(v.isoformat())} | {g} |")
            A("")

        if founding_notes or source_line:
            A('::: {.callout-note collapse="true" title="Notes on the founding document"}\n')
            if source_line:
                A(source_line + "\n")
            for n in founding_notes:
                A(f"- {n}")
            A("\n:::\n")

    # ================================================= one section each
    A("## The workstreams\n")
    for name, s in streams.items():
        meta = s["meta"]
        tasks_all = [t for lst in s["outputs"].values() for t in lst]
        done = sum(1 for t in tasks_all if t["Task status"] == "Done")
        aband = sum(1 for t in tasks_all if t["Task status"] == "Abandoned")
        n_out = len([o for o in s["outputs"] if o])

        A(f"### {name} {{#{slug(name)}}}\n")
        if name in NOTES:
            A(NOTES[name].format(n_twg=n_twg) + "\n")

        val, dv = meta["External validation"], meta["Validation date"]
        if val:
            A('::: {.callout-warning title="External validation"}')
            A(f"**When:** {readable(dv)}  ")
            A(f"**Status:** {status_badges(meta['Validation status'], dv)}\n")
            A(val)
            A(":::\n")
        else:
            A("*No external validation recorded: the test for this workstream "
              "is meeting the contractual deadline.*\n")

        line = f"{plural(n_out, 'output')} and {plural(len(tasks_all), 'task')}"
        span = sprint_span(tasks_all)
        if span:
            line += f", across {span}"
        line += ". " + ("None done yet." if done == 0 else f"{done} done.")
        if aband:
            line += f" {aband} abandoned."
        A(line + "\n")

        for output, tasks in s["outputs"].items():
            if output:
                t0 = tasks[0]
                A(f'::: {{.callout-note collapse="true" title="{callout_title(output)}"}}\n')
                A(f"**Due** {readable(t0['Output date']) or 'date to be set'} "
                  f"· **Status** {status_badges(t0['Output status'], t0['Output date'])}\n")
                if t0["Definition of done (output)"]:
                    A(f"**Definition of done.** {t0['Definition of done (output)']}\n")
            else:
                A('::: {.callout-note collapse="true" '
                  'title="Tasks serving the external validation directly"}\n')
                A("Tasks with no intermediate output: they serve this "
                  "workstream's external validation directly.\n")

            dl = Counter(t["Deadlock"] for t in tasks if t["Deadlock"])
            for text, n in dl.items():
                A(f"[Deadlock]{{.st-deadlock}} {text}"
                  f"{'' if n == len(tasks) else f' ({n} of {len(tasks)} tasks)'}\n")

            A("| Task | Sprint | Est. | Depends on | Status |")
            A("|:---|:---|---:|:---|:---|")
            for t in tasks:
                est = t["Estimate (days)"]
                A(f"| {cell(t['Task'])} | {cell(t['Sprint']) or '—'} "
                  f"| {est + ' d' if est else '—'} | {cell(t['Depends on']) or '—'} "
                  f"| {badge(t['Task status'])} |")
            A("")

            learned = [(t["Task"], t["Learning"]) for t in tasks if t.get("Learning")]
            if learned:
                A("**Learning recorded**\n")
                for task, l in learned:
                    A(f"- *{task}* — {l}")
                A("")
            A(":::\n")

    OUT.write_text("\n".join(out), encoding="utf-8")
    print(f"Written: {OUT}")
    print(f"  {len(streams)} workstreams, {n_outputs} outputs, {n_tasks} tasks "
          f"({st['Done']} done, {st['Abandoned']} abandoned), "
          f"{sum(sum(c.values()) for c in deadlocks.values())} deadlocked, "
          f"flags as of {TODAY}")


if __name__ == "__main__":
    main()
