import json, re, sys, os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import FormulaRule, CellIsRule
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

SRC, OUT = sys.argv[1], sys.argv[2]
cat = json.load(open(os.path.join(SRC, "NIST_SP-800-53_rev5_catalog.json")))["catalog"]
VER = cat["metadata"]["version"]
base = {}
for b in ["LOW", "MODERATE", "HIGH", "PRIVACY"]:
    p = json.load(open(os.path.join(SRC, f"NIST_SP-800-53_rev5_{b}-baseline_profile.json")))["profile"]
    base[b] = {i for im in p["imports"] for inc in im["include-controls"] for i in inc["with-ids"]}

# ---------- catalog parsing ----------
params, labels = {}, {}
def collect(ctl):
    for p in ctl.get("params", []):
        params[p["id"]] = p
    labels[ctl["id"]] = next(p["value"] for p in ctl["props"] if p["name"] == "label" and "class" not in p)
    for s in ctl.get("controls", []):
        collect(s)
for g in cat["groups"]:
    for c in g["controls"]:
        collect(c)

INS = re.compile(r"\{\{\s*insert:\s*param,\s*([^\s}]+)\s*\}\}")
def render_param(pid):
    p = params.get(pid)
    if not p:
        return f"[{pid}]"
    if "select" in p:
        s = p["select"]
        how = " (one or more)" if s.get("how-many") == "one-or-more" else ""
        return f"[Selection{how}: " + "; ".join(sub(ch) for ch in s.get("choice", [])) + "]"
    lab = p.get("label", pid)
    return f"[Assignment: {lab}]" if lab.lower().startswith("organization-defined") else f"[Assignment: organization-defined {lab}]"
def sub(text):
    return INS.sub(lambda m: render_param(m.group(1)), text or "")
def part_text(part, depth=0):
    lines = []
    lab = next((p["value"] for p in part.get("props", []) if p["name"] == "label"), "")
    prose = sub(part.get("prose", ""))
    if lab or prose:
        lines.append(("  " * depth) + (lab + " " if lab else "") + prose)
    for sp in part.get("parts", []):
        lines.append(part_text(sp, depth + (1 if (lab or prose) else 0)))
    return "\n".join(l for l in lines if l)

def tier(cid):
    if cid in base["LOW"]: return "1 - Low"
    if cid in base["MODERATE"]: return "2 - Moderate"
    if cid in base["HIGH"]: return "3 - High"
    return "4 - Not in Baseline"

rows = []
def add(ctl, fam_id, fam_name, parent=None):
    props = ctl.get("props", [])
    withdrawn = any(p["name"] == "status" and p["value"] == "withdrawn" for p in props)
    sort_id = next(p["value"] for p in props if p["name"] == "sort-id")
    lvl = sorted({p["value"].title() for p in props if p["name"] == "implementation-level"})
    assur = "Yes" if any(p["name"] == "contributes-to-assurance" for p in props) else ""
    links = ctl.get("links", [])
    disp = ""
    if withdrawn:
        inc = [labels.get(l["href"].lstrip("#"), l["href"]) for l in links if l["rel"] == "incorporated-into"]
        mov = [labels.get(l["href"].lstrip("#"), l["href"]) for l in links if l["rel"] == "moved-to"]
        disp = "; ".join(x for x in [("Incorporated into " + ", ".join(inc)) if inc else "",
                                     ("Moved to " + ", ".join(mov)) if mov else ""] if x) or "Withdrawn"
    related = ", ".join(labels.get(l["href"].lstrip("#"), "") for l in links if l["rel"] == "related")
    stmt = "\n".join(part_text(p) for p in ctl.get("parts", []) if p["name"] == "statement")
    disc = "\n".join(sub(p.get("prose", "")) for p in ctl.get("parts", []) if p["name"] == "guidance")
    title = ctl["title"] if parent is None else f"{parent['title']} | {ctl['title']}"
    cid = ctl["id"]
    rows.append(dict(
        sort=sort_id, id=labels[cid], fam=fam_id.upper(), famname=fam_name, name=title,
        type="Control" if parent is None else "Enhancement",
        parent="" if parent is None else labels[parent["id"]],
        status="Withdrawn" if withdrawn else "Active", disp=disp, lvl=" / ".join(lvl), assur=assur,
        L="X" if cid in base["LOW"] else "", M="X" if cid in base["MODERATE"] else "",
        H="X" if cid in base["HIGH"] else "", P="X" if cid in base["PRIVACY"] else "",
        tier="Withdrawn" if withdrawn else tier(cid), stmt=stmt, disc=disc, related=related))
    for s in ctl.get("controls", []):
        add(s, fam_id, fam_name, ctl)
for g in cat["groups"]:
    for c in g["controls"]:
        add(c, g["id"], g["title"])
rows.sort(key=lambda r: r["sort"])
families = [(g["id"].upper(), g["title"]) for g in cat["groups"]]

# ---------- styles ----------
F = "Arial"
NAVY, LIGHT, INPUT = "1F3864", "D9E1F2", "FFF2CC"
hfont = Font(name=F, bold=True, color="FFFFFF", size=10)
hfill = PatternFill("solid", fgColor=NAVY)
bfont = Font(name=F, size=10)
bold = Font(name=F, size=10, bold=True)
title_font = Font(name=F, size=14, bold=True, color=NAVY)
infill = PatternFill("solid", fgColor=INPUT)
calcfill = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="A6A6A6")
box = Border(left=thin, right=thin, top=thin, bottom=thin)
wrap = Alignment(wrap_text=True, vertical="top")
center = Alignment(horizontal="center", vertical="top")

def dv(ws, formula, rng, title=None, prompt=None, allow_blank=True):
    d = DataValidation(type="list", formula1=formula, allow_blank=allow_blank, showErrorMessage=True,
                       errorTitle="Invalid entry", error="Select a value from the list.")
    if prompt:
        d.promptTitle, d.prompt, d.showInputMessage = title, prompt, True
    ws.add_data_validation(d); d.add(rng)

wb = Workbook()

# ================= Instructions =================
ins = wb.active; ins.title = "Instructions"
ins.sheet_view.showGridLines = False
ins.column_dimensions["A"].width = 3; ins.column_dimensions["B"].width = 30; ins.column_dimensions["C"].width = 110
ins["B2"] = "Security Control Traceability Matrix (SCTM) — NIST SP 800-53 Rev 5"; ins["B2"].font = title_font
ins["B3"] = (f"Control catalog: NIST SP 800-53 Rev {VER} (OSCAL release, {len(rows)} controls & enhancements, "
             f"{sum(r['status']=='Withdrawn' for r in rows)} withdrawn). Baselines: NIST SP 800-53B as published in the "
             f"Rev {VER} OSCAL baseline profiles (LOW {len(base['LOW'])}, MODERATE {len(base['MODERATE'])}, "
             f"HIGH {len(base['HIGH'])}, PRIVACY {len(base['PRIVACY'])}).")
ins["B3"].font = Font(name=F, size=10, italic=True); ins["B3"].alignment = wrap; ins.merge_cells("B3:C3")
ins.row_dimensions[3].height = 30
ins["B4"] = "Source: https://github.com/usnistgov/oscal-content/tree/main/nist.gov/SP800-53/rev5/json"
ins["B4"].font = Font(name=F, size=9, color="0563C1", underline="single"); ins["B4"].hyperlink = "https://github.com/usnistgov/oscal-content/tree/main/nist.gov/SP800-53/rev5/json"

steps = [
    ("HOW TO USE", None),
    ("1. Categorize", "On 'Categorization', enter the information types the system processes (SP 800-60) with their C/I/A impact levels. "
     "The system C, I and A are derived as the high-water mark of the information types. Use the Override column to set a "
     "level directly. The overall impact (FIPS 199 / FIPS 200 high-water mark) selects the 800-53B baseline. Set 'Privacy baseline applies' to Yes if the system processes PII."),
    ("2. Define overlays", "On 'Overlays', name up to 5 overlays (e.g., CNSSI 1253 Privacy, Space Platform, Classified Information, "
     "organizational overlay) and set Enabled = Yes. Disabled overlays are ignored by every formula."),
    ("3. Apply overlays", "On 'SCTM', use the Overlay 1–5 columns for each control: Add = overlay selects the control; "
     "Remove = overlay removes the control; Modify = overlay changes parameters/supplemental guidance (record the detail in 'ODP Values / Overlay Parameters')."),
    ("4. Review determination", "'Selected by Baseline' and 'Final Determination' compute automatically. Final Determination values: "
     "Required · Required (Overlay Added) · Required (Overlay Modified) · Removed by Overlay · Not Selected · Conflict - Review "
     "(an enabled overlay adds and another removes the control) · Withdrawn · Pending Categorization."),
    ("5. Record status", "Set 'Implementation Status' for each control: Implemented, Not Implemented, Tailored Out, N/A, Planned, Not Planned. "
     "Give a justification in 'Tailoring Justification' for Tailored Out, N/A and Not Planned."),
    ("6. Track", "'Dashboard' summarizes required controls and their status by family. All columns on 'SCTM' are sortable and filterable "
     "(header drop-downs). To restore catalog order, sort by 'Sort Key' A→Z."),
    ("CELL KEY", None),
    ("Yellow cells", "Input cells — edit these."),
    ("Green cells", "Formula cells — calculated automatically; do not overwrite."),
    ("Blue / white cells", "NIST reference data (catalog and 800-53B baselines) — read only."),
    ("COLUMN NOTES", None),
    ("LOW / MODERATE / HIGH", "'X' = the control is in that SP 800-53B security control baseline."),
    ("PRIVACY", "'X' = the control is in the SP 800-53B privacy control baseline (applies independently of the impact level)."),
    ("Lowest Baseline", "The lowest baseline that selects the control (1 - Low, 2 - Moderate, 3 - High, 4 - Not in Baseline). Sort on this column to group controls by impact level."),
    ("Implementation Level", "NIST-designated implementation level: Organization and/or System."),
    ("Assurance", "Yes = NIST marks the control as contributing to assurance (SP 800-53 Rev 5, Appendix C)."),
    ("Control Statement", "Full control text; organization-defined parameters appear as [Assignment: …] / [Selection: …]."),
    ("CNSSI 1253 note", "SP 800-53B selects one baseline from the overall high-water mark. CNSSI 1253 selects controls separately for C, I and A. "
     "To use CNSSI 1253 for a National Security System, load its control selections as an overlay (Add/Remove) on top of the baseline, "
     "or set the baseline override on 'Categorization'."),
    ("EXAMPLE ROW", None),
    ("Example (AC-2)", "Overlay 1: Modify · Implementation Status: Implemented · Control Origination: Hybrid · Common Control Provider: Enterprise IdM (AD) · "
     "Responsible Role: ISSO · Implementation Description: 'Accounts are requested in ServiceNow, approved by the data owner, and provisioned by the IdM team…' · "
     "ODP Values: 'AC-02_ODP[01]: ISSM approval; inactivity disable = 35 days' · Evidence: 'Account Mgmt SOP v3.1; quarterly review records' · Assessment Result: Satisfied."),
]
r = 6
for k, v in steps:
    if v is None:
        r += 1
        ins.cell(r, 2, k).font = Font(name=F, size=11, bold=True, color=NAVY)
    else:
        a = ins.cell(r, 2, k); a.font = bold; a.alignment = wrap
        b = ins.cell(r, 3, v); b.font = bfont; b.alignment = wrap
        if k == "Yellow cells": a.fill = infill
        if k == "Green cells": a.fill = calcfill
        if k == "Blue / white cells": a.fill = PatternFill("solid", fgColor=LIGHT)
        ins.row_dimensions[r].height = max(15, 14 * (len(v) // 115 + 1))
    r += 1

# ================= Categorization =================
cz = wb.create_sheet("Categorization"); cz.sheet_view.showGridLines = False
for col, w in zip("ABCDEFG", [3, 34, 16, 16, 16, 16, 40]):
    cz.column_dimensions[col].width = w
cz["B2"] = "System Security Categorization (FIPS 199 / SP 800-60 / FIPS 200)"; cz["B2"].font = title_font
cz["B4"] = "System Name"; cz["B4"].font = bold
cz.merge_cells("C4:F4"); cz["C4"].fill = infill; cz["C4"].border = box; cz["C4"].font = bfont

hdr = ["Security Objective", "Derived from Info Types", "Override (optional)", "System Impact Level"]
for i, h in enumerate(hdr):
    c = cz.cell(6, 2 + i, h); c.font = hfont; c.fill = hfill; c.alignment = Alignment(horizontal="center", wrap_text=True)
IT0, IT1 = 19, 43  # information types table rows
objs = [("Confidentiality", "D"), ("Integrity", "E"), ("Availability", "F")]
for i, (o, col) in enumerate(objs):
    rr = 7 + i
    rng = f"${col}${IT0}:${col}${IT1}"
    cz.cell(rr, 2, o).font = bold
    cz.cell(rr, 3, f'=IF(COUNTIF({rng},"High")>0,"High",IF(COUNTIF({rng},"Moderate")>0,"Moderate",IF(COUNTIF({rng},"Low")>0,"Low","")))').fill = calcfill
    cz.cell(rr, 4).fill = infill
    cz.cell(rr, 5, f'=IF(D{rr}<>"",D{rr},C{rr})').fill = calcfill
    for cc in range(2, 6):
        x = cz.cell(rr, cc); x.border = box
        if cc > 2: x.alignment = Alignment(horizontal="center"); x.font = bold if cc == 5 else bfont
dv(cz, '"Low,Moderate,High"', "D7:D9", "Override", "Optional: set the impact level directly instead of deriving it from information types.")

cz["B11"] = "Overall System Impact (High-Water Mark)"; cz["B11"].font = bold
cz["E11"] = '=IF(COUNTIF(E7:E9,"?*")<3,"",IF(COUNTIF(E7:E9,"High")>0,"HIGH",IF(COUNTIF(E7:E9,"Moderate")>0,"MODERATE","LOW")))'
cz["B12"] = "Baseline Override (optional)"; cz["B12"].font = bold
cz["B13"] = "Effective Control Baseline"; cz["B13"].font = Font(name=F, size=11, bold=True, color=NAVY)
cz["E13"] = '=IF(E12<>"",E12,E11)'
cz["B14"] = "Privacy Baseline Applies (system processes PII)?"; cz["B14"].font = bold
cz["E14"] = "No"
for a, fill in [("E11", calcfill), ("E12", infill), ("E13", calcfill), ("E14", infill)]:
    cz[a].fill = fill; cz[a].border = box; cz[a].alignment = Alignment(horizontal="center"); cz[a].font = bold
cz["E13"].font = Font(name=F, size=11, bold=True, color=NAVY)
dv(cz, '"LOW,MODERATE,HIGH"', "E12", "Baseline override", "Leave blank to use the high-water mark. Use only with a documented, AO-approved rationale.")
dv(cz, '"Yes,No"', "E14", allow_blank=False)
cz["G12"] = "Leave blank unless the AO approves a different baseline."; cz["G12"].font = Font(name=F, size=9, italic=True)
cz["G14"] = "Yes adds the SP 800-53B privacy baseline."; cz["G14"].font = Font(name=F, size=9, italic=True)

cz["B17"] = "Information Types (SP 800-60 Vol. II)"; cz["B17"].font = Font(name=F, size=11, bold=True, color=NAVY)
for i, h in enumerate(["Information Type", "SP 800-60 ID", "Confidentiality", "Integrity", "Availability", "Rationale / Adjustment"]):
    c = cz.cell(18, 2 + i, h); c.font = hfont; c.fill = hfill; c.alignment = Alignment(horizontal="center", wrap_text=True)
for rr in range(IT0, IT1 + 1):
    for cc in range(2, 8):
        x = cz.cell(rr, cc); x.fill = infill; x.border = box; x.font = bfont
        if 4 <= cc <= 6: x.alignment = Alignment(horizontal="center")
dv(cz, '"Low,Moderate,High"', f"D{IT0}:F{IT1}")
cz.column_dimensions["G"].width = 46

# ================= Overlays =================
ov = wb.create_sheet("Overlays"); ov.sheet_view.showGridLines = False
for col, w in zip("ABCDEFG", [3, 10, 36, 11, 28, 12, 60]):
    ov.column_dimensions[col].width = w
ov["B2"] = "Custom Overlays"; ov["B2"].font = title_font
ov["B3"] = ("Define up to 5 overlays. Only overlays with Enabled = Yes affect 'Final Determination' on the SCTM. "
            "Per control, set the overlay column on 'SCTM' to Add, Remove or Modify.")
ov["B3"].font = Font(name=F, size=10, italic=True); ov.merge_cells("B3:G3"); ov["B3"].alignment = wrap; ov.row_dimensions[3].height = 28
for i, h in enumerate(["Overlay #", "Overlay Name", "Enabled", "Source / Authority", "Version", "Description / Applicability"]):
    c = ov.cell(5, 2 + i, h); c.font = hfont; c.fill = hfill; c.alignment = Alignment(horizontal="center", wrap_text=True)
NOV = 5
for k in range(NOV):
    rr = 6 + k
    ov.cell(rr, 2, k + 1).alignment = Alignment(horizontal="center")
    ov.cell(rr, 4, "No").alignment = Alignment(horizontal="center")
    for cc in range(2, 8):
        x = ov.cell(rr, cc); x.border = box
        x.font = bfont if cc != 2 else bold
        if cc > 2: x.fill = infill
dv(ov, '"Yes,No"', "D6:D10", allow_blank=False)

# ================= SCTM =================
ws = wb.create_sheet("SCTM")
cols = [  # (header, key/kind, width)
    ("Sort Key", "sort", 10), ("Control ID", "id", 12), ("Family", "fam", 7), ("Family Name", "famname", 22),
    ("Control Name", "name", 44), ("Type", "type", 11), ("Parent Control", "parent", 9),
    ("Catalog Status", "status", 10), ("Withdrawal Disposition", "disp", 24),
    ("Implementation Level", "lvl", 13), ("Assurance", "assur", 9),
    ("LOW", "L", 6), ("MODERATE", "M", 7), ("HIGH", "H", 6), ("PRIVACY", "P", 7),
    ("Lowest Baseline", "tier", 14), ("Selected by Baseline", "f_sel", 11),
] + [(f"Overlay {k+1}", f"ov{k}", 10) for k in range(NOV)] + [
    ("Final Determination", "f_final", 22), ("Implementation Status", "in_status", 17),
    ("Control Origination", "in_orig", 15), ("Common Control Provider / Inherited From", "in", 22),
    ("Responsible Role / POC", "in", 18), ("Implementation Description", "in", 50),
    ("ODP Values / Overlay Parameters", "in", 30), ("Tailoring Justification", "in", 30),
    ("Evidence / Artifacts", "in", 28), ("Assessment Result", "in_assess", 15), ("POA&M ID", "in", 11),
    ("Planned Completion Date", "in_date", 12), ("Comments", "in", 30),
    ("Control Statement (NIST)", "stmt", 70), ("Discussion (NIST)", "disc", 70), ("Related Controls", "related", 30),
]
ci = {h: i + 1 for i, (h, _, _) in enumerate(cols)}
CL = lambda h: L(ci[h])
HR, R0 = 3, 4
R1 = R0 + len(rows) - 1

ws["A1"] = f"NIST SP 800-53 Rev {VER} — Security Control Traceability Matrix"; ws["A1"].font = title_font
ws["E1"] = '="System: "&IF(Categorization!$C$4="","(not set)",Categorization!$C$4)&"   |   Effective baseline: "&IF(Categorization!$E$13="","(categorize system)",Categorization!$E$13)&"   |   Privacy baseline: "&Categorization!$E$14'
ws["E1"].font = Font(name=F, size=10, bold=True, color=NAVY)
ws["A2"] = "Overlay names →"; ws["A2"].font = Font(name=F, size=9, italic=True)
for k in range(NOV):
    c = ws.cell(2, ci[f"Overlay {k+1}"], f'=IF(Overlays!$C${6+k}="","(unnamed)",Overlays!$C${6+k})&IF(Overlays!$D${6+k}="Yes",""," [OFF]")')
    c.font = Font(name=F, size=8, bold=True, color=NAVY); c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="bottom")
ws.row_dimensions[2].height = 36

for i, (h, k, w) in enumerate(cols):
    c = ws.cell(HR, i + 1, h); c.font = hfont; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.column_dimensions[L(i + 1)].width = w
ws.row_dimensions[HR].height = 42

S, LOWc, MODc, HIc, PRc = CL("Catalog Status"), CL("LOW"), CL("MODERATE"), CL("HIGH"), CL("PRIVACY")
SEL, FIN = CL("Selected by Baseline"), CL("Final Determination")
OVc = [CL(f"Overlay {k+1}") for k in range(NOV)]
for n, row in enumerate(rows):
    rr = R0 + n
    for i, (h, k, w) in enumerate(cols):
        c = ws.cell(rr, i + 1); c.font = bfont
        if k in row:
            c.value = row[k]
            if k in ("L", "M", "H", "P", "assur", "fam", "type", "status", "tier", "parent", "sort", "lvl"):
                c.alignment = center
            else:
                c.alignment = Alignment(vertical="top", wrap_text=k in ("name", "disp"))
        elif k == "f_sel":
            c.value = (f'=IF(${S}{rr}="Withdrawn","N/A",IF(Categorization!$E$13="","Pending",'
                       f'IF(OR(AND(Categorization!$E$13="LOW",${LOWc}{rr}="X"),AND(Categorization!$E$13="MODERATE",${MODc}{rr}="X"),'
                       f'AND(Categorization!$E$13="HIGH",${HIc}{rr}="X"),AND(Categorization!$E$14="Yes",${PRc}{rr}="X")),"Yes","No")))')
            c.fill = calcfill; c.alignment = center
        elif k == "f_final":
            addc = "+".join(f'(Overlays!$D${6+j}="Yes")*(${OVc[j]}{rr}="Add")' for j in range(NOV))
            remc = "+".join(f'(Overlays!$D${6+j}="Yes")*(${OVc[j]}{rr}="Remove")' for j in range(NOV))
            modc = "+".join(f'(Overlays!$D${6+j}="Yes")*(${OVc[j]}{rr}="Modify")' for j in range(NOV))
            c.value = (f'=IF(${S}{rr}="Withdrawn","Withdrawn",IF(${SEL}{rr}="Pending","Pending Categorization",'
                       f'IF(AND({addc}>0,{remc}>0),"Conflict - Review",IF({remc}>0,IF(${SEL}{rr}="Yes","Removed by Overlay","Not Selected"),'
                       f'IF({addc}>0,IF(${SEL}{rr}="Yes",IF({modc}>0,"Required (Overlay Modified)","Required"),"Required (Overlay Added)"),'
                       f'IF(${SEL}{rr}="Yes",IF({modc}>0,"Required (Overlay Modified)","Required"),"Not Selected"))))))')
            c.fill = calcfill; c.alignment = Alignment(vertical="top")
        elif k.startswith("ov") or k.startswith("in"):
            c.fill = infill
            c.alignment = center if (k.startswith("ov") or k in ("in_status", "in_orig", "in_assess", "in_date")) else Alignment(vertical="top", wrap_text=True)
            if k == "in_date": c.number_format = "yyyy-mm-dd"
    for key in ("L", "M", "H", "P"):
        pass
# reference-data shading on baseline columns
light = PatternFill("solid", fgColor=LIGHT)
for col in (LOWc, MODc, HIc, PRc, CL("Lowest Baseline")):
    for rr in range(R0, R1 + 1):
        ws[f"{col}{rr}"].fill = light

ref = f"A{HR}:{L(len(cols))}{R1}"
t = Table(displayName="SCTM", ref=ref)
t.tableStyleInfo = TableStyleInfo(name="TableStyleLight15", showRowStripes=False)
ws.add_table(t)
ws.freeze_panes = ws[f"C{R0}"]

# hide sort key? keep visible but narrow; group NIST text columns
dv(ws, '"Implemented,Not Implemented,Tailored Out,N/A,Planned,Not Planned"', f"{CL('Implementation Status')}{R0}:{CL('Implementation Status')}{R1}",
   "Implementation Status", "Implemented · Not Implemented · Tailored Out · N/A · Planned · Not Planned")
dv(ws, '"Add,Remove,Modify"', f"{OVc[0]}{R0}:{OVc[-1]}{R1}", "Overlay action",
   "Add = overlay selects control; Remove = overlay removes it; Modify = overlay changes parameters/guidance.")
dv(ws, '"System-Specific,Common,Hybrid,Inherited"', f"{CL('Control Origination')}{R0}:{CL('Control Origination')}{R1}")
dv(ws, '"Satisfied,Other Than Satisfied,Not Assessed"', f"{CL('Assessment Result')}{R0}:{CL('Assessment Result')}{R1}")
dd = DataValidation(type="date", operator="greaterThan", formula1="36526", allow_blank=True, showErrorMessage=True,
                    error="Enter a date (YYYY-MM-DD).")
ws.add_data_validation(dd); dd.add(f"{CL('Planned Completion Date')}{R0}:{CL('Planned Completion Date')}{R1}")

# conditional formatting
ST = CL("Implementation Status"); rng = f"{ST}{R0}:{ST}{R1}"
for val, fg, fc in [("Implemented", "C6EFCE", "006100"), ("Not Implemented", "FFC7CE", "9C0006"),
                    ("Planned", "FFEB9C", "9C5700"), ("Not Planned", "F4B084", "833C0B"),
                    ("Tailored Out", "D9D9D9", "404040"), ("N/A", "EDEDED", "595959")]:
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=[f'"{val}"'],
                                  fill=PatternFill("solid", fgColor=fg, bgColor=fg), font=Font(color=fc, bold=True)))
frng = f"{FIN}{R0}:{FIN}{R1}"
for txt, fg, fc in [("Conflict", "FF0000", "FFFFFF"), ("Required", "C6EFCE", "006100"), ("Removed", "FCE4D6", "833C0B"),
                    ("Not Selected", "F2F2F2", "7F7F7F"), ("Pending", "FFEB9C", "9C5700"), ("Withdrawn", "D9D9D9", "7F7F7F")]:
    ws.conditional_formatting.add(frng, FormulaRule(formula=[f'LEFT(${FIN}{R0},{len(txt)})="{txt}"'], stopIfTrue=True,
                                  fill=PatternFill("solid", fgColor=fg, bgColor=fg), font=Font(color=fc, bold=True)))
srng = f"{SEL}{R0}:{SEL}{R1}"
ws.conditional_formatting.add(srng, CellIsRule(operator="equal", formula=['"Yes"'], font=Font(color="006100", bold=True)))
# withdrawn rows greyed / italic
ws.conditional_formatting.add(f"A{R0}:{FIN}{R1}", FormulaRule(formula=[f'${S}{R0}="Withdrawn"'],
                              font=Font(color="A6A6A6", italic=True)))
# warn: status needs justification
TJ = CL("Tailoring Justification")
ws.conditional_formatting.add(f"{TJ}{R0}:{TJ}{R1}", FormulaRule(
    formula=[f'AND(OR(${ST}{R0}="Tailored Out",${ST}{R0}="N/A",${ST}{R0}="Not Planned"),${TJ}{R0}="")'],
    fill=PatternFill("solid", fgColor="FFC7CE", bgColor="FFC7CE")))
ws[f"{TJ}{HR}"].comment = Comment("Turns red when status is Tailored Out, N/A or Not Planned and no justification is entered.", "SCTM")
ws[f"{FIN}{HR}"].comment = Comment("Calculated from 'Selected by Baseline' and the Add/Remove/Modify actions of ENABLED overlays.", "SCTM")
ws[f"{SEL}{HR}"].comment = Comment("Calculated: Yes when the control is in the effective baseline (Categorization!E13) or in the privacy baseline when Categorization!E14 = Yes.", "SCTM")

# ================= Dashboard =================
db = wb.create_sheet("Dashboard", 1); db.sheet_view.showGridLines = False
db["B2"] = "SCTM Dashboard"; db["B2"].font = title_font
db["B3"] = '="Effective baseline: "&IF(Categorization!$E$13="","(categorize system)",Categorization!$E$13)&"   |   Privacy baseline: "&Categorization!$E$14'
db["B3"].font = Font(name=F, size=10, bold=True, color=NAVY)
stats = ["Implemented", "Planned", "Not Implemented", "Not Planned", "Tailored Out", "N/A"]
hdr = ["Family", "Family Name", "Active Controls", "LOW", "MODERATE", "HIGH", "PRIVACY", "Required"] + stats + ["No Status", "% Implemented"]
HRD = 5
for i, h in enumerate(hdr):
    c = db.cell(HRD, 2 + i, h); c.font = hfont; c.fill = hfill; c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
db.row_dimensions[HRD].height = 32
widths = [8, 40, 9, 7, 10, 7, 9, 10] + [12] * len(stats) + [10, 12]
for i, w in enumerate(widths):
    db.column_dimensions[L(2 + i)].width = w
db.column_dimensions["A"].width = 3
rgf = lambda h: f"SCTM!${CL(h)}${R0}:${CL(h)}${R1}"
FAM, FINR, STR = rgf("Family"), rgf("Final Determination"), rgf("Implementation Status")
for n, (fid, fname) in enumerate(families):
    rr = HRD + 1 + n
    db.cell(rr, 2, fid); db.cell(rr, 3, fname)
    db.cell(rr, 4, f'=COUNTIFS({FAM},$B{rr},{rgf("Catalog Status")},"Active")')
    for j, h in enumerate(["LOW", "MODERATE", "HIGH", "PRIVACY"]):
        db.cell(rr, 5 + j, f'=COUNTIFS({FAM},$B{rr},{rgf(h)},"X")')
    db.cell(rr, 9, f'=COUNTIFS({FAM},$B{rr},{FINR},"Required*")')
    for j, s in enumerate(stats):
        db.cell(rr, 10 + j, f'=COUNTIFS({FAM},$B{rr},{FINR},"Required*",{STR},"{s}")')
    db.cell(rr, 10 + len(stats), f'=COUNTIFS({FAM},$B{rr},{FINR},"Required*",{STR},"")')
    pc = db.cell(rr, 11 + len(stats), f'=IF(I{rr}=0,"",J{rr}/I{rr})'); pc.number_format = "0.0%"
TR = HRD + 1 + len(families)
db.cell(TR, 2, "TOTAL"); db.cell(TR, 3, "")
for cc in range(4, 11 + len(stats)):
    db.cell(TR, cc, f"=SUM({L(cc)}{HRD+1}:{L(cc)}{TR-1})")
pc = db.cell(TR, 11 + len(stats), f'=IF(I{TR}=0,"",J{TR}/I{TR})'); pc.number_format = "0.0%"
for rr in range(HRD + 1, TR + 1):
    for cc in range(2, 12 + len(stats)):
        x = db.cell(rr, cc); x.border = box
        x.font = bold if rr == TR else bfont
        if cc != 3: x.alignment = Alignment(horizontal="center")
        if rr == TR: x.fill = PatternFill("solid", fgColor=LIGHT)
db.conditional_formatting.add(f"{L(11+len(stats))}{HRD+1}:{L(11+len(stats))}{TR}",
    __import__("openpyxl").formatting.rule.DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="63BE7B"))

# Determination summary
SR = TR + 3
db.cell(SR - 1, 2, "Final Determination Summary").font = Font(name=F, size=11, bold=True, color=NAVY)
for i, h in enumerate(["Determination", "", "Count"]):
    c = db.cell(SR, 2 + i, h); c.font = hfont; c.fill = hfill
dets = ["Required", "Required (Overlay Added)", "Required (Overlay Modified)", "Removed by Overlay", "Not Selected",
        "Conflict - Review", "Pending Categorization", "Withdrawn"]
for i, d in enumerate(dets):
    rr = SR + 1 + i
    db.cell(rr, 2, d); db.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
    db.cell(rr, 4, f'=COUNTIF({FINR},B{rr})')
    for cc in (2, 3, 4):
        db.cell(rr, cc).border = box; db.cell(rr, cc).font = bfont
    db.cell(rr, 4).alignment = Alignment(horizontal="center")
db.merge_cells(start_row=SR, start_column=2, end_row=SR, end_column=3)
db.freeze_panes = db[f"B{HRD+1}"]

wb.move_sheet("SCTM", offset=-(len(wb.sheetnames) - 4))
order = ["Instructions", "Categorization", "Overlays", "SCTM", "Dashboard"]
wb._sheets = [wb[n] for n in order]
wb.active = 3
for s in wb.worksheets:
    s.sheet_properties.tabColor = {"SCTM": NAVY, "Dashboard": "548235", "Categorization": "BF8F00", "Overlays": "BF8F00", "Instructions": "7F7F7F"}[s.title]
wb.save(OUT)
print("rows", len(rows), "R1", R1, "cols", len(cols))
