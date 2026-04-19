"""
Adds (or rebuilds) a 'Summary' tab in Program-Evals-2026.xlsx with live
formulas that calculate every number the smarties dashboard needs.

Run this once after initial setup, or again any time the workbook structure
changes (e.g. when TTTL2 or PrivateL2 sheets get added).

Creates a timestamped backup of the source workbook before modifying.
"""
import shutil
from datetime import datetime
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = Path(r"C:\Users\Amy Miller - TS\OneDrive - TalentSmart\Program-Evals-2026.xlsx")

# Column letters differ between TTTL1 and PrivateL1 — map each metric.
TTTL1_COLS = {
    "respondent_id": "B",
    "modality": "S",
    "content_relevant": "X",
    "fac_knowledge": "Y",
    "fac_engaged": "Z",
    "challenged": "AA",
    "worthwhile": "AB",
    "apply_on_job": "AC",
    "gained_knowledge": "AD",
    "nps": "AG",
    "ei_dev_pct": "AJ",
    "confidence_pct": "AK",
}
PRIVATEL1_COLS = {
    "respondent_id": "B",
    "modality": "O",
    "content_relevant": "S",
    "fac_knowledge": "T",
    "fac_engaged": "U",
    "challenged": "V",
    "worthwhile": "W",
    "apply_on_job": "X",
    "gained_knowledge": "Y",
    "nps": "AB",
    "ei_dev_pct": "AD",
    "confidence_pct": "AE",
}

NAVY = "002D61"
GREEN = "0ACC8B"
GOLD = "DFA351"
BG_SOFT = "F7F8FB"
BG_INPUT = "FFF9DB"  # soft yellow for manual-input cells

def nps_formula(sheet, col):
    return (
        f'=IFERROR(ROUND((COUNTIF({sheet}!{col}:{col},">=9")'
        f'-COUNTIF({sheet}!{col}:{col},"<=6"))'
        f'/COUNT({sheet}!{col}:{col})*100,0),0)'
    )

def combined_nps_formula(pairs):
    promoters = "+".join([f'COUNTIF({s}!{c}:{c},">=9")' for s, c in pairs])
    detractors = "+".join([f'COUNTIF({s}!{c}:{c},"<=6")' for s, c in pairs])
    total = "+".join([f'COUNT({s}!{c}:{c})' for s, c in pairs])
    return f"=IFERROR(ROUND(({promoters}-({detractors}))/({total})*100,0),0)"

def t2b_formula(sheet, col):
    return (
        f'=IFERROR(ROUND(COUNTIF({sheet}!{col}:{col},">=4")'
        f'/COUNT({sheet}!{col}:{col})*100,0),0)'
    )

def combined_t2b_formula(pairs):
    top = "+".join([f'COUNTIF({s}!{c}:{c},">=4")' for s, c in pairs])
    total = "+".join([f'COUNT({s}!{c}:{c})' for s, c in pairs])
    return f"=IFERROR(ROUND(({top})/({total})*100,0),0)"

def avg_formula(sheet, col):
    return f'=IFERROR(ROUND(AVERAGE({sheet}!{col}:{col}),0),0)'

def combined_avg_formula(pairs):
    total = "+".join([f'SUM({s}!{c}:{c})' for s, c in pairs])
    n = "+".join([f'COUNT({s}!{c}:{c})' for s, c in pairs])
    return f"=IFERROR(ROUND(({total})/({n}),0),0)"

def count_formula(sheet, col):
    return f'=COUNTA({sheet}!{col}:{col})-1'

def combined_count_formula(pairs):
    parts = "+".join([f'(COUNTA({s}!{c}:{c})-1)' for s, c in pairs])
    return f"={parts}"

def countif_text(sheet, col, value):
    return f'=COUNTIF({sheet}!{col}:{col},"{value}")'

def combined_countif_text(pairs, value):
    parts = "+".join([f'COUNTIF({s}!{c}:{c},"{value}")' for s, c in pairs])
    return f"={parts}"


def build_summary(wb):
    if "Summary" in wb.sheetnames:
        del wb["Summary"]

    ws = wb.create_sheet("Summary", 0)
    ws.sheet_view.showGridLines = False

    # Column widths
    widths = {"A": 42, "B": 14, "C": 14, "D": 14, "E": 2, "F": 42, "G": 16}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    section_font = Font(name="Calibri", size=11, bold=True, color=NAVY)
    label_font = Font(name="Calibri", size=11, color="333333")
    value_font = Font(name="Calibri", size=11, bold=True, color="000000")
    input_font = Font(name="Calibri", size=11, bold=True, color="8B6F00")
    title_font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
    sub_font = Font(name="Calibri", size=10, italic=True, color="555555")
    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    navy_fill = PatternFill("solid", fgColor=NAVY)
    soft_fill = PatternFill("solid", fgColor=BG_SOFT)
    input_fill = PatternFill("solid", fgColor=BG_INPUT)

    # Title
    ws.merge_cells("A1:G1")
    ws["A1"] = "Program Impact — Summary"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws["A1"].fill = navy_fill
    ws.row_dimensions[1].height = 38

    ws.merge_cells("A2:G2")
    ws["A2"] = (
        "Auto-calculated from the raw response sheets. Yellow cells are manual inputs. "
        "Copy the data.json values on the right into GitHub each week."
    )
    ws["A2"].font = sub_font
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[2].height = 22

    # Headers row
    header_row = 4
    headers = ["Metric", "TTT L1", "Private L1", "Combined"]
    for i, h in enumerate(headers):
        c = ws.cell(row=header_row, column=i + 1, value=h)
        c.font = bold
        c.fill = navy_fill
        c.alignment = center
    ws.row_dimensions[header_row].height = 24

    pairs_all = [("TTTL1", TTTL1_COLS["nps"]), ("PrivateL1", PRIVATEL1_COLS["nps"])]

    json_entries = []  # (label, cell ref, json path)

    def write_section(title, start_row):
        ws.cell(row=start_row, column=1, value=title).font = section_font
        ws.cell(row=start_row, column=1).fill = soft_fill
        for col in range(2, 5):
            ws.cell(row=start_row, column=col).fill = soft_fill
        ws.row_dimensions[start_row].height = 20
        return start_row + 1

    def write_metric(row, label, ttt_val, priv_val, comb_val, is_input=False, json_paths=None):
        ws.cell(row=row, column=1, value=label).font = label_font
        ws.cell(row=row, column=1).alignment = left
        for col, val in ((2, ttt_val), (3, priv_val), (4, comb_val)):
            cell = ws.cell(row=row, column=col, value=val)
            cell.alignment = center
            if is_input and val is None:
                cell.fill = input_fill
                cell.font = input_font
            else:
                cell.font = value_font
        if json_paths:
            # json_paths is dict like {"ttt": "views.ttt.participants", ...}
            for key, path in json_paths.items():
                col_letter = {"ttt": "B", "private": "C", "all": "D"}[key]
                json_entries.append((path, f"{col_letter}{row}"))

    r = write_section("OVERVIEW", 5)
    # Respondent counts (auto)
    write_metric(r, "Respondents",
                 count_formula("TTTL1", TTTL1_COLS["respondent_id"]),
                 count_formula("PrivateL1", PRIVATEL1_COLS["respondent_id"]),
                 combined_count_formula([("TTTL1", TTTL1_COLS["respondent_id"]),
                                         ("PrivateL1", PRIVATEL1_COLS["respondent_id"])]),
                 json_paths={"ttt": "views.ttt.participants",
                             "private": "views.private.participants",
                             "all": "views.all.participants"})
    r += 1
    # Manual: sessions
    write_metric(r, "Sessions delivered (enter manually)", None, None, "=B{0}+C{0}".format(r),
                 is_input=True,
                 json_paths={"ttt": "views.ttt.sessions",
                             "private": "views.private.sessions",
                             "all": "views.all.sessions"})
    ws.cell(row=r, column=2).value = 1   # seed values
    ws.cell(row=r, column=3).value = 1
    ws.cell(row=r, column=2).fill = input_fill
    ws.cell(row=r, column=3).fill = input_fill
    ws.cell(row=r, column=2).font = input_font
    ws.cell(row=r, column=3).font = input_font
    r += 1
    # Manual: clients
    write_metric(r, "Client companies (enter manually)", None, None, "=B{0}+C{0}".format(r),
                 is_input=True,
                 json_paths={"ttt": "views.ttt.clients",
                             "private": "views.private.clients",
                             "all": "views.all.clients"})
    ws.cell(row=r, column=2).value = 1
    ws.cell(row=r, column=3).value = 1
    ws.cell(row=r, column=2).fill = input_fill
    ws.cell(row=r, column=3).fill = input_fill
    ws.cell(row=r, column=2).font = input_font
    ws.cell(row=r, column=3).font = input_font
    r += 1
    # Modality
    write_metric(r, "Virtual respondents",
                 countif_text("TTTL1", TTTL1_COLS["modality"], "Virtual"),
                 countif_text("PrivateL1", PRIVATEL1_COLS["modality"], "Virtual"),
                 combined_countif_text([("TTTL1", TTTL1_COLS["modality"]),
                                        ("PrivateL1", PRIVATEL1_COLS["modality"])], "Virtual"),
                 json_paths={"ttt": "views.ttt.modality.virtual",
                             "private": "views.private.modality.virtual",
                             "all": "views.all.modality.virtual"})
    r += 1
    write_metric(r, "In person respondents",
                 countif_text("TTTL1", TTTL1_COLS["modality"], "In person"),
                 countif_text("PrivateL1", PRIVATEL1_COLS["modality"], "In person"),
                 combined_countif_text([("TTTL1", TTTL1_COLS["modality"]),
                                        ("PrivateL1", PRIVATEL1_COLS["modality"])], "In person"),
                 json_paths={"ttt": "views.ttt.modality.inPerson",
                             "private": "views.private.modality.inPerson",
                             "all": "views.all.modality.inPerson"})
    r += 2

    r = write_section("NPS", r)
    write_metric(r, "Net Promoter Score",
                 nps_formula("TTTL1", TTTL1_COLS["nps"]),
                 nps_formula("PrivateL1", PRIVATEL1_COLS["nps"]),
                 combined_nps_formula([("TTTL1", TTTL1_COLS["nps"]),
                                       ("PrivateL1", PRIVATEL1_COLS["nps"])]),
                 json_paths={"ttt": "views.ttt.nps",
                             "private": "views.private.nps",
                             "all": "views.all.nps"})
    r += 2

    r = write_section("TOP-2-BOX % (rated 4 or 5 out of 5)", r)
    t2b_metrics = [
        ("Content was relevant to my job", "content_relevant", "contentRelevant"),
        ("Facilitator's knowledge enhanced learning", "fac_knowledge", "facilitatorKnowledge"),
        ("Facilitator kept me engaged", "fac_engaged", "facilitatorEngaging"),
        ("Appropriately challenged", "challenged", None),
        ("Worthwhile investment of time", "worthwhile", "worthwhileInvestment"),
        ("Will apply on the job", "apply_on_job", "applyOnJob"),
        ("Gained new knowledge", "gained_knowledge", "gainedKnowledge"),
    ]
    for label, key, json_key in t2b_metrics:
        paths = None
        if json_key:
            paths = {
                "ttt": f"views.ttt.topBox.{json_key}",
                "private": f"views.private.topBox.{json_key}",
                "all": f"views.all.topBox.{json_key}",
            }
        write_metric(r, label,
                     t2b_formula("TTTL1", TTTL1_COLS[key]),
                     t2b_formula("PrivateL1", PRIVATEL1_COLS[key]),
                     combined_t2b_formula([("TTTL1", TTTL1_COLS[key]),
                                           ("PrivateL1", PRIVATEL1_COLS[key])]),
                     json_paths=paths)
        r += 1
    r += 1

    r = write_section("EI DEVELOPMENT ATTRIBUTION", r)
    write_metric(r, "Avg % of EI growth credited to training",
                 avg_formula("TTTL1", TTTL1_COLS["ei_dev_pct"]),
                 avg_formula("PrivateL1", PRIVATEL1_COLS["ei_dev_pct"]),
                 combined_avg_formula([("TTTL1", TTTL1_COLS["ei_dev_pct"]),
                                       ("PrivateL1", PRIVATEL1_COLS["ei_dev_pct"])]),
                 json_paths={"ttt": "views.ttt.eiDevelopmentAttributed",
                             "private": "views.private.eiDevelopmentAttributed",
                             "all": "views.all.eiDevelopmentAttributed"})
    r += 1
    write_metric(r, "Avg confidence in estimate",
                 avg_formula("TTTL1", TTTL1_COLS["confidence_pct"]),
                 avg_formula("PrivateL1", PRIVATEL1_COLS["confidence_pct"]),
                 combined_avg_formula([("TTTL1", TTTL1_COLS["confidence_pct"]),
                                       ("PrivateL1", PRIVATEL1_COLS["confidence_pct"])]),
                 json_paths={"ttt": "views.ttt.confidenceInEstimate",
                             "private": "views.private.confidenceInEstimate",
                             "all": "views.all.confidenceInEstimate"})
    r += 2

    # data.json copy-paste panel (column F/G)
    panel_start = 4
    ws.cell(row=panel_start, column=6, value="data.json value").font = bold
    ws.cell(row=panel_start, column=6).fill = navy_fill
    ws.cell(row=panel_start, column=6).alignment = center
    ws.cell(row=panel_start, column=7, value="Paste this").font = bold
    ws.cell(row=panel_start, column=7).fill = navy_fill
    ws.cell(row=panel_start, column=7).alignment = center

    for i, (path, cell_ref) in enumerate(json_entries):
        row_i = panel_start + 1 + i
        ws.cell(row=row_i, column=6, value=path).font = label_font
        v = ws.cell(row=row_i, column=7, value=f"={cell_ref}")
        v.font = value_font
        v.alignment = center

    # Freeze top two rows + header
    ws.freeze_panes = "A5"


def main():
    if not SRC.exists():
        raise SystemExit(f"Source file not found: {SRC}")

    ts = datetime.now().strftime("%Y%m%d-%H%M%S")
    backup = SRC.with_name(SRC.stem + f"-backup-{ts}.xlsx")
    shutil.copy2(SRC, backup)
    print(f"Backup created: {backup.name}")

    wb = openpyxl.load_workbook(SRC)
    build_summary(wb)
    wb.save(SRC)
    print(f"Summary tab added to: {SRC.name}")


if __name__ == "__main__":
    main()
