"""
Reads Program-Evals-2026.xlsx and writes the dashboard's data.json.

Run any time you've added new survey responses to the workbook.

    python compute-data.py
"""
import json
from collections import Counter
from datetime import datetime
from pathlib import Path
import openpyxl

SRC = Path(r"C:\Users\Amy Miller - TS\OneDrive - TalentSmart\Program-Evals-2026.xlsx")
OUT = Path(__file__).parent / "data.json"

# Per-sheet column mappings. None means the sheet doesn't have that question.
SHEET_COLS = {
    "TTTL1": {
        "session": "A", "modality": "S",
        "content_relevant": "X", "fac_knowledge": "Y", "fac_engaged": "Z",
        "worthwhile": "AB", "apply_on_job": "AC", "gained_knowledge": "AD",
        "nps": "AG", "ei_dev_pct": "AJ", "confidence_pct": "AK",
        "manager_exp": None, "quote": "AL",
    },
    "TTTTeams": {
        # NOTE: header row in this sheet is misaligned with data — data follows
        # the TTTL2 layout (one column to the left of TTTL1 headers).
        "session": "A", "modality": "R",
        "content_relevant": "W", "fac_knowledge": "X", "fac_engaged": "Y",
        "worthwhile": "AA", "apply_on_job": "AB", "gained_knowledge": "AC",
        "nps": "AF", "ei_dev_pct": "AI", "confidence_pct": "AJ",
        "manager_exp": None, "quote": "AK",
    },
    "TTTL2": {
        "session": "A", "modality": "R",
        "content_relevant": "W", "fac_knowledge": "X", "fac_engaged": "Y",
        "worthwhile": "AA", "apply_on_job": "AB", "gained_knowledge": "AC",
        "nps": "AF", "ei_dev_pct": "AI", "confidence_pct": "AJ",
        "manager_exp": None, "quote": "AK",
    },
    "PrivateL1": {
        "session": "A", "modality": "O",
        "content_relevant": "S", "fac_knowledge": "T", "fac_engaged": "U",
        "worthwhile": "W", "apply_on_job": "X", "gained_knowledge": "Y",
        "nps": "AB", "ei_dev_pct": "AD", "confidence_pct": "AE",
        "manager_exp": "AF", "quote": "AG",
    },
    "PrivateL2": {
        "session": "A", "modality": "O",
        "content_relevant": "S", "fac_knowledge": "T", "fac_engaged": "U",
        "worthwhile": "W", "apply_on_job": "X", "gained_knowledge": "Y",
        "nps": "AB", "ei_dev_pct": "AD", "confidence_pct": "AE",
        "manager_exp": "AF", "quote": "AG",
    },
    "PublicL1": {
        "session": "A", "modality": "O",
        "content_relevant": "S", "fac_knowledge": "T", "fac_engaged": "U",
        "worthwhile": "W", "apply_on_job": "X", "gained_knowledge": "Y",
        "nps": "AB", "ei_dev_pct": "AD", "confidence_pct": "AE",
        "manager_exp": None, "quote": "AG",  # column AF exists but is all blank
    },
    "Custom Programs": {
        "session": "A", "modality": "R",
        "content_relevant": "V", "fac_knowledge": "W", "fac_engaged": "X",
        "worthwhile": "Z", "apply_on_job": "AA", "gained_knowledge": "AB",
        "nps": "AF", "ei_dev_pct": "AI", "confidence_pct": "AJ",
        "manager_exp": "AK", "quote": "AL",
    },
}

CONFIDENCE_MAP = {
    "slightly confident": 1,
    "moderately confident": 2,
    "confident": 3,
    "fully confident": 4,
}
VALUABLE_TOP = {"very valuable", "extremely valuable"}


def col_values(ws, col):
    return [ws[f"{col}{r}"].value for r in range(2, ws.max_row + 1)]

def numeric(values):
    out = []
    for v in values:
        if v is None or v == "":
            continue
        try:
            out.append(float(v))
        except (TypeError, ValueError):
            pass
    return out

def nps(scores):
    nums = numeric(scores)
    if not nums:
        return 0
    promoters = sum(1 for s in nums if s >= 9)
    detractors = sum(1 for s in nums if s <= 6)
    return round((promoters - detractors) / len(nums) * 100)

def top2box(scores):
    nums = numeric(scores)
    if not nums:
        return 0
    return round(sum(1 for s in nums if s >= 4) / len(nums) * 100)

def avg(values):
    nums = numeric(values)
    return round(sum(nums) / len(nums)) if nums else 0

def count_text(values, target):
    return sum(1 for v in values if v and str(v).strip().lower() == target.lower())

def distinct_count(values):
    return len({v for v in values if v not in (None, "")})


def collect_sheet(ws, cols):
    """Pull all relevant column values from a sheet at once."""
    return {key: col_values(ws, c) if c else [] for key, c in cols.items()}


def standard_view(label, datasets):
    """Aggregate one or more dataset dicts into a standard view object."""
    def cat(key):
        out = []
        for ds in datasets:
            out.extend(ds.get(key, []))
        return out

    sessions_all = [s for s in cat("session") if s]
    nps_all = cat("nps")
    participants = sum(len(numeric(ds.get("nps", []))) for ds in datasets)
    if participants == 0:
        # fall back to count of rows that have anything in modality
        participants = sum(sum(1 for v in ds.get("modality", []) if v) for ds in datasets)

    view = {
        "label": label,
        "type": "standard",
        "nps": nps(nps_all),
        "participants": participants,
        "sessions": distinct_count(sessions_all),
        "clients": distinct_count(sessions_all),  # 1 client per session as proxy; editable
        "eiDevelopmentAttributed": avg(cat("ei_dev_pct")),
        "confidenceInEstimate": avg(cat("confidence_pct")),
        "topBox": {
            "applyOnJob": top2box(cat("apply_on_job")),
            "gainedKnowledge": top2box(cat("gained_knowledge")),
            "worthwhileInvestment": top2box(cat("worthwhile")),
            "contentRelevant": top2box(cat("content_relevant")),
            "facilitatorKnowledge": top2box(cat("fac_knowledge")),
            "facilitatorEngaging": top2box(cat("fac_engaged")),
        },
        "modality": {
            "virtual": sum(count_text(ds.get("modality", []), "Virtual") for ds in datasets),
            "inPerson": sum(count_text(ds.get("modality", []), "In person") for ds in datasets),
        },
    }

    # Manager expectations: % NO across datasets that asked the question.
    me_values = []
    for ds in datasets:
        me_values.extend(v for v in ds.get("manager_exp", []) if v not in (None, ""))
    if me_values:
        no_count = sum(1 for v in me_values if str(v).strip().lower() == "no")
        view["noManagerExpectationsPct"] = round(no_count / len(me_values) * 100)
        view["managerExpectationsResponses"] = len(me_values)

    return view


def refresher_view(ws):
    before = col_values(ws, "L")
    after = col_values(ws, "M")
    value_rating = col_values(ws, "AB")
    sessions = col_values(ws, "A")

    def to_num(v):
        if v is None: return None
        return CONFIDENCE_MAP.get(str(v).strip().lower())

    before_n = [n for n in (to_num(v) for v in before) if n is not None]
    after_n = [n for n in (to_num(v) for v in after) if n is not None]

    paired = [(to_num(b), to_num(a)) for b, a in zip(before, after)]
    paired = [(b, a) for b, a in paired if b is not None and a is not None]
    moved_up = sum(1 for b, a in paired if a > b)

    valuable = [v for v in value_rating if v]
    valuable_top = sum(1 for v in valuable if str(v).strip().lower() in VALUABLE_TOP)

    return {
        "label": "Refresher",
        "type": "refresher",
        "participants": len([v for v in sessions if v]),
        "sessions": distinct_count(sessions),
        "confidenceBefore": round(sum(before_n) / len(before_n), 2) if before_n else 0,
        "confidenceAfter": round(sum(after_n) / len(after_n), 2) if after_n else 0,
        "confidenceGrowth": round((sum(after_n)/len(after_n)) - (sum(before_n)/len(before_n)), 2) if before_n and after_n else 0,
        "pctMovedUpInConfidence": round(moved_up / len(paired) * 100) if paired else 0,
        "pctRatedValuable": round(valuable_top / len(valuable) * 100) if valuable else 0,
        "confidenceScale": "1=Slightly · 2=Moderately · 3=Confident · 4=Fully",
    }


def best_quotes(ds, program_label, max_n=4):
    """Pick a few non-trivial quotes from a dataset."""
    quotes = [str(q).strip() for q in ds.get("quote", []) if q]
    # prefer 30-260 char quotes (skip "n/a", super long rambles)
    good = [q for q in quotes if 30 <= len(q) <= 260]
    good.sort(key=lambda q: (-len(q),))  # longer first within range
    seen, picked = set(), []
    for q in good:
        key = q[:60].lower()
        if key in seen: continue
        seen.add(key)
        picked.append({"quote": q, "program": program_label})
        if len(picked) >= max_n: break
    return picked


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)

    raw = {}
    for sheet_name, cols in SHEET_COLS.items():
        if sheet_name not in wb.sheetnames:
            print(f"  skipping missing sheet: {sheet_name}")
            continue
        raw[sheet_name] = collect_sheet(wb[sheet_name], cols)

    views = {}

    # All Programs (everything except Refresher)
    views["all"] = standard_view("All Programs", list(raw.values()))

    # Train the Trainer family
    views["ttt-summary"] = standard_view("Train the Trainer — Summary",
        [raw["TTTL1"], raw["TTTL2"], raw["TTTTeams"]])
    views["ttt-l1"] = standard_view("Train the Trainer — Level 1", [raw["TTTL1"]])
    views["ttt-l2"] = standard_view("Train the Trainer — Level 2", [raw["TTTL2"]])
    views["ttt-teams"] = standard_view("Train the Trainer — Teams", [raw["TTTTeams"]])

    # Private family
    views["private-summary"] = standard_view("Private Program — Summary",
        [raw["PrivateL1"], raw["PrivateL2"]])
    views["private-l1"] = standard_view("Private Program — Level 1", [raw["PrivateL1"]])
    views["private-l2"] = standard_view("Private Program — Level 2", [raw["PrivateL2"]])

    # Public, Custom, Refresher
    views["public-l1"] = standard_view("Public Program", [raw["PublicL1"]])
    views["custom"] = standard_view("Custom Programs", [raw["Custom Programs"]])
    views["refresher"] = refresher_view(wb["Refresher"])

    # Quotes
    testimonials = []
    label_for = {
        "TTTL1": "Train the Trainer L1", "TTTL2": "Train the Trainer L2",
        "TTTTeams": "TTT for Teams", "PrivateL1": "Private Program L1",
        "PrivateL2": "Private Program L2", "PublicL1": "Public Program",
        "Custom Programs": "Custom Program",
    }
    view_for = {
        "TTTL1": "ttt-l1", "TTTL2": "ttt-l2", "TTTTeams": "ttt-teams",
        "PrivateL1": "private-l1", "PrivateL2": "private-l2",
        "PublicL1": "public-l1", "Custom Programs": "custom",
    }
    for sheet_name, ds in raw.items():
        for q in best_quotes(ds, label_for[sheet_name], max_n=3):
            q["view"] = view_for[sheet_name]
            testimonials.append(q)

    payload = {
        "meta": {
            "lastUpdated": datetime.now().strftime("%Y-%m-%d"),
            "reportingPeriod": "Q1 2026 to date",
        },
        "views": views,
        "testimonials": testimonials,
    }

    OUT.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    print(f"\nWrote: {OUT}")
    print("\nQuick summary of computed views:")
    for k, v in views.items():
        if v.get("type") == "refresher":
            print(f"  {k:18s} participants={v['participants']:>3}  growth={v['confidenceGrowth']}  valuable%={v['pctRatedValuable']}")
        else:
            extra = f"  noManager%={v.get('noManagerExpectationsPct','-')}" if "noManagerExpectationsPct" in v else ""
            print(f"  {k:18s} NPS={v['nps']:>3}  N={v['participants']:>3}  applyOnJob={v['topBox']['applyOnJob']}%{extra}")


if __name__ == "__main__":
    main()
