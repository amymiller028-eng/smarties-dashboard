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
        "session": "A", "modality": "S", "facilitator": "T",
        "content_relevant": "X", "fac_knowledge": "Y", "fac_engaged": "Z",
        "worthwhile": "AB", "apply_on_job": "AC", "gained_knowledge": "AD",
        "nps": "AG", "ei_dev_pct": "AJ", "confidence_pct": "AK",
        "manager_exp": None, "quote": "AL",
    },
    "TTTTeams": {
        # NOTE: header row in this sheet is misaligned with data — data follows
        # the TTTL2 layout (one column to the left of TTTL1 headers).
        "session": "A", "modality": "R", "facilitator": "S",
        "content_relevant": "W", "fac_knowledge": "X", "fac_engaged": "Y",
        "worthwhile": "AA", "apply_on_job": "AB", "gained_knowledge": "AC",
        "nps": "AF", "ei_dev_pct": "AI", "confidence_pct": "AJ",
        "manager_exp": None, "quote": "AK",
    },
    "TTTL2": {
        "session": "A", "modality": "R", "facilitator": "S",
        "content_relevant": "W", "fac_knowledge": "X", "fac_engaged": "Y",
        "worthwhile": "AA", "apply_on_job": "AB", "gained_knowledge": "AC",
        "nps": "AF", "ei_dev_pct": "AI", "confidence_pct": "AJ",
        "manager_exp": None, "quote": "AK",
    },
    "PrivateL1": {
        "session": "A", "modality": "O", "facilitator": "P",
        "content_relevant": "S", "fac_knowledge": "T", "fac_engaged": "U",
        "worthwhile": "W", "apply_on_job": "X", "gained_knowledge": "Y",
        "nps": "AB", "ei_dev_pct": "AD", "confidence_pct": "AE",
        "manager_exp": "AF", "quote": "AG",
    },
    "PrivateL2": {
        "session": "A", "modality": "O", "facilitator": "P",
        "content_relevant": "S", "fac_knowledge": "T", "fac_engaged": "U",
        "worthwhile": "W", "apply_on_job": "X", "gained_knowledge": "Y",
        "nps": "AB", "ei_dev_pct": "AD", "confidence_pct": "AE",
        "manager_exp": "AF", "quote": "AG",
    },
    "PublicL1": {
        "session": "A", "modality": "O", "facilitator": "P",
        "content_relevant": "S", "fac_knowledge": "T", "fac_engaged": "U",
        "worthwhile": "W", "apply_on_job": "X", "gained_knowledge": "Y",
        "nps": "AB", "ei_dev_pct": "AD", "confidence_pct": "AE",
        "manager_exp": None, "quote": "AG",
    },
    "Custom Programs": {
        "session": "A", "modality": "R", "facilitator": "S",
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

# --- Quote attribution (participant name + company) ---------------------------
# How names render on quotes: "first" → "Jane", "first_initial" → "Jane S.",
# "full" → "Jane Smith". On a public site "first" keeps exposure low; switch
# this one value to show fuller names.
NAME_STYLE = "first"

# Email domains that are personal, not a company — never shown as a "company" tag.
GENERIC_EMAIL_DOMAINS = {
    "gmail", "googlemail", "yahoo", "ymail", "rocketmail", "hotmail", "outlook",
    "live", "msn", "icloud", "me", "mac", "aol", "aim", "proton", "protonmail",
    "gmx", "mail", "zoho", "yandex", "comcast", "verizon", "att", "sbcglobal",
    "bellsouth", "cox", "charter",
}

# Header text used to locate the name/email/strengths columns per sheet.
# Detected by header (not fixed letters) so it survives the per-sheet column
# shifts. First match wins, so list more-specific needles first.
EXTRA_HEADER_NEEDLES = {
    "email": ["email address", "email"],
    "first": ["first name"],
    "last": ["last name"],
    "name_optional": ["name (optional)"],
    "strengths": ["facilitator's strengths", "facilitator strengths", "strengths"],
}


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


POSITIVE_WORDS = {
    "amazing", "appreciate", "appreciated", "authentic", "awesome", "beneficial",
    "best", "brilliant", "captivating", "dynamic", "effective", "empowered",
    "empowering", "engaging", "enjoyed", "enlightening", "excellent", "exceptional",
    "fabulous", "fantastic", "genuine", "good", "grateful", "great", "happy",
    "helpful", "highly", "impactful", "incredible", "inspired", "inspiring",
    "insightful", "invaluable", "love", "loved", "magnificent", "meaningful",
    "memorable", "motivated", "outstanding", "passionate", "perfect", "phenomenal",
    "powerful", "profound", "recommend", "rich", "thank", "thoughtful", "thrilled",
    "transformative", "valuable", "wonderful",
    # phrases will be matched separately below
}
POSITIVE_PHRASES = [
    "eye opening", "eye-opening", "must take", "must do", "must attend",
    "best in", "highly recommend", "well worth", "go for it", "top notch",
    "in the business",
]
EXCLUDE_PHRASES = [
    "log off", "block diary", "ask organiser", "ask organizer",
    "internal milestones", "should have been", "would have been",
    "needs to be improved", "n/a", "na ", "none ", "no comment",
    "wish there had been", "missed the mark", "didn't enjoy", "did not enjoy",
    # Advisory / cautionary tone — not testimonial material
    "get lost", "don't get lost", "dont get lost",
    "pay attention", "pay good attention",
    "be ready", "you'll need", "make sure you",
    "so you don't", "so you dont",
]

def quote_score(text):
    t = " " + text.lower() + " "
    if any(p in t for p in EXCLUDE_PHRASES):
        return -99
    pos = sum(1 for w in POSITIVE_WORDS if f" {w} " in t or f" {w}." in t or f" {w}!" in t or f" {w}," in t)
    pos += sum(2 for p in POSITIVE_PHRASES if p in t)
    return pos


def mentions_facilitator(text, facilitator):
    """True if the quote text mentions any token of the facilitator's name (>=3 chars)."""
    if not facilitator:
        return False
    t = text.lower()
    tokens = [tok for tok in facilitator.lower().split() if len(tok) >= 3]
    return any(tok in t for tok in tokens)


def company_from_email(email):
    """Derive a short company tag from an email's domain, e.g. jane@syf.com -> 'SYF'.
    Returns '' for blank/personal/generic domains (gmail, yahoo, ...)."""
    if not email or "@" not in str(email):
        return ""
    domain = str(email).split("@")[-1].strip().lower().strip(".")
    if "." not in domain:
        return ""
    label = domain.split(".")[0]
    if not label or label in GENERIC_EMAIL_DOMAINS:
        return ""
    return label.upper()


def _clean_name_token(s):
    """Keep letters/hyphens/apostrophes; drop stray punctuation and mojibake."""
    return "".join(ch for ch in s if ch.isalpha() or ch in "-'").strip("-'")


def display_name(first, last, name_optional):
    """Build the display name per NAME_STYLE. Falls back to the free-text
    'Name (optional)' field when First/Last aren't filled in. '' if unknown."""
    first = str(first).strip() if first else ""
    last = str(last).strip() if last else ""
    if not first and name_optional:
        parts = str(name_optional).strip().split()
        if parts:
            first = parts[0]
            last = parts[-1] if len(parts) > 1 else ""
    first = _clean_name_token(first)
    last = _clean_name_token(last)
    if not first:
        return ""
    if NAME_STYLE == "full":
        return (first + " " + last).strip()
    if NAME_STYLE == "first_initial" and last:
        return f"{first} {last[0]}."
    return first


def detect_extra_cols(ws):
    """Locate name/email/strengths columns by header text. Returns {key: 'G', ...}."""
    from openpyxl.utils import get_column_letter
    found = {}
    for c in range(1, min(ws.max_column, 60) + 1):
        h = ws.cell(row=1, column=c).value
        if not h:
            continue
        hl = str(h).strip().lower()
        for key, needles in EXTRA_HEADER_NEEDLES.items():
            if key in found:
                continue
            if any(n in hl for n in needles):
                found[key] = get_column_letter(c)
    return found


def best_quotes(ds, program_label, max_n=None):
    """Collect all quotes that are either clearly positive OR name the trainer.

    Pulls from two free-text sources: "share with future facilitators" and
    "facilitator's strengths". Each qualifying answer becomes a testimonial,
    attributed with the respondent's name (+ company when a corporate email
    is on file). No cap by default; excluded only on an EXCLUDE_PHRASES marker.
    """
    quotes_raw = ds.get("quote", [])
    strengths_raw = ds.get("strengths", [])
    facilitators = ds.get("facilitator", [])
    emails = ds.get("email", [])
    firsts = ds.get("first", [])
    lasts = ds.get("last", [])
    names_opt = ds.get("name_optional", [])

    def at(lst, i):
        return lst[i] if i < len(lst) else None

    n_rows = max((len(x) for x in (quotes_raw, strengths_raw, facilitators,
                                   emails, firsts, lasts, names_opt)), default=0)

    pairs = []
    for i in range(n_rows):
        fac = ""
        f = at(facilitators, i)
        if f:
            fac = str(f).strip()
        name = display_name(at(firsts, i), at(lasts, i), at(names_opt, i))
        company = company_from_email(at(emails, i))

        for src in (at(quotes_raw, i), at(strengths_raw, i)):
            if not src:
                continue
            text = str(src).strip()
            if not (25 <= len(text) <= 500): continue
            score = quote_score(text)
            if score == -99: continue  # killed by exclude phrases

            names_trainer = mentions_facilitator(text, fac)

            # Include if positive OR names the trainer (the two "good ones" criteria)
            if score < 1 and not names_trainer:
                continue

            pairs.append((score, text, fac, names_trainer, name, company))

    # Sort: positivity desc, with name-mentioning quotes promoted slightly
    pairs.sort(key=lambda x: (-(x[0] + (1 if x[3] else 0)), -len(x[1])))
    seen, picked = set(), []
    for _, text, fac, _, name, company in pairs:
        key = text[:60].lower()
        if key in seen: continue
        seen.add(key)
        picked.append({"quote": text, "program": program_label, "facilitator": fac,
                       "name": name, "company": company})
        if max_n is not None and len(picked) >= max_n:
            break
    return picked


PROGRAM_LABELS = {
    "TTTL1": "TTT L1", "TTTL2": "TTT L2", "TTTTeams": "TTT Teams",
    "PrivateL1": "Private L1", "PrivateL2": "Private L2",
    "PublicL1": "Public L1", "Custom Programs": "Custom Programs",
}


def update_trainers_tab():
    """Rewrites the Trainers tab. One OVERALL row per trainer plus a per-program
    drill-down row for each program they've taught."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from collections import defaultdict

    wb = openpyxl.load_workbook(SRC)
    if "Trainers" in wb.sheetnames:
        del wb["Trainers"]
    insert_idx = 2 if "Quarterly" in wb.sheetnames else (1 if "Summary" in wb.sheetnames else 0)
    ws = wb.create_sheet("Trainers", insert_idx)
    ws.sheet_view.showGridLines = False

    NAVY = "002D61"
    GOLD_LIGHT = "FFF4D6"
    BORDER = "E5E9F0"

    # Title + subtitle
    ws.merge_cells("A1:I1")
    ws["A1"] = "Trainer Performance — Year to Date"
    ws["A1"].font = Font(name="Calibri", size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=NAVY)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 36

    ws.merge_cells("A2:I2")
    ws["A2"] = ("OVERALL row = trainer's combined performance across all programs they've taught. "
                "Drill-down rows beneath show the same trainer per program. "
                "Refresh by running update-data.bat after adding survey rows.")
    ws["A2"].font = Font(name="Calibri", size=10, italic=True, color="555555")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center", indent=1, wrap_text=True)
    ws.row_dimensions[2].height = 30

    headers = ["Trainer", "Program / Scope", "Sessions", "Respondents",
               "NPS", "% Apply on Job", "% Engaging",
               "% Virtual", "% In-Person"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=4, column=i, value=h)
        c.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[4].height = 38

    widths = [26, 22, 11, 13, 9, 14, 14, 12, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    # Gather per-trainer × per-program data
    # trainer_data[trainer] = {program_label: {nps, apply, engage, virtual, inperson, sessions}}
    trainer_data = defaultdict(lambda: defaultdict(lambda: {
        "nps": [], "apply": [], "engage": [], "modality": [], "sessions": set()
    }))
    for sheet_name, cols in SHEET_COLS.items():
        if sheet_name not in wb.sheetnames or not cols.get("facilitator"):
            continue
        raw_ws = wb[sheet_name]
        program_label = PROGRAM_LABELS.get(sheet_name, sheet_name)
        for r in range(2, raw_ws.max_row + 1):
            fac = raw_ws[f"{cols['facilitator']}{r}"].value
            if not fac: continue
            fac = str(fac).strip()
            if not fac: continue
            g = trainer_data[fac][program_label]
            g["nps"].append(raw_ws[f"{cols['nps']}{r}"].value)
            g["apply"].append(raw_ws[f"{cols['apply_on_job']}{r}"].value)
            g["engage"].append(raw_ws[f"{cols['fac_engaged']}{r}"].value)
            mod = raw_ws[f"{cols['modality']}{r}"].value
            if mod: g["modality"].append(str(mod).strip())
            sess = raw_ws[f"{cols['session']}{r}"].value
            if sess: g["sessions"].add(str(sess).strip())

    def aggregate(groups_dict):
        """Combine all programs for a trainer into one row of stats."""
        agg = {"nps": [], "apply": [], "engage": [], "modality": [], "sessions": set()}
        for prog, g in groups_dict.items():
            agg["nps"].extend(g["nps"])
            agg["apply"].extend(g["apply"])
            agg["engage"].extend(g["engage"])
            agg["modality"].extend(g["modality"])
            agg["sessions"] |= g["sessions"]
        return agg

    def stats_row(label_program, g):
        n_resp = len(numeric(g["nps"]))
        if n_resp == 0:
            return None
        n_virtual = sum(1 for m in g["modality"] if m.lower() == "virtual")
        n_inperson = sum(1 for m in g["modality"] if m.lower() == "in person")
        n_total_mod = n_virtual + n_inperson
        v_pct = round(n_virtual / n_total_mod * 100) if n_total_mod else 0
        i_pct = round(n_inperson / n_total_mod * 100) if n_total_mod else 0
        return [
            label_program,
            len(g["sessions"]),
            n_resp,
            nps(g["nps"]),
            f'{top2box(g["apply"])}%',
            f'{top2box(g["engage"])}%',
            f'{v_pct}%' if n_total_mod else '-',
            f'{i_pct}%' if n_total_mod else '-',
        ]

    # Render: trainer (alpha order), OVERALL row, then per-program rows beneath.
    excel_row = 5
    bold_navy = Font(bold=True, color=NAVY, size=11)
    bold = Font(bold=True, size=11)
    light = Font(color="555555", size=11)
    overall_fill = PatternFill("solid", fgColor=GOLD_LIGHT)
    thin = Side(border_style="thin", color=BORDER)
    bot = Border(bottom=Side(border_style="thin", color="CCCCCC"))

    for trainer in sorted(trainer_data.keys(), key=str.lower):
        programs = trainer_data[trainer]
        # OVERALL row
        agg = aggregate(programs)
        row_vals = stats_row("OVERALL", agg)
        if not row_vals: continue
        ws.cell(row=excel_row, column=1, value=trainer).font = bold_navy
        for j, v in enumerate(row_vals, start=2):
            c = ws.cell(row=excel_row, column=j, value=v)
            c.alignment = Alignment(horizontal="center" if j > 2 else "left", vertical="center")
            c.font = bold
            c.fill = overall_fill
        ws.cell(row=excel_row, column=1).fill = overall_fill
        excel_row += 1

        # Per-program rows
        for prog in sorted(programs.keys()):
            row_vals = stats_row(prog, programs[prog])
            if not row_vals: continue
            ws.cell(row=excel_row, column=1, value="").alignment = Alignment(horizontal="left", indent=1)
            for j, v in enumerate(row_vals, start=2):
                c = ws.cell(row=excel_row, column=j, value=v)
                c.alignment = Alignment(horizontal="center" if j > 2 else "left", vertical="center", indent=1 if j == 2 else 0)
                c.font = light
            excel_row += 1
        # blank separator row
        excel_row += 1

    ws.freeze_panes = "A5"
    wb.save(SRC)
    print(f"Trainers tab updated: {len(trainer_data)} trainers")


def main():
    wb = openpyxl.load_workbook(SRC, data_only=True)

    raw = {}
    for sheet_name, cols in SHEET_COLS.items():
        if sheet_name not in wb.sheetnames:
            print(f"  skipping missing sheet: {sheet_name}")
            continue
        ws = wb[sheet_name]
        ds = collect_sheet(ws, cols)
        # Attribution + extra quote source, located by header text.
        extra = detect_extra_cols(ws)
        for key in ("email", "first", "last", "name_optional", "strengths"):
            ds[key] = col_values(ws, extra[key]) if extra.get(key) else []
        raw[sheet_name] = ds

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
        for q in best_quotes(ds, label_for[sheet_name]):
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

    # Refresh the Trainers tab in the workbook
    update_trainers_tab()
    print("\nQuick summary of computed views:")
    for k, v in views.items():
        if v.get("type") == "refresher":
            print(f"  {k:18s} participants={v['participants']:>3}  growth={v['confidenceGrowth']}  valuable%={v['pctRatedValuable']}")
        else:
            extra = f"  noManager%={v.get('noManagerExpectationsPct','-')}" if "noManagerExpectationsPct" in v else ""
            print(f"  {k:18s} NPS={v['nps']:>3}  N={v['participants']:>3}  applyOnJob={v['topBox']['applyOnJob']}%{extra}")


if __name__ == "__main__":
    main()
