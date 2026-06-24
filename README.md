# Smarties Dashboard — TalentSmartEQ

Interactive impact dashboard that reads from `data.json`. Update the JSON each week (or run the helper script) and the page refreshes live.

## Files

- `index.html` — page structure
- `styles.css` — brand styling
- `dashboard.js` — interactivity (tabs, drill-down, view rendering)
- `data.json` — **what the live site reads** (the only file you need to upload to GitHub)
- `compute-data.py` — Python script that reads your Excel and rebuilds `data.json`
- `update-data.bat` — double-click this to run `compute-data.py`
- `build-summary.py` — (optional) adds a Summary tab to the Excel workbook

## Tab structure

- **All Programs** — combined across everything except Refresher
- **Train the Trainer** ▾ Summary · Level 1 · Level 2 · Teams
- **Private Program** ▾ Summary · Level 1 · Level 2
- **Public Program**
- **Custom Programs**
- **Refresher** (different metrics — confidence growth and value rating)

## Weekly update workflow (~1 minute)

1. Paste new survey responses into the appropriate raw sheets in `Program-Evals-2026.xlsx` (TTTL1, PrivateL1, etc.)
2. Save and close the Excel file
3. Double-click **`update-data.bat`** in this folder. It rebuilds `data.json` from scratch
4. Upload the updated `data.json` to your GitHub repo:
   - Open repo → click `data.json` → pencil icon (✏️)
   - Delete all the existing content, paste the new file contents
   - OR: in your repo's main page, click **Add file → Upload files**, drag `data.json` in, commit
5. Live site updates within ~60 seconds

## Adding a new program type

If you add a brand-new program with its own sheet (e.g. "TTTL3" or "Coaching"), tell Claude. The new sheet's column layout needs to be added to `compute-data.py`'s `SHEET_COLS` mapping, and a new tab needs to be added to the dashboard.

## NPS quick formula

- **Promoter** = score 9 or 10
- **Passive** = 7 or 8 (ignored in the math)
- **Detractor** = 0 through 6
- **NPS = (Promoters ÷ Total × 100) − (Detractors ÷ Total × 100)**, rounded
- Anything **above 70 is considered world-class**.

## Top-2-box quick formula

For any 1–5 question:
**% Top-2-box = (count of 4s + count of 5s) ÷ total responses × 100**, rounded

## Refresher metrics

The Refresher survey doesn't ask the same questions, so its tab shows different stats:

- **Confidence growth** — average confidence on a 4-point scale (1=Slightly · 2=Moderately · 3=Confident · 4=Fully) before vs. after
- **% improved confidence** — share who moved up at least one level
- **% rated valuable or extremely valuable**

## Manager-expectations metric

Three program types ask "Has your manager communicated their expectations about how you will apply this training on the job?" — Private L1, Private L2, and Custom Programs.

The dashboard shows the **% who answered NO** as a contextual orange/gold tile on those views only. Hidden everywhere else (because the question wasn't asked).

## Testimonials / quotes

Quotes are pulled automatically by `compute-data.py` from **two** free-text columns on each sheet:

1. *"What would you want to share with future facilitators…"* (the original source)
2. *"What were your facilitator's strengths?"* (added — this is where most of the rich, named comments live)

How they're chosen and shown:

- **Included** if the comment reads positively *or* names the trainer. Advisory/negative comments are filtered out (see `EXCLUDE_PHRASES`).
- **Ranked** by substance, not just buzzwords — multi-sentence, longer, first-person reflections rise to the top (see `substance_bonus`). This keeps thoughtful comments from getting buried.
- **No cap** — every qualifying quote is shown, ordered best-first.
- **Attribution:** each quote shows the participant's **first name**, taken from the First/Last columns or, if blank, the free-text *"Name (optional)"* field. Style is controlled by `NAME_STYLE` in `compute-data.py` (`first` / `first_initial` / `full`).
- **Company tag** (e.g. "Jane, SYF") appears only when a **corporate email** is on file — generic domains (gmail, yahoo, …) are skipped. Currently no emails are in the workbook, so only names show; tags light up automatically once corporate emails are added.

## EQ-growth figure (confidence-adjusted)

The side card shows **% of EQ growth participants credit to this program** plus their **% confidence** in that estimate. The **ⓘ** tooltip explains how to use it: multiply the two for a *confidence-adjusted* figure (e.g. 71% × 84% ≈ 60%) — a conservative number that holds up to scrutiny. This is the credibility/isolation step from the **Phillips ROI Methodology**. The talk track tells reps to lead with the simple "X% credited" line and keep the adjusted number in reserve. Wording lives in `talkTrack()` in `dashboard.js`.

## Privacy note — names are now shown

This dashboard shows aggregated metrics plus **participant first names on quotes** (and a company abbreviation when a corporate email is on file). Emails themselves and individual scores are **not** exposed. Because it's a public site, keep names to first-name only (`NAME_STYLE = "first"`) unless you have a reason to show more.
