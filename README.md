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

## Safe to share

This dashboard shows only aggregated metrics and anonymous quotes. No respondent names, emails, or individual scores are exposed.
