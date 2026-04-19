# Smarties Dashboard — TalentSmartEQ

Interactive impact dashboard that reads from `data.json`. Update the JSON each week; the page refreshes live.

## Files

- `index.html` — page structure
- `styles.css` — brand styling
- `dashboard.js` — interactivity (tabs, chart)
- `data.json` — **the only file you edit weekly**

## Preview locally

Just double-click `index.html` or drag it into a browser tab. If the dashboard shows a red error bar, open it via a tiny local server (browsers block `fetch` on raw files in some setups):

- On Windows, from the project folder in PowerShell: `py -m http.server 8000` → open http://localhost:8000

## First-time setup on GitHub Pages

1. Go to https://github.com/new — create a new **public** repo named something like `smarties-dashboard`.
2. On the new repo page, click **"uploading an existing file"** (gray link near the top of the empty repo).
3. Drag `index.html`, `styles.css`, `dashboard.js`, and `data.json` into the upload box. Click **Commit changes**.
4. Go to the repo's **Settings → Pages** (left sidebar).
5. Under **Source**, pick **Deploy from a branch**. Branch: `main`, folder: `/ (root)`. Click **Save**.
6. Wait ~60 seconds, then refresh the Pages settings. You'll see your site URL (like `https://yourusername.github.io/smarties-dashboard/`).
7. Open that URL — your dashboard is live.
8. (Optional) Add a link to that URL on a SharePoint page so your team finds it where they already go.

## Weekly update workflow (~2 minutes)

1. Open your `Program-Evals-2026.xlsx` workbook.
2. Recalculate the metrics for each view (All / TTT / Private). If you want, keep a hidden `Summary` sheet that computes them automatically.
3. In your GitHub repo, click `data.json` → pencil icon (✏️ top right) to edit.
4. Update the numbers (see field guide below). Change `meta.lastUpdated` to today's date in `YYYY-MM-DD` format.
5. Scroll down, click **Commit changes**. Done — the live site updates within ~1 minute.

## data.json field guide

```
meta.lastUpdated             The date you refreshed the numbers
views.all                    Combined (TTT + Private) metrics
views.ttt                    Train the Trainer only
views.private                Private Programs only

For each view:
  nps                        Net Promoter Score (−100 to 100)
                             = %Promoters(9-10) − %Detractors(0-6)
  participants               Total respondents across all sessions
  sessions                   Number of sessions delivered
  clients                    Number of client companies
  eiDevelopmentAttributed    Average % of EI growth participants credit
                             to training (from the % development question)
  confidenceInEstimate       Average of the "confidence in estimate" question

  topBox: all values are the % of respondents who rated 4 or 5 on a 1–5 scale.
    applyOnJob               "I will be able to apply what I learned on the job"
    gainedKnowledge          "I have gained new knowledge"
    worthwhileInvestment     "This course was a worthwhile investment of my time"
    contentRelevant          "The program content was relevant to my job"
    facilitatorKnowledge     "My learning was enhanced by facilitator's knowledge"
    facilitatorEngaging      "The facilitator kept me engaged"

  modality.virtual           Count of respondents in virtual sessions
  modality.inPerson          Count of respondents in in-person sessions

trend[]                      One entry per month: { label, nps }
testimonials[]               Quotes to display. Tag program as
                             "Train the Trainer" or "Private Program".
```

## NPS quick formula

- **Promoter** = score 9 or 10
- **Passive** = 7 or 8 (ignored in the math)
- **Detractor** = 0 through 6
- **NPS = (Promoters ÷ Total × 100) − (Detractors ÷ Total × 100)**, rounded

## Top-2-box quick formula

For any 1–5 rated question:
**% Top-2-box = (count of 4s + count of 5s) ÷ total responses × 100**, rounded

## Safe to share

This dashboard shows only aggregated metrics and anonymous quotes. No respondent names, emails, or individual scores are exposed.
