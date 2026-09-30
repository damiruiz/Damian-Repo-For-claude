# Job search operating system: LinkedIn + recruiters in one place

## The problem (audit, 2026-09-26)

| What existed | Rows | State |
|---|---|---|
| `LinkedIn Recruiter Tracker - Damian Ruiz` | 155 | Active. One row **per message**, with ~50 different free-text statuses |
| `LinkedIn Recruiter Tracker — Damian Ruiz (2026)` | 27 | Abandoned in May. Duplicates the tracker above |
| `Damian — Job Search Tracker (Jobs)` | 25 | Last written 8/16. 22 roles still marked "new", 6 weeks old |
| `Damian — Job Search Tracker (Outreach)` | 0 | Header row only |
| `ARCHIVE …` copies | 4 | Clutter that shows up in Drive search |

What was going wrong:

1. **The best conversations weren't tracked at all.** The tracker only captured LinkedIn inbound. Email
   threads where things actually moved (an agency offer, two RTRs, a panel interview, a repeat recruiter
   asking for a call) were in none of the sheets.
2. **Drafts rot.** 17 replies were marked "Draft prepared" and about 15 sat unsent in Gmail. Several were
   for fully remote roles at or above the pay floor.
3. **There was no "next action" or "due date"**, so nothing surfaced when it went stale: RTRs silent for 6
   weeks, an interview with no logged outcome.
4. **Message log vs. relationships.** Because each row is a message, the same recruiter appears 2–3 times and
   there's no view of who your best recruiters are.
5. **Duplicate-req risk.** Two agencies pitched what looks like the same Contract Management Specialist IV
   req. If both submit you, the client disqualifies you.

## The fix: one workbook, four working tabs

`Damian - Job Search HQ` (built by `tools/build_job_search_hq.py`):

| Tab | Grain | Purpose |
|---|---|---|
| **Start Here** | — | Live counts (reply owed, waiting, submitted, interviewing, offers, overdue) plus today's action list |
| **Pipeline** | 1 row per opportunity | The only tab that drives your day. Fixed Stage list, Next action, Due date, Days idle, annualized pay, and an "at least $85K?" check. Overdue rows turn red |
| **Recruiters** | 1 row per person | CRM tiers: **A** (placed/submitted you, or a specialist in your field), **B** (pitches target roles), **C** (off-target), **DNC**. Next nurture date is automatic: A every 6 weeks, B every 90 days |
| **Inbox Log** | 1 row per message | Add-only history of every recruiter message, including declines. Has a normalized Stage column |
| **Job Leads** | 1 row per posting | Roles you sourced yourself (HiringCafe/Jobright). Promote a lead to Pipeline when you act on it |
| **Templates** | — | Seven replies: remote-only decline, clarify, RTR yes, 5-day nudge, post-interview nudge, rejection, A-tier nurture |

Stage list (dropdown): `1 New - triage` → `2 Reply owed (me)` → `3 Waiting on recruiter` →
`4 Submitted / RTR` → `5 Interviewing` → `6 Offer` → `7 Accepted`, plus `Nurture` and five `Closed - …` outcomes.

## Rules

- **24-hour reply rule.** Every recruiter gets an answer within one business day, even if it's a decline.
  Templates make that a 2-minute job.
- **Three automatic declines:** not fully remote, under $85K ($41/hr W2), or Enbridge.
- **One agency per req.** Before signing an RTR, search Pipeline for the same title and client.
- **Follow-up clock:** after a submission, nudge at 5 business days, again at 10, then mark `Closed - Went cold`.
  After an interview, send a thank-you the same day and nudge the recruiter at day 7.
- **Every open Pipeline row has a Due date.** Red rows are overdue and get worked first.
- **Nurture the A-tier.** A short "still looking" note every 6 weeks. Agency recruiters place the people they
  remember.

## Weekly rhythm

| When | Time | Do |
|---|---|---|
| Daily (weekday AM) | 10 min | Log new recruiter messages to the Inbox Log. Put anything worth pursuing in Pipeline. Clear `Reply owed (me)` |
| Mon / Wed / Fri | 30 min | Job Leads sweep → pick 1–3 → tailor, find a referral, apply (per the personal-recruiter skill) |
| Friday | 20 min | Clear red rows. Move or close anything idle 14+ days. Send nurture notes that are due |

## Automation hooks

- The daily personal-recruiter run should **read and write this workbook** instead of the four legacy sheets:
  Phase 0 reads Pipeline and Recruiters, Phase 3 appends to Job Leads, and outreach goes to Recruiters and Pipeline.
- **Gmail → Inbox Log.** Search `from:(linkedin.com OR inmail) newer_than:1d` plus replies from known
  recruiter domains. Append one row per new message.
- **Apollo (optional).** Use `apollo_people_match` to fill missing recruiter emails for A/B-tier rows.
  A "recruiter nurture" Apollo sequence could send the 6-week A-tier note automatically. Both spend credits
  and send from your account, so turn them on deliberately.
- **LinkedIn stays manual.** Drafts only, pasted by hand. Automating LinkedIn risks restricting the account.

## Rebuilding the workbook

```bash
pip install openpyxl
# data/ is gitignored: it holds recruiter contact details
python tools/build_job_search_hq.py --recruiters data/recruiters.csv --data data/data.json \
    --out "Damian - Job Search HQ.xlsx"
python tools/build_job_search_hq.py --out template.xlsx   # blank template
```

Upload the .xlsx to Google Drive with "Convert uploads" on. Dropdowns, conditional formatting and formulas
carry over to Google Sheets.
