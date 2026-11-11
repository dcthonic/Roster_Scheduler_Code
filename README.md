# Google Sheets roster scheduler

This folder contains the Apps Script scheduler (`Code.gs`) and its importable starter workbook (`Scheduler_Template_GoogleSheets.xlsx`). Use this guide for setup, routine use, and troubleshooting. The spreadsheet's **Instructions** and **Settings** tabs remain the source for live staffing requirements and configurable values.

## First-time setup

1. In Google Sheets, create a blank spreadsheet and choose **File > Import > Upload**. Select `Scheduler_Template_GoogleSheets.xlsx` and import it.
2. Open **Extensions > Apps Script**. Replace the editor contents with all of `Code.gs`, save, return to Sheets, and reload.
3. Choose **Roster > Generate schedule**. On first run, complete Google's authorization prompt, then run it again.

## Run a cycle

1. Update **Names**, **Settings**, and **Unavailability**.
2. On **Names**, optionally enter personal AM, PM, or Graveyard workdays/week in columns E:G (whole numbers 0–7). Blank uses the matching Settings default (or existing Perm-AM setting). In QUOTA mode the value sets the quota; in EXACT mode it sets the person's maximum. OFF/OSL/LL absences still reduce the affected week's quota. Enhanced Graveyard has separate no-cap rules.
3. On **Unavailability**, enter hard absences in columns A:D. Each date blocks assignment. An exact reason of `OFF`, `OSL`, or `LL` also reduces the member's expected shifts for that week by one and lowers the fair-share target. Other reasons do not reduce the quota. Enter soft preferred days off in columns F:G; list multiple dates separated by commas in `DD/MM` format (for example `02/11, 05/11`). The scheduler tries to honor these but can override them to meet staffing requirements.
4. In **Settings**, set **Plan Graveyard week (Yes/No)**. **Yes** includes Graveyard in a normal full-cycle generation. **No** leaves Graveyard pending while AM/PM are generated; use **Enhanced Graveyard** later. This setting affects new generations only.
5. Choose **Roster > Generate schedule** and review the completion message and **Log** sheet. Each run creates date-named Roster and Points sheets. Older output sheets are kept hidden. Where a prior roster exists, generation requires at least one assignment to differ.

The **Points Log** ranks members by their mean fair-share result across cycles in which they appear. Newcomers are not counted for earlier cycles. It also shows mean raw points and mean fair-share target. The log refreshes after relevant edits or generation; use **Roster > Refresh Points Log** to refresh it manually.

## Graveyard and roster imports

**Roster > Enhanced Graveyard** edits the open cycle in place: it retains existing Graveyard assignments, fills to 10 people per day with 2–3 `#` members, and applies no Graveyard weekly or weekend cap. It updates the matching Points sheet and Points Log; AM/PM weeks and other cycles stay as they are. Points Log percentages help prioritize members below 100% where constraints allow. Resolve any hard-absence conflict with an existing Graveyard assignment first. If the member list changed, generate a regular cycle before using this command.

The stable **Current Roster** tab follows the newest date-named roster, including manual edits. In another spreadsheet, import from it once with:

```text
=IMPORTRANGE("SOURCE_SPREADSHEET_ID","'Current Roster'!A1:Z32")
```

Allow access when prompted. The destination's imported values update with the source; formatting is managed separately.

## Menu commands

- **Generate schedule** — Creates a new cycle's Roster and Points sheets, updates the Points Log, and archives older output sheets as hidden tabs.
- **Enhanced Graveyard** — Fills the open cycle's Graveyard week and updates its matching Points sheet in place.
- **Fix roster colours** — Normalizes assignment-cell colours and formatting on roster sheets. Schedule entries other than `X`/`AM` use red fill with bold black text; summary and requirements rows retain their normal bold formatting.
- **Refresh Points Log** — Rebuilds the member ranking and averages from eligible Points sheets.

Other Apps Script functions are internal helpers for input validation, availability and preference parsing, points and fair-share calculations, optimization, output formatting, archiving, and schedule checks.

## Troubleshooting

| Issue | What to check |
|---|---|
| Roster menu is missing | Save the Apps Script project and reload the spreadsheet. |
| Required sheet or member is missing | Sheet names must be **Settings**, **Names**, and **Unavailability**. Check spelling in Unavailability against Names. |
| Not enough available staff or `#` members | Review hard absences and staffing requirements in Settings for the affected day. Enhanced Graveyard needs 2–3 `#` members each day. |
| Enhanced Graveyard reports an assignment conflict | Correct the hard absence or clear the conflicting existing Graveyard assignment, then retry. |
| Run is slow or times out | Set **Balance points when scheduling** to **No**, ensure the solver time limit is reasonable, and retry. Check **Log** for the limiting constraint. |
| Roster colours look wrong | Run **Roster > Fix roster colours**. |
| Points Log is stale | Run **Roster > Refresh Points Log**. |
| Current Roster import shows an error or old values | Confirm the source tab is named **Current Roster**, approve IMPORTRANGE access, and allow Sheets time to refresh. |

## Verification note

The scheduling logic was checked against a stand-in for Google's optimizer using November data, and the formulas were checked there. The script has not been verified in a live Google Sheets run; runtime and formula behavior may differ. If a live run fails, use the **Log** sheet details to diagnose it.
