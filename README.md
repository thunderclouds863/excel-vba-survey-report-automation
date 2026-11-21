# Excel VBA Survey Report Automation

An Excel macro workbook that processes a customer survey on education media (posters) at car dealer branches. It collects survey responses, calculates the survey rate per branch and dealer group each month, and exports the reports as PDF.

Built during my internship at the Customer Assistance Center, PT Toyota-Astra Motor (2025–2026).

> **Data removed.** This repository holds a template only. All survey responses, customer names, vehicle plate numbers and survey results were deleted before publishing, including the pivot cache records. Sheets keep their headers, branch list, layout, formulas, pivot table and buttons.

## Problem

Survey responses came in from a Google Form and were compiled by hand. For each dealer group, someone copied data from a pivot table, fixed the columns, recalculated the survey rate and exported a PDF, one by one. A full report run took about 1 hour.

## Workflow

1. **Add Data** imports the latest responses.
2. **Refresh** updates the pivot table and syncs the monthly numbers into the report.
3. **Olah Data** builds the survey-rate percentage sheet.
4. **Export to PDF** produces the reports.

Result: the report run went from about 1 hour to 5–10 minutes, with consistent formatting across all reports.

## Macros behind the dashboard buttons

| Macro | What it does |
|---|---|
| `AddData` | Picks an Excel or CSV file, loads it into an array, detects duplicates by Timestamp + survey date + branch, asks once whether to replace them, then merges and sorts by timestamp |
| `RefreshSemuaPivot` | Refreshes all pivot tables, then runs `SinkronkanSemuaData` |
| `SinkronkanSemuaData` | Reads the month columns from the pivot, adds or shifts month columns in CUSTOMER SURVEY RATE to match, and copies the monthly counts per branch |
| `OlahDataSurveyRate_Fixed` | Builds the SURVEY RATE sheet: marks each branch 1 or 0 per month (surveyed or not), then calculates each dealer's rate as surveyed branches ÷ total branches, with highlighting for 0 values |
| `ExportPDF_Dinamis_SatuHalaman` | Treats each **bold** row in CUSTOMER SURVEY RATE as a dealer group, builds a one-page report for it, and saves one PDF per dealer group. Asks for the education media type and the cut-off label |
| `ExportPDF_SatuFile_BoldSaja_HinggaBulanTerakhir` | Exports one PDF with the survey rate of all dealer groups, up to the latest month |
| `ClearAllData` | Empties the Data sheet after a confirmation |
| `Save` | Saves a copy of the workbook as a macro-enabled file |

## Workbook structure

| Sheet | Purpose |
|---|---|
| Dashboard Main Menu | Buttons for the workflow above |
| Data | Survey responses table (`Data`) |
| MONTHLY SURVEY ACHIEVEMENT | Pivot table: responses per branch per month |
| CUSTOMER SURVEY RATE | Monthly counts, number of outlets, outlets surveyed and survey rate per branch and dealer group |
| SURVEY RATE | Monthly survey-rate percentage per dealer group |
| CUSTOMER SURVEY RATE_PULAU | Hidden helper sheet |

## Files

- `Template_Pengolahan_Data_Survey_Meduk.xlsm`: the workbook template with macros
- `src/Module1.bas`: import, sync and per-dealer PDF export
- `src/Module2.bas`: pivot refresh and save
- `src/Module4.bas`: all-dealer PDF export
- `src/Module5.bas`: survey-rate calculation
- `src/deletemenu.frm`, `src/frmDeleteData.frm`, `src/frmDeleteMedia.frm`: delete-data forms from an earlier version, not linked to the dashboard

## Notes

- Enable macros when opening the file. Built for Microsoft Excel on Windows.
- PDFs are saved in a `Report…` folder next to the workbook.

## Tools

Excel VBA, Excel tables, pivot tables, formulas and conditional formatting.
