# Excel VBA Survey Report Automation

An Excel macro workbook for a customer survey on education media (posters) at car dealer branches. It collects survey responses into one table and turns them into a survey-rate report per dealer group, exported as PDF.

Built during my internship at the Customer Assistance Center, PT Toyota-Astra Motor (2025–2026).

> **Data removed.** This repository holds a template only. All survey responses, customer names, interviewer names, photo links, results and the pivot cache were deleted before publishing. Sheets keep their headers, layout, formulas and buttons.

## Problem

Survey responses came in from a Google Form and were compiled by hand. For each dealer group, someone copied data from a pivot table, fixed the columns, recalculated the survey rate and exported a PDF, one by one. A full report run took about 1 hour.

## What the workbook does

| Feature | Macro | How it works |
|---|---|---|
| Import responses | `AddData` | Picks an Excel or CSV file, loads it into an array, detects duplicates by Timestamp + survey date + branch, asks once whether to replace them, merges and sorts by timestamp |
| Delete data | `DeleteDataMenu` | Deletes by timestamp range, by education media type, by keyword, or by selected rows |
| Clear all data | `ClearAllData` | Empties the Data sheet after a confirmation |
| Export reports | `ExportPDF` | Treats each **bold** row as a dealer group header, collects the rows under it, builds a 2-page report (Customer Survey Rate + Monthly Survey Achievement) with a grand total, and saves one PDF per dealer group into a `Report` folder |

Result: the report run went from about 1 hour to 5–10 minutes, with consistent formatting across all branch reports.

## Workbook structure

| Sheet | Purpose |
|---|---|
| Dashboard Main Menu | Buttons for Add Data, Delete Data, Clear Data, Olah Data and Save |
| Data | Survey responses table (`Form_Responses13`) |
| Form Responses 1 | Raw Google Form export |
| CUSTOMER SURVEY RATE | Survey rate per branch, formula-based, with the Export to PDF button |
| MONTHLY SURVEY ACHIEVEMENT | Monthly achievement per dealer |
| udah belum / udah belum (2) | Tracking of branches that have or have not run the survey |

## Files

- `Survey_Meduk_Aksesoris_Template.xlsm`: the workbook template with macros
- `src/Module1.bas`: main macros (import, delete, export)
- `src/frmDeleteData.frm`: user form for choosing which data to delete

## Notes

- `ExportPDF` reads from a sheet named `export`. Add or rename a sheet to `export` with the report layout before running it.
- The Monthly Survey Achievement pivot table was removed together with its cached data. Recreate it from the `Form_Responses13` table if needed.
- The delete-by-media and delete-by-keyword options open a form named `frmDeleteDataFast`, while the form in this workbook is `frmDeleteData`. Rename one of them so the names match.
- The Olah Data button calls a `FillterData` macro that is not part of this workbook.
- Enable macros when opening the file. Built for Microsoft Excel on Windows.

## Tools

Excel VBA, Excel tables, formulas and pivot tables.
