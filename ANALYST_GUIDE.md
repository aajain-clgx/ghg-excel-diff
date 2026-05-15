# GHG Excel QA Tool — Analyst Guide

**Version 1.0 · Internal Use Only**

---

## Overview

The GHG Excel QA Tool compares two versions of a GHG inventory Excel file — such as Draft 1 vs Draft 2, or Prior Year vs Current Year — and shows you exactly what changed, site by site and column by column. It runs entirely on your laptop; no data is sent to any server or cloud service.

### What you can do with it

| Capability | Details |
|---|---|
| Identify changed sites | Every site with at least one value change is surfaced, sorted by severity |
| Understand magnitude | Absolute and percentage delta for every numeric change |
| Annotate each change | Mark each flag as **Expected**, **Error — fixing**, or **Needs Review** |
| Add reviewer notes | Free-text note per site+column (e.g. "Updated to EPA 2024 emission factors") |
| Export a QA Log | One-click Excel export with all changes, severities, and your annotations |
| Work across multi-sheet files | Choose which sheet to compare if a file contains multiple tabs |
| Template conformance check | Upload a reference template to verify required columns are present |

---

## Getting Started

### System requirements

- Windows 10 or later
- Excel files must be saved as `.xlsx` (not `.xls` or `.csv`)
- Edge or Chrome browser (opens automatically)

### First-time setup

Run **`setup_windows.bat`** once. It will download the Python packages the tool needs (takes 1–2 minutes). You only need to do this once.

### Launching the tool

1. Double-click **`run_app.bat`**
2. A black command window opens — leave it running
3. Your browser opens automatically at `http://localhost:8510`

To stop the tool, close the black command window or press `Ctrl+C` inside it. You can close and reopen the browser tab at any time without losing your work.

---

## Step-by-Step Walkthrough

### Step 1 — Project Setup

This is the first screen you see when the tool loads.

| Field | What to enter |
|---|---|
| **Your name** | Used to sign the exported QA Log |
| **Project label** | A short description, e.g. "GHG Inventory Q4 2025" |
| **V1 file** | The baseline/original file (the older or approved version) |
| **V1 label** | Short name for V1, e.g. "Draft 1" or "Prior Year" |
| **V2 file** | The revised file to check against V1 |
| **V2 label** | Short name for V2, e.g. "Draft 2" or "Current Year" |
| **Template file** *(optional)* | An `.xlsx` template whose first-row headers define the expected column set |

**Column template:** If you upload a template, the tool will warn you if any expected columns are missing from either V1 or V2, and flag any extra columns that shouldn't be there. This is useful for verifying that no columns were accidentally added or deleted between drafts.

**Multi-sheet files:** If either uploaded file contains more than one sheet, a sheet selector will appear before the comparison runs. Choose the tab that holds the inventory data.

Click **"Load Files →"** to proceed.

---

### Step 2 — Column Matching

The tool automatically pairs columns between V1 and V2 by normalising names (case-insensitive, whitespace-trimmed). The matching table shows:

| Status | Meaning |
|---|---|
| ✅ Matched | Same column found in both files — will be diffed |
| ⚠️ V1 only | Column exists in V1 but not V2 — will be shown as a structural removal |
| ⚠️ V2 only | Column exists in V2 but not V1 — will be shown as a structural addition |

**Manually pairing columns:** If a column was renamed between V1 and V2 (e.g. "EF NG" → "EF Natural Gas"), the tool won't match them automatically. Use the **"Manually pair columns"** section to select the V1 column and V2 column and click **"Add pair"**. Manual pairs are included in the diff.

**Site ID column:** Select the column that uniquely identifies each site (e.g. "Facility ID", "Site Code", "Plant ID"). The tool uses this to align rows between files.

**Materiality threshold:** Changes above this percentage are flagged as **High Impact 🔴**. The default is **5 %**. Adjust to match your team's QA policy. A live preview next to the slider shows how many sites would be flagged High, Warning, or Minor at the selected threshold, so you can fine-tune before running the full diff.

Click **"Run Diff →"** to proceed.

---

### Step 3 — Diff Results

This is the main review screen.

#### Summary banner

At the top of the page you will see a sentence like:

> *5 of 42 sites changed — 2 high-impact, 1 warning, 2 minor.*

This gives an instant overview before you dive into the detail.

#### Severity colour codes

| Icon | Severity | Trigger |
|---|---|---|
| 🔴 **High Impact** | Numeric change strictly above the materiality threshold |
| 🟡 **Warning** | A value went blank (or appeared from blank); a text field changed; a numeric change where V1 was zero; an emission factor column changed while the corresponding activity data did **not** change |
| 🟢 **Minor** | Numeric change at or below the threshold |

The results table is sorted: High Impact sites appear first, then Warnings, then Minor.

#### Filtering

Use the **"Filter by severity"** selector above the table to narrow the view to 🔴 High Impact, 🟡 Warning, or 🟢 Minor sites if you only want to focus on one category.

#### Reviewing a site

Click any row in the results table to expand it. You will see one card per changed column showing:

- Column name
- V1 value → V2 value
- Absolute delta and percentage change
- A hint if the change looks like an isolated emission-factor update (the activity data didn't change alongside it)

#### Annotating changes

Each column card has inline controls:

1. **Disposition** — choose one:
   - ✅ **Expected** — the change is correct and intentional
   - 🔧 **Error — fixing** — the change is wrong and needs correction
   - ❓ **Needs Review** — you're not sure yet; flag for follow-up
2. **Note** — free-text field (e.g. "Switched to EPA 2024 emission factors, approved by team lead")
3. Click **"Save"** to store the annotation

Annotations save immediately to a local sidecar file. You can close the browser, reopen it, and all your annotations will still be there (as long as the command window is still running).

#### Structural changes

If sites or columns appeared or disappeared between V1 and V2, they are listed separately at the bottom of the results in **yellow Structural rows**. These are informational — they don't need annotations.

---

### Exporting the QA Log

Once you have finished reviewing, click **"Generate & Download QA Log"**.

- If any **High Impact 🔴** sites still have no annotation, you will see a warning listing them. You can still export — this is a reminder, not a blocker.
- The downloaded file is a formatted `.xlsx` with:
  - A header block (project, analyst, run date, V1/V2 labels, materiality threshold, summary sentence)
  - One data row per changed column per site, colour-coded by severity
  - Structural change rows at the end
  - Freeze panes on the header row and an auto-filter so you can sort/filter in Excel

Attach this file to your QA report or send it to your reviewer.

---

## Severity Reference

### When a change is High Impact 🔴

A numeric value changed by more than the materiality threshold (default 5 %). For example: V1 = 100, V2 = 112 → +12 % → High Impact at a 5 % threshold.

### When a change is Warning 🟡

Any of the following apply, regardless of magnitude:

- A value that existed in V1 is blank in V2 (value disappeared)
- A value that was blank in V1 has a value in V2 (value appeared)
- A text column changed (non-numeric comparison)
- A numeric column changed but V1 was zero (percentage can't be computed)
- An emission-factor-style column changed, but the activity data column for the same site did **not** change — this often indicates a factor was revised without a corresponding activity update, which warrants a check

### When a change is Minor 🟢

A numeric change at or below the materiality threshold, with no blank or text anomalies.

---

## Emission Factor Hint

If the tool detects that a column whose name contains words like "EF", "emission factor", or "factor" changed, while the paired activity column (activity, consumption, usage, MMBtu, MWh, kWh, litres, gallons) for the same site did **not** change, it will display a hint:

> *Emission factor changed but corresponding activity data appears unchanged. Verify the factor source and version used.*

This is a prompt to confirm the factor revision was intentional and documented, not a hard error.

---

## Tips and Best Practices

- **Your source files are never modified.** The tool only reads V1 and V2; it never writes to them.
- **Work through High Impact sites first.** The results table is sorted by severity so you can process the most important flags before moving to minor ones.
- **Use notes liberally.** Even a short note like "EPA 2024 update" or "Corrected double-count" makes the QA Log much more useful for reviewers and auditors.
- **Reopen at any time.** Your annotations are persisted locally. As long as the command window is still running, refreshing the browser (F5) restores exactly where you left off.
- **Threshold preview.** Before running the diff, drag the materiality slider in Step 2 to see a live count of how many sites would fall into each severity bucket. This helps you choose the right threshold for your engagement.
- **Multi-reviewer workflow.** If two analysts are reviewing different site groups, each can run their own copy of the tool and annotate their portion. Merge the sidecar JSON files (they are human-readable, one key per site::column) or have one person consolidate and re-export.

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Blank screen or `ModuleNotFoundError` | Run `setup_windows.bat` again |
| Browser doesn't open automatically | Navigate manually to `http://localhost:8510` in Edge or Chrome |
| "Port already in use" error | Close the old command window first. Alternatively, edit `run_app.bat` and change `8510` to any unused port number. |
| "File is not a valid .xlsx file" | Save the file as **Excel Workbook (.xlsx)**, not `.xls` or `.csv`. Close the file in Excel before uploading. |
| Column matching looks wrong | Use **"Manually pair columns"** in Step 2 to correct mismatched or renamed columns |
| Site ID column not auto-detected | Select it manually from the dropdown in Step 2 |
| Annotations disappeared | Check that the black command window is still running. If it was closed and restarted, re-upload your files — the sidecar file will be found automatically if the project folder is the same. |
| Export button is greyed out | At least one file must have been loaded and a diff run first |

---

## Frequently Asked Questions

**Q: Does the tool connect to the internet?**
No. Everything runs locally on your laptop. No files, annotations, or results leave your machine.

**Q: What if my two files have different numbers of sites?**
Sites that exist in V1 but not V2 are flagged as **structural removals**. Sites that exist in V2 but not V1 are flagged as **structural additions**. Both appear in the Structural rows section of the results and in the exported QA Log.

**Q: Can I compare files with different column sets?**
Yes. Matched columns are diffed; unmatched columns are shown as structural additions/removals. You can manually pair any renamed columns so they are diffed rather than shown as structural changes.

**Q: What if a site ID appears more than once in a file?**
The tool will run and flag diffs, but duplicate site IDs may produce unexpected results. De-duplicate the source data before running the comparison if possible.

**Q: Can I change the materiality threshold after running the diff?**
Yes. Go back to Step 2 using the breadcrumb navigation at the top of the page, adjust the threshold slider, and click **"Run Diff →"** again. Your annotations from the previous run are preserved.

**Q: What file formats are supported?**
Only `.xlsx` (Excel Open XML) is supported. If your file is `.xls`, open it in Excel and **Save As → Excel Workbook (.xlsx)** first.

**Q: Is the QA Log auditable?**
Yes. The exported `.xlsx` includes the run date, analyst name, project label, both file labels, the materiality threshold used, and a full record of every change with its severity and your annotation. It is designed to be attached directly to a QA report or audit file.

---

## Glossary

| Term | Definition |
|---|---|
| **V1** | The baseline file — the original, approved, or prior-year version |
| **V2** | The revised file — the newer version being checked |
| **Materiality threshold** | The percentage change above which a numeric difference is treated as High Impact |
| **Disposition** | Your judgement on a change: Expected, Error, or Needs Review |
| **Annotation** | A disposition plus an optional free-text note stored for a specific site+column flag |
| **Site ID column** | The column used to match rows between V1 and V2 (e.g. Facility ID) |
| **Structural change** | A site or column that appeared or disappeared between V1 and V2 |
| **Sidecar file** | A local `.json` file where annotations are automatically saved |
| **QA Log** | The exported `.xlsx` file containing all diff results and annotations |
| **EF** | Emission Factor — a coefficient used to convert activity data into GHG emissions |

---

*For support, contact your GHG team lead or the person who distributed this tool.*
