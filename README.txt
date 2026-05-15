GHG EXCEL QA TOOL
=================
Version: 1.0  |  For internal use only

WHAT THIS TOOL DOES
-------------------
Compares two versions of a GHG inventory Excel file (e.g. Draft 1 vs Draft 2)
and shows you exactly what changed, site by site and column by column.

You can:
  • See which sites changed and by how much (%)
  • Flag changes as Expected, Error, or Needs Review
  • Add notes explaining each change
  • Export a signed-off QA Log Excel file

─────────────────────────────────────────────────────────────────────────────
FIRST TIME SETUP (do this once)
─────────────────────────────────────────────────────────────────────────────

1. Double-click:  setup_windows.bat
2. Wait for it to finish (1-2 minutes, downloads Python packages)
3. You're done — setup only needs to be run once.

If setup fails with a network error, ask your IT support desk to install:
  Python 3.9 (or later) + the packages in requirements.txt

─────────────────────────────────────────────────────────────────────────────
HOW TO USE THE TOOL
─────────────────────────────────────────────────────────────────────────────

STARTING THE TOOL
  1. Double-click run_app.bat
  2. A black command window will open — this is normal, leave it running
  3. Your browser (Edge or Chrome) will open automatically
     If it doesn't, type this in your browser:  http://localhost:8510

STOPPING THE TOOL
  Close the black command window, or press Ctrl+C inside it.
  You can safely close the browser tab at any time.

─────────────────────────────────────────────────────────────────────────────
STEP-BY-STEP WALKTHROUGH
─────────────────────────────────────────────────────────────────────────────

STEP 1 — Project Setup
  • Enter your name and a project label (e.g. "GHG Inventory Q4 2025")
  • Upload the BASELINE file (V1 — the older or approved version)
  • Give it a short label  (e.g. "Draft 1" or "Prior Year")
  • Upload the REVISED file (V2 — the newer version to check)
  • Give it a short label  (e.g. "Draft 2" or "Current Year")
  • (Optional) Upload a column template file to check for missing columns
  • Click "Load Files →"

STEP 2 — Column Matching
  • The tool automatically matches columns between V1 and V2
  • Review the matched pairs — they should all make sense
  • If a column wasn't matched, you can manually pair it
  • Select the "Site ID Column" — this is the column with facility codes
    (e.g. "Facility ID", "Site Code", "Plant ID")
  • Set the Materiality Threshold — changes above this % are flagged red
    Default is 5%. Adjust if your team uses a different threshold.
  • Click "Run Diff →"

STEP 3 — Review Results
  • The summary banner at the top shows total changes at a glance
  • The table lists all sites that changed, sorted by severity:
      🔴  High Impact  — change is above the materiality threshold
      🟡  Possible Error — a value went blank, or an EF changed without activity
      🟢  Minor         — small change, below threshold
  • Click any row to see exactly which columns changed and by how much

  ANNOTATING CHANGES
  • For each site, choose a disposition:
      ✅ Expected   — the change is correct and intentional
      🔧 Error      — the change is wrong and needs to be fixed
      ❓ Needs Review — you're not sure yet, flag for follow-up
  • Add a note explaining the change (e.g. "Updated to EPA 2024 EF")
  • Click "Save Annotation"

  EXPORTING THE QA LOG
  • Once you've reviewed the sites, click "Generate & Download QA Log"
  • This saves an Excel file with all changes, severities, and your notes
  • You'll be warned if any High Impact (🔴) sites are unannotated
  • Open the downloaded file in Excel and attach it to your QA report

─────────────────────────────────────────────────────────────────────────────
TIPS
─────────────────────────────────────────────────────────────────────────────

• Your files are NOT uploaded to any server — everything runs on your laptop.
• Annotations are saved automatically as you click Save on each site.
• You can close the browser and reopen it without losing your work
  (as long as the black command window is still open).
• If the browser tab goes blank, just refresh it (F5).
• The tool works with any .xlsx file — it does not modify your source files.

─────────────────────────────────────────────────────────────────────────────
TROUBLESHOOTING
─────────────────────────────────────────────────────────────────────────────

"ModuleNotFoundError" or blank screen on startup
  → Run setup_windows.bat again

Browser doesn't open automatically
  → Manually type  http://localhost:8510  in Edge or Chrome

"Port already in use" error
  → Another copy is already running. Close the old command window first.
  → Or change the port by editing run_app.bat (change 8510 to another number)

"File is not a valid .xlsx file"
  → Make sure the file is saved as Excel Workbook (.xlsx), not .xls or .csv
  → If the file is open in Excel, close it and try again

Column matching looks wrong
  → Use the "Manually pair columns" section in Step 2 to fix any mismatches

─────────────────────────────────────────────────────────────────────────────
SUPPORT
─────────────────────────────────────────────────────────────────────────────

For issues, contact your GHG team lead or the person who distributed this tool.
