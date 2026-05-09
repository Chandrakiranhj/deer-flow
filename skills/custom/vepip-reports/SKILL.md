---
name: vepip-reports
description: Generate funder-ready Word, PDF and PowerPoint reports from VEPIP project data using the bundled premium template scripts. Use whenever the user asks to "generate a report", "create a deck", "export Q3 narrative", or anything that produces a polished document for download.
license: private
---

# VEPIP Premium Funder Reports

You produce **Word (.docx)**, **PDF (.pdf)** and **PowerPoint (.pptx)** funder reports from live project data. The visual design (Vision Empower's dark-brown + gold + braille-texture brand) is already baked into three pre-built Python scripts shipped with this skill — you do **not** write rendering code yourself. Your only job is to assemble the data JSON and invoke the right script.

Output files written to `/mnt/user-data/outputs/` are surfaced as download links by the VEPIP UI via `present_files`.

---

## Workflow (every report)

1. **Resolve the project.** The user message contains a `<context>` block with `project_id`, `project_name`, `user_email`, and `today`. The orchestrator usually pre-embeds `project_context` and `report_data` JSON in the prompt — use those values directly. Only call `get_project_context` / `get_report_data` if the JSON is not already provided.

2. **Write the data file.** Build a single JSON file at `/mnt/user-data/workspace/report_data.json` with this top-level shape:

   ```json
   {
     "project": { ... },         // see "Project schema" below
     "draft": "Optional narrative text used for challenges/next-steps extraction",
     "periodStart": "YYYY-MM-DD",
     "periodEnd":   "YYYY-MM-DD"
   }
   ```

   Use `write_file` with the exact path above. Do **not** write Python code — the rendering scripts already exist.

3. **Run the matching builder.** One bash command, no piping, no heredocs:

   | Format | Command |
   |---|---|
   | PPTX | `python /mnt/skills/custom/vepip-reports/scripts/build_pptx.py --data /mnt/user-data/workspace/report_data.json --output /mnt/user-data/outputs/<filename>.pptx --report-type <quarterly\|full> --period-start <YYYY-MM-DD> --period-end <YYYY-MM-DD>` |
   | DOCX | `python /mnt/skills/custom/vepip-reports/scripts/build_docx.py --data /mnt/user-data/workspace/report_data.json --output /mnt/user-data/outputs/<filename>.docx --report-type <quarterly\|full> --period-start <YYYY-MM-DD> --period-end <YYYY-MM-DD>` |
   | PDF  | `python /mnt/skills/custom/vepip-reports/scripts/build_pdf.py --data /mnt/user-data/workspace/report_data.json --output /mnt/user-data/outputs/<filename>.pdf --report-type <quarterly\|full> --period-start <YYYY-MM-DD> --period-end <YYYY-MM-DD>` |

   The scripts auto-install `python-pptx` / `python-docx` / `reportlab` / `Pillow` on first run via `pip`. Stdout will end with `WROTE: <path> (<bytes>)` on success. Any other output ending with a Python traceback means the run failed — read the traceback, fix the data JSON (never the script), and re-run.

4. **Verify** with `ls /mnt/user-data/outputs/` — confirm the file exists and its size is non-zero.

5. **Present** with `present_files` paths=["/mnt/user-data/outputs/<filename>"].

6. **End your reply** with exactly one line in this form (and nothing after it):

   ```
   📎 [<filename>](artifact://outputs/<filename>)
   ```

---

## Project schema (the `project` field of `report_data.json`)

All fields are optional — the scripts substitute sensible defaults. Use the exact field names below.

```jsonc
{
  "name":            "STEM Teacher Enablement",
  "funderName":      "Wipro Foundation",
  "grantAmount":     1800000,         // INR
  "startDate":       "2024-04-01",
  "endDate":         "2025-03-31",
  "states":          ["Karnataka", "Tamil Nadu"],
  "summary":         "Two-sentence project pitch for the cover.",
  "status":          "active",
  "funderLogoUrl":   "https://...png",  // optional, embedded in cover

  "deliverables": [
    {
      "title":    "Teachers Trained",
      "target":   450,
      "achieved": 282,
      "unit":     "teachers",
      "dueDate":  "2025-03-31",
      "status":   "in_progress",        // completed | in_progress | overdue | not_started
      "description": "..."
    }
  ],

  "budgets": [
    { "name": "Travel",   "approvedAmount": 300000, "spentAmount": 245000 },
    { "name": "Materials","approvedAmount": 600000, "spentAmount": 410000 }
  ],

  "activities": [
    {
      "title":           "School visit, Mysore",
      "activityDate":    "2025-02-14",
      "state":           "Karnataka",
      "location":        "Mysore",
      "teachersReached": 48,
      "studentsReached": 200,
      "schoolsReached":  3,
      "notes":           "Conducted teacher training at three schools.",
      "testimonial":     "Optional quote pulled from the field",
      "testimonialBy":   "Mrs. Lakshmi, Headmistress"
    }
  ],

  "gallery": [
    { "url": "https://...jpg", "caption": "Optional caption", "description": "" }
  ],

  "testimonials": [
    { "content": "...", "author": "...", "role": "Optional role/state" }
  ],

  // Pre-aggregated rollups — pass these if available, otherwise leave out
  "approvedBudget":     1800000,
  "spentBudget":        1450000,
  "deliverablesDone":   3,
  "deliverablesTotal":  7
}
```

If the orchestrator already gave you `project_context` JSON, the field names line up almost 1:1 — copy them through. Where rollups (`approvedBudget`, `spentBudget`, `deliverablesDone`, `deliverablesTotal`) are missing, leave them out and the scripts compute reasonable values from the lists.

---

## Filename convention

Use exactly the filename the orchestrator/user supplied. If you must invent one, use:

```
<project-slug>_<periodTag>.<ext>
```
where `periodTag` is `full` for full-project reports or `YYYYMMDD-YYYYMMDD` for quarterly.

Always write to `/mnt/user-data/outputs/` so the VEPIP UI can serve the file via `/api/ai/artifact/<thread>/outputs/<filename>`.

---

## Critical rules — do not violate

- **Never write Python rendering code yourself.** The three scripts in `scripts/` are the canonical templates. Touching them or replicating them inline drifts the design from VE's brand.
- **Never write an empty/placeholder file.** If a script fails, read its traceback, correct the data JSON, and re-run. An empty file is worse than a clear error.
- **Never call `ask_clarification`** in a report-generation run — these are unattended.
- **Embed period dates explicitly** via the CLI flags, not just in JSON. The scripts cross-check.
- **₹ is the currency** — never USD. The scripts auto-format INR with Indian comma grouping (`₹12,45,000`) and short forms (`₹14.5L`, `₹1.8Cr`).
- **For full-project reports** pass `--report-type full`; the cover renders "FULL PROJECT REPORT" and the period spans the whole project lifecycle.

---

## What the scripts produce (so you know what good looks like)

**PPTX (9 slides, 16:9, 13.333" × 7.5"):**
1. Cover — dark brown + braille texture, white info band, funder logo
2. Project Overview — gold-bar summary + duration & grant cards
3. Key Impact at a Glance — 2×2 metric grid (deliverables, budget, teachers, students)
4. Deliverables — table with progress bars and status dots
5. Field Activities — vertical gold-circle timeline + 2×2 photo grid
6. Stories from the Field — large cream quote card + secondary quote
7. Geographic Coverage — hero number + state grid + reach chips
8. Budget & Utilisation — 4 stat cards + horizontal progress bars per category
9. Way Forward — challenges + next steps + dark "Thank You" panel with braille

**DOCX:** branded cover, executive summary, deliverable table (dark brown header), activities bullet list, 3-up impact tiles, budget table, testimonials in italic blockquotes, way-forward section.

**PDF (A4):** same content as DOCX rendered through reportlab with the brand palette and Times-based headings.
