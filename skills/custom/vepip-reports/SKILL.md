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
   | PPTX | `python /mnt/skills/custom/vepip-reports/scripts/build_pptx.py --data /mnt/user-data/workspace/report_data.json --output /mnt/user-data/outputs/<filename>.pptx --report-type <quarterly\|full> --period-start <YYYY-MM-DD> --period-end <YYYY-MM-DD> --vibe <vibe-key>` |
   | DOCX | `python /mnt/skills/custom/vepip-reports/scripts/build_docx.py --data /mnt/user-data/workspace/report_data.json --output /mnt/user-data/outputs/<filename>.docx --report-type <quarterly\|full> --period-start <YYYY-MM-DD> --period-end <YYYY-MM-DD> --vibe <vibe-key>` |
   | PDF  | `python /mnt/skills/custom/vepip-reports/scripts/build_pdf.py --data /mnt/user-data/workspace/report_data.json --output /mnt/user-data/outputs/<filename>.pdf --report-type <quarterly\|full> --period-start <YYYY-MM-DD> --period-end <YYYY-MM-DD> --vibe <vibe-key>` |

   `--vibe` accepts: `editorial-serif` (default — warm cream + gold, prose-heavy), `dark-premium` (numbers-led board pitch, ≤8 slides), `magazine-bold` (saffron + black, photo-led, ≤6 slides), `ocean-corporate` (deep blue, audit-friendly, data-led). All three formats honour the same vibe selection so a project's PPTX, DOCX and PDF read as one coherent document.

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

**PPTX (16:9, 13.333" × 7.5", up to 10 slides — manifest-driven, fully editable):**

**Architecture.** The PPTX builder uses a three-stage pipeline that produces real editable PowerPoint shapes, not rasterized image slides:

1. **Python composer** — builds one HTML document containing N stacked 1920×1080 slide sections, using the bundled `html-ppt` skill's design tokens (Inter, Playfair Display, real CSS gradients, real grids, real SVG charts).
2. **Headless Chrome via puppeteer-core** — loads the deck HTML, waits for webfonts to settle so text metrics are correct.
3. **dom-to-pptx** — walks each section's DOM, reads `getComputedStyle()`/`getBoundingClientRect()` in pixel space, and emits **native editable PowerPoint shapes** (text boxes, vector shapes, gradients, embedded fonts) through pptxgenjs. SVGs are preserved as PowerPoint vectors (`svgAsVector: true`).

The result: open the .pptx in PowerPoint and click any text — it's a real editable text box with the correct font, color, and position. Charts stay crisp at any zoom because they're vectors, not pixels.

**Dependencies the runtime needs:**
- Python: `python-pptx` + `Pillow` (auto-installed via `_bootstrap.ensure`).
- Node 18+ on PATH or at `VEPIP_NODE`. Packages `dom-to-pptx` and `puppeteer-core` must be installed at the repo root (`npm install dom-to-pptx puppeteer-core`). The Python script auto-detects the nearest `node_modules` that contains them.
- A Chromium-class browser on PATH (`google-chrome`, `chromium`, or `microsoft-edge`). The builder probes the obvious Windows/macOS/Linux install paths; set `VEPIP_CHROME=<path>` to override. `puppeteer-core` does NOT bundle Chrome — we use the system browser.
- The `html-ppt` skill present at the sibling `public/html-ppt/` directory (it ships the 36 themes, base.css, fonts.css).

**Slide manifest — what gets emitted:**

The manifest picks slides based on data quality AND the selected vibe's structural preferences (prose-loving vs number-led vs photo-led). A sparse project gets a tight 4-5 slide deck; a rich project gets up to 10. No placeholder slides — if the data isn't there, the slide isn't there.

1. **Cover** — always. Left accent rail, eyebrow + period + huge auto-sized project title, funder + grant + states anchored to bottom.
2. **Project Overview** — emitted only if `narrative.overview` or `summary` ≥ N chars (N=80 for prose vibes, 180 for number-led vibes). No "This report summarises progress…" boilerplate.
3. **Key Impact** — emitted if any of deliverables / budget / teachers / students has real numbers. Hero card with mega numeral + three supporting cards.
4. **Deliverables** — emitted if `deliverables.length ≥ 1`. Table with HTML progress bars and percent labels.
5. **Field Activities** — emitted if `activities.length ≥ 1`. Vertical timeline + 2×2 photo grid (Pillow-cropped to 16:10) when gallery photos exist.
6. **Stories from the Field** — emitted only if a testimonial ≥ N chars exists (N=40 for quote-led vibes, 100 for number-led). Large serif italic pull-quote + secondary quote.
7. **Geographic Coverage** — emitted only if `states.length ≥ 2`. Hero state count + 2-col state grid + 4-metric chip row.
8. **Budget & Utilisation** — emitted if `approvedBudget > 0`. **Real SVG doughnut chart** (CSS-themable, crisp at any zoom) + financial-card column + budget-breakdown bars per category.
9. **Way Forward** — emitted only if `narrative.challenges` or `narrative.way_forward` exists. Two-column challenges/next-steps cards.
10. **Closing** — always. Big serif italic "Thank you" + project name + funder line + brand mark.

**Vibe → html-ppt theme mapping** (drives palette, typography, web fonts):
- `editorial-serif` → `editorial-serif.css` (warm cream + gold serif)
- `dark-premium` → `tokyo-night.css` (deep navy + blue accent, bold sans)
- `magazine-bold` → `magazine-bold.css` (saffron + black, ultra-bold)
- `ocean-corporate` → `corporate-clean.css` (slate + deep blue, restrained)

Each vibe also has structural knobs (`slide_count_cap`, `prefers_prose`, `prefers_quotes`, `prefers_charts`, `prefers_photos`) that tune the manifest. Same project data → 4 distinctly-shaped decks.

**DOCX:** vibe-driven palette + typography, branded cover, gated sections (executive summary, deliverable table with vibe-coloured header, activities bullet list, 3-up impact tiles, budget table, testimonials in italic blockquotes, challenges/next-steps). Same content-gating as the PPTX manifest — no "No X recorded" placeholders, no "To be documented in consultation with the field team" filler.

**PDF (A4):** same content as DOCX rendered through reportlab with the vibe palette. Serif vibes (`editorial-serif`) use Times-based headings; sans vibes (`dark-premium`, `magazine-bold`, `ocean-corporate`) use Helvetica-based headings. Same gating rules.
