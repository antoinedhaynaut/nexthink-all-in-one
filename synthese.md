# Project Synthesis — SaaS Platform Infographic

> Hand-off document. Read this first if you are resuming the project in a new session.

---

## 1. Goal

Build an interactive visual representation of a SaaS solution that:

- Captures telemetry from many data sources (virtual sessions, laptops, desktops, mobile devices, external systems, user sentiment).
- Routes it through a **unified platform** with **foundation features** grouped into `collect`, `correlate`, `act`, `govern`.
- Delivers business outcomes through **6 value categories**, each broken down into **3–6 use cases**.
- Each use case relies on a subset of foundation features (M:N) and serves a subset of personas (M:N).
- Clicking a use case opens a right-hand drawer with: **value calculation**, **customer story**, **personas concerned**, **features used**, **maturity ladder (L1→L4)**.

The point of the infographic is to make it obvious that this is a **platform**, not a tool — one engine powering many outcomes for many teams.

---

## 2. Architecture decision

**Pattern: spreadsheet → JSON → HTML (single file, bundled inline).**

Three layers, each playing to its strength:

| Layer | File | Role |
|---|---|---|
| Editing | `platform_content.xlsx` | Source of truth. Anyone can edit it. Diffable in git. |
| Runtime data | `data.json` | Canonical normalized artifact. Validated. |
| Viewing | `platform_infographic.html` | Self-contained visualization. JSON is bundled inline so the file works from `file://` (no server needed). |

A single command refreshes everything:

```bash
python3 xlsx_to_json.py
```

### Why this combo and not alternatives

- **Pure JSON** — editing nested arrays + many IDs + M:N relationships is painful for non-technical contributors. Spreadsheets are universally understood.
- **Airtable / Notion / a CMS** — adds accounts, API keys, network round-trips for what is fundamentally a static catalog that changes weekly, not per second.
- **A Google Sheet** — would work, but breaks the "open the HTML offline" promise; the spreadsheet → JSON pipeline keeps everything local and shareable as files.

---

## 3. Data model

```
data_sources           — every system feeding the platform
foundation_features    — capabilities offered by the platform
                         (group ∈ {collect, correlate, act, govern})
value_categories       — the 6 logical value buckets
use_cases              — each belongs to ONE value_category
personas               — audiences

use_case_features      — M:N: which features each use case uses
use_case_personas      — M:N: which personas each use case targets
maturity_levels        — 1:N: 3–4 maturity steps per use case
```

### Conventions

- **IDs are immutable** and prefixed by entity type: `src_*`, `f_*`, `vc_*`, `uc_*`, `p_*`. Names can change freely; IDs cannot — they are the foreign-key target.
- **All M:N relationships** live in dedicated join sheets (one row per pair). This stays clean as the catalog grows.
- **`order`** columns control display order (lower first).
- **`color`** is a 6-character hex without the leading `#` (e.g. `22D3EE`).
- **Multi-line text** is allowed in description / pitch / example cells; the renderer respects line breaks.

### JSON shape produced by the converter

```jsonc
{
  "dataSources":        [{id, name, category, icon, description, order}],
  "foundationFeatures": [{id, name, group, description, order}],
  "valueCategories":    [{id, name, color, tagline, order}],
  "personas":           [{id, name, role, icon, order}],
  "useCases": [{
    id, name, valueCategoryId, shortPitch,
    valueCalculation: {metric, formula, example},
    customerStory:    {customer, industry, oneLiner, result},
    featureIds:       [...],
    personaIds:       [...],
    maturityLevels:   [{level, label, example}, ...],
    order
  }]
}
```

---

## 4. Files in `/outputs`

| File | Role | Hand-edit? |
|---|---|---|
| `platform_content.xlsx` | Source-of-truth content workbook | **Yes — primary editing surface** |
| `build_workbook.py` | Rebuilds the workbook from scratch (dummy content) | Only if reshaping the schema |
| `xlsx_to_json.py` | Converts xlsx → `data.json`, also bundles into HTML | Only if changing pipeline |
| `data.json` | Canonical runtime data (generated) | **No — generated** |
| `platform_infographic.html` | Self-contained visualization | UI tweaks only — never edit the bundled JSON block |
| `synthese.md` | This file | Update when context changes |

---

## 5. Workflow (day-to-day editing loop)

1. Open `platform_content.xlsx`.
2. Edit any sheet except `README`. Add rows, change names, tweak descriptions.
3. Run from the outputs folder:
   ```bash
   python3 xlsx_to_json.py
   ```
4. The script:
   - Reads every sheet.
   - **Validates referential integrity** — any dangling FK aborts with a clear error and non-zero exit.
   - Writes `data.json`.
   - Rewrites the inline `<script id="bundled-data">` block inside `platform_infographic.html`, between markers `/*BUNDLED_DATA_START*/` and `/*BUNDLED_DATA_END*/`.
5. Reload `platform_infographic.html` in a browser. Done.

To rebuild the workbook from scratch (e.g. after a schema change in `build_workbook.py`):

```bash
python3 build_workbook.py        # regenerate dummy content
python3 xlsx_to_json.py          # refresh JSON + HTML bundle
```

---

## 6. Visualization layout

Dark glass UI, single page, max-width ~1500px.

```
┌─ Header ───────────────────────────────────────────────┐
│ Title + lede                    [Value-cat filter btns]│
├─ Left col ─┬─ Stage ───────────────────────────────────┤
│ data       │ ┌── Hub: Foundation features ───────────┐ │
│ sources    │ │ collect | correlate | act | govern    │ │
│ list       │ └──────────────────────────────────────-─┘ │
│            │ ┌── Value categories (6 columns) ───────┐ │
│            │ │  vc1   vc2   vc3   vc4   vc5   vc6    │ │
│            │ └────────────────────────────────────-───┘ │
│            │ ┌── Use cases (stacked under each vc) ──┐ │
│            │ │  uc...  uc...  uc...  uc...           │ │
│            │ └─────────────────────────────────-──────┘ │
│            │ ┌── Personas band (6 cols) ────-────────-┐ │
│            │ └────────────────────────────-───────────┘ │
└────────────┴────────────────────────────────────────────┘
```

Curved gradient SVG paths connect:

- sources → hub
- hub → each value category card
- value category → first use case in its column
- (when a use case is selected) selected use case → each feature it uses

### Interactions

- **Click a use case** → adds `.selected` to it; adds `.hi` to its features and personas; `.dim` to the others; opens the drawer.
- **Click the same use case again, press Esc, or click ✕** → clears selection, closes drawer.
- **Top filter buttons** → "All" or one specific value category. Outside-category items get `.dim`.
- **Hover** → subtle lift + border-color change on cards.

### Drawer (right side, slides in)

Sections in order:

1. **Header** — value-category color tag, use-case name, short pitch.
2. **Value calculation** — metric / formula / worked example (3 KV cards).
3. **Customer story** — customer + industry + one-liner + result (single highlighted box).
4. **Personas concerned** — chips.
5. **Features used** — chips.
6. **Maturity ladder** — L1→L4 steps with label + example.

---

## 7. State of the work

### ✅ Done

- Workbook with full schema, README sheet, and seeded dummy content (6 sources, 13 features, 6 value categories, 12 use cases, 6 personas, maturity ladders for the most-used cases).
- Converter (`xlsx_to_json.py`) with full referential-integrity validation.
- Inline-bundling pipeline so the HTML works straight from `file://`.
- Visualization fully wired: filter, selection, drawer with all 5 sections, curved connection lines, responsive collapse below 1100px viewport width.

### 🟡 Placeholder content to replace with the user's real material

- Customer names (Acme, Globex, Initech, Stark, Wayne, Cyberdyne, Tyrell, etc.) — invent or use actual references.
- ROI numbers in every `value_example` cell — replace with the user's real worked examples.
- The 6 value category names (DEX / Proactive IT / Change / Security / Sustain / Workplace) — placeholder taxonomy; user may have a different one.
- Foundation feature list — 13 features in 4 groups; may need to reflect the user's actual platform capabilities.
- Data source list — 6 generic types; user may collect more / fewer / different sources.

### 🔵 Open questions / decisions to revisit

- **Source ↔ use case linkage**: currently data sources just "feed the platform" (no per-use-case wiring). If the user wants to highlight which sources each use case actually depends on, add a `use_case_sources` join sheet and a corresponding M:N highlight in the UI.
- **Print / PDF export**: this is a web-first interactive infographic. A static poster-style version would need a separate render path.
- **Multi-language**: all strings are English; the schema has no locale columns. Decide if needed.
- **The "Collect / Correlate / Act / Govern" framing** for foundation features — it's a placeholder grouping; easy to change via the `group` column, but the colors and labels are hard-coded in the HTML CSS (`--c-collect`, `--c-correlate`, `--c-act`, `--c-govern`) and in the `groupLabel` map in the JS. If you change the group keys, update those too.

---

## 8. Tech notes for whoever resumes

- **Stack**: pure HTML/CSS/JS — no build step, no framework, no npm. Single file. SVG for connection lines.
- **Python deps**: `openpyxl` only.
- **Bundling marker**: the inline data block in the HTML is delimited by `/*BUNDLED_DATA_START*/` and `/*BUNDLED_DATA_END*/`. Do not remove these markers — `xlsx_to_json.py` uses them to locate the rewrite target.
- **Boot order in the HTML**:
  1. Try the inline JSON block (works offline from `file://`).
  2. Fall back to `fetch('data.json')` if the inline block is empty (useful when developing with `python3 -m http.server`).
  3. If both fail, show a friendly error.
- **Theming**: CSS variables on `:root`. Each value category supplies its own color via the workbook's `color` column.
- **Icons**: small inline SVG paths kept in JS dictionaries (`ICONS` for personas / data sources, `FEAT_ICON` for feature groups). To add a new icon, add a key + path string.
- **Responsive**: the layout collapses to 2 columns at viewport width ≤ 1100px.

---

## 9. Suggested next moves (if the user resumes)

1. Replace placeholder content with the user's real taxonomy in the workbook, then re-run the converter.
2. Add `use_case_sources` join sheet + UI highlighting if source-to-use-case linkage matters.
3. Add a print stylesheet (`@media print`) for a poster export.
4. "Compare two use cases" mode — overlay two selections in the hub and union the highlighted features.
5. Add an "Export PNG" button via `html2canvas` or similar for sharing a static snapshot.
6. Optional: split the visualization into per-value-category mini-pages (one HTML per category) sharing the same `data.json`.

---

## 10. Quick reference — minimal commands

```bash
# from the outputs folder

# Refresh the visualization after editing the workbook:
python3 xlsx_to_json.py

# Rebuild the workbook itself from scratch (rarely):
python3 build_workbook.py

# Local dev server (only needed if not relying on the inline bundle):
python3 -m http.server
# then open http://localhost:8000/platform_infographic.html
```

That's all you need to keep moving.
