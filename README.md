# Manufacturing BOM Dashboard

A two-part manufacturing operations tool: a **Google Sheets workbook** (managed by an Apps Script companion) that stores all master data, and a **standalone HTML dashboard** (`bom-viewer.html`) that visualises it as an interactive tree.

No build step, no server, no dependencies beyond D3.js (loaded from CDN). The HTML file opens directly in a browser.

---

## Files

| File | Purpose |
|---|---|
| `bom-viewer.html` | Dashboard viewer — open in any browser |
| `appscript.gs` | Google Apps Script — paste into the Script Editor of your Google Sheet |

---

## Quick Start

### 1. Set up the Google Sheet

1. Create a new Google Sheet.
2. Open **Extensions → Apps Script**, paste the contents of `appscript.gs`, and save.
3. Reload the sheet. A **🏭 Dashboard** menu appears in the menu bar.
4. Run **🏭 Dashboard → Setup workbook (run once)**.  
   This creates all 12 sheets with correct headers, formatting, validation, and formulas.

### 2. Enter your data

Fill in the master sheets in order:

| Sheet | What to enter |
|---|---|
| `BOM_Nodes` | Product tree — nodes with `id`, `parent`, `type`, `name`, `sublabel` |
| `Doc_Nodes` | Procedures, tests, checklists, and references linked to BOM nodes |
| `Technicians` | Technician initials (one per row) |
| `Suppliers` | Supplier names and trust scores |

### 3. Sync

Run **🏭 Dashboard → Sync all sheets** after any change to master data.  
Sync updates derived sheets (`Overview`, `Cycle_BOM`, `Cycle_Docs`, `Supply`, `Training`, `Testing`, `Testing_Notes`), rebuilds the training and testing matrix columns, restores any accidentally-deleted auto-fill formulas, and refreshes supplier dropdowns.

### 4. Fill in the view sheets

After sync, complete the editable columns in the view sheets:

| Sheet | What to fill in |
|---|---|
| `Cycle_BOM` | Actual and goal cycle time (hours) per BOM node |
| `Cycle_Docs` | Actual and goal cycle time (hours) per assembly/test doc |
| `Supply` | Supplier selection and historical quality score (1–5) per stock/PCB node (up to 3 suppliers) |
| `Training` | Training score per technician per doc (see below) |
| `Testing` | Testing score per doc per BOM node it validates |
| `Testing_Notes` | Free-text notes per doc–BOM pair |

### 5. Export

Run **🏭 Dashboard → Export JSON**.  
The script selects the populated cells in `_Export`. Press **⌘C / Ctrl+C** to copy, then paste into the dashboard's **Import JSON** dialog.

---

## Data Reference

### BOM node types

| Type | Description |
|---|---|
| `product` | Top-level product |
| `assembly` | Sub-assembly |
| `machined` | Machined / fabricated part |
| `stock` | Bought-in stock item (appears in Supply sheet) |
| `pcb` | PCB / electronics module (appears in Supply sheet) |

### Doc node types

| Type | Description |
|---|---|
| `assembly` | Assembly procedure |
| `test` | Standalone test procedure |
| `assembly/test` | Procedure that combines assembly steps with an integrated test |
| `checklist` | Inspection or sign-off checklist |
| `reference` | Reference / supplemental document (no training or cycle time) |

Doc readiness `score`: `0` = not started, `1` = draft, `2` = in review, `3` = approved.

### Multi-parent BOM nodes

A BOM node can appear under multiple parents by entering parent IDs comma-separated in the `parent` column:

```
id    parent   type      name
SA1   P1,P2    assembly  Shared Sub-Assembly
```

The dashboard creates one independent sub-tree instance per parent (`SA1__P1`, `SA1__P2`). All docs linked to `SA1` automatically get instances in both sub-trees.

### Multi-parent docs

A doc can belong to multiple BOM nodes by entering BOM node IDs comma-separated in `bom_node_id`:

```
id     bom_node_id   type      label
DOC1   OA,AZ         assembly  Shared Procedure
```

The doc gets one instance per BOM node (`DOC1__OA`, `DOC1__AZ`).

`leads_to` and `linked_to` support a per-BOM split format for multi-parent docs:

```
leads_to = OA=DOC2,AZ=DOC3
```

---

## Training Matrix

The `Training` sheet has a matrix of docs (rows) × technicians (columns). Each cell holds a training score:

| Value | Meaning |
|---|---|
| blank | Not assigned |
| `1` | Awareness |
| `2` | Supervised |
| `3` | Qualified |

**Per-BOM split format** — if a technician's coverage differs between BOM contexts for the same doc, enter a split value in the cell instead of a plain number:

```
OA=3,AZ=2
```

The plain number format covers all BOM parents; split format overrides per-BOM where specified.

**Score thresholds** (sum of all technician scores for a doc or BOM node):

| Threshold | Colour | Meaning |
|---|---|---|
| ≥ 8 | 🟢 Green | Great coverage |
| ≥ 5 | 🟡 Amber | Meets minimum |
| > 3 | 🟠 Orange | At risk |
| ≤ 3 | 🔴 Red | Critical deficit |

---

## Testing Matrix

The `Testing` sheet has a matrix of test/checklist docs (rows) × BOM nodes (columns).  
A column turns white for the BOM nodes a doc validates (driven by `tests_node_id` in `Doc_Nodes`, falling back to `bom_node_id`). Fill in a score (1–3) in white cells.

The `Testing_Notes` sheet has the same structure — enter free-text notes per doc–BOM pair. These appear in the dashboard when you click a testing arrow.

---

## Dashboard Views

Open `bom-viewer.html` in a browser. Load data via **Import JSON** (paste from the `_Export` sheet) or by pasting a Google Sheets URL into the URL field.

### Overview
Default view. BOM nodes coloured by type. Badges showing doc readiness score.

### Training
BOM nodes and docs coloured by training coverage score using the thresholds above. In exploded view, each doc row is coloured by its own training score (not doc quality). Click a BOM node to drill into its training doc coverage.

### Supply Chain
Stock and PCB nodes coloured by supplier quality. Click a node to see its supplier entries and historical quality ratings.

### Testing & Lag
Testing and validation arrows drawn as **Hierarchical Edge Bundles** (HEB) — edges follow the tree path to the lowest common ancestor before fanning to their targets, creating natural bundling along shared hierarchy segments.

Arrow colour reflects combined risk of testing lag (hierarchy levels between doc and validated node) and cycle weight:

| Colour | Risk |
|---|---|
| 🟢 Green | Low lag, low cycle weight |
| 🟡 Amber | Moderate |
| 🔴 Red | High lag or high cycle weight |

**Click any arrow** to:
- Highlight the two connected BOM nodes
- Show a floating info card: doc name → tested node, score, lag, and testing note
- Dim all other arrows

Click the canvas background to dismiss.

Use the risk filter buttons (All / Green / Amber / Red) to focus on specific risk levels.

### Cycle Time
BOM nodes coloured by actual vs goal cycle time ratio. Exploded view shows individual doc cycle times. Overshoot is highlighted in red/amber.

### Docs
Click a BOM node to enter its doc flow — an ordered sequence of all linked docs shown with `leads_to` arrows and `linked_to` dashed branches. Doc nodes coloured by readiness score.

### Simulate
Staffing and capacity simulation. Assign technicians to BOM nodes and track progress against cycle time targets.

### Builds
Build tracking. Create named builds, log progress per BOM node, archive completed builds.

---

## Viewer Controls

| Control | Action |
|---|---|
| Click BOM node | Select node / drill into view |
| Click validation arrow | Show edge info dialog |
| Click canvas background | Deselect / dismiss |
| Scroll | Zoom in/out |
| Drag | Pan |
| **Fit to Screen** | Reset zoom to fit all nodes |
| **⊞ Explode** | Expand docs inline within each BOM node |
| Collapse toggle (▶ on node) | Collapse/expand a sub-tree |
| **+ Add Node / + Section** | Add BOM nodes or section dividers in-browser |

---

## Export Format

The `_Export` sheet stores one JSON chunk per row (one key of the top-level data object per row) to stay within Google Sheets' 50,000-character cell limit. The dashboard import handler merges the chunks automatically — just select all rows in column A below the header, copy, and paste.

---

## Architecture Notes

- **No backend.** `bom-viewer.html` is fully self-contained. Data lives in localStorage between sessions.
- **No build step.** D3 v7 is loaded from CDN. The file can be opened directly from disk or served from any static host.
- **appscript.gs** runs entirely within Google Apps Script — no external APIs or OAuth beyond the sheet itself.
- **GViz import** (paste a Google Sheets URL) uses JSONP via Google's Visualization API, which works from any origin without CORS issues on sheets shared as "Anyone with the link can view".
- **BOM multi-parent expansion** happens server-side in `expandBomRows()` during export, and client-side in `buildHierarchy()` as a fallback for the GViz import path.
- **Validation arrow deduplication**: for docs that appear in multiple BOM sub-trees, arrows are deduplicated by `(baseDocId, targetBomNode)` — the source instance with the deepest Lowest Common Ancestor with the target wins, preventing cross-tree crossings.
