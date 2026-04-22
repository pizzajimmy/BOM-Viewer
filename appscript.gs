// ═══════════════════════════════════════════════════════════════════
//  Manufacturing BOM Dashboard — Google Sheets companion script
//  Run setupDashboardSheets() once to build the full workbook.
// ═══════════════════════════════════════════════════════════════════

const SHEET_NAMES = {
  BOM:             'BOM_Nodes',
  DOCS:            'Doc_Nodes',
  TECHNICIANS:     'Technicians',
  SUPPLIERS:       'Suppliers',
  OVERVIEW:        'Overview',
  CYCLE_BOM:       'Cycle_BOM',
  CYCLE_DOCS:      'Cycle_Docs',
  SUPPLY:          'Supply',
  TRAINING:        'Training',
  TESTING:         'Testing',
  TESTING_NOTES:   'Testing_Notes',
  EXPORT:          '_Export',
};

const TRAINING_BASE_COLS = 3; // doc_id | bom_node_id | label
const TESTING_BASE_COLS  = 4; // doc_id | bom_node_id | label | type

// ── Visual constants ────────────────────────────────────────────────
const HDR_BG        = '#1e293b';
const HDR_FG        = '#f1f5f9';
const HDR_FONT      = 'Google Sans';
const DATA_FONT     = 'Roboto Mono';
const LOCKED_BG     = '#0f172a';   // kept for reference; light theme uses LOCKED_LIGHT_BG
const LOCKED_LIGHT_BG = '#e2e8f0'; // auto-fill columns on light-theme sheets
const TECH_HDR_BG   = '#dcfce7';   // light green header for technician columns
const TECH_HDR_FG   = '#14532d';   // dark green text
const TECH_COL_W    = 48;

// ── Sheet definitions ──────────────────────────────────────────────
// darkTheme: false → data rows white, alt row #f1f5f9, dark font (all sheets)
const SHEET_DEFS = {

  [SHEET_NAMES.BOM]: {
    tab: { color: '#4f46e5' },
    darkTheme: false,
    headers: ['id', 'parent', 'type', 'name', 'sublabel'],
    notes: {
      1: 'Unique node ID (e.g. P1, A1, M1)',
      2: 'Parent node ID — leave blank for root.\nFor nodes that appear under multiple parents, enter all parent IDs comma-separated (e.g. P1,P2).\nThe dashboard will create one tree instance per parent.',
      3: 'product | assembly | machined | stock | pcb',
      4: 'Human-readable display name',
      5: 'Part / SKU number shown under the name',
    },
  },

  [SHEET_NAMES.DOCS]: {
    tab: { color: '#7c3aed' },
    darkTheme: false,
    headers: ['id', 'bom_node_id', 'type', 'label', 'doc_num', 'score', 'leads_to', 'linked_to', 'tests_node_id'],
    notes: {
      1: 'Unique doc ID',
      2: 'BOM node(s) this doc is PART OF — drives tree placement and cycle time.\nComma-separated for docs that span multiple BOMs (e.g. OA,AZ).',
      3: 'assembly | test | assembly/test | checklist | reference\nassembly/test = procedure that includes both assembly steps and an integrated test',
      4: 'Document display name',
      5: 'Document number (e.g. SOP-011)',
      6: 'Readiness score 0–3',
      7: 'ID of the doc this one leads TO (next step in the flow).\nFor docs shared across multiple BOMs, use split format: OA=DOCX,AZ=DOCY\n(one target per BOM, comma-separated).',
      8: 'ID of a doc this one is side-linked to (dashed branch).\nFor docs shared across multiple BOMs, use split format: OA=DOCX,AZ=DOCY\n(one target per BOM, comma-separated).',
      9: 'BOM node(s) this doc VALIDATES — comma-separated (e.g. MOTOR,GEARBOX).\nControls which Testing matrix columns are active (white) for this doc.\nLeave blank to default to bom_node_id.\nDoes NOT affect cycle time or tree placement.',
    },
  },

  [SHEET_NAMES.TECHNICIANS]: {
    tab: { color: '#d97706' },
    darkTheme: false,
    headers: ['id', 'name'],
    notes: {
      1: 'Technician ID / initials used everywhere operationally (e.g. AC, BS)\nShown in the Training matrix column headers and used by live queue URLs.',
      2: 'Display name used in the dashboard UI (e.g. Alice, Ben)\nRun Dashboard → Sync after adding a new technician.',
    },
  },

  [SHEET_NAMES.SUPPLIERS]: {
    tab: { color: '#d97706' },
    darkTheme: false,
    headers: ['id', 'name', 'trust_score'],
    notes: {
      1: 'Supplier ID — referenced in Supply sheet dropdowns (e.g. SUP1)',
      2: 'Supplier display name',
      3: 'Trust score 1–5 (5 = most trusted)',
    },
  },

  [SHEET_NAMES.OVERVIEW]: {
    tab: { color: '#0ea5e9' },
    darkTheme: false,
    headers: ['id', 'name', 'sublabel'],
    autoFill: [
      { col: 1, formula: `=ARRAYFORMULA(IF(BOM_Nodes!A2:A="","",BOM_Nodes!A2:A))` },
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,BOM_Nodes!A:E,4,FALSE)))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,BOM_Nodes!A:E,5,FALSE)))` },
    ],
    notes: { 1: 'Auto-filled from BOM_Nodes — do not edit' },
  },

  [SHEET_NAMES.CYCLE_BOM]: {
    tab: { color: '#0d9488' },
    darkTheme: false,
    // fpy_pct removed
    headers: ['id', 'name', 'sublabel', 'cycle_time_hrs', 'goal_cycle_time_hrs'],
    autoFill: [
      { col: 1, formula: `=ARRAYFORMULA(IF(BOM_Nodes!A2:A="","",BOM_Nodes!A2:A))` },
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,BOM_Nodes!A:E,4,FALSE)))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,BOM_Nodes!A:E,5,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from BOM_Nodes — do not edit',
      4: 'Actual cycle time in hours',
      5: 'Target / goal cycle time in hours',
    },
  },

  [SHEET_NAMES.CYCLE_DOCS]: {
    tab: { color: '#0d9488' },
    darkTheme: false,
    headers: ['doc_id', 'bom_node_id', 'label', 'doc_num', 'type', 'cycle_time_hrs', 'goal_cycle_time_hrs'],
    autoFill: [
      // Only assembly and test docs have meaningful cycle times to track.
      // Checklists and reference/supplemental docs are excluded.
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="assembly")+(Doc_Nodes!C2:C="test")+(Doc_Nodes!C2:C="assembly/test")),"")` },
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE)))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
      { col: 4, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,5,FALSE)))` },
      { col: 5, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,3,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes (assembly + test types only) — do not edit',
      4: 'Auto-filled doc number — for reference',
      5: 'Auto-filled doc type — for reference',
      6: 'Cycle time in hours — editable. Takes priority over Doc_Nodes on export.\nFor docs shared across multiple BOM nodes, enter per-BOM split: OA=10,AZ=10,AH=15\nEach BOM node then receives only its portion in the cycle time view.\n(Leave 0 on Cycle_BOM for those nodes so the doc sum drives the display.)',
      7: 'Target / goal cycle time in hours for this document.\nSupports the same per-BOM split format as cycle_time_hrs: OA=10,AZ=10,AH=15\nEach BOM node receives only its portion in the cycle time explode view.',
    },
  },

  // Supply: one row per stock/pcb node (synced by syncSupplySheet).
  // id col populated by script, name/sublabel via VLOOKUP formula.
  // Up to 3 supplier slots — each supplier_N is a dropdown from Suppliers!A:A.
  [SHEET_NAMES.SUPPLY]: {
    tab: { color: '#0ea5e9' },
    darkTheme: false,
    headers: ['id', 'name', 'sublabel', 'supplier_1', 'quality_1', 'supplier_2', 'quality_2', 'supplier_3', 'quality_3'],
    autoFill: [
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,BOM_Nodes!A:E,4,FALSE)))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,BOM_Nodes!A:E,5,FALSE)))` },
    ],
    notes: {
      1: 'Stock/PCB node ID — auto-populated by Sync (stock and pcb types only)',
      2: 'Auto-filled from BOM_Nodes',
      3: 'Auto-filled from BOM_Nodes',
      4: 'Select supplier from Suppliers sheet',
      5: 'Historical quality score 1–5 for this supplier at this node',
      6: 'Second supplier (optional)',
      7: 'Historical quality score 1–5',
      8: 'Third supplier (optional)',
      9: 'Historical quality score 1–5',
    },
  },

  [SHEET_NAMES.TRAINING]: {
    tab: { color: '#059669' },
    darkTheme: false,
    headers: ['doc_id', 'bom_node_id', 'label'],
    autoFill: [
      // Only assembly and test docs require technician training records.
      // Checklists and reference/supplemental docs are excluded.
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="assembly")+(Doc_Nodes!C2:C="test")+(Doc_Nodes!C2:C="assembly/test")),"")` },
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE)))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes (assembly + test types only) — do not edit\nTechnician columns added automatically via Dashboard → Sync.',
    },
  },

  [SHEET_NAMES.TESTING]: {
    tab: { color: '#059669' },
    darkTheme: false,
    headers: ['doc_id', 'tests_node_id', 'label', 'type'],
    autoFill: [
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="test")+(Doc_Nodes!C2:C="checklist")+(Doc_Nodes!C2:C="assembly/test")),"")` },
      // col 2: use tests_node_id (Doc_Nodes col 9) if set; fall back to bom_node_id (col 2)
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",IF(IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),"")<>"",IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),""),IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE),""))))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
      { col: 4, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,3,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes (test + checklist + assembly/test types) — do not edit',
      2: 'Auto-filled: shows tests_node_id from Doc_Nodes (which BOM nodes this doc validates).\nFalls back to bom_node_id if tests_node_id is blank.\nMatrix columns turn white for matching BOM node IDs.',
      4: 'Auto-filled type — for reference only',
    },
  },

  [SHEET_NAMES.TESTING_NOTES]: {
    tab: { color: '#059669' },
    darkTheme: false,
    headers: ['doc_id', 'tests_node_id', 'label', 'type'],
    autoFill: [
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="test")+(Doc_Nodes!C2:C="checklist")+(Doc_Nodes!C2:C="assembly/test")),"")` },
      // col 2: use tests_node_id (Doc_Nodes col 9) if set; fall back to bom_node_id (col 2)
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",IF(IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),"")<>"",IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),""),IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE),""))))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
      { col: 4, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,3,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes (test + checklist + assembly/test types) — do not edit',
      2: 'Auto-filled: shows tests_node_id from Doc_Nodes (which BOM nodes this doc validates).\nFalls back to bom_node_id if tests_node_id is blank.',
      4: 'Auto-filled type — for reference only',
    },
  },

  [SHEET_NAMES.EXPORT]: {
    tab: { color: '#64748b' },
    darkTheme: false,
    headers: ['json_output'],
    notes: { 1: 'Auto-generated by Dashboard → Export JSON.\nThe export is split across multiple rows (one per data section).\nSelect all filled cells in column A below this header, copy, then paste into the dashboard Import dialog.' },
  },
};


// ═══════════════════════════════════════════════════════════════════
//  MENU
// ═══════════════════════════════════════════════════════════════════
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏭 Dashboard')
    .addItem('🔧 Setup workbook (run once)', 'setupDashboardSheets')
    .addSeparator()
    .addItem('🔄 Sync all sheets', 'syncViewSheets')
    .addItem('📤 Export JSON', 'exportJSON')
    .addToUi();
}


// ═══════════════════════════════════════════════════════════════════
//  ONE-TIME SETUP
// ═══════════════════════════════════════════════════════════════════
function setupDashboardSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const existing = ss.getSheets().map(s => s.getName());
  const toUpdate = Object.keys(SHEET_DEFS).filter(n => existing.includes(n));

  if (toUpdate.length) {
    const resp = ui.alert(
      'Sheets already exist',
      `These sheets will be reformatted (data rows kept):\n${toUpdate.join(', ')}\n\nContinue?`,
      ui.ButtonSet.OK_CANCEL
    );
    if (resp !== ui.Button.OK) return;
  }

  const ORDER = [
    SHEET_NAMES.BOM,
    SHEET_NAMES.DOCS,
    SHEET_NAMES.TECHNICIANS,
    SHEET_NAMES.SUPPLIERS,
    SHEET_NAMES.OVERVIEW,
    SHEET_NAMES.CYCLE_BOM,
    SHEET_NAMES.CYCLE_DOCS,
    SHEET_NAMES.SUPPLY,
    SHEET_NAMES.TRAINING,
    SHEET_NAMES.TESTING,
    SHEET_NAMES.TESTING_NOTES,
    SHEET_NAMES.EXPORT,
  ];

  ORDER.forEach((name, idx) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name, idx);
    applySheetDef(ss, sh, SHEET_DEFS[name]);
  });

  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) {
    ss.setActiveSheet(defaultSheet);
    ss.moveActiveSheet(ss.getSheets().length);
    defaultSheet.setTabColor('#94a3b8');
    defaultSheet.hideSheet();
  }

  syncTrainingMatrix(ss);
  syncTestingMatrix(ss);
  applySupplierDropdowns(ss);

  ss.setActiveSheet(ss.getSheetByName(SHEET_NAMES.BOM));
  ui.alert(
    '✅ Setup complete',
    'Next steps:\n1. Add BOM nodes to "BOM_Nodes"\n2. Add docs to "Doc_Nodes"\n3. Add suppliers to "Suppliers"\n4. Add technician IDs + names to "Technicians"\n5. Run Dashboard → Sync\n6. Fill in view columns, then Export JSON',
    ui.ButtonSet.OK
  );
}


// ── Apply a sheet definition ────────────────────────────────────────
// Safe to run on sheets that already have data: user-editable column
// data is saved by header name before the schema is applied, then
// written back to the correct new column positions.  This means adding,
// removing, or reordering schema columns never corrupts existing data.
function applySheetDef(ss, sh, def) {
  const numCols    = def.headers.length;
  const darkTheme  = def.darkTheme !== false;
  const baseBg     = darkTheme ? LOCKED_BG   : null;
  const altBg      = darkTheme ? HDR_BG      : '#f1f5f9';
  const dataFgCol  = darkTheme ? '#e2e8f0'   : '#1e293b';

  // ── Save existing user data keyed by column header name ──────────
  // Auto-fill columns are excluded — their formulas are always rewritten.
  // This makes the function idempotent even after schema changes.
  const autoFillColNums = new Set((def.autoFill || []).map(af => af.col));
  const savedByHeader   = {};   // { headerName → [rowValue, ...] }

  const existingLastRow = sh.getLastRow();
  const existingLastCol = sh.getLastColumn();

  if (existingLastRow > 1 && existingLastCol > 0) {
    const existingHeaders = sh.getRange(1, 1, 1, existingLastCol)
      .getValues()[0].map(h => String(h).trim());
    const allData = sh.getRange(2, 1, existingLastRow - 1, existingLastCol).getValues();

    existingHeaders.forEach((h, colIdx) => {
      if (!h) return;
      // Skip if this header maps to an auto-fill column in the new schema
      const newPos = def.headers.indexOf(h) + 1;
      if (newPos > 0 && autoFillColNums.has(newPos)) return;
      savedByHeader[h] = allData.map(row => row[colIdx]);
    });
  }

  if (def.tab?.color) sh.setTabColor(def.tab.color);

  if (sh.getMaxRows()    < 2)       sh.insertRowsAfter(sh.getMaxRows(), 2 - sh.getMaxRows());
  if (sh.getMaxColumns() < numCols) sh.insertColumnsAfter(sh.getMaxColumns(), numCols - sh.getMaxColumns());

  // Header row
  sh.getRange(1, 1, 1, numCols)
    .setValues([def.headers])
    .setBackground(HDR_BG).setFontColor(HDR_FG)
    .setFontFamily(HDR_FONT).setFontWeight('bold')
    .setFontSize(10).setHorizontalAlignment('left').setVerticalAlignment('middle');
  sh.setRowHeight(1, 32);
  sh.setFrozenRows(1);

  // Clear any header cells left over from a previously wider schema
  const lastCol = sh.getLastColumn();
  if (lastCol > numCols) {
    sh.getRange(1, numCols + 1, 1, lastCol - numCols).clearContent();
  }

  if (def.notes) {
    Object.entries(def.notes).forEach(([col, note]) => {
      sh.getRange(1, Number(col)).setNote(note);
    });
  }

  // Column widths
  def.headers.forEach((h, i) => {
    const w = ['id','doc_id','parent','type','score',
               'cycle_time_hrs','goal_cycle_time_hrs','trust_score',
               'quality_1','quality_2','quality_3'].includes(h) ? 90
            : h === 'json_output'    ? 600
            : ['supplier_1','supplier_2','supplier_3'].includes(h) ? 160
            : 180;
    sh.setColumnWidth(i + 1, w);
  });

  // Data rows — base background + font
  if (sh.getMaxRows() > 1) {
    const dataRange = sh.getRange(2, 1, sh.getMaxRows() - 1, numCols);
    if (baseBg) dataRange.setBackground(baseBg);
    else        dataRange.setBackground(null);
    dataRange.setFontColor(dataFgCol).setFontFamily(DATA_FONT).setFontSize(9).setVerticalAlignment('middle');
  }
  sh.setRowHeightsForced(2, Math.max(sh.getMaxRows() - 1, 1), 26);

  // ── Restore user data to correct column positions ─────────────────
  // Clear writable columns first (avoids stale values from old column
  // positions), then write each header's saved data to its new column.
  const numDataRows = Math.max(existingLastRow - 1, 0);
  if (numDataRows > 0 && Object.keys(savedByHeader).length) {
    const blank = Array.from({ length: numDataRows }, () => ['']);
    def.headers.forEach((h, i) => {
      const col = i + 1;
      if (autoFillColNums.has(col)) return;
      // Clear the column before writing (removes stale data from old positions)
      sh.getRange(2, col, numDataRows, 1).setValues(blank);
      const saved = savedByHeader[h];
      if (!saved?.length) return;
      const rows = Math.min(saved.length, numDataRows);
      sh.getRange(2, col, rows, 1).setValues(saved.slice(0, rows).map(v => [v]));
    });
  }

  // Auto-fill formulas + lock styling
  const lockedBg    = darkTheme ? LOCKED_BG : LOCKED_LIGHT_BG;
  const lockedFgCol = darkTheme ? '#94a3b8'  : '#64748b';
  if (def.autoFill?.length) {
    def.autoFill.forEach(af => {
      const cell = sh.getRange(2, af.col);
      cell.setFormula(af.formula);
      sh.getRange(2, af.col, Math.max(sh.getMaxRows() - 1, 1), 1)
        .setBackground(lockedBg).setFontColor(lockedFgCol).setFontStyle('italic');
    });

    sh.getProtections(SpreadsheetApp.ProtectionType.RANGE)
      .filter(p => p.getDescription() === 'auto-fill')
      .forEach(p => p.remove());

    def.autoFill.forEach(af => {
      sh.getRange(2, af.col, sh.getMaxRows() - 1, 1)
        .protect().setDescription('auto-fill').setWarningOnly(true);
    });
  }

  // Alternating row shading
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ROW()>1,MOD(ROW(),2)=0,A2<>"")')
      .setBackground(altBg)
      .setRanges([sh.getRange(2, 1, sh.getMaxRows() - 1, numCols)])
      .build(),
  ]);

  applyValidations(sh, def.headers);
}


// ── Data validation ─────────────────────────────────────────────────
function applyValidations(sh, headers) {
  const maxDataRow = Math.max(sh.getMaxRows() - 1, 1);
  headers.forEach((h, i) => {
    const col = i + 1;
    let rule = null;
    if (h === 'type' && sh.getName() === SHEET_NAMES.BOM) {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['product','assembly','machined','stock','pcb'], true)
        .setAllowInvalid(false).build();
    } else if (h === 'type' && sh.getName() === SHEET_NAMES.DOCS) {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['assembly','test','assembly/test','checklist','reference'], true)
        .setAllowInvalid(false).build();
    } else if (h === 'score') {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['0','1','2','3'], true)
        .setAllowInvalid(false).build();
    } else if (h === 'trust_score') {
      rule = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 5).setAllowInvalid(false).build();
    } else if (['cycle_time_hrs','goal_cycle_time_hrs'].includes(h)) {
      if (sh.getName() === SHEET_NAMES.CYCLE_DOCS) {
        // No validation — both columns accept plain numbers and "OA=10,AZ=10" split format
        sh.getRange(2, col, maxDataRow, 1).clearDataValidations();
        return;
      }
      rule = SpreadsheetApp.newDataValidation()
        .requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build();
    } else if (/^quality_\d+$/.test(h)) {
      rule = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 5).setAllowInvalid(true).build();
    }
    if (rule) sh.getRange(2, col, maxDataRow, 1).setDataValidation(rule);
  });
}


// ── Supplier dropdowns in Supply sheet ─────────────────────────────
// Called at setup and sync. References Suppliers!A:A live so new
// suppliers appear in the dropdown automatically.
function applySupplierDropdowns(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const supplySh     = ss.getSheetByName(SHEET_NAMES.SUPPLY);
  const suppliersSh  = ss.getSheetByName(SHEET_NAMES.SUPPLIERS);
  if (!supplySh || !suppliersSh) return;

  const maxRows = Math.max(supplySh.getMaxRows() - 1, 1);
  // Reference the entire id column so new suppliers appear automatically
  const idRange = suppliersSh.getRange('A2:A');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(idRange, true)   // true = show dropdown arrow
    .setAllowInvalid(true)                // allow blank (not all slots needed)
    .build();

  // supplier_1, supplier_2, supplier_3 are cols 4, 6, 8
  [4, 6, 8].forEach(col => {
    supplySh.getRange(2, col, maxRows, 1).setDataValidation(rule);
  });
}


// ═══════════════════════════════════════════════════════════════════
//  TRAINING MATRIX
// ═══════════════════════════════════════════════════════════════════
// Full keyed rebuild: saves existing scores by [docId][initials],
// deletes all technician columns, re-adds from current Technicians
// sheet in order, then writes saved scores back.  Removed technicians
// lose their column; renamed ones effectively start fresh.
function syncTrainingMatrix(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const techSh  = ss.getSheetByName(SHEET_NAMES.TECHNICIANS);
  const trainSh = ss.getSheetByName(SHEET_NAMES.TRAINING);
  if (!techSh || !trainSh) return;

  const techData = techSh.getLastRow() > 1
    ? techSh.getRange(2, 1, techSh.getLastRow() - 1, 1).getValues().flat()
        .map(v => String(v).trim()).filter(Boolean)
    : [];

  const matrixStart = TRAINING_BASE_COLS + 1;

  // ── Save existing scores keyed by [docId][initials] ──────────────
  const saved = {};   // { docId: { initials: score } }
  const curLastCol = trainSh.getLastColumn();
  if (curLastCol > TRAINING_BASE_COLS && trainSh.getLastRow() > 1) {
    const numDataRows   = trainSh.getLastRow() - 1;
    const existingTechs = trainSh.getRange(1, matrixStart, 1, curLastCol - TRAINING_BASE_COLS)
      .getValues()[0].map(h => String(h).trim());
    const docIds    = trainSh.getRange(2, 1, numDataRows, 1).getValues().flat().map(v => String(v).trim());
    const matrixVals = trainSh.getRange(2, matrixStart, numDataRows, existingTechs.length).getValues();

    docIds.forEach((docId, ri) => {
      if (!docId) return;
      existingTechs.forEach((initials, ci) => {
        if (!initials) return;
        const val = matrixVals[ri][ci];
        if (val !== '' && val !== null && val !== undefined) {
          if (!saved[docId]) saved[docId] = {};
          saved[docId][initials] = val;
        }
      });
    });
  }

  // ── Delete all existing technician columns ────────────────────────
  const oldMatrixCols = trainSh.getLastColumn() - TRAINING_BASE_COLS;
  if (oldMatrixCols > 0) trainSh.deleteColumns(matrixStart, oldMatrixCols);

  if (!techData.length) { trainSh.setFrozenColumns(TRAINING_BASE_COLS); return; }

  // ── Add current technician columns in order ───────────────────────
  trainSh.insertColumnsAfter(TRAINING_BASE_COLS, techData.length);
  const maxRows = Math.max(trainSh.getMaxRows() - 1, 1);

  techData.forEach((initials, i) => {
    const col = matrixStart + i;
    trainSh.getRange(1, col)
      .setValue(initials)
      .setBackground(TECH_HDR_BG).setFontColor(TECH_HDR_FG)
      .setFontFamily(HDR_FONT).setFontWeight('bold')
      .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setNote(`Training score for ${initials}\nEnter 0–3:\n  blank = not assigned\n  1 = awareness\n  2 = supervised\n  3 = qualified`);
    trainSh.setColumnWidth(col, TECH_COL_W);

    trainSh.getRange(2, col, maxRows, 1)
      .setBackground(null).setFontColor('#1e293b')
      .setHorizontalAlignment('center').setFontFamily(DATA_FONT).setFontSize(10)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, 3).setAllowInvalid(true)
          .setHelpText('0 = not trained, 1 = awareness, 2 = supervised, 3 = qualified').build()
      );
  });

  // ── Restore saved scores ──────────────────────────────────────────
  if (trainSh.getLastRow() > 1) {
    const numDataRows = trainSh.getLastRow() - 1;
    const currentDocIds = trainSh.getRange(2, 1, numDataRows, 1).getValues().flat().map(v => String(v).trim());
    techData.forEach((initials, i) => {
      const col = matrixStart + i;
      const colVals = currentDocIds.map(docId => {
        const val = saved[docId]?.[initials];
        return [val !== undefined ? val : ''];
      });
      trainSh.getRange(2, col, numDataRows, 1).setValues(colVals);
    });
  }

  trainSh.setFrozenColumns(TRAINING_BASE_COLS);
}


// ═══════════════════════════════════════════════════════════════════
//  TESTING MATRIX (validates scores + notes)
// ═══════════════════════════════════════════════════════════════════
// Full keyed rebuild: saves existing data by [docId][bomId], deletes
// all BOM columns, re-adds from current BOM_Nodes in order, writes
// data back.  Removed/renamed BOM nodes lose their column cleanly.
function syncTestingMatrix(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const bomSh = ss.getSheetByName(SHEET_NAMES.BOM);
  if (!bomSh || bomSh.getLastRow() < 2) return;

  const bomNodes = bomSh.getRange(2, 1, bomSh.getLastRow() - 1, 4).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => ({ id: String(r[0]).trim(), name: String(r[3]).trim() }));
  if (!bomNodes.length) return;

  [SHEET_NAMES.TESTING, SHEET_NAMES.TESTING_NOTES].forEach(shName => {
    const sh = ss.getSheetByName(shName);
    if (!sh) return;

    const isNotes    = shName === SHEET_NAMES.TESTING_NOTES;
    const matrixStart = TESTING_BASE_COLS + 1;

    // ── Save existing matrix data keyed by [docId][bomId] ────────
    const saved = {};   // { docId: { bomId: value } }
    const curLastCol = sh.getLastColumn();
    if (curLastCol > TESTING_BASE_COLS && sh.getLastRow() > 1) {
      const numDataRows   = sh.getLastRow() - 1;
      const existingBomCols = sh.getRange(1, matrixStart, 1, curLastCol - TESTING_BASE_COLS)
        .getValues()[0].map(h => String(h).trim());
      const docIds     = sh.getRange(2, 1, numDataRows, 1).getValues().flat().map(v => String(v).trim());
      const matrixVals = sh.getRange(2, matrixStart, numDataRows, existingBomCols.length).getValues();

      docIds.forEach((docId, ri) => {
        if (!docId) return;
        existingBomCols.forEach((bomId, ci) => {
          if (!bomId) return;
          const val = matrixVals[ri]?.[ci];
          if (val !== '' && val !== null && val !== undefined) {
            if (!saved[docId]) saved[docId] = {};
            saved[docId][bomId] = val;
          }
        });
      });
    }

    // ── Delete all existing BOM columns then add current ones ─────
    const oldMatrixCols = sh.getLastColumn() - TESTING_BASE_COLS;
    if (oldMatrixCols > 0) sh.deleteColumns(matrixStart, oldMatrixCols);
    sh.insertColumnsAfter(TESTING_BASE_COLS, bomNodes.length);
    SpreadsheetApp.flush();

    const maxDataRows = Math.max(sh.getMaxRows() - 1, 1);

    // ── Write new BOM column headers + formatting ─────────────────
    bomNodes.forEach(({ id, name }, idx) => {
      const col = matrixStart + idx;
      sh.getRange(1, col)
        .setValue(id)
        .setBackground(isNotes ? '#dbeafe' : '#ede9fe')
        .setFontColor(isNotes ? '#1e40af' : '#4c1d95')
        .setFontFamily(HDR_FONT).setFontWeight('bold')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNote(`BOM node: ${name} (${id})\n${isNotes
          ? 'Optional note for this validates edge.\nLeave blank if none.'
          : 'Validation score 0–3:\n  blank = not validated\n  1 = partial\n  2 = functional\n  3 = fully validated'}`);
      sh.setColumnWidth(col, isNotes ? 160 : TECH_COL_W);

      const dataRange = sh.getRange(2, col, maxDataRows, 1);
      dataRange.setBackground('#e2e8f0').setFontColor('#1e293b');
      if (isNotes) {
        dataRange.setFontFamily(HDR_FONT).setFontSize(9).setWrap(true).setVerticalAlignment('top');
      } else {
        dataRange.setHorizontalAlignment('center').setFontFamily(DATA_FONT).setFontSize(10)
          .setDataValidation(
            SpreadsheetApp.newDataValidation()
              .requireNumberBetween(0, 3).setAllowInvalid(true)
              .setHelpText('0 = not validated, 1 = partial, 2 = functional, 3 = fully validated').build()
          );
      }
    });

    // ── Restore saved data into correct new columns ───────────────
    if (sh.getLastRow() > 1) {
      SpreadsheetApp.flush();
      const numDataRows2 = sh.getLastRow() - 1;
      const currentDocIds = sh.getRange(2, 1, numDataRows2, 1).getValues().flat().map(v => String(v).trim());
      bomNodes.forEach(({ id: bomId }, idx) => {
        const col = matrixStart + idx;
        const colVals = currentDocIds.map(docId => {
          const val = saved[docId]?.[bomId];
          return [val !== undefined ? val : ''];
        });
        sh.getRange(2, col, numDataRows2, 1).setValues(colVals);
      });
    }

    // ── Rebuild conditional format rules ──────────────────────────
    // White match rules first so they beat the alternating-row rule.
    const matchRules = bomNodes.map(({ id: bomId }, idx) => {
      const colIdx = matrixStart + idx;
      return SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=ISNUMBER(SEARCH(","&"${bomId}"&",",","&TRIM($B2)&","))`)
        .setBackground('#ffffff')
        .setRanges([sh.getRange(2, colIdx, maxDataRows, 1)])
        .build();
    });

    const altRule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=AND(ROW()>1,MOD(ROW(),2)=0,A2<>"")')
      .setBackground('#f1f5f9')
      .setRanges([sh.getRange(2, 1, maxDataRows, sh.getLastColumn())])
      .build();

    sh.setConditionalFormatRules([...matchRules, altRule]);
    sh.setFrozenColumns(TESTING_BASE_COLS);
  });
}


// ═══════════════════════════════════════════════════════════════════
//  SYNC
// ═══════════════════════════════════════════════════════════════════

// Re-applies any missing autoFill formula cells across all sheet defs.
// Called during every sync so that accidentally-deleted formulas are
// restored without requiring a full "Setup Sheets" run.
function reapplyAutoFills(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(SHEET_DEFS).forEach(([shName, def]) => {
    if (!def.autoFill?.length) return;
    const sh = ss.getSheetByName(shName);
    if (!sh) return;
    def.autoFill.forEach(af => {
      const cell = sh.getRange(2, af.col);
      if (!cell.getFormula()) {
        cell.setFormula(af.formula);
      }
    });
  });
}

function syncViewSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Restore any missing auto-fill anchors before any keyed matrix rebuild.
  // Training/Testing snapshot existing values by the auto-filled ID columns;
  // if those formulas were accidentally deleted, a sync would otherwise read
  // blank row keys and then write back an empty matrix.
  reapplyAutoFills(ss);
  SpreadsheetApp.flush();

  _syncMasterToViews(ss, SHEET_NAMES.BOM, [SHEET_NAMES.OVERVIEW, SHEET_NAMES.CYCLE_BOM]);
  // All doc-based sheets are driven by FILTER formulas — no row sync needed:
  // Cycle_Docs:     assembly + test types only
  // Training:       assembly + test types only
  // Testing:        test + checklist types only
  // Testing_Notes:  test + checklist types only
  syncSupplySheet(ss);
  syncTrainingMatrix(ss);
  syncTestingMatrix(ss);
  reapplyAutoFills(ss);   // final pass in case a later structural edit cleared a formula anchor
  SpreadsheetApp.flush();
  applySupplierDropdowns(ss);

  SpreadsheetApp.getUi().alert('✅ Sync complete.');
}

// Syncs only stock/pcb BOM nodes into the Supply sheet
function syncSupplySheet(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const bomSh    = ss.getSheetByName(SHEET_NAMES.BOM);
  const supplySh = ss.getSheetByName(SHEET_NAMES.SUPPLY);
  if (!bomSh || !supplySh || bomSh.getLastRow() < 2) return;

  const bomData = bomSh.getRange(2, 1, bomSh.getLastRow() - 1, 3).getValues();
  const stockPcbIds = bomData
    .filter(r => ['stock', 'pcb'].includes(String(r[2]).trim().toLowerCase()))
    .map(r => String(r[0]).trim())
    .filter(Boolean);

  const existing = supplySh.getLastRow() > 1
    ? supplySh.getRange(2, 1, supplySh.getLastRow() - 1, 1)
        .getValues().flat().map(String).filter(Boolean)
    : [];

  const missing = stockPcbIds.filter(id => !existing.includes(id));
  if (missing.length) {
    supplySh.getRange(supplySh.getLastRow() + 1, 1, missing.length, 1)
      .setValues(missing.map(id => [id]));
  }
}

function _syncMasterToViews(ss, masterName, viewNames) {
  const master = ss.getSheetByName(masterName);
  if (!master || master.getLastRow() < 2) return;

  const masterIds = master
    .getRange(2, 1, master.getLastRow() - 1, 1)
    .getValues().flat().map(String).filter(Boolean);

  viewNames.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;
    const existing = sh.getLastRow() > 1
      ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().map(String).filter(Boolean)
      : [];
    const missing = masterIds.filter(id => !existing.includes(id));
    if (missing.length) {
      sh.getRange(sh.getLastRow() + 1, 1, missing.length, 1)
        .setValues(missing.map(id => [id]));
    }
  });
}


// ═══════════════════════════════════════════════════════════════════
//  EXPORT
// ═══════════════════════════════════════════════════════════════════
function exportJSON() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const data = buildDashboardJSON(ss);

  // Split into one row per top-level key so no single cell exceeds the
  // 50,000-character Google Sheets limit.  Each row is a self-contained
  // {"key": value} JSON object.  The viewer merges them on import.
  const keys   = Object.keys(data);
  const chunks = keys.map(k => JSON.stringify({ [k]: data[k] }));

  let exportSh = ss.getSheetByName(SHEET_NAMES.EXPORT);
  if (!exportSh) exportSh = ss.insertSheet(SHEET_NAMES.EXPORT);
  exportSh.clearContents();
  exportSh.getRange('A1').setValue('json_output');
  chunks.forEach((chunk, i) => {
    exportSh.getRange(i + 2, 1).setValue(chunk).setWrap(false);
  });
  exportSh.setColumnWidth(1, 800);

  const lastDataRow = chunks.length + 1;
  ss.setActiveSheet(exportSh);
  exportSh.setActiveRange(exportSh.getRange(2, 1, chunks.length, 1));

  SpreadsheetApp.getUi().alert(
    '✅ Export ready',
    `Cells A2:A${lastDataRow} are selected in "_Export" (${chunks.length} chunk${chunks.length > 1 ? 's' : ''}).\nPress ⌘C / Ctrl+C, then paste into the dashboard Import dialog.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ── BOM multi-parent expansion ──────────────────────────────────────────────
// When a BOM node has a comma-separated parent column (e.g. "P1,P2") it is
// replicated once per parent — each copy gets the suffix "__PARENTID" appended
// to its id, and all of its descendants are replicated with the same suffix so
// each copy carries a full independent sub-tree.
//
// Returns { rows, expansionMap } where:
//   rows         – the expanded flat array ready to feed buildDashboardJSON
//   expansionMap – { origId: [expandedId1, expandedId2, ...] }
//                  (only multi-parent nodes and their descendants are keyed here)
function expandBomRows(rows) {
  const multiParentIds = new Set(
    rows.filter(n => String(n.parent || '').includes(','))
        .map(n => String(n.id))
  );
  if (!multiParentIds.size) return { rows, expansionMap: {} };

  // Build a single-parent children map (children of multi-parent nodes
  // point to the original id which still has a single parent at this stage)
  const childrenOf = {};
  rows.forEach(n => {
    const p = String(n.parent || '').trim();
    if (p && !p.includes(',')) {
      (childrenOf[p] = childrenOf[p] || []).push(String(n.id));
    }
  });

  // Mark every node that is a descendant of a multi-parent node
  // (they will be produced by expand() so must be skipped in the main loop)
  const descendants = new Set();
  function markDesc(id) {
    (childrenOf[id] || []).forEach(c => {
      if (!descendants.has(c)) { descendants.add(c); markDesc(c); }
    });
  }
  multiParentIds.forEach(id => markDesc(id));

  const rowMap = Object.fromEntries(rows.map(r => [String(r.id), r]));
  const expansionMap = {}; // origId → [newId, ...]

  // Recursively clone origId and all its single-parent descendants,
  // rewriting ids to origId+suffix and parent to newParentId.
  function expand(origId, newParentId, suffix) {
    const row = rowMap[origId];
    if (!row) return [];
    const newId = origId + suffix;
    (expansionMap[origId] = expansionMap[origId] || []).push(newId);
    const kids = (childrenOf[origId] || []).flatMap(cid => expand(cid, newId, suffix));
    return [{ ...row, id: newId, parent: newParentId }, ...kids];
  }

  const result = [];
  rows.forEach(n => {
    const origId = String(n.id);
    const p      = String(n.parent || '').trim();
    if (multiParentIds.has(origId)) {
      // One copy per parent
      p.split(',').map(x => x.trim()).filter(Boolean).forEach(parentId => {
        result.push(...expand(origId, parentId, `__${parentId}`));
      });
    } else if (!descendants.has(origId)) {
      // Normal node — pass through unchanged
      result.push({ ...n, parent: p || null });
    }
    // descendants are handled recursively by expand(); skip them here
  });

  return { rows: result, expansionMap };
}

function buildDashboardJSON(ss) {

  function sheetToObjects(name) {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return [];
    const data    = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const headers = data[0].map(h => String(h).trim());
    return data.slice(1)
      .filter(r => String(r[0]).trim() !== '')
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => { if (h) obj[h] = r[i]; });
        return obj;
      });
  }

  const techRows = sheetToObjects(SHEET_NAMES.TECHNICIANS);
  const techNameById = Object.fromEntries(
    techRows
      .map(r => [String(r.id || '').trim(), String(r.name || '').trim()])
      .filter(([id]) => !!id)
  );
  const techRegistry = techRows
    .map(r => ({
      id: String(r.id || '').trim(),
      name: String(r.name || '').trim() || String(r.id || '').trim(),
    }))
    .filter(t => t.id);

  // ── BOM nodes ───────────────────────────────────────────────────
  const bomRows    = sheetToObjects(SHEET_NAMES.BOM);
  const cycleRows  = sheetToObjects(SHEET_NAMES.CYCLE_BOM);
  const supplyRows = sheetToObjects(SHEET_NAMES.SUPPLY);
  const cycleById  = Object.fromEntries(cycleRows.map(r  => [String(r.id), r]));
  const supplyById = Object.fromEntries(supplyRows.map(r => [String(r.id), r]));

  // Expand any BOM nodes that have comma-separated parents into one clone per parent.
  // expansionMap lets us resolve doc bom_node_id references to expanded ids later.
  const { rows: expandedBomRows, expansionMap: bomExpansionMap } = expandBomRows(bomRows);

  // Build a reverse map: expandedId → origId, so cycle/supply lookups still
  // hit their data after the id has been suffixed (e.g. NODE_A__P1 → NODE_A).
  const origIdOfBomNode = {};
  Object.entries(bomExpansionMap).forEach(([origId, expandedIds]) => {
    expandedIds.forEach(eid => { origIdOfBomNode[eid] = origId; });
  });

  const nodes = expandedBomRows.map(n => {
    const lookupId = origIdOfBomNode[String(n.id)] || String(n.id);
    const cy  = cycleById[lookupId]  ?? {};
    const sup = supplyById[lookupId] ?? {};

    // Parse supplier_1/quality_1 … supplier_3/quality_3
    const supplier_entries = [1, 2, 3].flatMap(i => {
      const sid = String(sup[`supplier_${i}`] ?? '').trim();
      const q   = Number(sup[`quality_${i}`]) || 0;
      return sid ? [{ supplierId: sid, historical_quality: q }] : [];
    });

    return {
      id:                  String(n.id),
      parent:              n.parent ? String(n.parent) : null,
      type:                String(n.type),
      label:               String(n.name),
      sublabel:            String(n.sublabel || ''),
      cycle_time_hrs:      Number(cy.cycle_time_hrs)      || 0,
      goal_cycle_time_hrs: Number(cy.goal_cycle_time_hrs) || 0,
      supplier_entries,
      validates:           [],
      technicians:         [],
    };
  });

  // ── Training matrix → technicians per doc ───────────────────────
  // Supports two score formats per technician cell:
  //   Plain number  "3"           → applies to every BOM parent of this doc
  //   Split format  "OA=3,AZ=2"  → per-BOM breakdown (multi-parent docs)
  //
  // trainingByDocId[docId] = { plain: [{name,score}], byBom: {bomId:[{name,score}]} }
  // getTechsForDocBom(td, bomNodeId) merges these: per-BOM scores win; plain fills the rest.
  const trainSh         = ss.getSheetByName(SHEET_NAMES.TRAINING);
  const trainingByDocId = {};

  if (trainSh && trainSh.getLastRow() > 1 && trainSh.getLastColumn() > TRAINING_BASE_COLS) {
    const trainData    = trainSh.getRange(1, 1, trainSh.getLastRow(), trainSh.getLastColumn()).getValues();
    const techInitials = trainData[0].map(h => String(h).trim()).slice(TRAINING_BASE_COLS);

    trainData.slice(1).forEach(row => {
      const docId = String(row[0]).trim();
      if (!docId) return;
      const plain  = [];
      const byBom  = {};

      techInitials.forEach((initials, i) => {
        const raw = String(row[TRAINING_BASE_COLS + i] ?? '').trim();
        if (!raw) return;

        if (raw.includes('=')) {
          // Per-BOM: "OA=3,AZ=2"
          raw.split(',').forEach(part => {
            const eq = part.indexOf('=');
            if (eq === -1) return;
            const bomId = part.slice(0, eq).trim();
            const score = Number(part.slice(eq + 1).trim());
            if (bomId && !isNaN(score)) {
              (byBom[bomId] = byBom[bomId] || []).push({
                id: initials,
                name: techNameById[initials] || initials,
                score,
              });
            }
          });
        } else if (!isNaN(Number(raw))) {
          plain.push({
            id: initials,
            name: techNameById[initials] || initials,
            score: Number(raw)
          });
        }
      });

      trainingByDocId[docId] = { plain, byBom };
    });
  }

  // Returns the technician list for a specific BOM expansion of a doc.
  // Per-BOM entries override plain ones for the same technician name.
  function getTechsForDocBom(td, bomNodeId) {
    if (!td) return [];
    const bomSpecific = td.byBom?.[bomNodeId] ?? [];
    const bomIds      = new Set(bomSpecific.map(t => t.id || t.name));
    const plain       = (td.plain ?? []).filter(t => !bomIds.has(t.id || t.name));
    return [...plain, ...bomSpecific];
  }

  // ── Testing matrices → validates per doc ────────────────────────
  function readTestingMatrix(shName) {
    const sh = ss.getSheetByName(shName);
    if (!sh || sh.getLastRow() < 2 || sh.getLastColumn() <= TESTING_BASE_COLS) return {};
    const data      = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
    const bomColIds = data[0].slice(TESTING_BASE_COLS).map(h => String(h).trim());
    const result    = {};
    data.slice(1).forEach(row => {
      const docId = String(row[0]).trim();
      if (!docId) return;
      result[docId] = {};
      bomColIds.forEach((bomId, i) => {
        const val = row[TESTING_BASE_COLS + i];
        if (val !== '' && val !== null) result[docId][bomId] = val;
      });
    });
    return result;
  }

  const testingScores = readTestingMatrix(SHEET_NAMES.TESTING);
  const testingNotes  = readTestingMatrix(SHEET_NAMES.TESTING_NOTES);

  // ── Doc cycle times (Cycle_Docs overrides Doc_Nodes) ────────────
  const cycleDocRows = sheetToObjects(SHEET_NAMES.CYCLE_DOCS);
  const cycleDocById = Object.fromEntries(cycleDocRows.map(r => [String(r.doc_id), r]));

  // Parses a leads_to / linked_to cell value which is either:
  //   - a plain doc ID  → { single: 'DOCX', byBom: null }
  //   - split format "OA=DOCX,AZ=DOCY" → { single: null, byBom: { OA: 'DOCX', AZ: 'DOCY' } }
  // The split format lets a multi-parent doc point to different targets in each BOM.
  function parseLinkField(raw) {
    const str = String(raw ?? '').trim();
    if (!str) return { single: null, byBom: null };
    if (!str.includes('=')) return { single: str, byBom: null };
    const byBom = {};
    str.split(',').forEach(part => {
      const eq = part.indexOf('=');
      if (eq === -1) return;
      const bomId = part.slice(0, eq).trim();
      const docId = part.slice(eq + 1).trim();
      if (bomId && docId) byBom[bomId] = docId;
    });
    return Object.keys(byBom).length ? { single: null, byBom } : { single: str, byBom: null };
  }

  // Parses a cycle_time_hrs cell value which is either:
  //   - a plain number → { total: n, byBom: null }
  //   - split format "A1=3,A2=4" → { total: 7, byBom: { A1: 3, A2: 4 } }
  function parseCycleTime(raw) {
    const str = String(raw ?? '').trim();
    if (!str) return { total: 0, byBom: null };
    const asNum = Number(str);
    if (!isNaN(asNum)) return { total: asNum, byBom: null };
    // Split format
    const byBom = {};
    let total = 0;
    str.split(',').forEach(part => {
      const eq = part.indexOf('=');
      if (eq === -1) return;
      const key = part.slice(0, eq).trim();
      const val = Number(part.slice(eq + 1).trim());
      if (key && !isNaN(val)) { byBom[key] = val; total += val; }
    });
    return Object.keys(byBom).length ? { total, byBom } : { total: 0, byBom: null };
  }

  // ── Doc nodes ───────────────────────────────────────────────────
  const docMaster = sheetToObjects(SHEET_NAMES.DOCS);

  // Pre-build a map of origDocId → { expandedBomNodeId → suffixedDocId }.
  // Accounts for both:
  //   (a) docs with multiple raw bom_node_id entries (comma-separated)
  //   (b) docs whose single bom_node_id was expanded due to BOM multi-parent expansion
  // Used in the post-pass to rewrite leads_to / linked_to references.
  const docIdMap = {};
  docMaster.forEach(d => {
    const origId    = String(d.id);
    const rawBomIds = String(d.bom_node_id).split(',').map(s => s.trim()).filter(Boolean);
    // Resolve each raw BOM id through the expansion map
    const bomIds    = rawBomIds.flatMap(id => bomExpansionMap[id] || [id]);
    if (bomIds.length > 1) {
      docIdMap[origId] = {};
      bomIds.forEach(bomId => { docIdMap[origId][bomId] = `${origId}__${bomId}`; });
    }
  });

  const docNodes = docMaster.flatMap(d => {
    const origId    = String(d.id);
    const rawBomIds = String(d.bom_node_id).split(',').map(s => s.trim()).filter(Boolean);
    // Expand raw BOM node ids through the BOM multi-parent expansion map.
    // e.g. if NODE_A was expanded to NODE_A__P1 / NODE_A__P2, a doc linked
    // to NODE_A will automatically get instances in both sub-trees.
    const bomIds = rawBomIds.flatMap(id => bomExpansionMap[id] || [id]);
    const multi  = bomIds.length > 1;
    const cy     = cycleDocById[origId] ?? {};

    const scoreMap  = testingScores[origId] ?? {};
    const noteMap   = testingNotes[origId]  ?? {};
    const allBomIds = new Set([...Object.keys(scoreMap), ...Object.keys(noteMap)]);
    const validates = [...allBomIds].flatMap(bomId => {
      const score = Number(scoreMap[bomId]) || 0;
      const note  = String(noteMap[bomId]  ?? '').trim();
      if (!score && !note) return [];
      return [{ id: bomId, score, note }];
    });

    // Prefer Cycle_Docs value; fall back to Doc_Nodes plain number
    const rawCt = (cy.cycle_time_hrs !== undefined && String(cy.cycle_time_hrs).trim() !== '')
      ? cy.cycle_time_hrs
      : (d.cycle_time_hrs ?? '');
    const { total: ctTotal, byBom: ctByBom } = parseCycleTime(rawCt);

    // Goal cycle time: same split format as actual (OA=10,AZ=10)
    const rawGt = String(cy.goal_cycle_time_hrs ?? '').trim();
    const { total: gtTotal, byBom: gtByBom } = parseCycleTime(rawGt);

    // Per-BOM leads_to / linked_to (format: "OA=DOCX,AZ=DOCY" or plain "DOCX")
    const { single: ltSingle, byBom: ltByBom } = parseLinkField(d.leads_to);
    const { single: liSingle, byBom: liByBom } = parseLinkField(d.linked_to);

    return bomIds.map(bomNodeId => {
      const nodeId = multi ? `${origId}__${bomNodeId}` : origId;
      // For multi-parent docs, use the per-BOM split cycle time if available
      const ct = (multi && ctByBom) ? (ctByBom[bomNodeId] ?? ctTotal) : ctTotal;
      const gt = (multi && gtByBom) ? (gtByBom[bomNodeId] ?? gtTotal) : gtTotal;
      // Pick the per-BOM link target, falling back to the single (plain) value
      const leadsToRaw  = ltByBom ? (ltByBom[bomNodeId] ?? null) : ltSingle;
      const linkedToRaw = liByBom ? (liByBom[bomNodeId] ?? null) : liSingle;
      return {
        id:             nodeId,
        bomNodeId,
        // tests_node_id: which BOM nodes this doc validates (may differ from bomNodeId
        // for test/checklist docs that validate sub-assemblies of their parent BOM).
        // Empty string means "same as bom_node_id" (single-parent, self-validates).
        tests_node_id:  String(d.tests_node_id || '').trim(),
        type:           String(d.type),
        label:          String(d.label),
        doc_num:        String(d.doc_num  || ''),
        score:          Number(d.score)   || 0,
        // leads_to / linked_to resolved in post-pass below
        leads_to:       leadsToRaw  || null,
        linked_to:      linkedToRaw || null,
        cycle_time_hrs:      ct,
        goal_cycle_time_hrs: gt,
        // cycle_time_by_bom only emitted for single-parent docs using split format
        ...((!multi && ctByBom) ? { cycle_time_by_bom: ctByBom } : {}),
        validates,
        technicians:    getTechsForDocBom(trainingByDocId[origId], bomNodeId),
      };
    });
  });

  // Post-pass: rewrite leads_to / linked_to for multi-parent doc references.
  // If doc A leads_to doc B, and both are expanded for the same bomNodeId,
  // A__OA should lead_to B__OA rather than the raw B id.
  docNodes.forEach(dn => {
    ['leads_to', 'linked_to'].forEach(field => {
      const raw = dn[field];
      if (raw && docIdMap[raw]?.[dn.bomNodeId]) {
        dn[field] = docIdMap[raw][dn.bomNodeId];
      }
    });
  });

  // ── Suppliers ───────────────────────────────────────────────────
  const suppliers = sheetToObjects(SHEET_NAMES.SUPPLIERS).map(s => ({
    id:          String(s.id),
    name:        String(s.name),
    trust_score: Number(s.trust_score) || 0,
  }));

  return { nodes, docNodes, supplierRegistry: suppliers, techRegistry };
}
