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
      2: 'Parent node ID — leave blank for root',
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
      3: 'assembly | test | checklist | reference',
      4: 'Document display name',
      5: 'Document number (e.g. SOP-011)',
      6: 'Readiness score 0–3',
      7: 'ID of the doc this one leads TO (next step in the flow)',
      8: 'ID of a doc this one is side-linked to (dashed branch)',
      9: 'BOM node(s) this doc VALIDATES — comma-separated (e.g. MOTOR,GEARBOX).\nControls which Testing matrix columns are active (white) for this doc.\nLeave blank to default to bom_node_id.\nDoes NOT affect cycle time or tree placement.',
    },
  },

  [SHEET_NAMES.TECHNICIANS]: {
    tab: { color: '#d97706' },
    darkTheme: false,
    headers: ['id', 'initials'],
    notes: {
      1: 'Technician ID (e.g. T1, T2) — used internally',
      2: 'Short initials shown as column header in Training matrix (e.g. AC, BS)\nRun Dashboard → Sync after adding a new technician.',
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
    headers: ['doc_id', 'bom_node_id', 'label', 'doc_num', 'type', 'cycle_time_hrs'],
    autoFill: [
      { col: 1, formula: `=ARRAYFORMULA(IF(Doc_Nodes!A2:A="","",Doc_Nodes!A2:A))` },
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE)))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
      { col: 4, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,5,FALSE)))` },
      { col: 5, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,3,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes — do not edit',
      4: 'Auto-filled doc number — for reference',
      5: 'Auto-filled doc type — for reference',
      6: 'Cycle time in hours — editable. Takes priority over Doc_Nodes on export.\nFor docs shared across multiple BOM nodes, enter per-BOM split: OA=10,AZ=10,AH=15\nEach BOM node then receives only its portion in the cycle time view.\n(Leave 0 on Cycle_BOM for those nodes so the doc sum drives the display.)',
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
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="assembly")+(Doc_Nodes!C2:C="test")),"")` },
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
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="test")+(Doc_Nodes!C2:C="checklist")),"")` },
      // col 2: use tests_node_id (Doc_Nodes col 10) if set; fall back to bom_node_id (col 2)
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",IF(IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),"")<>"",IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),""),IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE),""))))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
      { col: 4, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,3,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes (test + checklist types only) — do not edit',
      2: 'Auto-filled: shows tests_node_id from Doc_Nodes (which BOM nodes this doc validates).\nFalls back to bom_node_id if tests_node_id is blank.\nMatrix columns turn white for matching BOM node IDs.',
      4: 'Auto-filled type — always test or checklist',
    },
  },

  [SHEET_NAMES.TESTING_NOTES]: {
    tab: { color: '#059669' },
    darkTheme: false,
    headers: ['doc_id', 'tests_node_id', 'label', 'type'],
    autoFill: [
      { col: 1, formula: `=IFERROR(FILTER(Doc_Nodes!A2:A,(Doc_Nodes!C2:C="test")+(Doc_Nodes!C2:C="checklist")),"")` },
      // col 2: use tests_node_id (Doc_Nodes col 10) if set; fall back to bom_node_id (col 2)
      { col: 2, formula: `=ARRAYFORMULA(IF(A2:A="","",IF(IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),"")<>"",IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,9,FALSE),""),IFERROR(VLOOKUP(A2:A,Doc_Nodes!A:I,2,FALSE),""))))` },
      { col: 3, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,4,FALSE)))` },
      { col: 4, formula: `=ARRAYFORMULA(IF(A2:A="","",VLOOKUP(A2:A,Doc_Nodes!A:I,3,FALSE)))` },
    ],
    notes: {
      1: 'Auto-filled from Doc_Nodes (test + checklist types only) — do not edit',
      2: 'Auto-filled: shows tests_node_id from Doc_Nodes (which BOM nodes this doc validates).\nFalls back to bom_node_id if tests_node_id is blank.',
      4: 'Auto-filled type — for reference only',
    },
  },

  [SHEET_NAMES.EXPORT]: {
    tab: { color: '#64748b' },
    darkTheme: false,
    headers: ['json_output'],
    notes: { 1: 'Auto-generated by Dashboard → Export JSON. Copy cell A2 and paste into the dashboard Import dialog.' },
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
    'Next steps:\n1. Add BOM nodes to "BOM_Nodes"\n2. Add docs to "Doc_Nodes"\n3. Add suppliers to "Suppliers"\n4. Add technician initials to "Technicians"\n5. Run Dashboard → Sync\n6. Fill in view columns, then Export JSON',
    ui.ButtonSet.OK
  );
}


// ── Apply a sheet definition ────────────────────────────────────────
function applySheetDef(ss, sh, def) {
  const numCols    = def.headers.length;
  const darkTheme  = def.darkTheme !== false; // default true if omitted
  const baseBg     = darkTheme ? LOCKED_BG   : null;   // null = sheet default (white)
  const altBg      = darkTheme ? HDR_BG      : '#f1f5f9';
  const dataFgCol  = darkTheme ? '#e2e8f0'   : '#1e293b';

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
    dataRange
      .setFontColor(dataFgCol)
      .setFontFamily(DATA_FONT)
      .setFontSize(9)
      .setVerticalAlignment('middle');
  }
  sh.setRowHeightsForced(2, Math.max(sh.getMaxRows() - 1, 1), 26);

  // Auto-fill formulas + lock styling
  const lockedBg     = darkTheme ? LOCKED_BG      : LOCKED_LIGHT_BG;
  const lockedFgCol  = darkTheme ? '#94a3b8'       : '#64748b';
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
        .requireValueInList(['assembly','test','checklist','reference'], true)
        .setAllowInvalid(false).build();
    } else if (h === 'score') {
      rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(['0','1','2','3'], true)
        .setAllowInvalid(false).build();
    } else if (h === 'trust_score') {
      rule = SpreadsheetApp.newDataValidation()
        .requireNumberBetween(1, 5).setAllowInvalid(false).build();
    } else if (['cycle_time_hrs','goal_cycle_time_hrs'].includes(h)) {
      if (h === 'cycle_time_hrs' && sh.getName() === SHEET_NAMES.CYCLE_DOCS) {
        // No validation — column accepts plain numbers and "A1=3,A2=4" split format
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
function syncTrainingMatrix(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const techSh  = ss.getSheetByName(SHEET_NAMES.TECHNICIANS);
  const trainSh = ss.getSheetByName(SHEET_NAMES.TRAINING);
  if (!techSh || !trainSh) return;

  const techData = techSh.getLastRow() > 1
    ? techSh.getRange(2, 2, techSh.getLastRow() - 1, 1).getValues().flat()
        .map(v => String(v).trim()).filter(Boolean)
    : [];
  if (!techData.length) return;

  const headerRow     = trainSh.getRange(1, 1, 1, trainSh.getLastColumn()).getValues()[0];
  const existingTechs = headerRow.slice(TRAINING_BASE_COLS).map(h => String(h).trim());

  // ── Pass 1: add any missing technician columns ────────────────
  techData.forEach(initials => {
    if (existingTechs.includes(initials)) return;
    const newCol = trainSh.getLastColumn() + 1;
    if (newCol > trainSh.getMaxColumns()) trainSh.insertColumnsAfter(trainSh.getMaxColumns(), 1);

    trainSh.getRange(1, newCol)
      .setValue(initials)
      .setBackground(TECH_HDR_BG).setFontColor(TECH_HDR_FG)
      .setFontFamily(HDR_FONT).setFontWeight('bold')
      .setFontSize(10).setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setNote(`Training score for ${initials}\nEnter 0–3:\n  blank = not assigned\n  1 = awareness\n  2 = supervised\n  3 = qualified`);

    trainSh.setColumnWidth(newCol, TECH_COL_W);
    existingTechs.push(initials);
  });

  // ── Re-format ALL technician data columns (overwrites old dark styling) ──
  const maxRows = Math.max(trainSh.getMaxRows() - 1, 1);
  existingTechs.forEach((initials, i) => {
    if (!initials) return;
    const colIdx = TRAINING_BASE_COLS + 1 + i;
    trainSh.getRange(2, colIdx, maxRows, 1)
      .setBackground(null).setFontColor('#1e293b')
      .setHorizontalAlignment('center').setFontFamily(DATA_FONT).setFontSize(10)
      .setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireNumberBetween(0, 3).setAllowInvalid(true)
          .setHelpText('0 = not trained, 1 = awareness, 2 = supervised, 3 = qualified').build()
      );
  });

  trainSh.setFrozenColumns(TRAINING_BASE_COLS);
}


// ═══════════════════════════════════════════════════════════════════
//  TESTING MATRIX (validates scores + notes)
// ═══════════════════════════════════════════════════════════════════
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

    const isNotes      = shName === SHEET_NAMES.TESTING_NOTES;
    const headerRow    = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const existingCols = headerRow.slice(TESTING_BASE_COLS).map(h => String(h).trim());

    // ── Pass 1: add any missing BOM node columns (headers + validation only) ──
    bomNodes.forEach(({ id, name }) => {
      if (existingCols.includes(id)) return;
      const newCol = sh.getLastColumn() + 1;
      if (newCol > sh.getMaxColumns()) sh.insertColumnsAfter(sh.getMaxColumns(), 1);

      sh.getRange(1, newCol)
        .setValue(id)
        .setBackground(isNotes ? '#dbeafe' : '#ede9fe')
        .setFontColor(isNotes ? '#1e40af' : '#4c1d95')
        .setFontFamily(HDR_FONT).setFontWeight('bold')
        .setFontSize(9).setHorizontalAlignment('center').setVerticalAlignment('middle')
        .setNote(`BOM node: ${name} (${id})\n${isNotes
          ? 'Optional note for this validates edge.\nLeave blank if none.'
          : 'Validation score 0–3:\n  blank = not validated\n  1 = partial\n  2 = functional\n  3 = fully validated'}`);

      sh.setColumnWidth(newCol, isNotes ? 160 : TECH_COL_W);
      if (!isNotes) {
        sh.getRange(2, newCol, Math.max(sh.getMaxRows() - 1, 1), 1)
          .setDataValidation(
            SpreadsheetApp.newDataValidation()
              .requireNumberBetween(0, 3).setAllowInvalid(true)
              .setHelpText('0 = not validated, 1 = partial, 2 = functional, 3 = fully validated').build()
          );
      }
      existingCols.push(id);
    });

    // ── Re-format ALL matrix data columns (overwrites any old dark styling) ──
    // Grey base; conditional rules in Pass 2 override to white for active cells.
    const maxDataRows = Math.max(sh.getMaxRows() - 1, 1);
    existingCols.forEach((bomId, i) => {
      if (!bomId) return;
      const colIdx  = TESTING_BASE_COLS + 1 + i;
      const dataRange = sh.getRange(2, colIdx, maxDataRows, 1);
      dataRange.setBackground('#e2e8f0').setFontColor('#1e293b');
      if (isNotes) {
        dataRange.setFontFamily(HDR_FONT).setFontSize(9).setWrap(true).setVerticalAlignment('top');
      } else {
        dataRange.setHorizontalAlignment('center').setFontFamily(DATA_FONT).setFontSize(10);
      }
    });

    // ── Pass 2: rebuild all conditional format rules ──────────────
    // White match rules come first so they beat the alternating-row rule.
    // A cell turns white when its row's bom_node_id contains this column's ID
    // (bom_node_id may be comma-separated, e.g. "A1,A2").
    const matchRules = existingCols.map((bomId, i) => {
      if (!bomId) return null;
      const colIdx = TESTING_BASE_COLS + 1 + i;
      return SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=ISNUMBER(SEARCH(","&"${bomId}"&",",","&TRIM($B2)&","))`)
        .setBackground('#ffffff')
        .setRanges([sh.getRange(2, colIdx, maxDataRows, 1)])
        .build();
    }).filter(Boolean);

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
function syncViewSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  _syncMasterToViews(ss, SHEET_NAMES.BOM,  [SHEET_NAMES.OVERVIEW, SHEET_NAMES.CYCLE_BOM]);
  _syncMasterToViews(ss, SHEET_NAMES.DOCS, [SHEET_NAMES.CYCLE_DOCS]);
  // Training/Testing/Testing_Notes rows are all driven by FILTER formulas — no row sync needed.
  // Training:       assembly + test types only
  // Testing:        test + checklist types only
  // Testing_Notes:  test + checklist types only
  syncSupplySheet(ss);
  syncTrainingMatrix(ss);
  syncTestingMatrix(ss);
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
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const out = buildDashboardJSON(ss);

  let exportSh = ss.getSheetByName(SHEET_NAMES.EXPORT);
  if (!exportSh) exportSh = ss.insertSheet(SHEET_NAMES.EXPORT);
  exportSh.clearContents();
  exportSh.getRange('A1').setValue('json_output');
  exportSh.getRange('A2').setValue(out).setWrap(false);
  exportSh.setColumnWidth(1, 800);

  ss.setActiveSheet(exportSh);
  exportSh.setActiveRange(exportSh.getRange('A2'));

  SpreadsheetApp.getUi().alert(
    '✅ Export ready',
    'Cell A2 is selected in "_Export".\nPress ⌘C / Ctrl+C, then paste into the dashboard Import dialog.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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

  // ── BOM nodes ───────────────────────────────────────────────────
  const bomRows    = sheetToObjects(SHEET_NAMES.BOM);
  const cycleRows  = sheetToObjects(SHEET_NAMES.CYCLE_BOM);
  const supplyRows = sheetToObjects(SHEET_NAMES.SUPPLY);
  const cycleById  = Object.fromEntries(cycleRows.map(r  => [String(r.id), r]));
  const supplyById = Object.fromEntries(supplyRows.map(r => [String(r.id), r]));

  const nodes = bomRows.map(n => {
    const cy  = cycleById[String(n.id)]  ?? {};
    const sup = supplyById[String(n.id)] ?? {};

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
  const trainSh         = ss.getSheetByName(SHEET_NAMES.TRAINING);
  const trainingByDocId = {};

  if (trainSh && trainSh.getLastRow() > 1 && trainSh.getLastColumn() > TRAINING_BASE_COLS) {
    const trainData    = trainSh.getRange(1, 1, trainSh.getLastRow(), trainSh.getLastColumn()).getValues();
    const techInitials = trainData[0].map(h => String(h).trim()).slice(TRAINING_BASE_COLS);

    trainData.slice(1).forEach(row => {
      const docId = String(row[0]).trim();
      if (!docId) return;
      const technicians = [];
      techInitials.forEach((initials, i) => {
        const raw = row[TRAINING_BASE_COLS + i];
        if (raw !== '' && raw !== null && !isNaN(Number(raw))) {
          technicians.push({ name: initials, score: Number(raw) });
        }
      });
      trainingByDocId[docId] = technicians;
    });
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

  // Pre-build a map of origId → { bomNodeId → suffixedId } for multi-parent docs.
  // Used in the post-pass to rewrite leads_to / linked_to references.
  const docIdMap = {};
  docMaster.forEach(d => {
    const origId = String(d.id);
    const bomIds = String(d.bom_node_id).split(',').map(s => s.trim()).filter(Boolean);
    if (bomIds.length > 1) {
      docIdMap[origId] = {};
      bomIds.forEach(bomId => { docIdMap[origId][bomId] = `${origId}__${bomId}`; });
    }
  });

  const docNodes = docMaster.flatMap(d => {
    const origId = String(d.id);
    const bomIds = String(d.bom_node_id).split(',').map(s => s.trim()).filter(Boolean);
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

    return bomIds.map(bomNodeId => {
      const nodeId = multi ? `${origId}__${bomNodeId}` : origId;
      // For multi-parent docs, use the per-BOM split cycle time if available
      const ct = (multi && ctByBom) ? (ctByBom[bomNodeId] ?? ctTotal) : ctTotal;
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
        leads_to:       d.leads_to  ? String(d.leads_to)  : null,
        linked_to:      d.linked_to ? String(d.linked_to) : null,
        cycle_time_hrs: ct,
        // cycle_time_by_bom only emitted for single-parent docs using split format
        ...((!multi && ctByBom) ? { cycle_time_by_bom: ctByBom } : {}),
        validates,
        technicians:    trainingByDocId[origId] ?? [],
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

  return JSON.stringify({ nodes, docNodes, supplierRegistry: suppliers }, null, 2);
}
