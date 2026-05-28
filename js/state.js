/* ═══════════════════════════════════════════
   HeriExcel.Pro — state.js  v7.3
   Estado global, almacenes de formato y
   constantes compartidas por todos los módulos

   FIXES v7.3:
   - Añadido isFillDragging correcto (bool)
   - cellNotes inicializado como objeto vacío
   - fillAnchorR/C para el fill-handle
═══════════════════════════════════════════ */
'use strict';

const STATE = {
  workbook:       null,
  sheetNames:     [],
  activeSheet:    '',
  data:           [],
  wsRaw:          null,
  mergeMap:       {},
  fileName:       '',
  selectedRow:    -1,
  activeCellR:    -1,
  activeCellC:    -1,
  selectionStart: { r: -1, c: -1 },
  selectionEnd:   { r: -1, c: -1 },
  selectedCells:  new Set(),
  editingCell:    null,
  dirty:          false,
  frozen:         false,
  zoom:           100,
  undoStack:      [],
  redoStack:      [],
  isDragging:     false,
  isFillDragging: false,
  fillStart:      null,
  fillEnd:        null,
  fillAnchorR:    -1,
  fillAnchorC:    -1,
  borderColor:    '#7c8cf8',
  borderStyle:    'solid',
  clipboard:      null,
  activeFilters:  {},
  colWidths:      {},
  rowHeights:     {},
  vsRowHeight:    28,
  vsBuffer:       10,
  vsStart:        0,
  vsEnd:          50,
  cellNotes:      {},
  sortOrders:     [],
  datePicker:     null,
};

/* Almacén de formato por celda  { "hoja:r,c": { bold, italic, ... } } */
const FMT = {};

/* Almacén de bordes por celda   { "hoja:r,c": { top, bottom, left, right } } */
const BORDERS = {};

/* Reglas de formato condicional */
let COND_RULES = [];
