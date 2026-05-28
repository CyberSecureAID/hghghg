/* ===============================================
   HeriExcel.Pro — app-core.js  v7.5
   Tema, toast, loading, dropdowns,
   undo/redo, portapapeles, seleccion,
   edicion de celda, autocompletado,
   teclado, barra de formato

   FIXES v7.5 sobre v7.4:
   - B-12: selectEntireRow resetea activeCellR/C a -1
     al deseleccionar fila (toggle OFF).
   - B-13: exitEdit(false) restaura fxInput al valor
     original de la celda al cancelar con Escape.
   - B-16: _autocompleteEl inicializado a null con
     guard en hideAutocomplete().
   - B-53: clearSelection limpia clases td-selected Y
     td-active de todos los TD del DOM antes de limpiar
     el Set, evitando estilos huérfanos tras re-render.
   - B-54: activateCell ya no llama exitEdit(true)
     incondicionalmente al inicio; solo lo hace si hay
     una celda en edición diferente a (r,c).
   - B-55: syncFormatBarToCell maneja correctamente el
     caso r<0/c<0 (limpia la barra de fórmula).
   - B-56: updateColHighlight llama a querySelectorAll
     con selector más robusto.
   - B-57: handleCellMousedown previene doble llamada a
     exitEdit cuando se hace clic en la celda en edición.
   - B-58: fxInput.input no intenta modificar el editor
     si no hay celda activa.
   - B-59: getTargetCells retorna [] de forma segura
     cuando no hay selección ni celda activa.
=============================================== */
'use strict';

/* ================================================
   THEME
================================================ */
function applyTheme(theme) {
  document.documentElement.setAttribute('data-theme', theme);
  localStorage.setItem('heriexcel-theme', theme);
  $('iconSun').style.display  = theme === 'dark' ? ''     : 'none';
  $('iconMoon').style.display = theme === 'dark' ? 'none' : '';
}

(function initTheme() {
  applyTheme(localStorage.getItem('heriexcel-theme') || 'dark');
})();

$('themeToggle').addEventListener('click', function() {
  applyTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark');
});

/* ================================================
   TOAST
================================================ */
var _toastTimer = null;
function showToast(msg, type) {
  if (type === undefined) type = '';
  var el = $('toast');
  if (!el) return;
  el.textContent = msg;
  el.className   = 'toast ' + type + ' show';
  clearTimeout(_toastTimer);
  _toastTimer = setTimeout(function() { el.classList.remove('show'); }, 3000);
}

/* ================================================
   LOADING
================================================ */
function showLoading(msg) {
  var ov = $('loadingOverlay');
  if (!ov) {
    ov = document.createElement('div');
    ov.id = 'loadingOverlay';
    ov.className = 'loading-overlay';
    ov.innerHTML = '<div class="loading-spinner"></div><div class="loading-text" id="loadingText"></div>';
    document.body.appendChild(ov);
  }
  var lt = ov.querySelector('#loadingText');
  if (lt) lt.textContent = msg || 'Cargando...';
  ov.classList.remove('hidden');
}
function hideLoading() {
  var ov = $('loadingOverlay');
  if (ov) ov.classList.add('hidden');
}

/* ================================================
   DROPDOWNS
================================================ */
var DD_LIST = [
  { btnId: 'btnRowsMenu',      menuId: 'ddRowsMenu'       },
  { btnId: 'btnColsMenu',      menuId: 'ddColsMenu'       },
  { btnId: 'btnReplaceMenu',   menuId: 'ddReplaceMenu'    },
  { btnId: 'btnExportMenu',    menuId: 'ddExportMenu'     },
  { btnId: 'btnCellStyleMenu', menuId: 'ddCellStylesMenu' },
  { btnId: 'btnBordersMenu',   menuId: 'ddBordersMenu'    },
  { btnId: 'btnCondFmtMenu',   menuId: 'ddCondFmtMenu'    },
];

function closeAllDropdowns(except) {
  DD_LIST.forEach(function(d) {
    var el = $(d.menuId);
    if (el && el !== except) el.classList.remove('open');
  });
}

function openDropdownAt(btnEl, menuEl) {
  var rect   = btnEl.getBoundingClientRect();
  var scrollX = window.scrollX || window.pageXOffset || 0;
  var scrollY = window.scrollY || window.pageYOffset || 0;
  menuEl.style.top   = (rect.bottom + 4) + 'px';
  menuEl.style.left  = rect.left + 'px';
  menuEl.style.right = 'auto';
  menuEl.classList.add('open');
  requestAnimationFrame(function() {
    var mRect = menuEl.getBoundingClientRect();
    if (mRect.right  > window.innerWidth  - 10) menuEl.style.left = Math.max(8, rect.right - mRect.width) + 'px';
    if (mRect.bottom > window.innerHeight - 10) menuEl.style.top  = Math.max(4, rect.top - mRect.height - 4) + 'px';
  });
}

DD_LIST.forEach(function(d) {
  var btn  = $(d.btnId);
  var menu = $(d.menuId);
  if (!btn || !menu) return;
  btn.addEventListener('click', function(e) {
    e.stopPropagation();
    var wasOpen = menu.classList.contains('open');
    closeAllDropdowns(null);
    if (!wasOpen) openDropdownAt(btn, menu);
  });
});

document.addEventListener('click', function(e) {
  if (!e.target.closest('.ribbon-dropdown')) {
    closeAllDropdowns(null);
  }
});

/* ================================================
   UNDO / REDO
================================================ */
function pushUndo() {
  STATE.undoStack.push(JSON.stringify(STATE.data));
  if (STATE.undoStack.length > 200) STATE.undoStack.shift();
  STATE.redoStack = [];
}

function undo() {
  if (!STATE.undoStack.length) return showToast('Nada que deshacer', '');
  STATE.redoStack.push(JSON.stringify(STATE.data));
  STATE.data = JSON.parse(STATE.undoStack.pop());
  exitEdit(false); renderTable(); updateStats(); markDirty();
  showToast('Acción deshecha', 'info');
}

function redo() {
  if (!STATE.redoStack.length) return showToast('Nada que rehacer', '');
  STATE.undoStack.push(JSON.stringify(STATE.data));
  STATE.data = JSON.parse(STATE.redoStack.pop());
  exitEdit(false); renderTable(); updateStats(); markDirty();
  showToast('Acción rehecha', 'info');
}

$('btnUndo').addEventListener('click', undo);
$('btnRedo').addEventListener('click', redo);

/* ================================================
   CLIPBOARD
================================================ */
function copyCells(cut) {
  if (!STATE.selectedCells.size && STATE.activeCellR < 0) {
    return showToast('Selecciona celdas primero', '');
  }
  var cellsToUse = STATE.selectedCells.size > 0
    ? Array.from(STATE.selectedCells).map(function(k) { return k.split(',').map(Number); })
    : [[STATE.activeCellR, STATE.activeCellC]];

  var minR = Math.min.apply(null, cellsToUse.map(function(rc) { return rc[0]; }));
  var minC = Math.min.apply(null, cellsToUse.map(function(rc) { return rc[1]; }));

  var data = cellsToUse.map(function(rc) {
    var r = rc[0], c = rc[1];
    return {
      r:   r - minR,
      c:   c - minC,
      val: (STATE.data[r] ? STATE.data[r][c] : '') || '',
      fmt: Object.assign({}, getCellFmt(r, c))
    };
  });

  STATE.clipboard = { data: data, mode: cut ? 'cut' : 'copy', minR: minR, minC: minC };

  var maxR2  = Math.max.apply(null, cellsToUse.map(function(rc) { return rc[0]; }));
  var maxC2  = Math.max.apply(null, cellsToUse.map(function(rc) { return rc[1]; }));
  var rows2  = maxR2 - minR + 1, cols2 = maxC2 - minC + 1;
  var grid   = [];
  for (var i = 0; i < rows2; i++) {
    grid.push([]);
    for (var j = 0; j < cols2; j++) grid[i].push('');
  }
  cellsToUse.forEach(function(rc) {
    var r = rc[0], c = rc[1];
    grid[r - minR][c - minC] = (STATE.data[r] ? STATE.data[r][c] : '') || '';
  });

  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(grid.map(function(row) { return row.join('\t'); }).join('\n')).catch(function() {});
  }
  showToast(cut ? ('Cortadas: ' + cellsToUse.length + ' celda(s)') : ('Copiadas: ' + cellsToUse.length + ' celda(s)'), 'info');
}

function pasteCells(valuesOnly) {
  if (valuesOnly === undefined) valuesOnly = false;
  if (STATE.activeCellR < 0) return showToast('Selecciona una celda destino', '');
  if (navigator.clipboard && navigator.clipboard.readText) {
    navigator.clipboard.readText().then(function(text) {
      if (text && text.includes('\t')) {
        pushUndo();
        var r0 = STATE.activeCellR, c0 = STATE.activeCellC;
        text.split('\n').forEach(function(line, ri) {
          line.split('\t').forEach(function(val, ci) {
            var tr = r0 + ri, tc = c0 + ci;
            if (tr < STATE.data.length && tc < (STATE.data[0] ? STATE.data[0].length : 0)) {
              STATE.data[tr][tc] = val;
            }
          });
        });
        renderTable(); markDirty();
        showToast('Pegado desde portapapeles', 'success');
      } else {
        pasteFromInternal(valuesOnly);
      }
    }).catch(function() { pasteFromInternal(valuesOnly); });
  } else {
    pasteFromInternal(valuesOnly);
  }
}

function pasteFromInternal(valuesOnly) {
  if (valuesOnly === undefined) valuesOnly = false;
  if (!STATE.clipboard) return showToast('Portapapeles vacío', '');
  pushUndo();
  var r0 = STATE.activeCellR, c0 = STATE.activeCellC;
  STATE.clipboard.data.forEach(function(item) {
    var tr = r0 + item.r, tc = c0 + item.c;
    if (tr < STATE.data.length && tc < (STATE.data[0] ? STATE.data[0].length : 0)) {
      STATE.data[tr][tc] = item.val;
      if (!valuesOnly) setCellFmt(tr, tc, item.fmt);
    }
  });
  if (STATE.clipboard.mode === 'cut') {
    STATE.clipboard.data.forEach(function(item) {
      var tr = STATE.clipboard.minR + item.r, tc = STATE.clipboard.minC + item.c;
      if (tr < STATE.data.length && tc < (STATE.data[0] ? STATE.data[0].length : 0)) {
        STATE.data[tr][tc] = '';
      }
    });
    STATE.clipboard = null;
  }
  renderTable(); markDirty();
  showToast(valuesOnly ? 'Valores pegados' : 'Celdas pegadas', 'success');
}

/* ================================================
   SELECTION

   B-53 FIX: clearSelection también limpia las clases
   td-selected y td-active de todos los TD del DOM
   antes de limpiar el Set, evitando estilos huérfanos.
================================================ */
function getCellEl(r, c) {
  var tbody = $('tableBody'); if (!tbody) return null;
  var tr = tbody.querySelector('tr[data-row="' + r + '"]'); if (!tr) return null;
  return tr.querySelector('td[data-col="' + c + '"]') || null;
}
function getCellSpan(r, c) { var td = getCellEl(r, c); return td ? td.querySelector('.cell-display') : null; }

function clearSelection() {
  /* B-53 FIX: limpiar clases del DOM antes de vaciar el Set */
  var tbody = $('tableBody');
  if (tbody) {
    tbody.querySelectorAll('.td-selected, .td-active').forEach(function(el) {
      el.classList.remove('td-selected', 'td-active');
    });
  }
  STATE.selectedCells.clear();
}

function addToSelection(r, c) {
  var key = r + ',' + c;
  STATE.selectedCells.add(key);
  var td = getCellEl(r, c);
  if (td) td.classList.add('td-selected');
}

function markActiveCell(r, c) {
  var td = getCellEl(r, c);
  if (td) td.classList.add('td-active');
}

function selectRange(r1, c1, r2, c2) {
  clearSelection();
  var minR = Math.min(r1, r2), maxR = Math.max(r1, r2);
  var minC = Math.min(c1, c2), maxC = Math.max(c1, c2);
  for (var r = minR; r <= maxR; r++) {
    for (var c = minC; c <= maxC; c++) {
      if (STATE.mergeMap[r + ',' + c] !== 'skip') addToSelection(r, c);
    }
  }
  updateSelectionInfo();
}

function updateSelectionInfo() {
  var count = STATE.selectedCells.size;
  var sep   = $('selInfoSep'), info = $('statSelInfo');
  if (count > 1) {
    if (sep)  sep.style.display  = '';
    if (info) info.style.display = '';
    var nums = [];
    STATE.selectedCells.forEach(function(key) {
      var parts = key.split(',').map(Number);
      var v = parseFloat((STATE.data[parts[0]] ? STATE.data[parts[0]][parts[1]] : '') || '');
      if (!isNaN(v)) nums.push(v);
    });
    if (nums.length > 0) {
      var sum = nums.reduce(function(a, b) { return a + b; }, 0);
      var min = Math.min.apply(null, nums);
      var max = Math.max.apply(null, nums);
      if (info) info.textContent = count + ' celdas · Suma: ' + sum.toFixed(2) +
        ' · Promedio: ' + (sum / nums.length).toFixed(2) + ' · Min: ' + min + ' · Max: ' + max;
    } else {
      if (info) info.textContent = count + ' celdas seleccionadas';
    }
  } else {
    if (sep)  sep.style.display  = 'none';
    if (info) info.style.display = 'none';
  }
}

function selectAll() {
  if (!STATE.data.length) return;
  clearSelection();
  var startRow = STATE.frozen ? 1 : 0;
  for (var r = startRow; r < STATE.data.length; r++) {
    for (var c = 0; c < (STATE.data[0] ? STATE.data[0].length : 0); c++) addToSelection(r, c);
  }
  STATE.activeCellR    = startRow;
  STATE.activeCellC    = 0;
  STATE.selectionStart = { r: startRow, c: 0 };
  updateSelectionInfo();
}

/*
 * B-12 FIX: selectEntireRow resetea activeCellR/C cuando se
 * deselecciona la fila (toggle OFF).
 */
function selectEntireRow(ri) {
  STATE.selectedRow = (STATE.selectedRow === ri) ? -1 : ri;
  clearSelection();
  if (STATE.selectedRow >= 0) {
    var cols = STATE.data[0] ? STATE.data[0].length : 0;
    for (var c = 0; c < cols; c++) addToSelection(ri, c);
    STATE.activeCellR = ri; STATE.activeCellC = 0;
    syncFormatBarToCell();
  } else {
    /* B-12 FIX: deselección → limpiar celda activa */
    STATE.activeCellR    = -1;
    STATE.activeCellC    = -1;
    STATE.selectionStart = { r: -1, c: -1 };
    STATE.selectionEnd   = { r: -1, c: -1 };
    syncFormatBarToCell();
  }
  var tbody2 = $('tableBody');
  if (tbody2) {
    tbody2.querySelectorAll('tr[data-row]').forEach(function(tr) {
      tr.classList.toggle('selected-row', parseInt(tr.dataset.row) === STATE.selectedRow);
    });
  }
  updateColHighlight();
  updateSelectionInfo();
}

function selectEntireCol(ci) {
  clearSelection();
  var startRow = STATE.frozen ? 1 : 0;
  for (var r = startRow; r < STATE.data.length; r++) addToSelection(r, ci);
  STATE.activeCellR    = startRow;
  STATE.activeCellC    = ci;
  STATE.selectionStart = { r: startRow, c: ci };
  updateColHighlight();
  updateSelectionInfo();
}

/* ================================================
   CELL ACTIVATION & EDIT

   B-54 FIX: activateCell solo llama exitEdit(true) si
   hay celda en edición diferente a (r,c), evitando
   guardar cambios prematuramente en doble clic.
================================================ */
function activateCell(r, c, extendSel, addSel) {
  if (extendSel === undefined) extendSel = false;
  if (addSel    === undefined) addSel    = false;

  /* B-54 FIX: solo salir del modo edición si la celda editada es diferente */
  if (STATE.editingCell && !(STATE.editingCell.r === r && STATE.editingCell.c === c)) {
    exitEdit(true);
  }

  if (extendSel && STATE.selectionStart.r >= 0) {
    STATE.selectionEnd = { r: r, c: c };
    selectRange(STATE.selectionStart.r, STATE.selectionStart.c, r, c);
    STATE.activeCellR = r; STATE.activeCellC = c;
  } else if (addSel) {
    STATE.activeCellR = r; STATE.activeCellC = c;
    addToSelection(r, c);
    updateSelectionInfo();
  } else {
    clearSelection();
    STATE.selectionStart = { r: r, c: c };
    STATE.selectionEnd   = { r: r, c: c };
    STATE.activeCellR    = r;
    STATE.activeCellC    = c;
    addToSelection(r, c);
    markActiveCell(r, c);
  }
  updateColHighlight();
  syncFormatBarToCell();
  scrollCellIntoView(r, c);
}

function scrollCellIntoView(r, c) {
  var td = getCellEl(r, c);
  if (td) td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/*
 * B-55 FIX: syncFormatBarToCell limpia correctamente la
 * barra cuando r<0 o c<0 (sin celda activa).
 */
function syncFormatBarToCell() {
  var r   = STATE.activeCellR, c = STATE.activeCellC;
  var fmt = (r >= 0 && c >= 0) ? getCellFmt(r, c) : {};

  $('fmtBold').classList.toggle('active',      !!fmt.bold);
  $('fmtItalic').classList.toggle('active',    !!fmt.italic);
  $('fmtUnderline').classList.toggle('active', !!fmt.underline);
  $('fmtStrike').classList.toggle('active',    !!fmt.strike);
  $('fmtFont').value = fmt.font || '';
  $('fmtSize').value = fmt.size || '';

  if (fmt.color)   {
    $('fmtColorInput').value   = toHex(fmt.color);
    $('fmtColorBar').style.background = fmt.color;
  }
  if (fmt.bgColor) {
    $('fmtBgColorInput').value = toHex(fmt.bgColor);
    $('fmtBgColorBar').style.background = fmt.bgColor;
  }

  ['Left','Center','Right'].forEach(function(a) {
    $('fmtAlign' + a).classList.toggle('active', (fmt.align || 'left') === a.toLowerCase());
  });
  $('fmtNumFormat').value = fmt.numFormat || 'general';

  var ref = (r >= 0 && c >= 0) ? (columnLabel(c) + (r + 1)) : '—';
  $('fxCellRef').textContent = ref;

  var rawVal = '';
  if (r >= 0 && c >= 0) {
    rawVal = (STATE.data[r] ? STATE.data[r][c] : '') || '';
  }
  $('fxInput').value = rawVal;
  $('statCell').textContent = (r >= 0 && c >= 0) ? (columnLabel(c) + (r + 1)) : 'Sin celda activa';
}

/*
 * B-56 FIX: updateColHighlight usa selector más robusto.
 */
function updateColHighlight() {
  var thead = $('tableHead'); if (!thead) return;
  thead.querySelectorAll('th[data-col]').forEach(function(th) {
    var colIdx = parseInt(th.getAttribute('data-col'));
    th.classList.toggle('col-active', !isNaN(colIdx) && colIdx === STATE.activeCellC);
  });
}

/* ────────────────────────────────────────────────
   enterEdit
   Antes de abrir el editor de texto, comprueba
   si la celda es de tipo fecha.
──────────────────────────────────────────────── */
function enterEdit(r, c) {
  /* Date Picker hook */
  if (typeof maybShowDatePicker === 'function' && maybShowDatePicker(r, c)) {
    return;
  }

  /* No volver a entrar en edición si ya estamos editando esta celda */
  if (STATE.editingCell && STATE.editingCell.r === r && STATE.editingCell.c === c) return;

  exitEdit(true);
  var td = getCellEl(r, c); if (!td) return;
  var display = td.querySelector('.cell-display'); if (!display) return;

  STATE.editingCell = { r: r, c: c };
  td.classList.add('td-editing');
  display.style.display = 'none';

  var input = document.createElement('div');
  input.className       = 'cell-editor';
  input.contentEditable = 'true';
  input.spellcheck      = false;

  var rawVal = (STATE.data[r] ? STATE.data[r][c] : '') || '';
  input.textContent = rawVal;

  var fmt = getCellFmt(r, c);
  if (fmt.font)   input.style.fontFamily = fmt.font;
  if (fmt.size)   input.style.fontSize   = fmt.size;
  if (fmt.bold)   input.style.fontWeight = '700';
  if (fmt.italic) input.style.fontStyle  = 'italic';
  if (fmt.color)  input.style.color      = fmt.color;
  if (fmt.align)  input.style.textAlign  = fmt.align;

  td.appendChild(input);
  input.focus();

  /* Mover cursor al final */
  var sel   = window.getSelection();
  var range = document.createRange();
  range.selectNodeContents(input);
  range.collapse(false);
  sel.removeAllRanges();
  sel.addRange(range);

  var targetR = r, targetC = c;
  input.addEventListener('keydown', function(e) { handleEditKeydown(e, targetR, targetC); });
  input.addEventListener('blur', function() {
    setTimeout(function() {
      if (STATE.editingCell && STATE.editingCell.r === targetR && STATE.editingCell.c === targetC) {
        exitEdit(true);
      }
    }, 120);
  }, { once: true });
  input.addEventListener('input', function() {
    $('fxInput').value = input.textContent;
    showAutocomplete(input, r, c);
  });

  $('fxCellRef').textContent = columnLabel(c) + (r + 1);
  $('fxInput').value = rawVal;
}

/*
 * B-13 FIX: exitEdit(false) restaura fxInput al valor original.
 */
function exitEdit(save) {
  if (!STATE.editingCell) return;
  var r = STATE.editingCell.r, c = STATE.editingCell.c;
  var td = getCellEl(r, c);
  STATE.editingCell = null;
  hideAutocomplete();
  if (!td) return;

  var input   = td.querySelector('.cell-editor');
  var display = td.querySelector('.cell-display');

  if (input) {
    if (save) {
      var newVal = input.textContent;
      var oldVal = (STATE.data[r] ? STATE.data[r][c] : '') || '';
      if (oldVal !== newVal) {
        pushUndo();
        if (!STATE.data[r]) STATE.data[r] = [];
        STATE.data[r][c] = newVal;
        markDirty();
      }
      if (display) {
        var fmt        = getCellFmt(r, c);
        var displayVal = newVal;
        var isFormula  = typeof newVal === 'string' && newVal.startsWith('=');
        if (isFormula) {
          displayVal = evaluateFormula(newVal, r, c);
          display.classList.toggle('formula-error', displayVal === '#ERROR!' || displayVal === '#DIV/0!' || displayVal === '#N/A' || displayVal === '#NUM!');
        } else {
          display.classList.remove('formula-error');
        }
        display.textContent = formatCellValue(displayVal, fmt.numFormat);
        applyFmtToSpan(display, fmt);
        applyCondFmtToCell(display, r, c, STATE.data[r] ? STATE.data[r][c] : '');
      }
    } else {
      /* B-13 FIX: al cancelar, restaurar fxInput al valor guardado */
      var savedVal = (STATE.data[r] ? STATE.data[r][c] : '') || '';
      $('fxInput').value = savedVal;

      /* Restaurar el display sin cambios */
      if (display) {
        var fmtCancel = getCellFmt(r, c);
        var rawCancel = savedVal;
        var dispCancel = rawCancel;
        if (typeof rawCancel === 'string' && rawCancel.startsWith('=')) {
          dispCancel = evaluateFormula(rawCancel, r, c);
        }
        display.textContent = formatCellValue(dispCancel, fmtCancel.numFormat);
        applyFmtToSpan(display, fmtCancel);
      }
    }
    input.remove();
  }

  if (display) display.style.display = '';
  td.classList.remove('td-editing');
  renderNoteIndicator(r, c);
}

function handleEditKeydown(e, r, c) {
  var rows = STATE.data.length, cols = STATE.data[0] ? STATE.data[0].length : 0;

  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    if (!_autocompleteEl || _autocompleteEl.classList.contains('hidden')) {
      exitEdit(true);
      if (r + 1 < rows) activateCell(r + 1, c);
      return;
    }
    var selAC = _autocompleteEl.querySelector('.autocomplete-item.selected');
    if (selAC) { applyAutocomplete(selAC.dataset.value, r, c); return; }
    exitEdit(true);
    if (r + 1 < rows) activateCell(r + 1, c);
    return;
  }
  if (e.key === 'Enter' && e.shiftKey) {
    e.preventDefault(); exitEdit(true);
    if (r > 0) activateCell(r - 1, c);
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault(); exitEdit(true);
    if (!e.shiftKey) { c < cols - 1 ? activateCell(r, c + 1) : (r + 1 < rows && activateCell(r + 1, 0)); }
    else             { c > 0        ? activateCell(r, c - 1) : (r > 0 && activateCell(r - 1, cols - 1)); }
    return;
  }
  if (e.key === 'Escape') {
    exitEdit(false);
    hideAutocomplete();
    return;
  }
  if (_autocompleteEl && !_autocompleteEl.classList.contains('hidden')) {
    if (e.key === 'ArrowDown') { e.preventDefault(); moveAutocomplete(1);  return; }
    if (e.key === 'ArrowUp')   { e.preventDefault(); moveAutocomplete(-1); return; }
  }
}

/* ================================================
   AUTOCOMPLETE
   B-16 FIX: _autocompleteEl inicializado a null.
================================================ */
var _autocompleteEl = null;

function initAutocomplete() {
  _autocompleteEl = document.createElement('div');
  _autocompleteEl.className = 'autocomplete-dropdown hidden';
  document.body.appendChild(_autocompleteEl);

  _autocompleteEl.addEventListener('mousedown', function(e) { e.preventDefault(); });
  _autocompleteEl.addEventListener('click', function(e) {
    var item = e.target.closest('.autocomplete-item');
    if (item && STATE.editingCell) {
      applyAutocomplete(item.dataset.value, STATE.editingCell.r, STATE.editingCell.c);
    }
  });
}

function showAutocomplete(inputEl, r, c) {
  if (!_autocompleteEl) return;
  var val = inputEl.textContent;
  if (!val || val.startsWith('=') || val.length < 1) { hideAutocomplete(); return; }

  var colVals = new Set();
  for (var ri = 0; ri < STATE.data.length; ri++) {
    if (ri !== r) {
      var v = String((STATE.data[ri] ? STATE.data[ri][c] : '') || '').trim();
      if (v && v.toLowerCase().startsWith(val.toLowerCase()) && v !== val) colVals.add(v);
    }
  }
  if (!colVals.size) { hideAutocomplete(); return; }

  var rect = inputEl.getBoundingClientRect();
  _autocompleteEl.innerHTML = '';
  var i = 0;
  colVals.forEach(function(v) {
    if (i >= 8) return;
    var item = document.createElement('div');
    item.className    = 'autocomplete-item' + (i === 0 ? ' selected' : '');
    item.dataset.value = v;
    var matchSpan = document.createElement('span');
    matchSpan.className   = 'autocomplete-match';
    matchSpan.textContent = v.slice(0, val.length);
    item.appendChild(matchSpan);
    item.appendChild(document.createTextNode(v.slice(val.length)));
    _autocompleteEl.appendChild(item);
    i++;
  });

  _autocompleteEl.style.left     = rect.left + 'px';
  _autocompleteEl.style.top      = (rect.bottom + 2) + 'px';
  _autocompleteEl.style.minWidth = Math.max(150, rect.width) + 'px';
  _autocompleteEl.classList.remove('hidden');
}

/* B-16 FIX: guard para llamadas antes de initAutocomplete */
function hideAutocomplete() {
  if (_autocompleteEl) _autocompleteEl.classList.add('hidden');
}

function moveAutocomplete(dir) {
  if (!_autocompleteEl) return;
  var items = Array.from(_autocompleteEl.querySelectorAll('.autocomplete-item'));
  if (!items.length) return;
  var cur  = _autocompleteEl.querySelector('.autocomplete-item.selected');
  var idx  = cur ? items.indexOf(cur) : -1;
  var next = items[(idx + dir + items.length) % items.length];
  items.forEach(function(i2) { i2.classList.remove('selected'); });
  next.classList.add('selected');
  next.scrollIntoView({ block: 'nearest' });
}

function applyAutocomplete(val, r, c) {
  var td = getCellEl(r, c); if (!td) return;
  var editor = td.querySelector('.cell-editor');
  if (editor) {
    editor.textContent = val;
    $('fxInput').value = val;
    var range2 = document.createRange();
    range2.selectNodeContents(editor);
    range2.collapse(false);
    var sel2 = window.getSelection();
    sel2.removeAllRanges();
    sel2.addRange(range2);
  }
  hideAutocomplete();
}

/* ================================================
   MOUSE DRAG SELECTION

   B-57 FIX: handleCellMousedown evita conflicto cuando
   se hace clic en la celda que ya está en edición.
================================================ */
function handleCellMousedown(e, ri, ci) {
  if (e.button !== 0) return;

  /* B-57 FIX: si ya estamos editando esta celda, no interferir */
  if (STATE.editingCell && STATE.editingCell.r === ri && STATE.editingCell.c === ci) return;

  e.preventDefault();
  e.stopPropagation();

  closeAllDropdowns(null);
  hideContextMenu();

  activateCell(ri, ci, e.shiftKey, e.ctrlKey || e.metaKey);

  STATE.isDragging     = true;
  STATE.selectionStart = { r: ri, c: ci };
}

function handleCellMouseenter(e, ri, ci) {
  if (e.buttons === 0) {
    STATE.isDragging = false;
    return;
  }
  if (!STATE.isDragging || e.buttons !== 1) return;

  selectRange(STATE.selectionStart.r, STATE.selectionStart.c, ri, ci);
  STATE.activeCellR  = ri;
  STATE.activeCellC  = ci;
  STATE.selectionEnd = { r: ri, c: ci };
  syncFormatBarToCell();
}

/* ================================================
   KEYBOARD
================================================ */
document.addEventListener('keydown', function(e) {
  var tag      = document.activeElement ? document.activeElement.tagName.toLowerCase() : '';
  var isInput  = (tag === 'input' || tag === 'textarea' || (document.activeElement && document.activeElement.id === 'fxInput'));
  var isEditing = !!STATE.editingCell;

  if (e.key === 'Escape') {
    closeAllDropdowns(null);
    hideContextMenu();
    hideFindReplace();
    hideFilterDropdown();
    if (typeof dpHide === 'function') dpHide();
    return;
  }
  if ((e.ctrlKey || e.metaKey) && e.key === 's')                              { e.preventDefault(); exportXLSX(); return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'z' && !e.shiftKey && !isInput)  { e.preventDefault(); undo(); return; }
  if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.key === 'z' && e.shiftKey)) && !isInput) { e.preventDefault(); redo(); return; }
  if ((e.ctrlKey || e.metaKey) && (e.key === 'h' || e.key === 'f'))           { e.preventDefault(); showFindReplace(); return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'c' && !isEditing && !isInput)   { e.preventDefault(); copyCells(false); return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'x' && !isEditing && !isInput)   { e.preventDefault(); copyCells(true);  return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'v' && !isEditing && !isInput)   { e.preventDefault(); pasteCells();     return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'a' && !isEditing && !isInput)   { e.preventDefault(); selectAll();      return; }
  if ((e.ctrlKey || e.metaKey) && e.key === 'p' && !isEditing)               { e.preventDefault(); printPreview();   return; }

  if (!isInput) {
    if ((e.ctrlKey || e.metaKey) && e.key === 'b') { e.preventDefault(); toggleFmt('bold');      return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'i') { e.preventDefault(); toggleFmt('italic');    return; }
    if ((e.ctrlKey || e.metaKey) && e.key === 'u') { e.preventDefault(); toggleFmt('underline'); return; }
  }

  if (STATE.activeCellR < 0 || isEditing || isInput) return;

  var r    = STATE.activeCellR, c = STATE.activeCellC;
  var rows = STATE.data.length, cols = STATE.data[0] ? STATE.data[0].length : 0;

  if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); clearActiveCells(); return; }
  if (e.key === 'F2' || e.key === 'Enter')         { e.preventDefault(); enterEdit(r, c); return; }
  if (e.key === 'Tab') {
    e.preventDefault();
    var next = !e.shiftKey
      ? (c < cols - 1 ? [r, c + 1] : (r + 1 < rows ? [r + 1, 0] : null))
      : (c > 0        ? [r, c - 1] : (r > 0         ? [r - 1, cols - 1] : null));
    if (next) activateCell(next[0], next[1]);
    return;
  }

  var navArrow = function(dr, dc) {
    e.preventDefault();
    var nr = r, nc = c;
    if (e.ctrlKey) {
      if (dr !== 0) {
        var rr = r + dr;
        while (rr >= 0 && rr < rows) {
          if (STATE.data[rr] && STATE.data[rr][c]) { nr = rr; break; }
          rr += dr;
        }
        if (nr === r) nr = dr > 0 ? rows - 1 : 0;
      } else {
        var cc = c + dc;
        while (cc >= 0 && cc < cols) {
          if (STATE.data[r] && STATE.data[r][cc]) { nc = cc; break; }
          cc += dc;
        }
        if (nc === c) nc = dc > 0 ? cols - 1 : 0;
      }
    } else {
      nr = Math.max(0, Math.min(rows - 1, r + dr));
      nc = Math.max(0, Math.min(cols - 1, c + dc));
    }
    if (e.shiftKey) {
      selectRange(STATE.selectionStart.r, STATE.selectionStart.c, nr, nc);
      STATE.activeCellR = nr; STATE.activeCellC = nc;
    } else {
      activateCell(nr, nc);
    }
    syncFormatBarToCell();
  };

  if (e.key === 'ArrowDown')  { navArrow(1,  0); return; }
  if (e.key === 'ArrowUp')    { navArrow(-1, 0); return; }
  if (e.key === 'ArrowRight') { navArrow(0,  1); return; }
  if (e.key === 'ArrowLeft')  { navArrow(0, -1); return; }
  if (e.key === 'Home')       { e.preventDefault(); activateCell(r, 0);       return; }
  if (e.key === 'End')        { e.preventDefault(); activateCell(r, cols-1);  return; }
  if (e.key === 'PageDown')   { e.preventDefault(); activateCell(Math.min(rows - 1, r + 25), c); return; }
  if (e.key === 'PageUp')     { e.preventDefault(); activateCell(Math.max(0, r - 25), c);        return; }
  if (e.ctrlKey && e.key === 'Home') { e.preventDefault(); activateCell(STATE.frozen ? 1 : 0, 0); return; }
  if (e.ctrlKey && e.key === 'End')  { e.preventDefault(); activateCell(rows - 1, cols - 1);      return; }

  if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
    enterEdit(r, c);
    var editor2 = getCellEl(r, c) ? getCellEl(r, c).querySelector('.cell-editor') : null;
    if (editor2) {
      editor2.textContent = e.key;
      $('fxInput').value  = e.key;
      var range2 = document.createRange();
      range2.selectNodeContents(editor2);
      range2.collapse(false);
      var sel2 = window.getSelection();
      sel2.removeAllRanges();
      sel2.addRange(range2);
    }
  }
});

/* ================================================
   TOOLBAR FORMAT

   B-59 FIX: getTargetCells retorna [] de forma segura.
================================================ */
function getTargetCells() {
  if (STATE.selectedCells.size > 0) {
    return Array.from(STATE.selectedCells).map(function(k) { return k.split(',').map(Number); });
  }
  if (STATE.activeCellR >= 0 && STATE.activeCellC >= 0) {
    return [[STATE.activeCellR, STATE.activeCellC]];
  }
  return [];
}

function toggleFmt(prop) {
  var cells = getTargetCells();
  if (!cells.length) return showToast('Haga clic en una celda primero', '');
  var first  = cells[0];
  var newVal = !getCellFmt(first[0], first[1])[prop];
  var fmtObj = {};
  fmtObj[prop] = newVal;
  cells.forEach(function(rc) { setCellFmt(rc[0], rc[1], fmtObj); });
  cells.forEach(function(rc) { refreshCellFmt(rc[0], rc[1]); });
  syncFormatBarToCell();
  markDirty();
}

$('fmtBold').addEventListener('click',      function() { toggleFmt('bold'); });
$('fmtItalic').addEventListener('click',    function() { toggleFmt('italic'); });
$('fmtUnderline').addEventListener('click', function() { toggleFmt('underline'); });
$('fmtStrike').addEventListener('click',    function() { toggleFmt('strike'); });

$('fmtFont').addEventListener('change', function() {
  var fontVal = $('fmtFont').value;
  getTargetCells().forEach(function(rc) { setCellFmt(rc[0], rc[1], { font: fontVal }); refreshCellFmt(rc[0], rc[1]); });
  markDirty();
});
$('fmtSize').addEventListener('change', function() {
  var sizeVal = $('fmtSize').value;
  getTargetCells().forEach(function(rc) { setCellFmt(rc[0], rc[1], { size: sizeVal }); refreshCellFmt(rc[0], rc[1]); });
  markDirty();
});
$('fmtColorInput').addEventListener('input', function() {
  var color = $('fmtColorInput').value;
  $('fmtColorBar').style.background = color;
  getTargetCells().forEach(function(rc) { setCellFmt(rc[0], rc[1], { color: color }); refreshCellFmt(rc[0], rc[1]); });
  markDirty();
});
$('fmtBgColorInput').addEventListener('input', function() {
  var bgColor = $('fmtBgColorInput').value;
  $('fmtBgColorBar').style.background = bgColor;
  getTargetCells().forEach(function(rc) { setCellFmt(rc[0], rc[1], { bgColor: bgColor }); refreshCellFmt(rc[0], rc[1]); });
  markDirty();
});
['Left','Center','Right'].forEach(function(dir) {
  $('fmtAlign' + dir).addEventListener('click', function() {
    getTargetCells().forEach(function(rc) { setCellFmt(rc[0], rc[1], { align: dir.toLowerCase() }); refreshCellFmt(rc[0], rc[1]); });
    ['Left','Center','Right'].forEach(function(d) { $('fmtAlign' + d).classList.remove('active'); });
    $('fmtAlign' + dir).classList.add('active');
    markDirty();
  });
});
$('fmtNumFormat').addEventListener('change', function() {
  var numFormat = $('fmtNumFormat').value;
  getTargetCells().forEach(function(rc) {
    setCellFmt(rc[0], rc[1], { numFormat: numFormat });
    var span2 = getCellSpan(rc[0], rc[1]);
    if (span2) span2.textContent = formatCellValue((STATE.data[rc[0]] ? STATE.data[rc[0]][rc[1]] : '') || '', numFormat);
  });
  markDirty();
  var sel = $('fmtNumFormat');
  showToast('Formato: ' + sel.options[sel.selectedIndex].text, 'info');
});

$('ddCellStylesMenu').addEventListener('click', function(e) {
  var btn = e.target.closest('[data-cell-style]'); if (!btn) return;
  var cells = getTargetCells();
  if (!cells.length) { closeAllDropdowns(null); return showToast('Haga clic en una celda primero', ''); }
  var style = btn.dataset.cellStyle;
  if (style === 'clear-style') {
    cells.forEach(function(rc) {
      setCellFmt(rc[0], rc[1], { cellStyle: 'normal' });
      var tdS = getCellEl(rc[0], rc[1]);
      if (tdS) {
        ['cell-style-good','cell-style-bad','cell-style-neutral','cell-style-header','cell-style-total','cell-style-warning']
          .forEach(function(s) { tdS.classList.remove(s); });
      }
    });
    showToast('Estilo eliminado', 'info');
  } else {
    cells.forEach(function(rc) { setCellFmt(rc[0], rc[1], { cellStyle: style }); refreshCellFmt(rc[0], rc[1]); });
    showToast('Estilo aplicado: ' + style, 'success');
  }
  markDirty(); closeAllDropdowns(null);
});

$('fmtClearAll').addEventListener('click', function() {
  var cells = getTargetCells();
  if (!cells.length) return showToast('Haga clic en una celda primero', '');
  cells.forEach(function(rc) {
    var k = STATE.activeSheet + ':' + rc[0] + ',' + rc[1];
    FMT[k]     = {};
    BORDERS[k] = { top: null, bottom: null, left: null, right: null };
    var span3 = getCellSpan(rc[0], rc[1]);
    if (span3) {
      span3.removeAttribute('style');
      var td2 = span3.closest('td');
      if (td2) {
        td2.removeAttribute('style');
        ['cell-style-good','cell-style-bad','cell-style-neutral','cell-style-header','cell-style-total','cell-style-warning']
          .forEach(function(s) { td2.classList.remove(s); });
      }
    }
  });
  $('fmtFont').value     = '';
  $('fmtSize').value     = '';
  $('fmtNumFormat').value = 'general';
  ['fmtBold','fmtItalic','fmtUnderline','fmtStrike'].forEach(function(id) { $(id).classList.remove('active'); });
  markDirty();
  showToast('Formato borrado', 'info');
});

/* Borders toolbar */
if ($('borderColorInput')) {
  $('borderColorInput').addEventListener('input', function(e) { STATE.borderColor = e.target.value; });
}
document.querySelectorAll('.border-style-option').forEach(function(btn) {
  btn.addEventListener('click', function() {
    document.querySelectorAll('.border-style-option').forEach(function(b) { b.classList.remove('active'); });
    btn.classList.add('active');
    STATE.borderStyle = btn.dataset.borderStyle;
  });
});
var borderGrid = document.querySelector('#ddBordersMenu .border-grid');
if (borderGrid) {
  borderGrid.addEventListener('click', function(e) {
    var btn = e.target.closest('.border-grid-btn'); if (!btn) return;
    var cells = getTargetCells();
    if (!cells.length) { closeAllDropdowns(null); return showToast('Selecciona celdas primero', ''); }
    applyBorderType(cells, btn.dataset.border);
    markDirty(); closeAllDropdowns(null);
  });
}

/* Cond format panel listeners */
$('condFmtOp').addEventListener('change', function() {
  var op = $('condFmtOp').value;
  $('condFmtVal2Row').style.display = (op === 'between') ? '' : 'none';
  $('condFmtValRow').style.display  = (op === 'empty' || op === 'notempty') ? 'none' : '';
});
$('btnAddCondRule').addEventListener('click', function() {
  COND_RULES.push({
    op:   $('condFmtOp').value,
    val1: $('condFmtVal1').value,
    val2: $('condFmtVal2').value,
    bg:   $('condFmtBg').value,
    fg:   $('condFmtFg').value,
    bold: $('condFmtBold').checked
  });
  renderCondRuleList();
  showToast('Regla de formato condicional agregada', 'success');
});
$('btnApplyCondRules').addEventListener('click', function() {
  applyAllCondRules();
  showToast('Reglas aplicadas a toda la tabla', 'success');
  closeAllDropdowns(null);
});
$('btnClearCondRules').addEventListener('click', function() {
  COND_RULES = [];
  renderCondRuleList();
  /* Restaurar estilos de celda desde FMT */
  var tbody3 = $('tableBody');
  if (tbody3) {
    tbody3.querySelectorAll('.cell-display').forEach(function(span4) {
      var td3 = span4.closest('td'); if (!td3) return;
      var ri3  = parseInt(td3.getAttribute('data-row'));
      var ci3  = parseInt(td3.getAttribute('data-col'));
      if (isNaN(ri3) || isNaN(ci3) || ri3 < 0 || ci3 < 0) return;
      var fmt2 = getCellFmt(ri3, ci3);
      span4.style.backgroundColor = fmt2.bgColor || '';
      span4.style.color            = fmt2.color   || '';
      span4.style.fontWeight       = fmt2.bold    ? '700' : '';
    });
  }
  showToast('Reglas condicionales borradas', 'info');
  closeAllDropdowns(null);
});

/* fx bar
   B-58 FIX: fxInput.input no modifica editor si no hay celda activa.
*/
$('fxInput').addEventListener('keydown', function(e) {
  if (e.key === 'Enter') {
    e.preventDefault();
    var r2 = STATE.activeCellR, c2 = STATE.activeCellC;
    if (r2 < 0 || c2 < 0) return;
    var newVal2 = $('fxInput').value;
    var oldVal2 = (STATE.data[r2] ? STATE.data[r2][c2] : '') || '';
    if (oldVal2 !== newVal2) {
      pushUndo();
      if (!STATE.data[r2]) STATE.data[r2] = [];
      STATE.data[r2][c2] = newVal2;
      markDirty();
    }
    var span5 = getCellSpan(r2, c2);
    if (span5) {
      var fmt3       = getCellFmt(r2, c2);
      var dv         = newVal2;
      var isFormula5 = typeof newVal2 === 'string' && newVal2.startsWith('=');
      if (isFormula5) dv = evaluateFormula(newVal2, r2, c2);
      span5.textContent = formatCellValue(dv, fmt3.numFormat);
      applyFmtToSpan(span5, fmt3);
      applyCondFmtToCell(span5, r2, c2, STATE.data[r2][c2]);
    }
    showToast('Valor actualizado', 'info');
  }
  if (e.key === 'Escape') {
    var rEsc = STATE.activeCellR, cEsc = STATE.activeCellC;
    $('fxInput').value = (rEsc >= 0 && cEsc >= 0 && STATE.data[rEsc]) ? (STATE.data[rEsc][cEsc] || '') : '';
  }
});

$('fxInput').addEventListener('input', function() {
  var r3 = STATE.activeCellR, c3 = STATE.activeCellC;
  if (r3 < 0 || c3 < 0) return; /* B-58 FIX */
  if (STATE.editingCell && STATE.editingCell.r === r3 && STATE.editingCell.c === c3) {
    var editor3 = getCellEl(r3, c3) ? getCellEl(r3, c3).querySelector('.cell-editor') : null;
    if (editor3) editor3.textContent = $('fxInput').value;
  }
});

$('fxCellRef').addEventListener('click', function() {
  if (STATE.activeCellR >= 0) $('fxInput').focus();
});
