/* ═══════════════════════════════════════════
   HeriExcel.Pro — format-engine.js  v7.3
   Motor de formato de celda: texto, bordes
   y formato condicional

   FIXES v7.3:
   - B-09: applyBorderType 'inside' detecta vecinos
     en la selección y solo aplica bordes internos.
   - B-48: getCellBorders ahora garantiza que todas
     las claves (top/bottom/left/right) existen antes
     de iterar, evitando errores con objetos parciales.
   - B-49: applyFmtToSpan comprueba que `span` y `td`
     no sean null antes de operar.
   - B-50: applyBordersToTd usa Object.assign para
     garantizar estructura completa antes de iterar.
   - B-51: matchesRule maneja correctamente el caso
     empty/notempty cuando val es 0 (falsy).
   - B-52: renderCondRuleList usa escapeHtml en val2
     también.
═══════════════════════════════════════════ */
'use strict';

/* ── Claves ── */
function fmtKey(r, c)    { return STATE.activeSheet + ':' + r + ',' + c; }
function borderKey(r, c) { return STATE.activeSheet + ':' + r + ',' + c; }

/* ── Acceso al formato ── */
function getCellFmt(r, c) {
  var k = fmtKey(r, c);
  if (!FMT[k]) FMT[k] = {};
  return FMT[k];
}
function setCellFmt(r, c, props) {
  var k = fmtKey(r, c);
  if (!FMT[k]) FMT[k] = {};
  Object.assign(FMT[k], props);
}

/*
 * B-48 FIX: getCellBorders garantiza que todas las claves
 * (top, bottom, left, right) existen en el objeto devuelto.
 */
function getCellBorders(r, c) {
  var k = borderKey(r, c);
  if (!BORDERS[k]) {
    BORDERS[k] = { top: null, bottom: null, left: null, right: null };
  } else {
    /* Asegurar que todas las claves existen aunque el objeto sea parcial */
    if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'top'))    BORDERS[k].top    = null;
    if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'bottom')) BORDERS[k].bottom = null;
    if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'left'))   BORDERS[k].left   = null;
    if (!Object.prototype.hasOwnProperty.call(BORDERS[k], 'right'))  BORDERS[k].right  = null;
  }
  return BORDERS[k];
}

/*
 * B-49 FIX: Aplicar formato a <span class="cell-display">.
 * Comprueba null antes de operar.
 */
function applyFmtToSpan(span, fmt) {
  if (!span || !fmt) return;
  span.style.fontWeight     = fmt.bold      ? '700'    : '';
  span.style.fontStyle      = fmt.italic    ? 'italic' : '';
  var decorations = [];
  if (fmt.underline) decorations.push('underline');
  if (fmt.strike)    decorations.push('line-through');
  span.style.textDecoration  = decorations.join(' ') || '';
  span.style.fontFamily      = fmt.font    || '';
  span.style.fontSize        = fmt.size    || '';
  span.style.color           = fmt.color   || '';
  span.style.backgroundColor = fmt.bgColor || '';
  span.style.textAlign       = fmt.align   || '';

  var td = span.closest('td');
  if (td) {
    var ALL_STYLES = [
      'cell-style-good', 'cell-style-bad', 'cell-style-neutral',
      'cell-style-header', 'cell-style-total', 'cell-style-warning',
    ];
    ALL_STYLES.forEach(function(s) { td.classList.remove(s); });
    if (fmt.cellStyle && fmt.cellStyle !== 'normal') {
      td.classList.add('cell-style-' + fmt.cellStyle);
    }
  }
}

/*
 * B-50 FIX: Aplicar bordes a <td>.
 * Garantiza estructura completa antes de iterar.
 */
function applyBordersToTd(td, r, c) {
  if (!td) return;
  var b = getCellBorders(r, c); /* getCellBorders ya garantiza todas las claves */
  ['top', 'bottom', 'left', 'right'].forEach(function(side) {
    var bdef = b[side];
    var propName = 'border' + cap(side);
    td.style[propName] = bdef
      ? ((bdef.width || '1px') + ' ' + (bdef.style || 'solid') + ' ' + (bdef.color || '#7c8cf8'))
      : '';
  });
}

/* Refresca display y bordes de una celda concreta */
function refreshCellFmt(r, c) {
  var span = getCellSpan(r, c);
  if (span) applyFmtToSpan(span, getCellFmt(r, c));
  var td = getCellEl(r, c);
  if (td)   applyBordersToTd(td, r, c);
}

/* ════════════════════════════════════════
   FORMATO CONDICIONAL
════════════════════════════════════════ */

function applyCondFmtToCell(span, r, c, val) {
  if (!span) return;
  var fmt = getCellFmt(r, c);
  span.style.backgroundColor = fmt.bgColor || '';
  span.style.color           = fmt.color   || '';
  span.style.fontWeight      = fmt.bold    ? '700' : '';

  for (var i = 0; i < COND_RULES.length; i++) {
    var rule = COND_RULES[i];
    if (matchesRule(val, rule)) {
      span.style.backgroundColor = rule.bg;
      span.style.color           = rule.fg;
      if (rule.bold) span.style.fontWeight = '700';
      break;
    }
  }
}

/*
 * B-51 FIX: matchesRule maneja correctamente empty/notempty
 * cuando val es 0 (falsy pero no vacío).
 */
function matchesRule(val, rule) {
  var num = parseFloat(val);
  var v1  = parseFloat(rule.val1);
  var v2  = parseFloat(rule.val2);
  switch (rule.op) {
    case 'gt':       return !isNaN(num) && !isNaN(v1) && num > v1;
    case 'lt':       return !isNaN(num) && !isNaN(v1) && num < v1;
    case 'eq':       return String(val) === String(rule.val1);
    case 'gte':      return !isNaN(num) && !isNaN(v1) && num >= v1;
    case 'lte':      return !isNaN(num) && !isNaN(v1) && num <= v1;
    case 'neq':      return String(val) !== String(rule.val1);
    case 'between':  return !isNaN(num) && !isNaN(v1) && !isNaN(v2) && num >= v1 && num <= v2;
    case 'contains': return String(val).toLowerCase().indexOf(String(rule.val1).toLowerCase()) !== -1;
    /* B-51: empty → '' o null o undefined (no 0) */
    case 'empty':    return val === '' || val === null || val === undefined;
    case 'notempty': return val !== '' && val !== null && val !== undefined;
    default:         return false;
  }
}

function applyAllCondRules() {
  var tbody = $('tableBody'); if (!tbody) return;
  tbody.querySelectorAll('tr[data-row]').forEach(function(tr) {
    var ri = parseInt(tr.dataset.row);
    tr.querySelectorAll('td:not(.row-num)').forEach(function(td) {
      var ci   = parseInt(td.getAttribute('data-col'));
      if (isNaN(ci) || ci < 0) return;
      var span = td.querySelector('.cell-display'); if (!span) return;
      var rawV = (STATE.data[ri] ? STATE.data[ri][ci] : '') || '';
      applyCondFmtToCell(span, ri, ci, rawV);
    });
  });
}

/* B-52 FIX: escapeHtml en ambos valores de la regla */
function renderCondRuleList() {
  var list  = $('condRuleList');
  var count = $('condRuleCount');
  if (!list) return;
  list.innerHTML = '';
  if (count) count.textContent = COND_RULES.length ? '(' + COND_RULES.length + ')' : '';

  var opLabels = {
    gt: '>', lt: '<', eq: '=', gte: '≥', lte: '≤', neq: '≠',
    between: 'entre', contains: 'contiene', empty: 'vacía', notempty: 'no vacía',
  };

  COND_RULES.forEach(function(rule, i) {
    var div = document.createElement('div');
    div.className = 'cond-rule-item';
    var label   = opLabels[rule.op] || rule.op;
    var val2str = rule.val2 ? ' y ' + escapeHtml(String(rule.val2)) : '';
    div.innerHTML =
      '<span class="cond-rule-preview" style="background:' + rule.bg + ';color:' + rule.fg + ';font-weight:' + (rule.bold ? '700' : '400') + '">Aa</span>' +
      '<span class="cond-rule-text">'  + escapeHtml(label) + ' ' + escapeHtml(String(rule.val1 || '')) + val2str + '</span>' +
      '<button class="cond-rule-del" data-idx="' + i + '" title="Eliminar">\u00D7</button>';
    list.appendChild(div);
  });

  list.querySelectorAll('.cond-rule-del').forEach(function(btn) {
    btn.addEventListener('click', function() {
      COND_RULES.splice(parseInt(btn.dataset.idx), 1);
      renderCondRuleList();
    });
  });
}

/*
 * Aplica tipo de borde a un conjunto de celdas.
 * B-09 FIX: 'inside' detecta vecinos en la selección.
 */
function applyBorderType(cells, type) {
  var color    = STATE.borderColor || '#7c8cf8';
  var bStyle   = STATE.borderStyle || 'solid';
  var bdef     = { color: color, style: bStyle, width: '1px' };
  var thickDef = { color: color, style: bStyle, width: '3px' };

  var cellSet = new Set(cells.map(function(rc) { return rc[0] + ',' + rc[1]; }));
  var inSel   = function(r, c) { return cellSet.has(r + ',' + c); };

  var setBorderSide = function(r, c, side, def) {
    if (r < 0 || c < 0 || r >= STATE.data.length || c >= (STATE.data[0] ? STATE.data[0].length : 0)) return;
    var b2 = getCellBorders(r, c);
    b2[side] = def ? Object.assign({}, def) : null;
    var td2 = getCellEl(r, c); if (td2) applyBordersToTd(td2, r, c);
  };

  cells.forEach(function(rc) {
    var r = rc[0], c = rc[1];
    var b = getCellBorders(r, c);
    switch (type) {
      case 'all':
        b.top = b.bottom = b.left = b.right = Object.assign({}, bdef);
        setBorderSide(r - 1, c, 'bottom', bdef);
        setBorderSide(r + 1, c, 'top',    bdef);
        setBorderSide(r, c - 1, 'right',  bdef);
        setBorderSide(r, c + 1, 'left',   bdef);
        break;

      case 'outside':
        b.top = b.bottom = b.left = b.right = Object.assign({}, bdef);
        break;

      /* B-09 FIX: 'inside' sólo pone bordes donde hay vecino en selección */
      case 'inside':
        if (inSel(r - 1, c)) b.top    = Object.assign({}, bdef);
        if (inSel(r + 1, c)) b.bottom = Object.assign({}, bdef);
        if (inSel(r, c - 1)) b.left   = Object.assign({}, bdef);
        if (inSel(r, c + 1)) b.right  = Object.assign({}, bdef);
        break;

      case 'none':
      case 'clear-borders':
        b.top = b.bottom = b.left = b.right = null;
        break;

      case 'top':
        b.top = Object.assign({}, bdef);
        setBorderSide(r - 1, c, 'bottom', bdef);
        break;

      case 'bottom':
        b.bottom = Object.assign({}, bdef);
        setBorderSide(r + 1, c, 'top', bdef);
        break;

      case 'left':
        b.left = Object.assign({}, bdef);
        setBorderSide(r, c - 1, 'right', bdef);
        break;

      case 'right':
        b.right = Object.assign({}, bdef);
        setBorderSide(r, c + 1, 'left', bdef);
        break;

      case 'thick-outside':
        b.top = b.bottom = b.left = b.right = Object.assign({}, thickDef);
        break;

      case 'double-bottom':
        b.bottom = Object.assign({}, bdef, { style: 'double', width: '3px' });
        setBorderSide(r + 1, c, 'top', bdef);
        break;

      case 'dashed-all':
        b.top = b.bottom = b.left = b.right = Object.assign({}, bdef, { style: 'dashed' });
        break;
    }

    var td = getCellEl(r, c);
    if (td) applyBordersToTd(td, r, c);
  });

  showToast('Borde "' + type + '" aplicado', 'success');
}
