/* ═══════════════════════════════════════════
   HeriExcel.Pro — utils.js  v7.3
   Funciones de utilidad: DOM, formato de
   valores, fechas, secuencias de relleno

   FIXES v7.3:
   - B-02: formatCellValue 'percent' no divide por 100.
   - B-04: parseDisplayDate valida con regex antes
     de usar el constructor Date genérico.
   - B-40: looksLikeDate requiere 4 dígitos en el año
     para evitar falsos positivos con "1/2/3".
   - B-41: cap() exportada correctamente para uso
     en format-engine.js (mismo scope global).
   - B-42: formatCellValue maneja null/undefined
     explícitamente en todos los casos.
═══════════════════════════════════════════ */
'use strict';

/* Shortcut getElementById */
const $ = function(id) { return document.getElementById(id); };

/* A → 0, B → 1, AA → 26, … */
function columnLabel(index) {
  var label = '', i = index;
  do {
    label = String.fromCharCode(65 + (i % 26)) + label;
    i = Math.floor(i / 26) - 1;
  } while (i >= 0);
  return label;
}

/* "A" → 0, "B" → 1, "AA" → 26, … */
function colLetterToIndex(letters) {
  var index = 0;
  for (var i = 0; i < letters.length; i++) {
    index = index * 26 + (letters.charCodeAt(i) - 64);
  }
  return index - 1;
}

/* Capitaliza la primera letra */
function cap(s) { return s.charAt(0).toUpperCase() + s.slice(1); }

/* Escapa caracteres HTML peligrosos */
function escapeHtml(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/* Convierte cualquier color CSS a #rrggbb */
function toHex(color) {
  if (!color) return '#ffffff';
  if (/^#[0-9a-f]{6}$/i.test(color)) return color;
  var m = color.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
  if (m) return '#' + [m[1], m[2], m[3]]
    .map(function(n) { return parseInt(n).toString(16).padStart(2, '0'); }).join('');
  return '#ffffff';
}

/*
 * Aplica el formato numérico a un valor crudo.
 * FIX B-02: 'percent' muestra el número directamente como %
 *   sin dividir por 100 (25 → "25.00 %", no "0.25 %").
 * FIX B-42: devuelve '' cuando val es null/undefined.
 */
function formatCellValue(val, numFormat) {
  if (val === null || val === undefined) return '';
  if (!numFormat || numFormat === 'general' || numFormat === 'text') return val;
  var num = parseFloat(String(val).replace(/[^0-9.\-]/g, ''));
  if (isNaN(num)) return val;
  switch (numFormat) {
    case 'number':
      return num.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 'integer':
      return Math.round(num).toLocaleString('es-ES');
    case 'currency':
      return num.toLocaleString('es-ES', { style: 'currency', currency: 'USD', minimumFractionDigits: 2 });
    case 'percent':
      /* B-02 FIX: mostramos num% directamente, sin dividir por 100. */
      return num.toLocaleString('es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' %';
    case 'scientific':
      return num.toExponential(2).toUpperCase();
    case 'date':
      try {
        var d = new Date(val);
        if (!isNaN(d)) return d.toLocaleDateString('es-ES');
      } catch (_) {}
      return val;
    default:
      return val;
  }
}

/*
 * Detecta si un string parece una fecha.
 * FIX B-40: requiere año de 4 dígitos para evitar
 *   falsos positivos con strings como "1/2/3".
 */
function looksLikeDate(val) {
  if (!val || typeof val !== 'string') return false;
  var s = val.trim();
  if (!s) return false;
  var datePatterns = [
    /^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}$/,
    /^\d{4}[\/\-\.]\d{1,2}[\/\-\.]\d{1,2}$/,
    /^\d{1,2}\s+(ene|feb|mar|abr|may|jun|jul|ago|sep|oct|nov|dic)\s+\d{4}$/i,
    /^\d{1,2}\s+(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s+\d{4}$/i,
  ];
  return datePatterns.some(function(p) { return p.test(s); });
}

/* Normaliza una fecha a DD/MM/YYYY */
function normalizeDateDisplay(val) {
  if (!val) return val;
  var s = String(val).trim();
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    var day   = String(d.getDate()).padStart(2, '0');
    var month = String(d.getMonth() + 1).padStart(2, '0');
    var year  = d.getFullYear();
    return day + '/' + month + '/' + year;
  }
  return s;
}

/*
 * Parsea un string DD/MM/YYYY o YYYY-MM-DD a objeto Date.
 * Retorna null si no puede parsear.
 *
 * FIX B-04: valida con regex explícitamente antes de usar el
 * constructor Date genérico, evitando que "100" sea interpretado
 * como 100 milisegundos desde el epoch de Unix.
 */
function parseDisplayDate(str) {
  if (!str) return null;
  str = String(str).trim();

  /* DD/MM/YYYY o DD-MM-YYYY o DD.MM.YYYY */
  var m = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (m) {
    var d = new Date(+m[3], +m[2] - 1, +m[1]);
    return isNaN(d.getTime()) ? null : d;
  }

  /* YYYY-MM-DD o YYYY/MM/DD */
  m = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (m) {
    var d2 = new Date(+m[1], +m[2] - 1, +m[3]);
    return isNaN(d2.getTime()) ? null : d2;
  }

  /* DD Mon YYYY  (ej: "15 ene 2024") */
  m = str.match(/^(\d{1,2})\s+([a-záéíóú]+)\s+(\d{4})$/i);
  if (m) {
    var MONTH_MAP = {
      ene:1,feb:2,mar:3,abr:4,may:5,jun:6,jul:7,ago:8,sep:9,oct:10,nov:11,dic:12,
      jan:1,apr:4,aug:8,dec:12,
    };
    var mo = MONTH_MAP[m[2].toLowerCase().slice(0,3)];
    if (mo) {
      var d3 = new Date(+m[3], mo - 1, +m[1]);
      return isNaN(d3.getTime()) ? null : d3;
    }
  }

  /*
   * Sólo intentar el constructor genérico si parece una fecha ISO
   * (contiene letras o al menos 2 guiones/slashes con año de 4 dígitos).
   */
  if (/[a-zA-Z]/.test(str) || (/\d{4}/.test(str) && (str.match(/[\/\-]/g) || []).length >= 2)) {
    var d4 = new Date(str);
    if (!isNaN(d4.getTime())) return d4;
  }

  return null;
}

/* Detecta una secuencia numérica o de lista en un array de valores */
function detectSequence(values) {
  if (!values || values.length < 2) return null;
  var nums = values.map(function(v) { return parseFloat(v); });
  if (nums.every(function(n) { return !isNaN(n); })) {
    var step = nums[1] - nums[0];
    if (nums.every(function(n, i) { return i === 0 || Math.abs(n - nums[i - 1] - step) < 0.0001; })) {
      return { type: 'numeric', step: step, last: nums[nums.length - 1] };
    }
  }
  var LISTS = [
    ['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'],
    ['january','february','march','april','may','june','july','august','september','october','november','december'],
    ['lunes','martes','miércoles','jueves','viernes','sábado','domingo'],
    ['monday','tuesday','wednesday','thursday','friday','saturday','sunday'],
  ];
  for (var li = 0; li < LISTS.length; li++) {
    var list    = LISTS[li];
    var indices = values.map(function(v) { return list.indexOf(String(v).toLowerCase()); });
    if (indices.every(function(i) { return i >= 0; })) {
      return { type: 'list', list: list, lastIdx: indices[indices.length - 1] };
    }
  }
  return null;
}

/* Devuelve el siguiente valor de una secuencia detectada */
function getNextSequenceValue(seq, offset) {
  if (!seq) return '';
  if (seq.type === 'numeric') {
    return String(Math.round((seq.last + seq.step * offset) * 1e10) / 1e10);
  }
  if (seq.type === 'list') {
    var idx = (seq.lastIdx + offset) % seq.list.length;
    var v   = seq.list[idx < 0 ? idx + seq.list.length : idx];
    return v.charAt(0).toUpperCase() + v.slice(1);
  }
  return '';
}
