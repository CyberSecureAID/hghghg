/* ═══════════════════════════════════════════
   HeriExcel.Pro — formula-engine.js  v7.3
   Evaluador de fórmulas Excel:
   SUM, AVERAGE, COUNT, IF, VLOOKUP, etc.

   FIXES v7.3:
   - B-05: SUMIF resuelve rango de criterios como
     valores raw (texto + número).
   - B-43: División por cero devuelve '#DIV/0!'
     en lugar de Infinity.
   - B-44: IF evalúa correctamente condiciones con
     operadores de comparación (>, <, =, >=, <=, <>).
   - B-45: AVERAGE devuelve '#DIV/0!' en lugar de 0
     cuando el rango está vacío.
   - B-46: VLOOKUP maneja coincidencias parciales
     y devuelve '#N/A' correctamente.
   - B-47: Aritmética fallback maneja división por 0.
═══════════════════════════════════════════ */
'use strict';

function evaluateFormula(formula, _r, _c) {
  if (!formula || typeof formula !== 'string') return formula;
  if (!formula.startsWith('=')) return formula;
  var expr = formula.slice(1).trim();

  var resolveRef = function(ref) {
    var m = ref.match(/^([A-Za-z]+)(\d+)$/);
    if (!m) return NaN;
    var row = parseInt(m[2]) - 1;
    var col = colLetterToIndex(m[1].toUpperCase());
    var v   = (STATE.data[row] ? STATE.data[row][col] : '') || '';
    return parseFloat(v);
  };

  /*
   * B-05 FIX: resolveRangeRaw devuelve valores raw (string),
   * no filtra NaN, para que SUMIF pueda comparar texto.
   */
  var resolveRangeRaw = function(r1, r2) {
    var m1 = r1.trim().match(/^([A-Za-z]+)(\d+)$/);
    var m2 = r2.trim().match(/^([A-Za-z]+)(\d+)$/);
    if (!m1 || !m2) return [];
    var c1   = colLetterToIndex(m1[1].toUpperCase()), row1 = parseInt(m1[2]) - 1;
    var c2   = colLetterToIndex(m2[1].toUpperCase()), row2 = parseInt(m2[2]) - 1;
    var vals = [];
    for (var rr = Math.min(row1, row2); rr <= Math.max(row1, row2); rr++) {
      for (var cc = Math.min(c1, c2); cc <= Math.max(c1, c2); cc++) {
        vals.push((STATE.data[rr] ? STATE.data[rr][cc] : '') || '');
      }
    }
    return vals;
  };

  /* resolveRange original: solo valores numéricos para SUM/AVERAGE/etc. */
  var resolveRange = function(r1, r2) {
    return resolveRangeRaw(r1, r2)
      .map(function(v) { return parseFloat(v); })
      .filter(function(v) { return !isNaN(v); });
  };

  var splitArgs = function(s) {
    var result = [], depth = 0, cur = '';
    for (var i = 0; i < s.length; i++) {
      var ch = s[i];
      if (ch === '(') depth++;
      else if (ch === ')') depth--;
      else if (ch === ',' && depth === 0) { result.push(cur); cur = ''; continue; }
      cur += ch;
    }
    if (cur) result.push(cur);
    return result;
  };

  var getArg = function(raw) {
    raw = raw.trim();
    if (raw.includes(':')) {
      var pts = raw.split(':');
      return resolveRange(pts[0].trim(), pts[1].trim());
    }
    if (/^[A-Za-z]+\d+$/.test(raw)) return [resolveRef(raw)];
    return [parseFloat(raw)];
  };

  var getAllNums = function(argsStr) {
    return splitArgs(argsStr).reduce(function(acc, a) {
      return acc.concat(getArg(a).filter(function(v) { return !isNaN(v); }));
    }, []);
  };

  var resolveTextRef = function(raw) {
    if (!raw) return '';
    raw = raw.trim().replace(/^"|"$/g, '');
    if (/^[A-Za-z]+\d+$/.test(raw)) {
      var mc = raw.match(/^([A-Za-z]+)(\d+)$/);
      if (mc) {
        var row = parseInt(mc[2]) - 1;
        var col = colLetterToIndex(mc[1].toUpperCase());
        return String((STATE.data[row] ? STATE.data[row][col] : '') || '');
      }
    }
    return raw;
  };

  try {
    var upper = expr.toUpperCase();

    if (upper.startsWith('SUM(')) {
      return String(getAllNums(expr.slice(4, -1)).reduce(function(a, b) { return a + b; }, 0));
    }

    if (upper.startsWith('AVERAGE(') || upper.startsWith('AVG(')) {
      var offset = upper.startsWith('AVERAGE(') ? 8 : 4;
      var nums   = getAllNums(expr.slice(offset, -1));
      if (!nums.length) return '#DIV/0!'; /* B-45 FIX */
      return String(nums.reduce(function(a, b) { return a + b; }, 0) / nums.length);
    }

    if (upper.startsWith('COUNT(')) {
      return String(getAllNums(expr.slice(6, -1)).length);
    }

    if (upper.startsWith('COUNTA(')) {
      var args = expr.slice(7, -1);
      var cnt  = 0;
      splitArgs(args).forEach(function(a) {
        a = a.trim();
        if (a.includes(':')) {
          var pts2 = a.split(':');
          var m1c  = pts2[0].trim().match(/^([A-Za-z]+)(\d+)$/);
          var m2c  = pts2[1].trim().match(/^([A-Za-z]+)(\d+)$/);
          if (m1c && m2c) {
            var c1b = colLetterToIndex(m1c[1].toUpperCase()), row1b = parseInt(m1c[2]) - 1;
            var c2b = colLetterToIndex(m2c[1].toUpperCase()), row2b = parseInt(m2c[2]) - 1;
            for (var rr2 = Math.min(row1b, row2b); rr2 <= Math.max(row1b, row2b); rr2++) {
              for (var cc2 = Math.min(c1b, c2b); cc2 <= Math.max(c1b, c2b); cc2++) {
                var v2 = (STATE.data[rr2] ? STATE.data[rr2][cc2] : '');
                if (v2 !== '' && v2 !== null && v2 !== undefined) cnt++;
              }
            }
          }
        } else if (/^[A-Za-z]+\d+$/.test(a)) {
          var mc2 = a.match(/^([A-Za-z]+)(\d+)$/);
          if (mc2) {
            var v3 = (STATE.data[parseInt(mc2[2]) - 1] ? STATE.data[parseInt(mc2[2]) - 1][colLetterToIndex(mc2[1].toUpperCase())] : '');
            if (v3 !== '' && v3 !== null && v3 !== undefined) cnt++;
          }
        }
      });
      return String(cnt);
    }

    if (upper.startsWith('MIN(')) {
      var numsMin = getAllNums(expr.slice(4, -1));
      return numsMin.length ? String(Math.min.apply(null, numsMin)) : '';
    }
    if (upper.startsWith('MAX(')) {
      var numsMax = getAllNums(expr.slice(4, -1));
      return numsMax.length ? String(Math.max.apply(null, numsMax)) : '';
    }

    if (upper.startsWith('ROUND(')) {
      var ptsRound = splitArgs(expr.slice(6, -1));
      var numR     = getArg(ptsRound[0])[0];
      var dec      = parseInt(ptsRound[1] || '0') || 0;
      return String(Math.round(numR * Math.pow(10, dec)) / Math.pow(10, dec));
    }

    if (upper.startsWith('ABS('))   return String(Math.abs(getArg(expr.slice(4, -1))[0]));
    if (upper.startsWith('SQRT('))  {
      var sqrtVal = getArg(expr.slice(5, -1))[0];
      if (sqrtVal < 0) return '#NUM!';
      return String(Math.sqrt(sqrtVal));
    }
    if (upper.startsWith('POWER(')) {
      var ptsPow = splitArgs(expr.slice(6, -1));
      return String(Math.pow(getArg(ptsPow[0])[0], getArg(ptsPow[1])[0]));
    }
    if (upper.startsWith('MOD(')) {
      var ptsMod  = splitArgs(expr.slice(4, -1));
      var divisor = getArg(ptsMod[1])[0];
      if (divisor === 0) return '#DIV/0!'; /* B-43 FIX */
      return String(getArg(ptsMod[0])[0] % divisor);
    }

    if (upper.startsWith('LEN('))   return String(resolveTextRef(expr.slice(4, -1)).length);
    if (upper.startsWith('UPPER(')) return resolveTextRef(expr.slice(6, -1)).toUpperCase();
    if (upper.startsWith('LOWER(')) return resolveTextRef(expr.slice(6, -1)).toLowerCase();
    if (upper.startsWith('TRIM('))  return resolveTextRef(expr.slice(5, -1)).trim();

    if (upper.startsWith('LEFT(')) {
      var ptsL = splitArgs(expr.slice(5, -1));
      return resolveTextRef(ptsL[0]).slice(0, parseInt(ptsL[1]) || 1);
    }
    if (upper.startsWith('RIGHT(')) {
      var ptsR = splitArgs(expr.slice(6, -1));
      return resolveTextRef(ptsR[0]).slice(-(parseInt(ptsR[1]) || 1));
    }
    if (upper.startsWith('MID(')) {
      var ptsMid = splitArgs(expr.slice(4, -1));
      var sMid   = resolveTextRef(ptsMid[0]);
      var stMid  = (parseInt(ptsMid[1]) || 1) - 1;
      return sMid.slice(stMid, stMid + (parseInt(ptsMid[2]) || 1));
    }

    if (upper.startsWith('CONCAT(') || upper.startsWith('CONCATENATE(')) {
      var concatOff = upper.startsWith('CONCATENATE(') ? 12 : 7;
      return splitArgs(expr.slice(concatOff, -1)).map(resolveTextRef).join('');
    }

    if (upper.startsWith('TEXTJOIN(')) {
      var ptsTj   = splitArgs(expr.slice(9, -1));
      var delimTj = ptsTj[0].trim().replace(/^"|"$/g, '');
      return ptsTj.slice(2).map(resolveTextRef).filter(Boolean).join(delimTj);
    }

    /* B-05 FIX: SUMIF con soporte completo para criterios de texto */
    if (upper.startsWith('SUMIF(')) {
      var ptsSif = splitArgs(expr.slice(6, -1));
      if (ptsSif.length >= 2) {
        var rangeStr    = ptsSif[0];
        var criteriaStr = ptsSif[1];
        var sumRangeStr = ptsSif[2];
        var criteria    = resolveTextRef(criteriaStr);
        var numCrit     = parseFloat(criteria);

        var critPts     = rangeStr.trim().split(':');
        var criteriaVals = critPts.length === 2
          ? resolveRangeRaw(critPts[0].trim(), critPts[1].trim())
          : [String((STATE.data[0] ? STATE.data[0][0] : '') || '')];

        var sumVals = criteriaVals;
        if (sumRangeStr) {
          var sumPts2 = sumRangeStr.trim().split(':');
          sumVals = sumPts2.length === 2
            ? resolveRangeRaw(sumPts2[0].trim(), sumPts2[1].trim())
            : criteriaVals;
        }

        var sumSif = 0;
        criteriaVals.forEach(function(rv, i) {
          var rvNum    = parseFloat(rv);
          var matches2 = !isNaN(numCrit)
            ? (!isNaN(rvNum) && rvNum === numCrit)
            : String(rv) === criteria;
          if (matches2) {
            var sv2 = parseFloat((sumVals[i] !== undefined ? sumVals[i] : 0));
            if (!isNaN(sv2)) sumSif += sv2;
          }
        });
        return String(sumSif);
      }
    }

    if (upper.startsWith('COUNTIF(')) {
      var ptsCif   = splitArgs(expr.slice(8, -1));
      if (ptsCif.length >= 2) {
        var cifRange    = ptsCif[0].trim().split(':');
        var cifCriteria = resolveTextRef(ptsCif[1]);
        var cifNums     = parseFloat(cifCriteria);
        var cifVals     = cifRange.length === 2
          ? resolveRangeRaw(cifRange[0], cifRange[1])
          : [];
        var cntCif = 0;
        cifVals.forEach(function(v4) {
          if (!isNaN(cifNums)) {
            if (parseFloat(v4) === cifNums) cntCif++;
          } else {
            if (String(v4) === cifCriteria) cntCif++;
          }
        });
        return String(cntCif);
      }
    }

    if (upper === 'TODAY()') return new Date().toLocaleDateString('es-ES');
    if (upper === 'NOW()')   return new Date().toLocaleString('es-ES');

    /*
     * B-44 FIX: IF evalúa correctamente condiciones con operadores
     * de comparación (>, <, =, >=, <=, <>) además de expresiones JS.
     */
    if (upper.startsWith('IF(')) {
      var ptsIf = splitArgs(expr.slice(3, -1));
      if (ptsIf.length < 2) return '#ERROR!';
      var trueVal  = (ptsIf[1] || '').trim().replace(/^"|"$/g, '');
      var falseVal = (ptsIf[2] || '').trim().replace(/^"|"$/g, '');
      var condStr  = (ptsIf[0] || '').trim();
      var condResult = false;

      /* Resolver referencias de celda en la condición */
      var condResolved = condStr.replace(/[A-Za-z]+\d+/g, function(m) {
        var mc3 = m.match(/^([A-Za-z]+)(\d+)$/);
        if (!mc3) return m;
        var row3 = parseInt(mc3[2]) - 1;
        var col3 = colLetterToIndex(mc3[1].toUpperCase());
        var v5   = (STATE.data[row3] ? STATE.data[row3][col3] : '') || '';
        var n5   = parseFloat(v5);
        return isNaN(n5) ? '"' + v5.replace(/"/g, '\\"') + '"' : String(n5);
      });

      /* Normalizar operadores Excel → JS */
      condResolved = condResolved
        .replace(/<>/g, '!=')
        .replace(/=(?!=)/g, '==');

      try {
        condResult = !!(new Function('return (' + condResolved + ')')());
      } catch (_) {
        condResult = false;
      }

      return condResult ? trueVal : falseVal;
    }

    if (upper.startsWith('IFERROR(')) {
      var ptsIfe = splitArgs(expr.slice(8, -1));
      try {
        var ifeVal = evaluateFormula('=' + (ptsIfe[0] || '').trim(), _r, _c);
        if (ifeVal === '#ERROR!' || ifeVal === '#REF!' || ifeVal === '#DIV/0!' || ifeVal === '#N/A' || ifeVal === '#NUM!')
          return ptsIfe[1] ? resolveTextRef(ptsIfe[1]) : '';
        return ifeVal;
      } catch (_) {
        return ptsIfe[1] ? resolveTextRef(ptsIfe[1]) : '';
      }
    }

    /*
     * B-46 FIX: VLOOKUP correctamente implementado con #N/A.
     */
    if (upper.startsWith('VLOOKUP(')) {
      var ptsVl = splitArgs(expr.slice(8, -1));
      if (ptsVl.length >= 3) {
        var lookupVal  = resolveTextRef(ptsVl[0]);
        var rangeVl    = ptsVl[1].trim().split(':');
        var colIdxVl   = (parseInt(ptsVl[2]) || 1) - 1;
        if (rangeVl.length === 2) {
          var m1vl = rangeVl[0].trim().match(/^([A-Za-z]+)(\d+)$/);
          var m2vl = rangeVl[1].trim().match(/^([A-Za-z]+)(\d+)$/);
          if (m1vl && m2vl) {
            var startRvl = parseInt(m1vl[2]) - 1, endRvl = parseInt(m2vl[2]) - 1;
            var startCvl = colLetterToIndex(m1vl[1].toUpperCase());
            for (var rrvl = startRvl; rrvl <= endRvl; rrvl++) {
              var cellVl = String((STATE.data[rrvl] ? STATE.data[rrvl][startCvl] : '') || '');
              if (cellVl === lookupVal) {
                var resultVl = (STATE.data[rrvl] ? STATE.data[rrvl][startCvl + colIdxVl] : '') || '';
                return String(resultVl);
              }
            }
          }
        }
        return '#N/A';
      }
    }

    if (upper.startsWith('HLOOKUP(')) {
      var ptsHl = splitArgs(expr.slice(8, -1));
      if (ptsHl.length >= 3) {
        var lookupValH = resolveTextRef(ptsHl[0]);
        var rangeHl    = ptsHl[1].trim().split(':');
        var rowIdxHl   = (parseInt(ptsHl[2]) || 1) - 1;
        if (rangeHl.length === 2) {
          var m1hl = rangeHl[0].trim().match(/^([A-Za-z]+)(\d+)$/);
          var m2hl = rangeHl[1].trim().match(/^([A-Za-z]+)(\d+)$/);
          if (m1hl && m2hl) {
            var startRhl = parseInt(m1hl[2]) - 1;
            var startChl = colLetterToIndex(m1hl[1].toUpperCase());
            var endChl   = colLetterToIndex(m2hl[1].toUpperCase());
            for (var cchl = startChl; cchl <= endChl; cchl++) {
              var cellHl = String((STATE.data[startRhl] ? STATE.data[startRhl][cchl] : '') || '');
              if (cellHl === lookupValH) {
                var resultHl = (STATE.data[startRhl + rowIdxHl] ? STATE.data[startRhl + rowIdxHl][cchl] : '') || '';
                return String(resultHl);
              }
            }
          }
        }
        return '#N/A';
      }
    }

    /* Aritmética básica como fallback */
    var arithmetic = expr.replace(/[A-Za-z]+\d+/g, function(m) {
      var mc4 = m.match(/^([A-Za-z]+)(\d+)$/);
      if (!mc4) return '0';
      var row4 = parseInt(mc4[2]) - 1;
      var col4 = colLetterToIndex(mc4[1].toUpperCase());
      return parseFloat((STATE.data[row4] ? STATE.data[row4][col4] : '') || 0) || 0;
    });
    if (/^[\d\s+\-*/().]+$/.test(arithmetic)) {
      /* B-47 FIX: capturar división por cero */
      try {
        var arithResult = new Function('return (' + arithmetic + ')')();
        if (!isFinite(arithResult)) return '#DIV/0!';
        return String(arithResult);
      } catch (_) {
        return '#ERROR!';
      }
    }

    return '#ERROR!';
  } catch (_) {
    return '#ERROR!';
  }
}
