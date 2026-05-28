/* ═══════════════════════════════════════════
   HeriExcel.Pro — datepicker-eyedropper.js  v1.0
   Módulo 2 & 3:
     · Date Picker con calendario emergente (sin deps)
     · EyeDropper API + mejora del color picker nativo
   
   Llama a initDatePicker() y initEyeDropper()
   desde DOMContentLoaded (en app.js).
═══════════════════════════════════════════ */
'use strict';

/* ══════════════════════════════════════════
   ██████  DATE PICKER
   Calendario emergente propio, sin dependencias.
   Se activa cuando el usuario hace doble clic
   en una celda cuyo valor luce como fecha,
   o cuando el numFormat de la celda es "date".
══════════════════════════════════════════ */

const DP = {
  el:        null,   // contenedor del calendario
  year:      0,
  month:     0,      // 0-11
  targetR:   -1,
  targetC:   -1,
  onSelect:  null,
};

/** Inicializa el Date Picker (llamar una sola vez desde DOMContentLoaded) */
function initDatePicker() {
  // Crear el elemento del calendario en el DOM
  const el = document.createElement('div');
  el.id = 'hxDatePicker';
  el.className = 'hx-dp hidden';
  el.setAttribute('role', 'dialog');
  el.setAttribute('aria-label', 'Selector de fecha');
  el.innerHTML = `
    <div class="hx-dp-header">
      <button class="hx-dp-nav" id="hxDpPrevYear"  title="Año anterior">«</button>
      <button class="hx-dp-nav" id="hxDpPrevMonth" title="Mes anterior">‹</button>
      <span   class="hx-dp-title" id="hxDpTitle"></span>
      <button class="hx-dp-nav" id="hxDpNextMonth" title="Mes siguiente">›</button>
      <button class="hx-dp-nav" id="hxDpNextYear"  title="Año siguiente">»</button>
    </div>
    <div class="hx-dp-weekdays">
      <span>Lu</span><span>Ma</span><span>Mi</span>
      <span>Ju</span><span>Vi</span><span>Sa</span><span>Do</span>
    </div>
    <div class="hx-dp-grid" id="hxDpGrid"></div>
    <div class="hx-dp-footer">
      <button class="hx-dp-today" id="hxDpToday">Hoy</button>
      <button class="hx-dp-clear" id="hxDpClear">Limpiar</button>
    </div>
  `;
  document.body.appendChild(el);
  DP.el = el;

  // Navegación
  document.getElementById('hxDpPrevYear').addEventListener('click',  () => { DP.year--;  dpRender(); });
  document.getElementById('hxDpNextYear').addEventListener('click',  () => { DP.year++;  dpRender(); });
  document.getElementById('hxDpPrevMonth').addEventListener('click', () => { dpShiftMonth(-1); });
  document.getElementById('hxDpNextMonth').addEventListener('click', () => { dpShiftMonth(+1); });
  document.getElementById('hxDpToday').addEventListener('click', () => {
    const d = new Date();
    dpSelectDate(d.getFullYear(), d.getMonth(), d.getDate());
  });
  document.getElementById('hxDpClear').addEventListener('click', () => {
    if (DP.onSelect) DP.onSelect('');
    dpHide();
  });

  // Cierre al click fuera
  document.addEventListener('mousedown', (e) => {
    if (!DP.el.classList.contains('hidden') && !DP.el.contains(e.target)) {
      dpHide();
    }
  }, true);

  // Tecla Escape cierra
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !DP.el.classList.contains('hidden')) dpHide();
  });
}

function dpShiftMonth(delta) {
  DP.month += delta;
  if (DP.month < 0)  { DP.month = 11; DP.year--; }
  if (DP.month > 11) { DP.month = 0;  DP.year++; }
  dpRender();
}

const DP_MONTHS_ES = [
  'Enero','Febrero','Marzo','Abril','Mayo','Junio',
  'Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'
];

function dpRender() {
  document.getElementById('hxDpTitle').textContent =
    `${DP_MONTHS_ES[DP.month]} ${DP.year}`;

  const grid = document.getElementById('hxDpGrid');
  grid.innerHTML = '';

  // Qué día de la semana es el 1 del mes (0=Dom → ajustamos a Lun=0)
  const firstDay = new Date(DP.year, DP.month, 1).getDay(); // 0=Dom
  const startOffset = (firstDay === 0) ? 6 : firstDay - 1;  // Lun como inicio

  const daysInMonth   = new Date(DP.year, DP.month + 1, 0).getDate();
  const daysInPrevMon = new Date(DP.year, DP.month, 0).getDate();

  const today = new Date();
  const todayY = today.getFullYear(), todayM = today.getMonth(), todayD = today.getDate();

  // Valor actual de la celda para resaltarla
  let selY = -1, selM = -1, selD = -1;
  if (DP.targetR >= 0 && DP.targetC >= 0) {
    const raw = STATE.data[DP.targetR]?.[DP.targetC] || '';
    const d   = parseDisplayDate(raw);
    if (d) { selY = d.getFullYear(); selM = d.getMonth(); selD = d.getDate(); }
  }

  let total = 0;

  // Días del mes anterior (relleno)
  for (let i = startOffset - 1; i >= 0; i--) {
    const btn = dpDayBtn(daysInPrevMon - i, true, false, false);
    grid.appendChild(btn); total++;
  }

  // Días del mes actual
  for (let d = 1; d <= daysInMonth; d++) {
    const isToday    = (DP.year === todayY && DP.month === todayM && d === todayD);
    const isSelected = (DP.year === selY   && DP.month === selM   && d === selD);
    const btn = dpDayBtn(d, false, isToday, isSelected);
    btn.addEventListener('click', () => dpSelectDate(DP.year, DP.month, d));
    grid.appendChild(btn); total++;
  }

  // Días del mes siguiente (relleno para completar cuadrícula de 6 filas)
  let next = 1;
  while (total % 7 !== 0) {
    const btn = dpDayBtn(next++, true, false, false);
    grid.appendChild(btn); total++;
  }
}

function dpDayBtn(day, muted, isToday, isSelected) {
  const btn = document.createElement('button');
  btn.type = 'button';
  btn.className = 'hx-dp-day';
  btn.textContent = day;
  if (muted)      btn.classList.add('muted');
  if (isToday)    btn.classList.add('today');
  if (isSelected) btn.classList.add('selected');
  return btn;
}

function dpSelectDate(y, m, d) {
  const str = `${String(d).padStart(2,'0')}/${String(m+1).padStart(2,'0')}/${y}`;
  if (DP.onSelect) DP.onSelect(str);
  dpHide();
}

/**
 * Muestra el Date Picker anclado a una celda.
 * @param {number} r  fila
 * @param {number} c  columna
 * @param {DOMRect} rect  getBoundingClientRect() de la celda
 * @param {function} onSelect  callback(dateString)
 */
function dpShow(r, c, rect, onSelect) {
  DP.targetR = r;
  DP.targetC = c;
  DP.onSelect = onSelect;

  // Fecha inicial: si la celda ya tiene fecha, abrir en ese mes
  const raw = STATE.data[r]?.[c] || '';
  const existing = parseDisplayDate(raw);
  const now = new Date();
  DP.year  = existing ? existing.getFullYear() : now.getFullYear();
  DP.month = existing ? existing.getMonth()    : now.getMonth();

  dpRender();

  // Posicionar
  const dp = DP.el;
  dp.classList.remove('hidden');

  // Intentar debajo de la celda; si no cabe, arriba
  let top  = rect.bottom + 4;
  let left = rect.left;
  const dpH = 290, dpW = 252;

  if (top + dpH > window.innerHeight - 8)  top  = Math.max(4, rect.top - dpH - 4);
  if (left + dpW > window.innerWidth  - 8)  left = Math.max(4, rect.right - dpW);

  dp.style.top  = top  + 'px';
  dp.style.left = left + 'px';
}

function dpHide() {
  if (DP.el) DP.el.classList.add('hidden');
  DP.targetR = -1; DP.targetC = -1; DP.onSelect = null;
}

/**
 * Punto de entrada: abre el picker si la celda es de tipo fecha.
 * Llamar desde el dblclick de celda (en app-core.js / enterEdit).
 */
function maybShowDatePicker(r, c) {
  const fmt = getCellFmt(r, c);
  const raw = (STATE.data[r]?.[c] || '');
  const isDateFmt  = fmt.numFormat === 'date';
  const looksDate  = looksLikeDate(String(raw));

  if (!isDateFmt && !looksDate) return false;  // no es fecha → no abrir

  const td   = getCellEl(r, c);
  if (!td) return false;
  const rect = td.getBoundingClientRect();

  dpShow(r, c, rect, (dateStr) => {
    // Guardar el valor en la celda
    if (STATE.editingCell) exitEdit(false);
    pushUndo();
    if (!STATE.data[r]) STATE.data[r] = [];
    STATE.data[r][c] = dateStr;
    markDirty();

    const span = getCellSpan(r, c);
    if (span) {
      span.textContent = dateStr
        ? formatCellValue(dateStr, getCellFmt(r, c).numFormat || 'date')
        : '';
      span.classList.remove('formula-error');
      applyFmtToSpan(span, getCellFmt(r, c));
    }
    $('fxInput').value = dateStr;
    syncFormatBarToCell();
    showToast(dateStr ? `Fecha: ${dateStr}` : 'Fecha limpiada', 'success');
  });

  return true;  // se abrió el picker, no iniciar edición normal
}


/* ══════════════════════════════════════════
   ██████  EYEDROPPER + COLOR PICKER MEJORADO
   Añade botón "cuentagotas" junto a cada
   input[type=color] de la barra de formato.
   Usa la EyeDropper API (Chrome 95+, Edge 95+).
   En navegadores sin soporte, el botón se oculta.
══════════════════════════════════════════ */

/** Inicializa EyeDropper para los color pickers de la barra de formato */
function initEyeDropper() {
  const supported = ('EyeDropper' in window);

  // Pares: [inputId, colorBarId]
  const COLOR_PAIRS = [
    { inputId: 'fmtColorInput',   barId: 'fmtColorBar',   prop: 'color'   },
    { inputId: 'fmtBgColorInput', barId: 'fmtBgColorBar', prop: 'bgColor' },
  ];

  COLOR_PAIRS.forEach(({ inputId, barId, prop }) => {
    const input = document.getElementById(inputId);
    if (!input) return;

    const wrapper = input.closest('.fmt-color-btn');
    if (!wrapper) return;

    if (!supported) return;  // sin soporte → no añadir botón

    // Crear botón eyedropper
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'hx-eyedrop-btn';
    btn.title = 'Capturar color de la pantalla';
    btn.innerHTML = `<svg viewBox="0 0 20 20" fill="currentColor" width="11" height="11">
      <path d="M13.586 1.586a2 2 0 012.828 2.828l-1.172 1.172
               a1 1 0 000 1.414l.586.586a1 1 0 010 1.414L13.414 11
               a1 1 0 01-1.414 0l-.586-.586a1 1 0 00-1.414 0
               L4.293 16.121A1 1 0 013 15.414V13a1 1 0 01.293-.707
               l5.707-5.707a1 1 0 000-1.414l-.586-.586
               a1 1 0 010-1.414l2.414-2.414a1 1 0 011.414 0
               l.586.586a1 1 0 001.414 0l1.344-1.344z"/>
    </svg>`;

    // Insertar después del wrapper en el formato bar
    wrapper.insertAdjacentElement('afterend', btn);

    btn.addEventListener('click', async () => {
      try {
        const dropper = new EyeDropper();
        const result  = await dropper.open();
        const hex     = result.sRGBHex;

        input.value = hex;
        const bar = document.getElementById(barId);
        if (bar) bar.style.background = hex;

        // Aplicar a las celdas seleccionadas
        const cells = getTargetCells();
        if (cells.length) {
          const fmtProp = {};
          fmtProp[prop] = hex;
          cells.forEach(([r, c]) => {
            setCellFmt(r, c, fmtProp);
            refreshCellFmt(r, c);
          });
          markDirty();
          showToast('Color capturado: ' + hex, 'success');
        }
      } catch (err) {
        // El usuario canceló la selección o hubo un error
        if (err.name !== 'AbortError') {
          showToast('No se pudo capturar el color', 'error');
        }
      }
    });
  });
}


/* ══════════════════════════════════════════
   CSS DEL DATE PICKER (inyectado inline)
   Se inyecta aquí para mantener el módulo
   autocontenido; no requiere editar CSS.
══════════════════════════════════════════ */
(function injectDatePickerStyles() {
  const style = document.createElement('style');
  style.id = 'hx-dp-styles';
  style.textContent = `
/* ── HeriExcel Date Picker ── */
.hx-dp {
  position: fixed;
  z-index: 199999;
  background: var(--bg-dropdown, #252638);
  border: 1px solid var(--border, #353654);
  border-radius: var(--radius-lg, 12px);
  box-shadow: var(--shadow-lg, 0 12px 40px rgba(0,0,0,.7));
  width: 252px;
  user-select: none;
  animation: dropIn .14s ease forwards;
  overflow: hidden;
}
.hx-dp.hidden { display: none; }

.hx-dp-header {
  display: flex;
  align-items: center;
  gap: 2px;
  padding: 10px 10px 6px;
  background: var(--bg-panel, #1f2030);
  border-bottom: 1px solid var(--border, #353654);
}
.hx-dp-nav {
  background: transparent;
  border: 1px solid transparent;
  border-radius: var(--radius, 5px);
  color: var(--text-muted, #6b7099);
  cursor: pointer;
  font-size: 14px;
  line-height: 1;
  padding: 3px 6px;
  transition: all .13s ease;
}
.hx-dp-nav:hover {
  background: var(--bg-hover, #2a2c40);
  color: var(--text-primary, #e8eaf6);
  border-color: var(--border-light, #4a4c6a);
}
.hx-dp-title {
  flex: 1;
  text-align: center;
  font-family: var(--font-ui);
  font-size: 12.5px;
  font-weight: 700;
  color: var(--text-primary, #e8eaf6);
  letter-spacing: .01em;
}

.hx-dp-weekdays {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  padding: 6px 8px 2px;
  gap: 2px;
}
.hx-dp-weekdays span {
  text-align: center;
  font-family: var(--font-ui);
  font-size: 10px;
  font-weight: 700;
  letter-spacing: .06em;
  text-transform: uppercase;
  color: var(--text-muted, #6b7099);
  padding: 2px 0;
}

.hx-dp-grid {
  display: grid;
  grid-template-columns: repeat(7, 1fr);
  gap: 2px;
  padding: 2px 8px 8px;
}
.hx-dp-day {
  aspect-ratio: 1;
  background: transparent;
  border: 1px solid transparent;
  border-radius: var(--radius, 5px);
  color: var(--text-primary, #e8eaf6);
  cursor: pointer;
  font-family: var(--font-mono);
  font-size: 11.5px;
  line-height: 1;
  padding: 0;
  transition: all .1s ease;
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 28px;
}
.hx-dp-day:hover:not(.muted) {
  background: var(--accent-dim, rgba(124,140,248,.14));
  border-color: var(--accent-border, rgba(124,140,248,.4));
  color: var(--accent, #7c8cf8);
}
.hx-dp-day.muted {
  color: var(--text-muted, #6b7099);
  opacity: .45;
  cursor: default;
  pointer-events: none;
}
.hx-dp-day.today {
  border-color: var(--accent-border, rgba(124,140,248,.4));
  color: var(--accent, #7c8cf8);
  font-weight: 700;
}
.hx-dp-day.selected {
  background: var(--accent, #7c8cf8);
  border-color: var(--accent, #7c8cf8);
  color: var(--text-on-accent, #fff);
  font-weight: 700;
}
.hx-dp-day.selected:hover {
  background: var(--accent-hover, #6679f0);
}

.hx-dp-footer {
  display: flex;
  gap: 6px;
  padding: 6px 8px 10px;
  border-top: 1px solid var(--border, #353654);
  background: var(--bg-panel, #1f2030);
}
.hx-dp-today, .hx-dp-clear {
  flex: 1;
  height: 26px;
  border-radius: var(--radius, 5px);
  border: 1px solid var(--border, #353654);
  font-family: var(--font-ui);
  font-size: 11.5px;
  font-weight: 500;
  cursor: pointer;
  transition: all .13s ease;
}
.hx-dp-today {
  background: var(--accent-dim, rgba(124,140,248,.14));
  color: var(--accent, #7c8cf8);
  border-color: var(--accent-border, rgba(124,140,248,.4));
}
.hx-dp-today:hover { background: var(--accent, #7c8cf8); color: #fff; }
.hx-dp-clear {
  background: transparent;
  color: var(--text-muted, #6b7099);
}
.hx-dp-clear:hover {
  background: var(--danger-dim, rgba(248,113,113,.12));
  color: var(--danger, #f87171);
  border-color: var(--danger-border, rgba(248,113,113,.35));
}

/* ── EyeDropper button ── */
.hx-eyedrop-btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  width: 22px;
  height: 22px;
  border-radius: var(--radius, 5px);
  border: 1px solid var(--border, #353654);
  background: var(--bg-panel, #252638);
  color: var(--text-muted, #6b7099);
  cursor: pointer;
  flex-shrink: 0;
  transition: all .13s ease;
  margin-left: 2px;
  padding: 0;
}
.hx-eyedrop-btn:hover {
  background: var(--accent-dim, rgba(124,140,248,.14));
  border-color: var(--accent-border, rgba(124,140,248,.4));
  color: var(--accent, #7c8cf8);
}
`;
  document.head.appendChild(style);
})();
