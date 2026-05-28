/* ===============================================
   HeriExcel.Pro — templates.js  v4.0
   Selector y generador de plantillas.
   25+ plantillas profesionales.

   CAMBIOS v4.0:
   - Pantalla de inicio: las 3 tarjetas usan
     exactamente el mismo wrapper tornasol y la
     misma estructura (hx-mode-card-wrap /
     hx-mode-card-inner / hx-mode-icon / hx-mode-title
     / hx-mode-sub / hx-mode-cta). Mismo tamaño fijo.
   - Grid de plantillas: grid-auto-rows fijo (230px).
     Todas las tarjetas idénticas en altura.
     Eliminado tpl-featured / grid-column span 2.
   - Estructura de tarjeta: preview (72px fijo) +
     body (flex-grow) + footer (44px fijo).

   Expone: initTemplateSelector()
=============================================== */
'use strict';

/* ─────────────────────────────────────────────
   COLORES POR CATEGORÍA
───────────────────────────────────────────── */
const CAT_META = {
  general:   { label: 'General',   color: '#4ade80' },
  finanzas:  { label: 'Finanzas',  color: '#7c8cf8' },
  almacen:   { label: 'Almacén',   color: '#fbbf24' },
  proyectos: { label: 'Proyectos', color: '#c084fc' },
  analisis:  { label: 'Análisis',  color: '#f87171' },
  rrhh:      { label: 'RRHH',      color: '#22d3ee' },
  marketing: { label: 'Marketing', color: '#fb923c' },
  legal:     { label: 'Legal',     color: '#a78bfa' },
};

/* ─────────────────────────────────────────────
   25+ PLANTILLAS
───────────────────────────────────────────── */
const TEMPLATES = [

  /* ════ GENERAL ════ */
  {
    id: 'basic', name: 'Hoja básica',
    desc: 'Tabla limpia con encabezados de ejemplo, lista para usar de inmediato.',
    icon: '📋', cat: 'general', fileName: 'hoja-basica.xlsx',
    sheets: [
      {
        name: 'Datos',
        widths: { 0: 180, 1: 200, 2: 140, 3: 130 },
        data: [
          ['Nombre', 'Descripción', 'Valor', 'Estado'],
          ['Elemento 1', 'Descripción del elemento 1', '100', 'Activo'],
          ['Elemento 2', 'Descripción del elemento 2', '250', 'Inactivo'],
          ['Elemento 3', 'Descripción del elemento 3', '180', 'Activo'],
          ['','','',''], ['','','',''], ['','','',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:2, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
        ],
      },
    ],
  },

  {
    id: 'todo', name: 'Lista de tareas',
    desc: 'Organiza pendientes con prioridades, fechas límite y estado de avance.',
    icon: '✅', cat: 'general', fileName: 'lista-tareas.xlsx',
    sheets: [
      {
        name: 'Tareas',
        widths: { 0:40, 1:260, 2:100, 3:110, 4:90, 5:100 },
        data: [
          ['#', 'Tarea', 'Prioridad', 'Fecha límite', 'Avance %', 'Estado'],
          ['1','Definir alcance del proyecto','Alta','15/01/2025','100','Completado'],
          ['2','Diseñar interfaz de usuario','Alta','22/01/2025','75','En progreso'],
          ['3','Desarrollar backend API','Media','05/02/2025','30','En progreso'],
          ['4','Pruebas unitarias','Media','15/02/2025','0','Pendiente'],
          ['5','Documentación técnica','Baja','20/02/2025','0','Pendiente'],
          ['6','Despliegue en producción','Alta','28/02/2025','0','Pendiente'],
          ['','','','','',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:0,c:2, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:4, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:0,c:5, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:1,c:2, fmt:{color:'#f87171',bold:true} },
          { r:2,c:2, fmt:{color:'#f87171',bold:true} },
          { r:3,c:2, fmt:{color:'#fbbf24',bold:true} },
          { r:1,c:5, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:5, fmt:{color:'#fbbf24',bold:true} },
          { r:3,c:5, fmt:{color:'#fbbf24',bold:true} },
        ],
      },
    ],
  },

  {
    id: 'contact', name: 'Directorio de contactos',
    desc: 'Lista de contactos con teléfono, email, empresa y notas adicionales.',
    icon: '📇', cat: 'general', fileName: 'directorio-contactos.xlsx',
    sheets: [
      {
        name: 'Contactos',
        widths: { 0:160, 1:160, 2:130, 3:180, 4:120, 5:150 },
        data: [
          ['Nombre', 'Apellido', 'Teléfono', 'Email', 'Empresa', 'Notas'],
          ['Ana','Martínez','+1 555 001 0001','ana@empresa.com','TechCorp S.A.','Cliente VIP'],
          ['Carlos','López','+1 555 001 0002','carlos@freelance.com','Freelancer','Diseñador web'],
          ['María','García','+1 555 001 0003','maria@startup.io','StartupXYZ','Socia estratégica'],
          ['Pedro','Ruiz','+1 555 001 0004','pedro@proveedor.net','Proveedores SA','Proveedor principal'],
          ['','','','','',''],
          ['','','','','',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80'} },
          { r:0,c:2, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80'} },
          { r:0,c:4, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80'} },
          { r:0,c:5, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80'} },
        ],
      },
    ],
  },

  /* ════ FINANZAS ════ */
  {
    id: 'finance', name: 'Reporte financiero',
    desc: 'Ingresos, gastos y saldo neto trimestral con fórmulas automáticas.',
    icon: '💰', cat: 'finanzas', fileName: 'reporte-financiero.xlsx',
    sheets: [
      {
        name: 'Ingresos',
        widths: { 0:200, 1:120, 2:120, 3:120, 4:120 },
        data: [
          ['REPORTE DE INGRESOS','','','',''],
          ['Concepto','Ene','Feb','Mar','Total'],
          ['Ventas producto A','15000','18000','21000','=SUM(B3:D3)'],
          ['Ventas producto B','8500','9200','11000','=SUM(B4:D4)'],
          ['Servicios','4000','4000','5500','=SUM(B5:D5)'],
          ['Otros ingresos','1200','900','1500','=SUM(B6:D6)'],
          ['TOTAL INGRESOS','=SUM(B3:B6)','=SUM(C3:C6)','=SUM(D3:D6)','=SUM(E3:E6)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'14px',color:'#4ade80'} },
          { r:1,c:0, fmt:{bold:true,bgColor:'#1e2a1e',color:'#4ade80',align:'center'} },
          { r:1,c:1, fmt:{bold:true,bgColor:'#1e2a1e',color:'#4ade80',align:'center'} },
          { r:1,c:2, fmt:{bold:true,bgColor:'#1e2a1e',color:'#4ade80',align:'center'} },
          { r:1,c:3, fmt:{bold:true,bgColor:'#1e2a1e',color:'#4ade80',align:'center'} },
          { r:1,c:4, fmt:{bold:true,bgColor:'#1e2a1e',color:'#4ade80',align:'center'} },
          { r:6,c:0, fmt:{bold:true,bgColor:'#0d1a0d',color:'#4ade80'} },
          { r:6,c:4, fmt:{bold:true,bgColor:'#0d1a0d',color:'#4ade80',numFormat:'currency'} },
        ],
      },
      {
        name: 'Gastos',
        widths: { 0:200, 1:120, 2:120, 3:120, 4:120 },
        data: [
          ['REPORTE DE GASTOS','','','',''],
          ['Concepto','Ene','Feb','Mar','Total'],
          ['Nómina','9000','9000','9000','=SUM(B3:D3)'],
          ['Alquiler','2500','2500','2500','=SUM(B4:D4)'],
          ['Proveedores','5200','4800','6100','=SUM(B5:D5)'],
          ['Servicios básicos','800','850','900','=SUM(B6:D6)'],
          ['Marketing','1500','2000','2500','=SUM(B7:D7)'],
          ['TOTAL GASTOS','=SUM(B3:B7)','=SUM(C3:C7)','=SUM(D3:D7)','=SUM(E3:E7)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'14px',color:'#f87171'} },
          { r:1,c:0, fmt:{bold:true,bgColor:'#2a1e1e',color:'#f87171',align:'center'} },
          { r:7,c:0, fmt:{bold:true,bgColor:'#1a0d0d',color:'#f87171'} },
          { r:7,c:4, fmt:{bold:true,bgColor:'#1a0d0d',color:'#f87171',numFormat:'currency'} },
        ],
      },
      {
        name: 'Resumen',
        widths: { 0:220, 1:150 },
        data: [
          ['RESUMEN FINANCIERO',''],
          ['Concepto','Total trimestre'],
          ['Total ingresos','28700'],
          ['Total gastos','19000'],
          ['Utilidad neta','9700'],
          ['Margen (%)','33.8'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'15px',color:'#7c8cf8'} },
          { r:1,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:1,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:2,c:1, fmt:{color:'#4ade80',numFormat:'currency'} },
          { r:3,c:1, fmt:{color:'#f87171',numFormat:'currency'} },
          { r:4,c:1, fmt:{bold:true,color:'#7c8cf8',numFormat:'currency'} },
          { r:5,c:1, fmt:{numFormat:'percent'} },
        ],
      },
    ],
  },

  {
    id: 'budget', name: 'Presupuesto mensual',
    desc: 'Control de ingresos y gastos personales con seguimiento de desviaciones.',
    icon: '💳', cat: 'finanzas', fileName: 'presupuesto-mensual.xlsx',
    sheets: [
      {
        name: 'Presupuesto',
        widths: { 0:200, 1:130, 2:130, 3:130 },
        data: [
          ['PRESUPUESTO MENSUAL','','',''],
          ['Categoría','Presupuestado','Real','Diferencia'],
          ['── INGRESOS ──','','',''],
          ['Salario principal','3500','3500','=C4-B4'],
          ['Ingresos adicionales','500','320','=C5-B5'],
          ['Total ingresos','=SUM(B4:B5)','=SUM(C4:C5)','=SUM(D4:D5)'],
          ['── GASTOS FIJOS ──','','',''],
          ['Alquiler / hipoteca','900','900','=C8-B8'],
          ['Servicios (luz, agua)','150','175','=C9-B9'],
          ['Internet y teléfono','80','80','=C10-B10'],
          ['Seguro','120','120','=C11-B11'],
          ['Total gastos fijos','=SUM(B8:B11)','=SUM(C8:C11)','=SUM(D8:D11)'],
          ['── GASTOS VARIABLES ──','','',''],
          ['Alimentación','400','450','=C14-B14'],
          ['Transporte','150','130','=C15-B15'],
          ['Entretenimiento','100','180','=C16-B16'],
          ['Ropa y personal','80','45','=C17-B17'],
          ['Total gastos variables','=SUM(B14:B17)','=SUM(C14:C17)','=SUM(D14:D17)'],
          ['BALANCE','=B6-B12-B18','=C6-C12-C18','=D6-D12-D18'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'15px',color:'#7c8cf8'} },
          { r:1,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:18,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:18,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',numFormat:'currency'} },
        ],
      },
    ],
  },

  {
    id: 'invoice', name: 'Plantilla de factura',
    desc: 'Factura profesional con subtotales, IVA y datos del cliente.',
    icon: '🧾', cat: 'finanzas', fileName: 'plantilla-factura.xlsx',
    sheets: [
      {
        name: 'Factura',
        widths: { 0:50, 1:280, 2:80, 3:110, 4:110 },
        data: [
          ['FACTURA','','','',''],
          ['','','','',''],
          ['Emisor:','Tu Empresa S.A.','','Factura N°:','F-2025-001'],
          ['Dirección:','Calle Principal 123','','Fecha:','01/01/2025'],
          ['Teléfono:','+1 (555) 000-0000','','Vencimiento:','31/01/2025'],
          ['Email:','contacto@tuempresa.com','','',''],
          ['','','','',''],
          ['Cliente:','Nombre del Cliente','','',''],
          ['Dirección:','Dirección del cliente','','',''],
          ['','','','',''],
          ['#','Descripción del producto / servicio','Cantidad','Precio unit.','Total'],
          ['1','Servicio de consultoría','10','150.00','=C12*D12'],
          ['2','Desarrollo de módulo web','1','800.00','=C13*D13'],
          ['3','Soporte técnico mensual','1','200.00','=C14*D14'],
          ['','','','',''],
          ['','','','Subtotal','=SUM(E12:E14)'],
          ['','','','IVA (16%)','=E16*0.16'],
          ['','','','TOTAL A PAGAR','=E16+E17'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'20px',color:'#7c8cf8'} },
          { r:10,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:10,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:10,c:2, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:10,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'right'} },
          { r:10,c:4, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'right'} },
          { r:17,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80',align:'right'} },
          { r:17,c:4, fmt:{bold:true,bgColor:'#1e1f2e',color:'#4ade80',numFormat:'currency'} },
        ],
      },
    ],
  },

  {
    id: 'cashflow', name: 'Flujo de caja',
    desc: 'Control de entradas y salidas de efectivo semana a semana.',
    icon: '💵', cat: 'finanzas', fileName: 'flujo-caja.xlsx',
    sheets: [
      {
        name: 'Flujo de Caja',
        widths: { 0:180, 1:110, 2:110, 3:110, 4:110, 5:110 },
        data: [
          ['FLUJO DE CAJA — ENERO 2025','','','','',''],
          ['Concepto','Sem 1','Sem 2','Sem 3','Sem 4','Total mes'],
          ['── ENTRADAS ──','','','','',''],
          ['Cobros clientes','8500','12000','9800','15000','=SUM(B4:E4)'],
          ['Anticipo proyecto','5000','0','0','0','=SUM(B5:E5)'],
          ['Otras entradas','500','300','400','200','=SUM(B6:E6)'],
          ['TOTAL ENTRADAS','=SUM(B4:B6)','=SUM(C4:C6)','=SUM(D4:D6)','=SUM(E4:E6)','=SUM(F4:F6)'],
          ['── SALIDAS ──','','','','',''],
          ['Nómina','0','9000','0','0','=SUM(B9:E9)'],
          ['Proveedores','3200','2100','4500','1800','=SUM(B10:E10)'],
          ['Alquiler','0','2500','0','0','=SUM(B11:E11)'],
          ['Otros gastos','400','600','350','700','=SUM(B12:E12)'],
          ['TOTAL SALIDAS','=SUM(B9:B12)','=SUM(C9:C12)','=SUM(D9:D12)','=SUM(E9:E12)','=SUM(F9:F12)'],
          ['SALDO NETO','=B7-B13','=C7-C13','=D7-D13','=E7-E13','=F7-F13'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'14px',color:'#4ade80'} },
          { r:1,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:13,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
        ],
      },
    ],
  },

  {
    id: 'loan', name: 'Tabla de amortización',
    desc: 'Cuotas de préstamo con capital, intereses y saldo pendiente por mes.',
    icon: '🏦', cat: 'finanzas', fileName: 'tabla-amortizacion.xlsx',
    sheets: [
      {
        name: 'Amortización',
        widths: { 0:70, 1:110, 2:120, 3:120, 4:130 },
        data: [
          ['TABLA DE AMORTIZACIÓN DE PRÉSTAMO','','','',''],
          ['','','','',''],
          ['Capital:','10000','','Tasa anual:','12%'],
          ['Plazo (meses):','12','','Cuota mensual:','888.49'],
          ['','','','',''],
          ['Cuota','Fecha','Capital','Intereses','Saldo'],
          ['1','01/02/2025','788.49','100.00','9211.51'],
          ['2','01/03/2025','796.38','92.12','8415.13'],
          ['3','01/04/2025','804.34','84.15','7610.79'],
          ['4','01/05/2025','812.38','76.11','6798.41'],
          ['5','01/06/2025','820.50','67.98','5977.91'],
          ['6','01/07/2025','828.71','59.78','5149.20'],
          ['7','01/08/2025','837.00','51.49','4312.20'],
          ['8','01/09/2025','845.37','43.12','3466.83'],
          ['9','01/10/2025','853.82','34.67','2613.01'],
          ['10','01/11/2025','862.36','26.13','1750.65'],
          ['11','01/12/2025','870.98','17.51','879.67'],
          ['12','01/01/2026','879.67','8.80','0.00'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'13px',color:'#7c8cf8'} },
          { r:5,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:5,c:1, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:5,c:2, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'right'} },
          { r:5,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'right'} },
          { r:5,c:4, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'right'} },
        ],
      },
    ],
  },

  /* ════ ALMACÉN ════ */
  {
    id: 'inventory', name: 'Inventario de productos',
    desc: 'Control de stock, costos unitarios, valor total y alertas de reabastecimiento.',
    icon: '📦', cat: 'almacen', fileName: 'inventario.xlsx',
    sheets: [
      {
        name: 'Inventario',
        widths: { 0:60, 1:220, 2:120, 3:80, 4:90, 5:110, 6:120 },
        data: [
          ['ID','Producto','Categoría','Stock','Stock mín.','Costo unit.','Valor total'],
          ['001','Laptop Dell XPS 13','Electrónica','12','5','850.00','=D2*F2'],
          ['002','Mouse inalámbrico','Accesorios','45','10','25.00','=D3*F3'],
          ['003','Teclado mecánico','Accesorios','8','10','65.00','=D4*F4'],
          ['004','Monitor 24"','Electrónica','6','3','320.00','=D5*F5'],
          ['005','Silla ergonómica','Mobiliario','3','2','280.00','=D6*F6'],
          ['006','Escritorio ajustable','Mobiliario','2','1','450.00','=D7*F7'],
          ['007','Auriculares Bluetooth','Accesorios','20','8','45.00','=D8*F8'],
          ['','','','','','TOTAL INVENTARIO','=SUM(G2:G8)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24'} },
          { r:0,c:2, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24',align:'center'} },
          { r:0,c:4, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24',align:'center'} },
          { r:0,c:5, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24',align:'right'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24',align:'right'} },
          { r:8,c:5, fmt:{bold:true,color:'#fbbf24'} },
          { r:8,c:6, fmt:{bold:true,color:'#fbbf24',numFormat:'currency'} },
        ],
      },
    ],
  },

  {
    id: 'purchase', name: 'Órdenes de compra',
    desc: 'Registro de órdenes a proveedores con estado de entrega y costos.',
    icon: '🛒', cat: 'almacen', fileName: 'ordenes-compra.xlsx',
    sheets: [
      {
        name: 'Órdenes',
        widths: { 0:90, 1:170, 2:100, 3:120, 4:110, 5:100, 6:110 },
        data: [
          ['Nro. OC','Proveedor','Fecha OC','Producto / Servicio','Cantidad','Precio','Estado'],
          ['OC-001','Proveedor Alpha','02/01/2025','Laptops Dell XPS','10','8500.00','Recibido'],
          ['OC-002','Tech Supplies','05/01/2025','Mousepads XL','50','250.00','En tránsito'],
          ['OC-003','Ergonomics Co.','10/01/2025','Sillas de oficina','20','5600.00','Pendiente'],
          ['OC-004','Proveedor Beta','12/01/2025','Teclados mecánicos','30','1950.00','Recibido'],
          ['OC-005','NetWork Corp','15/01/2025','Switches 24 puertos','5','2250.00','Pendiente'],
          ['','','','','','TOTAL','=SUM(F2:F6)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#1e1a10',color:'#fbbf24'} },
          { r:1,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:6, fmt:{color:'#fbbf24',bold:true} },
          { r:4,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:6,c:5, fmt:{bold:true,color:'#fbbf24'} },
          { r:6,c:6, fmt:{bold:true,color:'#fbbf24',numFormat:'currency'} },
        ],
      },
    ],
  },

  /* ════ PROYECTOS ════ */
  {
    id: 'schedule', name: 'Cronograma de proyecto',
    desc: 'Tareas, responsables, fechas de inicio/fin, estado y progreso.',
    icon: '📅', cat: 'proyectos', fileName: 'cronograma-proyecto.xlsx',
    sheets: [
      {
        name: 'Cronograma',
        widths: { 0:50, 1:240, 2:150, 3:110, 4:110, 5:110, 6:100 },
        data: [
          ['#','Tarea','Responsable','Inicio','Fin','Estado','Avance %'],
          ['1','Reunión de inicio','Ana Martínez','01/01/2025','01/01/2025','Completado','100'],
          ['2','Análisis de requerimientos','Carlos López','02/01/2025','08/01/2025','Completado','100'],
          ['3','Diseño de arquitectura','María García','09/01/2025','20/01/2025','En progreso','65'],
          ['4','Desarrollo módulo A','Pedro Ruiz','21/01/2025','10/02/2025','Pendiente','0'],
          ['5','Desarrollo módulo B','Laura Torres','21/01/2025','15/02/2025','Pendiente','0'],
          ['6','Pruebas de integración','Carlos López','16/02/2025','25/02/2025','Pendiente','0'],
          ['7','Capacitación al usuario','Ana Martínez','26/02/2025','05/03/2025','Pendiente','0'],
          ['8','Puesta en producción','María García','06/03/2025','10/03/2025','Pendiente','0'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc'} },
          { r:0,c:5, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc',align:'center'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc',align:'center'} },
          { r:1,c:5, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:5, fmt:{color:'#4ade80',bold:true} },
          { r:3,c:5, fmt:{color:'#fbbf24',bold:true} },
        ],
      },
    ],
  },

  {
    id: 'risk', name: 'Matriz de riesgos',
    desc: 'Identifica, clasifica y gestiona riesgos con plan de mitigación.',
    icon: '⚠️', cat: 'proyectos', fileName: 'matriz-riesgos.xlsx',
    sheets: [
      {
        name: 'Riesgos',
        widths: { 0:40, 1:220, 2:100, 3:100, 4:90, 5:150, 6:130 },
        data: [
          ['#','Riesgo','Probabilidad','Impacto','Nivel','Mitigación','Responsable'],
          ['1','Retrasos en entrega de proveedor','Alta','Alto','Crítico','Negociar SLA y alternativas','Pedro Ruiz'],
          ['2','Rotación de personal clave','Media','Alto','Alto','Plan de capacitación cruzada','Ana Martínez'],
          ['3','Cambios de alcance frecuentes','Alta','Medio','Alto','Change request formal','Carlos López'],
          ['4','Fallo de infraestructura cloud','Baja','Alto','Medio','Redundancia y backups','María García'],
          ['5','Problemas de integración API','Media','Medio','Medio','Pruebas tempranas de integración','Laura Torres'],
          ['6','Desvío presupuestario','Baja','Medio','Bajo','Monitoreo semanal de gastos','Ana Martínez'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc',align:'center'} },
          { r:1,c:4, fmt:{bold:true,color:'#f87171'} },
          { r:2,c:4, fmt:{bold:true,color:'#f87171'} },
          { r:3,c:4, fmt:{bold:true,color:'#f87171'} },
          { r:4,c:4, fmt:{bold:true,color:'#fbbf24'} },
          { r:5,c:4, fmt:{bold:true,color:'#fbbf24'} },
          { r:6,c:4, fmt:{bold:true,color:'#4ade80'} },
        ],
      },
    ],
  },

  {
    id: 'sprint', name: 'Sprint backlog (Scrum)',
    desc: 'Tablero de sprint con story points, prioridad y estado de cada historia.',
    icon: '🏃', cat: 'proyectos', fileName: 'sprint-backlog.xlsx',
    sheets: [
      {
        name: 'Sprint 1',
        widths: { 0:40, 1:280, 2:80, 3:100, 4:110, 5:130 },
        data: [
          ['ID','Historia de usuario','Puntos','Prioridad','Estado','Asignado a'],
          ['US-01','Como usuario quiero registrarme con email','5','Alta','Completado','María García'],
          ['US-02','Como usuario quiero iniciar sesión con OAuth','8','Alta','En progreso','Pedro Ruiz'],
          ['US-03','Como admin quiero gestionar usuarios','8','Alta','En progreso','Laura Torres'],
          ['US-04','Como usuario quiero ver mi dashboard','5','Media','Pendiente','Carlos López'],
          ['US-05','Como usuario quiero exportar mis datos','3','Media','Pendiente','María García'],
          ['US-06','Como admin quiero ver reportes de uso','5','Baja','Pendiente','Ana Martínez'],
          ['US-07','Como usuario quiero notificaciones push','3','Baja','Pendiente','Pedro Ruiz'],
          ['','','=SUM(C2:C8)','','',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1d1527',color:'#c084fc'} },
          { r:1,c:4, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:4, fmt:{color:'#fbbf24',bold:true} },
          { r:3,c:4, fmt:{color:'#fbbf24',bold:true} },
          { r:8,c:2, fmt:{bold:true,color:'#c084fc',bgColor:'#1d1527'} },
        ],
      },
    ],
  },

  /* ════ ANÁLISIS ════ */
  {
    id: 'analytics', name: 'Análisis de datos',
    desc: 'Tabla con estadísticas descriptivas calculadas automáticamente.',
    icon: '📊', cat: 'analisis', fileName: 'analisis-datos.xlsx',
    sheets: [
      {
        name: 'Datos',
        widths: { 0:60, 1:180, 2:100, 3:120, 4:100 },
        data: [
          ['ID','Nombre','Región','Ventas','Crecimiento %'],
          ['1','García Martínez','Norte','45200','12.5'],
          ['2','López Herrera','Sur','38500','8.2'],
          ['3','Rodríguez Silva','Este','52100','15.7'],
          ['4','Martínez Ruiz','Oeste','29800','-3.1'],
          ['5','Sánchez Torres','Norte','61400','22.4'],
          ['6','Pérez Morales','Sur','44700','10.9'],
          ['7','González Flores','Este','37200','6.3'],
          ['8','Ramírez Cruz','Oeste','55800','18.6'],
          ['','','','',''],
          ['── ESTADÍSTICAS ──','','','Ventas','Crecimiento'],
          ['Total','','','=SUM(D2:D9)','=SUM(E2:E9)'],
          ['Promedio','','','=AVERAGE(D2:D9)','=AVERAGE(E2:E9)'],
          ['Máximo','','','=MAX(D2:D9)','=MAX(E2:E9)'],
          ['Mínimo','','','=MIN(D2:D9)','=MIN(E2:E9)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1f1520',color:'#f87171',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#1f1520',color:'#f87171'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#1f1520',color:'#f87171',align:'right'} },
          { r:0,c:4, fmt:{bold:true,bgColor:'#1f1520',color:'#f87171',align:'right'} },
          { r:10,c:0, fmt:{bold:true,italic:true,color:'#f87171'} },
        ],
      },
    ],
  },

  {
    id: 'kpi', name: 'KPIs y métricas',
    desc: 'Panel de indicadores clave con metas, valores actuales y % cumplimiento.',
    icon: '📈', cat: 'analisis', fileName: 'kpis-metricas.xlsx',
    sheets: [
      {
        name: 'KPIs',
        widths: { 0:250, 1:130, 2:130, 3:140, 4:110 },
        data: [
          ['PANEL DE KPIs — 2025','','','',''],
          ['Indicador','Meta','Actual','Cumplimiento %','Estado'],
          ['Ingresos mensuales ($)','50000','48300','=C3/B3*100','=IF(C3>=B3,"Logrado","En riesgo")'],
          ['Nuevos clientes','30','27','=C4/B4*100','=IF(C4>=B4,"Logrado","En riesgo")'],
          ['Tasa de retención (%)','85','88','=C5/B5*100','=IF(C5>=B5,"Logrado","En riesgo")'],
          ['Tickets resueltos','200','215','=C6/B6*100','=IF(C6>=B6,"Logrado","En riesgo")'],
          ['NPS (satisfacción)','70','74','=C7/B7*100','=IF(C7>=B7,"Logrado","En riesgo")'],
          ['Costo de adquisición ($)','150','138','=B8/C8*100','=IF(C8<=B8,"Logrado","En riesgo")'],
          ['Tasa de conversión (%)','3.5','3.1','=C9/B9*100','=IF(C9>=B9,"Logrado","En riesgo")'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'15px',color:'#7c8cf8'} },
          { r:1,c:0, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8'} },
          { r:1,c:3, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
          { r:1,c:4, fmt:{bold:true,bgColor:'#1e1f2e',color:'#7c8cf8',align:'center'} },
        ],
      },
    ],
  },

  {
    id: 'survey', name: 'Análisis de encuesta',
    desc: 'Tabulación de respuestas con conteos y porcentajes por opción.',
    icon: '📝', cat: 'analisis', fileName: 'analisis-encuesta.xlsx',
    sheets: [
      {
        name: 'Resultados',
        widths: { 0:280, 1:100, 2:100, 3:100, 4:100, 5:100 },
        data: [
          ['ANÁLISIS DE ENCUESTA DE SATISFACCIÓN','','','','',''],
          ['Total de encuestados:','150','','','',''],
          ['','','','','',''],
          ['Pregunta: Calificación del servicio','Muy malo','Malo','Regular','Bueno','Excelente'],
          ['Respuestas (cantidad)','3','8','25','65','49'],
          ['Porcentaje (%)','=B5/150*100','=C5/150*100','=D5/150*100','=E5/150*100','=F5/150*100'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,fontSize:'14px',color:'#f87171'} },
          { r:3,c:0, fmt:{bold:true,bgColor:'#1f1520',color:'#f87171'} },
        ],
      },
    ],
  },

  /* ════ RRHH ════ */
  {
    id: 'hr', name: 'Registro de empleados',
    desc: 'Ficha de empleados con cargo, departamento, fecha de ingreso y salario.',
    icon: '👥', cat: 'rrhh', fileName: 'registro-empleados.xlsx',
    sheets: [
      {
        name: 'Empleados',
        widths: { 0:80, 1:180, 2:150, 3:130, 4:110, 5:120, 6:100 },
        data: [
          ['ID','Nombre completo','Cargo','Departamento','Fecha ingreso','Salario','Estado'],
          ['EMP001','Ana Sofía Martínez','Directora General','Dirección','15/03/2018','5500.00','Activo'],
          ['EMP002','Carlos Eduardo López','Gerente de Ventas','Ventas','01/06/2019','3800.00','Activo'],
          ['EMP003','María José García','Dev Senior','TI','20/09/2020','4200.00','Activo'],
          ['EMP004','Pedro Antonio Ruiz','Contador','Finanzas','10/01/2021','2900.00','Activo'],
          ['EMP005','Laura Valentina Torres','Diseñadora UX','TI','05/04/2021','3100.00','Activo'],
          ['EMP006','Roberto Sánchez','Analista de Datos','TI','18/11/2021','3400.00','Activo'],
          ['EMP007','Patricia Morales','Ejecutiva de Ventas','Ventas','02/03/2022','2500.00','Activo'],
          ['EMP008','Daniel Pérez','Soporte Técnico','TI','14/07/2022','2200.00','Inactivo'],
          ['','','','','','NÓMINA TOTAL','=SUM(F2:F9)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee'} },
          { r:0,c:5, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'right'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'center'} },
          { r:9,c:5, fmt:{bold:true,color:'#22d3ee',align:'right'} },
          { r:9,c:6, fmt:{bold:true,color:'#22d3ee',numFormat:'currency'} },
        ],
      },
    ],
  },

  {
    id: 'payroll', name: 'Planilla de nómina',
    desc: 'Cálculo de salarios netos con deducciones, bonos e impuestos.',
    icon: '💼', cat: 'rrhh', fileName: 'nomina.xlsx',
    sheets: [
      {
        name: 'Nómina Enero',
        widths: { 0:80, 1:180, 2:110, 3:100, 4:100, 5:100, 6:110 },
        data: [
          ['ID','Empleado','Salario bruto','Bonos','Deducciones','IRPF (15%)','Neto a pagar'],
          ['EMP001','Ana Sofía Martínez','5500','500','150','=C2*0.15','=C2+D2-E2-F2'],
          ['EMP002','Carlos Eduardo López','3800','200','120','=C3*0.15','=C3+D3-E3-F3'],
          ['EMP003','María José García','4200','300','130','=C4*0.15','=C4+D4-E4-F4'],
          ['EMP004','Pedro Antonio Ruiz','2900','0','100','=C5*0.15','=C5+D5-E5-F5'],
          ['EMP005','Laura Valentina Torres','3100','150','110','=C6*0.15','=C6+D6-E6-F6'],
          ['EMP006','Roberto Sánchez','3400','200','120','=C7*0.15','=C7+D7-E7-F7'],
          ['EMP007','Patricia Morales','2500','100','90','=C8*0.15','=C8+D8-E8-F8'],
          ['','TOTALES','=SUM(C2:C8)','=SUM(D2:D8)','=SUM(E2:E8)','=SUM(F2:F8)','=SUM(G2:G8)'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'center'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'right'} },
          { r:8,c:1, fmt:{bold:true,color:'#22d3ee'} },
          { r:8,c:6, fmt:{bold:true,color:'#4ade80',numFormat:'currency'} },
        ],
      },
    ],
  },

  {
    id: 'vacations', name: 'Control de vacaciones',
    desc: 'Seguimiento de días de vacaciones tomados, pendientes y programados.',
    icon: '🏖️', cat: 'rrhh', fileName: 'control-vacaciones.xlsx',
    sheets: [
      {
        name: 'Vacaciones 2025',
        widths: { 0:80, 1:180, 2:110, 3:100, 4:100, 5:130 },
        data: [
          ['ID','Empleado','Días asignados','Tomados','Pendientes','Fecha próximas'],
          ['EMP001','Ana Sofía Martínez','30','5','25','15/07/2025'],
          ['EMP002','Carlos Eduardo López','25','10','15','01/08/2025'],
          ['EMP003','María José García','25','0','25',''],
          ['EMP004','Pedro Antonio Ruiz','20','15','5','10/06/2025'],
          ['EMP005','Laura Valentina Torres','25','8','17','20/09/2025'],
          ['EMP006','Roberto Sánchez','25','3','22',''],
          ['EMP007','Patricia Morales','20','0','20',''],
          ['','PROMEDIO','=AVERAGE(C2:C8)','=AVERAGE(D2:D8)','=AVERAGE(E2:E8)',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee'} },
          { r:8,c:1, fmt:{bold:true,color:'#22d3ee'} },
        ],
      },
    ],
  },

  {
    id: 'org', name: 'Organigrama / Directorio',
    desc: 'Estructura organizacional con áreas, cargos, responsables y contactos.',
    icon: '🏢', cat: 'rrhh', fileName: 'organigrama.xlsx',
    sheets: [
      {
        name: 'Organización',
        widths: { 0:160, 1:180, 2:130, 3:180, 4:80 },
        data: [
          ['Área / Departamento','Responsable','Cargo','Email corporativo','Extensión'],
          ['Dirección General','Ana Sofía Martínez','CEO','a.martinez@empresa.com','100'],
          ['Finanzas y Contabilidad','Pedro Antonio Ruiz','CFO','p.ruiz@empresa.com','101'],
          ['Tecnología e Innovación','María José García','CTO','m.garcia@empresa.com','102'],
          ['Ventas y Comercial','Carlos Eduardo López','Director de Ventas','c.lopez@empresa.com','103'],
          ['Marketing y Comunicación','Laura Valentina Torres','Directora Marketing','l.torres@empresa.com','104'],
          ['Recursos Humanos','Patricia Isabel Morales','Directora RRHH','p.morales@empresa.com','105'],
          ['Soporte al Cliente','Roberto Miguel Sánchez','Jefe de Soporte','r.sanchez@empresa.com','106'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee'} },
          { r:0,c:4, fmt:{bold:true,bgColor:'#0d1a1a',color:'#22d3ee',align:'center'} },
        ],
      },
    ],
  },

  /* ════ MARKETING ════ */
  {
    id: 'campaign', name: 'Seguimiento de campañas',
    desc: 'Métricas de campañas digitales: impresiones, clics, conversiones y ROAS.',
    icon: '📣', cat: 'marketing', fileName: 'campanas-marketing.xlsx',
    sheets: [
      {
        name: 'Campañas Q1',
        widths: { 0:180, 1:100, 2:110, 3:80, 4:90, 5:80, 6:100, 7:90 },
        data: [
          ['Campaña','Canal','Presupuesto','Impr.','Clics','CTR %','Conversiones','ROAS'],
          ['Verano 2025 — Google','Google Ads','3500','120000','4800','=E2/D2*100','180','=H2/(C2/1000)'],
          ['Verano 2025 — Facebook','Facebook Ads','2000','85000','2900','=E3/D3*100','95','=H3/(C3/1000)'],
          ['Lanzamiento producto A','Instagram Ads','1500','60000','3100','=E4/D4*100','120','=H4/(C4/1000)'],
          ['Email newsletter abril','Email','200','0','12500','=E5/45000*100','320','=H5/(C5/1000)'],
          ['Remarketing Q1','Google Display','800','200000','1200','=E6/D6*100','45','=H6/(C6/1000)'],
          ['','','=SUM(C2:C6)','','','','=SUM(G2:G6)',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1a120a',color:'#fb923c'} },
          { r:0,c:7, fmt:{bold:true,bgColor:'#1a120a',color:'#fb923c',align:'center'} },
          { r:6,c:2, fmt:{bold:true,color:'#fb923c',numFormat:'currency'} },
          { r:6,c:6, fmt:{bold:true,color:'#fb923c'} },
        ],
      },
    ],
  },

  {
    id: 'social', name: 'Calendario de contenidos',
    desc: 'Planificación de publicaciones en redes sociales por semana y canal.',
    icon: '📱', cat: 'marketing', fileName: 'calendario-contenidos.xlsx',
    sheets: [
      {
        name: 'Enero 2025',
        widths: { 0:60, 1:120, 2:200, 3:100, 4:100, 5:100, 6:120 },
        data: [
          ['Semana','Fecha','Contenido / Tema','Instagram','Facebook','LinkedIn','Estado'],
          ['Sem 1','06/01/2025','Año nuevo — retrospectiva 2024','Reels','Post + Carrusel','Artículo','Publicado'],
          ['Sem 1','08/01/2025','Tendencias del sector 2025','Story','Post','Post','Publicado'],
          ['Sem 2','13/01/2025','Caso de éxito cliente A','Carrusel','Post + Video','Artículo','Programado'],
          ['Sem 2','15/01/2025','Tip de productividad','Reels','Post','—','Borrador'],
          ['Sem 3','20/01/2025','Lanzamiento nueva funcionalidad','Post + Story','Post','Comunicado','Pendiente'],
          ['Sem 3','22/01/2025','Concurso de engagement','Carrusel','Post','—','Pendiente'],
          ['Sem 4','27/01/2025','Resumen del mes','Reels + Story','Post','Artículo','Pendiente'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#1a120a',color:'#fb923c',align:'center'} },
          { r:0,c:2, fmt:{bold:true,bgColor:'#1a120a',color:'#fb923c'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#1a120a',color:'#fb923c',align:'center'} },
          { r:1,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:3,c:6, fmt:{color:'#fbbf24',bold:true} },
        ],
      },
    ],
  },

  /* ════ LEGAL ════ */
  {
    id: 'contracts', name: 'Registro de contratos',
    desc: 'Contratos activos con fechas, partes, valor y estado de vigencia.',
    icon: '📜', cat: 'legal', fileName: 'registro-contratos.xlsx',
    sheets: [
      {
        name: 'Contratos',
        widths: { 0:90, 1:200, 2:160, 3:110, 4:110, 5:110, 6:90 },
        data: [
          ['N° Contrato','Objeto','Contraparte','Inicio','Vencimiento','Valor','Estado'],
          ['CT-2024-001','Servicios de desarrollo web','TechCorp S.A.','01/01/2024','31/12/2025','24000.00','Activo'],
          ['CT-2024-002','Suministro de equipos TI','Tech Supplies S.L.','15/02/2024','14/02/2025','15000.00','Vencido'],
          ['CT-2024-003','Consultoría estratégica','BizConsult LLC','01/04/2024','31/03/2026','36000.00','Activo'],
          ['CT-2024-004','Mantenimiento instalaciones','Limpieza Pro SA','01/01/2024','31/12/2024','8400.00','Vencido'],
          ['CT-2025-001','Licencias software ERP','SAP Partners','01/01/2025','31/12/2027','45000.00','Activo'],
          ['CT-2025-002','Publicidad digital','Digital Agency','01/02/2025','31/07/2025','12000.00','Activo'],
          ['','','','','TOTAL ACTIVOS','=SUM(F2:F7)',''],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#140e1e',color:'#a78bfa'} },
          { r:0,c:6, fmt:{bold:true,bgColor:'#140e1e',color:'#a78bfa',align:'center'} },
          { r:1,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:3,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:5,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:6,c:6, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:6, fmt:{color:'#f87171',bold:true} },
          { r:4,c:6, fmt:{color:'#f87171',bold:true} },
          { r:7,c:4, fmt:{bold:true,color:'#a78bfa'} },
          { r:7,c:5, fmt:{bold:true,color:'#a78bfa',numFormat:'currency'} },
        ],
      },
    ],
  },

  {
    id: 'compliance', name: 'Checklist de cumplimiento',
    desc: 'Lista de verificación normativa con estado, responsable y fecha de revisión.',
    icon: '⚖️', cat: 'legal', fileName: 'checklist-cumplimiento.xlsx',
    sheets: [
      {
        name: 'Cumplimiento 2025',
        widths: { 0:40, 1:280, 2:120, 3:100, 4:150, 5:120 },
        data: [
          ['#','Requisito normativo','Norma / Ley','Estado','Responsable','Próxima revisión'],
          ['1','Política de privacidad actualizada','GDPR / LOPD','Cumplido','Ana Martínez','01/07/2025'],
          ['2','Registro de actividades de tratamiento','GDPR Art. 30','Cumplido','Pedro Ruiz','01/04/2025'],
          ['3','Cláusulas de protección en contratos','GDPR Art. 28','En revisión','María García','15/03/2025'],
          ['4','Evaluación de impacto (EIPD)','GDPR Art. 35','Pendiente','Ana Martínez','30/06/2025'],
          ['5','Seguridad de sistemas certificada','ISO 27001','Cumplido','Carlos López','01/01/2026'],
          ['6','Protocolo de brechas de seguridad','GDPR Art. 33','En revisión','María García','28/02/2025'],
          ['7','Formación anual en protección de datos','Interna','Pendiente','RRHH','31/03/2025'],
          ['8','Auditoría de proveedores cloud','GDPR Art. 28','Pendiente','Carlos López','30/04/2025'],
        ],
        fmts: [
          { r:0,c:0, fmt:{bold:true,bgColor:'#140e1e',color:'#a78bfa',align:'center'} },
          { r:0,c:1, fmt:{bold:true,bgColor:'#140e1e',color:'#a78bfa'} },
          { r:0,c:3, fmt:{bold:true,bgColor:'#140e1e',color:'#a78bfa',align:'center'} },
          { r:1,c:3, fmt:{color:'#4ade80',bold:true} },
          { r:2,c:3, fmt:{color:'#4ade80',bold:true} },
          { r:5,c:3, fmt:{color:'#4ade80',bold:true} },
          { r:3,c:3, fmt:{color:'#fbbf24',bold:true} },
          { r:6,c:3, fmt:{color:'#fbbf24',bold:true} },
        ],
      },
    ],
  },

];

/* ─────────────────────────────────────────────
   PREVIEW VISUAL — mini spreadsheet uniforme
───────────────────────────────────────────── */
const CAT_COLORS = {
  general: '#4ade80', finanzas: '#7c8cf8', almacen: '#fbbf24',
  proyectos: '#c084fc', analisis: '#f87171', rrhh: '#22d3ee',
  marketing: '#fb923c', legal: '#a78bfa',
};

function buildPreviewHTML(tpl) {
  var c    = CAT_COLORS[tpl.cat] || '#7c8cf8';
  var cols = 4;
  var rows = 3;
  var result = '';

  /* Fila de encabezado */
  var headerRow = '';
  for (var j = 0; j < cols; j++) {
    headerRow +=
      '<div class="tpl-preview-cell head" style="flex:' + (j === 0 ? '2' : '1') +
      ';background:' + c + '22;border-color:' + c + '55;"></div>';
  }
  result += '<div class="tpl-preview-row">' + headerRow + '</div>';

  /* Filas de datos */
  for (var i = 0; i < rows; i++) {
    var dataRow = '';
    for (var j2 = 0; j2 < cols; j2++) {
      var isTotal = (i === rows - 1) && (j2 === cols - 1);
      dataRow +=
        '<div class="tpl-preview-cell data" style="flex:' + (j2 === 0 ? '2' : '1') +
        ';background:' + (isTotal ? c + '44' : (i % 2 === 0 ? 'var(--bg-hover)' : 'var(--bg-panel)')) +
        ';"></div>';
    }
    result += '<div class="tpl-preview-row">' + dataRow + '</div>';
  }

  return (
    '<div class="tpl-card-preview">' +
      result +
      '<div class="tpl-card-icon-overlay">' + tpl.icon + '</div>' +
    '</div>'
  );
}

/* ─────────────────────────────────────────────
   GENERADOR DE WORKBOOK
───────────────────────────────────────────── */
function generateTemplateWorkbook(tpl) {
  var wb = XLSX.utils.book_new();
  tpl.sheets.forEach(function(sd) {
    var ws = XLSX.utils.aoa_to_sheet(sd.data);
    XLSX.utils.book_append_sheet(wb, ws, sd.name);
  });
  return wb;
}

function loadTemplateIntoApp(tpl) {
  var overlay = document.getElementById('tplGenerating');
  var label   = document.getElementById('tplGeneratingLabel');
  if (overlay) overlay.classList.remove('hidden');
  if (label)   label.textContent = 'Generando: ' + tpl.name + '…';

  setTimeout(function() {
    try {
      var wb     = generateTemplateWorkbook(tpl);
      var wbOut  = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      var data   = new Uint8Array(wbOut);
      var wbRead = XLSX.read(data, { type: 'array', cellDates: true });

      STATE.workbook   = wbRead;
      STATE.sheetNames = wbRead.SheetNames;
      STATE.fileName   = tpl.fileName;

      loadSheet(wbRead.SheetNames[0]);
      showEditor();

      tpl.sheets.forEach(function(sd, si) {
        var sheetName = wbRead.SheetNames[si];
        if (!sheetName || !sd.fmts) return;
        var prev = STATE.activeSheet;
        STATE.activeSheet = sheetName;
        sd.fmts.forEach(function(f) { setCellFmt(f.r, f.c, f.fmt); });
        STATE.activeSheet = prev;
      });

      if (tpl.sheets[0] && tpl.sheets[0].widths) {
        Object.keys(tpl.sheets[0].widths).forEach(function(col) {
          STATE.colWidths[parseInt(col)] = tpl.sheets[0].widths[col];
        });
      }

      STATE.activeSheet = wbRead.SheetNames[0];
      renderTable();
      markDirty();
      showToast('Plantilla cargada: ' + tpl.name, 'success');
      hideTplPanel();
    } catch (err) {
      console.error('Error generando plantilla:', err);
      showToast('Error al generar plantilla: ' + (err.message || err), 'error');
    }
    if (overlay) overlay.classList.add('hidden');
  }, 60);
}

/* ─────────────────────────────────────────────
   HOJA EN BLANCO
───────────────────────────────────────────── */
function loadBlankSheet() {
  var overlay = document.getElementById('tplGenerating');
  var label   = document.getElementById('tplGeneratingLabel');
  if (overlay) overlay.classList.remove('hidden');
  if (label)   label.textContent = 'Creando hoja en blanco…';

  setTimeout(function() {
    try {
      var rows = 50, cols = 20;
      var emptyData = [];
      for (var r = 0; r < rows; r++) {
        var row = [];
        for (var c = 0; c < cols; c++) row.push('');
        emptyData.push(row);
      }

      var wb  = XLSX.utils.book_new();
      var ws  = XLSX.utils.aoa_to_sheet(emptyData);
      XLSX.utils.book_append_sheet(wb, ws, 'Hoja1');

      var wbOut  = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      var data2  = new Uint8Array(wbOut);
      var wbRead = XLSX.read(data2, { type: 'array', cellDates: true });

      STATE.workbook   = wbRead;
      STATE.sheetNames = wbRead.SheetNames;
      STATE.fileName   = 'libro-nuevo.xlsx';

      loadSheet(wbRead.SheetNames[0]);
      showEditor();
      renderTable();
      markDirty();
      showToast('Hoja en blanco creada', 'success');
      hideTplPanel();
    } catch (err) {
      showToast('Error al crear hoja: ' + (err.message || err), 'error');
    }
    if (overlay) overlay.classList.add('hidden');
  }, 40);
}

/* ─────────────────────────────────────────────
   FILTRADO Y RENDER DEL GRID
───────────────────────────────────────────── */
var _activeFilter = 'todos';
var _searchQuery  = '';

function getFilteredTemplates() {
  return TEMPLATES.filter(function(t) {
    var matchCat    = _activeFilter === 'todos' || t.cat === _activeFilter;
    var matchSearch = !_searchQuery ||
      t.name.toLowerCase().indexOf(_searchQuery) !== -1 ||
      t.desc.toLowerCase().indexOf(_searchQuery) !== -1;
    return matchCat && matchSearch;
  });
}

function renderTplGrid() {
  var grid = document.getElementById('tplGrid');
  if (!grid) return;
  grid.innerHTML = '';

  var list = getFilteredTemplates();

  if (!list.length) {
    grid.innerHTML =
      '<div class="tpl-empty">' +
        '<svg width="40" height="40" viewBox="0 0 20 20" fill="currentColor">' +
          '<path fill-rule="evenodd" d="M8 4a4 4 0 100 8 4 4 0 000-8zM2 8a6 6 0 1110.89 3.476l4.817 4.817a1 1 0 01-1.414 1.414l-4.816-4.816A6 6 0 012 8z" clip-rule="evenodd"/>' +
        '</svg>' +
        '<div class="tpl-empty-title">Sin resultados</div>' +
        '<div class="tpl-empty-sub">Prueba con otra búsqueda o categoría</div>' +
      '</div>';
    return;
  }

  var catClass = {
    general: 'cat-general', finanzas: 'cat-finanzas', almacen: 'cat-almacen',
    proyectos: 'cat-proyectos', analisis: 'cat-analisis', rrhh: 'cat-rrhh',
    marketing: 'cat-marketing', legal: 'cat-legal',
  };

  list.forEach(function(tpl) {
    var card = document.createElement('div');
    /* Sin tpl-featured — todas son iguales */
    card.className   = 'tpl-card';
    card.dataset.cat = tpl.cat;
    card.title       = tpl.desc;

    var cm  = CAT_META[tpl.cat] || { label: tpl.cat };
    var ccl = catClass[tpl.cat] || 'cat-general';

    var sheetsInfo = tpl.sheets && tpl.sheets.length > 1
      ? '<span class="tpl-card-sheets">' +
          '<svg width="10" height="10" viewBox="0 0 20 20" fill="currentColor">' +
            '<path d="M9 4.804A7.968 7.968 0 005.5 4c-1.255 0-2.443.29-3.5.804v10A7.969 7.969 0 015.5 14c1.69 0 3.254.776 4.5 2.015V4.804zM10.5 16.015A7.967 7.967 0 0115 14c.966 0 1.874.18 2.5.398V5.202A7.967 7.967 0 0015 4c-1.255 0-2.443.29-3.5.804v11.211z"/>' +
          '</svg>' +
          tpl.sheets.length + ' hojas' +
        '</span>'
      : '';

    card.innerHTML =
      buildPreviewHTML(tpl) +
      '<div class="tpl-card-body">' +
        '<div class="tpl-card-name">' + escapeHtml(tpl.name) + '</div>' +
        '<div class="tpl-card-desc">' + escapeHtml(tpl.desc) + '</div>' +
      '</div>' +
      '<div class="tpl-card-footer">' +
        '<span class="tpl-card-badge ' + ccl + '">' + escapeHtml(cm.label) + '</span>' +
        sheetsInfo +
      '</div>';

    card.addEventListener('click', function() { loadTemplateIntoApp(tpl); });
    grid.appendChild(card);
  });

  updateSidebarCounts();
}

function updateSidebarCounts() {
  var btns = document.querySelectorAll('.tpl-filter-btn[data-filter]');
  btns.forEach(function(btn) {
    var f   = btn.dataset.filter;
    var cnt = f === 'todos'
      ? TEMPLATES.length
      : TEMPLATES.filter(function(t) { return t.cat === f; }).length;
    var el = btn.querySelector('.tpl-filter-count');
    if (el) el.textContent = cnt;
  });
}

/* ─────────────────────────────────────────────
   SHOW / HIDE PANEL
───────────────────────────────────────────── */
function showTplPanel() {
  var uploadModes = document.getElementById('uploadModes');
  var tplPanel    = document.getElementById('tplPanel');
  if (uploadModes) uploadModes.style.display = 'none';
  if (tplPanel)    tplPanel.classList.add('tpl-panel-open');
  renderTplGrid();
}

function hideTplPanel() {
  var uploadModes = document.getElementById('uploadModes');
  var tplPanel    = document.getElementById('tplPanel');
  if (uploadModes) uploadModes.style.display = '';
  if (tplPanel)    tplPanel.classList.remove('tpl-panel-open');
}

/* ─────────────────────────────────────────────
   SIDEBAR — construcción de filtros
───────────────────────────────────────────── */
function _buildSidebarFilters() {
  var cats = [
    { id: 'todos',     label: 'Todas',      color: 'var(--text-muted)' },
    { id: 'general',   label: 'General',    color: CAT_COLORS.general   },
    { id: 'finanzas',  label: 'Finanzas',   color: CAT_COLORS.finanzas  },
    { id: 'almacen',   label: 'Almacén',    color: CAT_COLORS.almacen   },
    { id: 'proyectos', label: 'Proyectos',  color: CAT_COLORS.proyectos },
    { id: 'analisis',  label: 'Análisis',   color: CAT_COLORS.analisis  },
    { id: 'rrhh',      label: 'RRHH',       color: CAT_COLORS.rrhh      },
    { id: 'marketing', label: 'Marketing',  color: CAT_COLORS.marketing },
    { id: 'legal',     label: 'Legal',      color: CAT_COLORS.legal     },
  ];

  return cats.map(function(cat) {
    var cnt = cat.id === 'todos'
      ? TEMPLATES.length
      : TEMPLATES.filter(function(t) { return t.cat === cat.id; }).length;
    var active = cat.id === 'todos' ? ' tpl-filter-active' : '';
    return (
      '<button class="tpl-filter-btn' + active + '" data-filter="' + cat.id + '">' +
        '<span class="tpl-filter-dot" style="background:' + cat.color + '"></span>' +
        cat.label +
        '<span class="tpl-filter-count">' + cnt + '</span>' +
      '</button>'
    );
  }).join('');
}

/* ─────────────────────────────────────────────
   INICIALIZACIÓN PRINCIPAL
───────────────────────────────────────────── */
function initTemplateSelector() {
  var section = document.getElementById('uploadSection');
  if (!section) return;

  /* ══ Reconstruir el HTML de la sección de carga ══ */
  section.innerHTML =

    /* ══ MODOS DE INICIO — 3 tarjetas uniformes ══ */
    '<div class="upload-modes" id="uploadModes">' +

      /* — Tarjeta A: Abrir archivo — */
      '<div class="upload-card-wrapper" id="dropZone">' +
        '<div class="hx-mode-card-wrap">' +
          '<div class="hx-mode-card-inner">' +
            '<div class="hx-mode-icon hx-mode-icon-snowflake" style="font-size:28px;">❄️</div>' +
            '<h2 class="hx-mode-title">Abrir archivo Excel</h2>' +
            '<p class="hx-mode-sub">Sube un archivo .xlsx, .xls o .csv desde tu equipo o arrástralo aquí.</p>' +
            '<label class="hx-mode-cta" for="fileInput" style="pointer-events:auto;cursor:pointer;">' +
              '<svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">' +
                '<path d="M3 17a1 1 0 011-1h12a1 1 0 110 2H4a1 1 0 01-1-1zM6.293 6.707a1 1 0 010-1.414l3-3a1 1 0 011.414 0l3 3a1 1 0 01-1.414 1.414L11 5.414V13a1 1 0 11-2 0V5.414L7.707 6.707a1 1 0 01-1.414 0z"/>' +
              '</svg>' +
              'Seleccionar archivo' +
              '<input type="file" id="fileInput" accept=".xlsx,.xls,.csv" hidden />' +
            '</label>' +
          '</div>' +
        '</div>' +
      '</div>' +

      '<span class="upload-mode-divider">o</span>' +

      /* — Tarjeta B: Hoja en blanco — */
      '<div id="btnBlankSheet" role="button" tabindex="0" ' +
           'title="Crear una hoja de cálculo vacía de 50×20 celdas">' +
        '<div class="hx-mode-card-wrap">' +
          '<div class="hx-mode-card-inner">' +
            '<div class="hx-mode-icon hx-mode-icon-blank">' +
              '<svg viewBox="0 0 20 20" fill="currentColor" width="28" height="28">' +
                '<path fill-rule="evenodd" d="M10 5a1 1 0 011 1v3h3a1 1 0 110 2h-3v3a1 1 0 11-2 0v-3H6a1 1 0 110-2h3V6a1 1 0 011-1z" clip-rule="evenodd"/>' +
              '</svg>' +
            '</div>' +
            '<h2 class="hx-mode-title">Hoja en blanco</h2>' +
            '<p class="hx-mode-sub">Comienza desde cero con una hoja vacía de 50 filas × 20 columnas.</p>' +
            '<span class="hx-mode-cta">Crear ahora →</span>' +
          '</div>' +
        '</div>' +
      '</div>' +

      '<span class="upload-mode-divider">o</span>' +

      /* — Tarjeta C: Plantillas — */
      '<div id="btnCreateTemplate" role="button" tabindex="0" ' +
           'title="Ver y usar plantillas profesionales">' +
        '<div class="hx-mode-card-wrap">' +
          '<div class="hx-mode-card-inner">' +
            '<div class="hx-mode-icon hx-mode-icon-tpl">' +
              '<svg viewBox="0 0 20 20" fill="currentColor" width="28" height="28">' +
                '<path d="M5 3a2 2 0 00-2 2v10a2 2 0 002 2h10a2 2 0 002-2V5a2 2 0 00-2-2H5zm0 2h4v4H5V5zm6 0h4v4h-4V5zm-6 6h4v4H5v-4zm6 0h4v4h-4v-4z"/>' +
              '</svg>' +
            '</div>' +
            '<h2 class="hx-mode-title">Usar plantilla</h2>' +
            '<p class="hx-mode-sub">' + TEMPLATES.length + '+ plantillas profesionales listas para editar.</p>' +
            '<span class="hx-mode-cta">Ver plantillas →</span>' +
          '</div>' +
        '</div>' +
      '</div>' +

    '</div>' + /* fin #uploadModes */

    /* ══ PANEL DE PLANTILLAS ══ */
    '<div class="tpl-panel" id="tplPanel" aria-label="Selector de plantillas">' +

      /* Sidebar */
      '<nav class="tpl-sidebar">' +
        '<div class="tpl-sidebar-logo">' +
          '<div class="tpl-sidebar-title">📐 Plantillas</div>' +
          '<div class="tpl-sidebar-subtitle">Elige una y personalízala</div>' +
        '</div>' +
        '<button class="tpl-back-btn" id="tplBackBtn">' +
          '<svg viewBox="0 0 20 20" fill="currentColor" width="12" height="12">' +
            '<path fill-rule="evenodd" d="M7.707 14.707a1 1 0 01-1.414 0l-4-4a1 1 0 010-1.414l4-4a1 1 0 011.414 1.414L5.414 9H17a1 1 0 110 2H5.414l2.293 2.293a1 1 0 010 1.414z" clip-rule="evenodd"/>' +
          '</svg>' +
          'Volver' +
        '</button>' +
        '<span class="tpl-sidebar-section">Categorías</span>' +
        _buildSidebarFilters() +
      '</nav>' +

      /* Contenido */
      '<div class="tpl-content">' +
        '<div class="tpl-content-header">' +
          '<div class="tpl-content-title-row">' +
            '<div class="tpl-content-title" id="tplContentTitle">Todas las plantillas</div>' +
            '<div class="tpl-content-subtitle" id="tplContentSubtitle">' + TEMPLATES.length + ' plantillas disponibles</div>' +
          '</div>' +
          '<div class="tpl-search-wrap">' +
            '<svg viewBox="0 0 20 20" fill="currentColor" width="13" height="13">' +
              '<path fill-rule="evenodd" d="M8 4a4 4 0 100 8 4 4 0 000-8zM2 8a6 6 0 1110.89 3.476l4.817 4.817a1 1 0 01-1.414 1.414l-4.816-4.816A6 6 0 012 8z" clip-rule="evenodd"/>' +
            '</svg>' +
            '<input type="text" class="tpl-search-input" id="tplSearchInput" placeholder="Buscar plantillas…" autocomplete="off" />' +
          '</div>' +
        '</div>' +
        '<div class="tpl-grid-scroll">' +
          '<div class="tpl-grid" id="tplGrid" role="list"></div>' +
        '</div>' +
      '</div>' +

    '</div>'; /* fin #tplPanel */

  /* ── Event listeners ── */

  /* Botón hoja en blanco */
  var btnBlank = document.getElementById('btnBlankSheet');
  if (btnBlank) {
    btnBlank.addEventListener('click', loadBlankSheet);
    btnBlank.addEventListener('keydown', function(e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); loadBlankSheet(); }
    });
  }

  /* Botón "Usar plantilla" */
  var btnCreate = document.getElementById('btnCreateTemplate');
  if (btnCreate) {
    btnCreate.addEventListener('click', showTplPanel);
    btnCreate.addEventListener('keydown', function(e) {
      if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); showTplPanel(); }
    });
  }

  /* Botón Volver */
  var btnBack = document.getElementById('tplBackBtn');
  if (btnBack) btnBack.addEventListener('click', hideTplPanel);

  /* Filtros del sidebar */
  var sidebar = document.querySelector('.tpl-sidebar');
  if (sidebar) {
    sidebar.addEventListener('click', function(e) {
      var btn = e.target.closest('.tpl-filter-btn[data-filter]');
      if (!btn) return;
      _activeFilter = btn.dataset.filter;

      sidebar.querySelectorAll('.tpl-filter-btn').forEach(function(b) {
        b.classList.remove('tpl-filter-active');
      });
      btn.classList.add('tpl-filter-active');

      var cm    = CAT_META[_activeFilter];
      var title = cm ? cm.label : 'Todas las plantillas';
      var titleEl = document.getElementById('tplContentTitle');
      if (titleEl) titleEl.textContent = _activeFilter === 'todos' ? 'Todas las plantillas' : title;

      var cnt = _activeFilter === 'todos'
        ? TEMPLATES.length
        : TEMPLATES.filter(function(t) { return t.cat === _activeFilter; }).length;
      var subtitleEl = document.getElementById('tplContentSubtitle');
      if (subtitleEl) subtitleEl.textContent = cnt + ' plantilla' + (cnt === 1 ? '' : 's');

      renderTplGrid();
    });
  }

  /* Búsqueda en tiempo real */
  var searchInput = document.getElementById('tplSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', function() {
      _searchQuery = searchInput.value.trim().toLowerCase();
      renderTplGrid();
    });
  }

  /* Drag & Drop — dropZone */
  var dropZone = document.getElementById('dropZone');
  if (dropZone) {
    dropZone.addEventListener('dragover', function(e) {
      e.preventDefault();
      dropZone.classList.add('drag-over');
    });
    dropZone.addEventListener('dragleave', function() {
      dropZone.classList.remove('drag-over');
    });
    dropZone.addEventListener('drop', function(e) {
      e.preventDefault();
      dropZone.classList.remove('drag-over');
      var file = e.dataTransfer.files[0];
      if (file && typeof loadFile === 'function') loadFile(file);
    });
  }

  /* fileInput change */
  var fileInput = document.getElementById('fileInput');
  if (fileInput) {
    fileInput.addEventListener('change', function(e) {
      var file = e.target.files[0];
      if (file && typeof loadFile === 'function') loadFile(file);
      fileInput.value = '';
    });
  }

  /* Render inicial del grid */
  renderTplGrid();
}
