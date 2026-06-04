// ============================================================
//  CMMS FLOTA MARÍTIMA — Sistema tipo SAP PM
//  Versión 3.0 — Multi-usuario (20+ personas)
//  Con control de roles, caché y auditoría
// ============================================================

// ============================================================
//  CMMS FLOTA MARÍTIMA — Sistema tipo SAP PM
//  Versión 3.1 — Stock de Repuestos + Combustible por Equipo
// ============================================================

const CONFIG = {
  SHEETS: {
    EMBARCACIONES:    'Embarcaciones',
    SISTEMAS:         'Sistemas',
    EQUIPOS:          'Equipos',
    PLAN_MANT:        'Plan_Mantenimiento',
    ORDENES:          'Ordenes_Trabajo',
    REGISTRO_FALLAS:  'Registro_Fallas',
    PARAMETROS:       'Registro_Parametros',
    KPIS:             'KPIs',
    CERTIFICADOS:     'Certificados',
    COMBUSTIBLE:      'Combustible',
    USUARIOS:         'Usuarios',
    CATALOGO_ACCIONES:'Catalogo_Acciones',
    AUDITORIA:        'Auditoria',
    CONFIG:           'Configuracion',
    STOCK_ITEMS:        'Stock_Items',
    STOCK_MOVIMIENTOS:  'Stock_Movimientos',
    COMBUSTIBLE_DIARIO: 'Combustible_Diario',
    G125_SEGUIMIENTO:   'G125_Seguimiento',
    G126_ACTIVIDADES:   'G126_Actividades'
  },
  // ── ROLES ──────────────────────────────────────────────────
  // ADMIN      → acceso total, gestión de usuarios
  // SUPERVISOR → lectura total + crear OTs + KPIs + certificados
  // TECNICO    → parámetros, fallas, combustible, OTs
  // OPERADOR   → solo lectura + registrar parámetros básicos
  ROLES: {
    ADMIN:      { nivel: 4, label: 'Administrador',  color: '#b71c1c' },
    SUPERVISOR: { nivel: 3, label: 'Supervisor',     color: '#e65100' },
    TECNICO:    { nivel: 2, label: 'Técnico',        color: '#1565c0' },
    OPERADOR:   { nivel: 1, label: 'Operador',       color: '#2e7d32' }
  },
  PERMISOS: {
    'embarcaciones.ver':    1,
    'embarcaciones.editar': 4,
    'equipos.ver':          1,
    'equipos.crear':        3,
    'equipos.editar':       3,
    'ordenes.ver':          1,
    'ordenes.crear':        2,
    'ordenes.cerrar':       3,
    'fallas.ver':           1,
    'fallas.crear':         2,
    'parametros.ver':       1,
    'parametros.crear':     1,
    'combustible.ver':      1,
    'combustible.crear':    2,
    'certificados.ver':     1,
    'certificados.crear':   3,
    'kpis.ver':             2,
    'usuarios.ver':         4,
    'usuarios.gestionar':   4,
    'plan.ver':             1,
    'plan.crear':           3,
    'stock.ver':            1,
    'stock.mover':          2,
    'stock.admin':          3
  },
  VERSION:   '3.1',
  APP_NAME:  'CMMS Flota Marítima',
  CACHE_TTL: 120  // segundos de vida del caché
};

// ════════════════════════════════════════════════════════════
//  PUNTO DE ENTRADA WEB
// ════════════════════════════════════════════════════════════
function doGet(e) {
  const html = HtmlService.createTemplateFromFile('index');
  html.appName = CONFIG.APP_NAME;
  html.version = CONFIG.VERSION;
  return html.evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ════════════════════════════════════════════════════════════
//  AUTENTICACIÓN Y CONTROL DE ROLES
// ════════════════════════════════════════════════════════════
function getUsuarioActual() {
  const email = Session.getActiveUser().getEmail();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const h     = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
  if (!h) return { email, nombre: email.split('@')[0], rol: 'TECNICO', embarcaciones: 'TODAS' };

  const data = h.getDataRange().getValues();
  const fila = data.slice(1).find(r => r[2] === email && r[5] === true);

  if (!fila) {
    const nombre = email.split('@')[0];
    const id = _genId('USR');
    h.appendRow([id, nombre, email, 'OPERADOR', 'TODAS', true, new Date()]);
    return { email, nombre, rol: 'OPERADOR', embarcaciones: 'TODAS', esNuevo: true };
  }
  return {
    email,
    nombre:        fila[1],
    rol:           fila[3],
    embarcaciones: fila[4],
    activo:        fila[5]
  };
}

function tienePermiso(accion) {
  try {
    const usuario  = getUsuarioActual();
    const nivelReq = CONFIG.PERMISOS[accion] || 99;
    const nivelUsr = (CONFIG.ROLES[usuario.rol] || { nivel: 1 }).nivel;
    return nivelUsr >= nivelReq;
  } catch(e) { return false; }
}

function _checkPermiso(accion) {
  if (!tienePermiso(accion))
    throw new Error('Sin permiso para: ' + accion + '. Contacta al administrador.');
}

function getUserInfo() {
  const u = getUsuarioActual();
  return {
    email:        u.email,
    nombre:       u.nombre,
    rol:          u.rol,
    rolLabel:     (CONFIG.ROLES[u.rol] || { label: 'Operador' }).label,
    rolColor:     (CONFIG.ROLES[u.rol] || { color: '#1565c0' }).color,
    permisos:     Object.keys(CONFIG.PERMISOS).filter(p => {
                    const niv = (CONFIG.ROLES[u.rol] || { nivel: 1 }).nivel;
                    return niv >= CONFIG.PERMISOS[p];
                  }),
    embarcaciones: u.embarcaciones,
    esNuevo:      u.esNuevo || false
  };
}

// ── Gestión de usuarios (solo ADMIN) ──
function getUsuarios() {
  _checkPermiso('usuarios.ver');
  return _leerHoja(CONFIG.SHEETS.USUARIOS);
}

function guardarUsuario(datos) {
  _checkPermiso('usuarios.gestionar');
  const h    = _ss().getSheetByName(CONFIG.SHEETS.USUARIOS);
  const data = h.getDataRange().getValues();
  const idx  = data.findIndex(r => r[0] === datos.ID_USUARIO);
  if (idx > 0) {
    h.getRange(idx+1, 4).setValue(datos.ROL);
    h.getRange(idx+1, 5).setValue(datos.EMBARCACIONES_ASIGNADAS || 'TODAS');
    h.getRange(idx+1, 6).setValue(datos.ACTIVO !== undefined ? datos.ACTIVO : true);
    _auditoria('USUARIO_ACTUALIZADO', datos.ID_USUARIO, datos.EMAIL);
    return { success: true, mensaje: 'Usuario actualizado' };
  }
  const id = _genId('USR');
  h.appendRow([id, datos.NOMBRE, datos.EMAIL, datos.ROL,
    datos.EMBARCACIONES_ASIGNADAS || 'TODAS', true, new Date()]);
  _auditoria('USUARIO_CREADO', id, datos.EMAIL);
  return { success: true, id, mensaje: 'Usuario creado: ' + datos.EMAIL };
}

// ════════════════════════════════════════════════════════════
//  INICIALIZACIÓN DEL SISTEMA
// ════════════════════════════════════════════════════════════

// ── Verificación rápida del sistema (para carga inicial de la web) ──
// No crea hojas ni configura nada — solo verifica que el sistema existe
function getSystemStatus() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) return { success: false, error: 'Sin hoja activa' };
    // Solo verificar que las hojas principales existen
    const hE = ss.getSheetByName(CONFIG.SHEETS.EMBARCACIONES);
    const hF = ss.getSheetByName(CONFIG.SHEETS.REGISTRO_FALLAS);
    if (!hE) {
      // Primera vez: inicializar
      return inicializarSistema();
    }
    return { success: true, version: CONFIG.VERSION,
             mensaje: 'CMMS v' + CONFIG.VERSION + ' activo' };
  } catch(e) { return { error: e.message }; }
}

function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('No hay hoja de cálculo activa. Ejecuta desde una Google Sheet.');
  Object.values(CONFIG.SHEETS).forEach(nombre => {
    if (!ss.getSheetByName(nombre)) ss.insertSheet(nombre);
  });
  _configurarHojasBase(ss);
  _inicializarHojasStock(ss);
  _inicializarHojaCombustibleDiario(ss);
  _inicializarFormatosG(ss);
  return { success: true, mensaje: 'Sistema CMMS v' + CONFIG.VERSION + ' inicializado', version: CONFIG.VERSION };
}

function _configurarHojasBase(ss) {
  const AZ = '#1a3a5c', BL = '#ffffff';
  function cab(h, cols) {
    if (h.getLastRow() > 0) return;
    h.appendRow(cols);
    h.getRange(1, 1, 1, cols.length)
      .setFontWeight('bold').setBackground(AZ).setFontColor(BL).setFontSize(11);
    h.setFrozenRows(1);
  }

  // EMBARCACIONES
  const hE = ss.getSheetByName(CONFIG.SHEETS.EMBARCACIONES);
  cab(hE, ['ID','NOMBRE','CODIGO','TIPO','ARMADOR','CENTRO_MANT','ESLORA','MANGA',
           'CALADO','POTENCIA_KW','ANO_CONSTRUCCION','BANDERA','ESTADO','FECHA_REGISTRO']);
  if (hE.getLastRow() <= 1) {
    [['EMB001','ALCATRAZ','ALC','Remolcador','Agrorural','CFMYT',28,8,3.2,800,2005,'PERU','ACTIVO',new Date()],
     ['EMB002','GUANAY','GUA','Barcaza','Agrorural','CFMYT',45,12,2.8,1200,2008,'PERU','ACTIVO',new Date()],
     ['EMB003','ISLA CHINCHA','ICH','Nave de carga','Agrorural','CFMYT',52,14,3.5,1600,2010,'PERU','ACTIVO',new Date()],
     ['EMB004','DELFIN 11','D11','Lancha','Agrorural','CFMYT',18,5,1.5,400,2012,'PERU','ACTIVO',new Date()],
     ['EMB005','DELFIN 12','D12','Lancha','Agrorural','CFMYT',18,5,1.5,400,2013,'PERU','ACTIVO',new Date()],
     ['EMB006','PELICANO','PEL','Remolcador','Agrorural','CFMYT',30,8.5,3,900,2007,'PERU','ACTIVO',new Date()],
     ['EMB007','CISTERNA 1','CIS','Cisterna','Agrorural','CFMYT',35,10,2.8,600,2009,'PERU','ACTIVO',new Date()],
     ['EMB008','PANGA','PAN','Panga','Agrorural','CFMYT',8,2.5,0.8,90,2015,'PERU','ACTIVO',new Date()]
    ].forEach(r => hE.appendRow(r));
  }

  // SISTEMAS
  const hS = ss.getSheetByName(CONFIG.SHEETS.SISTEMAS);
  cab(hS, ['ID_SISTEMA','CODIGO','NOMBRE','DESCRIPCION','CRITICIDAD']);
  if (hS.getLastRow() <= 1) {
    [['SIS001','SP','Sistema de Propulsion','Motores principales y cajas reductoras',5],
     ['SIS002','SG','Sistema de Gobierno','Equipos de gobierno hidráulico y timones',4],
     ['SIS003','SE','Sistema de Energia','Grupos electrógenos y tableros',4],
     ['SIS004','SGR','Sistema de Grua','Grúas y sistemas hidráulicos de carga',3],
     ['SIS005','SF','Sistema de Fondeo','Unidades hidráulicas, anclas y cadenas',3],
     ['SIS006','SAL','Sistema de Achique/Lastre','Electrobombas y sistemas de achique',4],
     ['SIS007','SAC','Sistema de Aire Comprimido','Compresores y depósitos',2],
     ['SIS008','EA','Equipos Auxiliares','Equipos auxiliares varios',2],
     ['SIS009','CC','Cubierta y Casco','Estructura, luces y equipos de cubierta',3]
    ].forEach(r => hS.appendRow(r));
  }

  // CATÁLOGO DE ACCIONES
  const hA = ss.getSheetByName(CONFIG.SHEETS.CATALOGO_ACCIONES);
  cab(hA, ['CODIGO','NOMBRE','DESCRIPCION','TIPO']);
  if (hA.getLastRow() <= 1) {
    [['IV','INSPECCION VISUAL','Inspección ocular sin desmontaje','Preventivo'],
     ['IM','INSPECCION MANUAL','Inspección táctil y funcional','Preventivo'],
     ['PR','PRUEBAS','Pruebas de funcionamiento y rendimiento','Preventivo'],
     ['CF','CAMBIO DE FILTROS','Cambio según intervalo programado','Preventivo'],
     ['MP','MANTENIMIENTO PREVENTIVO','Mantenimiento completo programado','Preventivo'],
     ['MC','MANTENIMIENTO CORRECTIVO','Reparación post-falla','Correctivo'],
     ['SOV','SEMI OVERHAUL','Revisión mayor parcial','Overhaul'],
     ['OVT','OVER HAUL TOTAL','Revisión mayor total del equipo','Overhaul']
    ].forEach(r => hA.appendRow(r));
  }

  // Tablas transaccionales — solo cabeceras
  cab(ss.getSheetByName(CONFIG.SHEETS.EQUIPOS),
    ['ID_EQUIPO','ID_EMBARCACION','ID_SISTEMA','CODIGO_EQUIPO','NOMBRE_EQUIPO','MARCA',
     'MODELO','SERIE','ANO_FABRICACION','POTENCIA','TIPO_INSPECCION','HOROMETRO_ACTUAL',
     'HOROMETRO_ULT_MANT','PROX_MANT_HORAS','ESTADO','CRITICIDAD','FECHA_INSTALACION']);
  cab(ss.getSheetByName(CONFIG.SHEETS.ORDENES),
    ['ID_OT','NUMERO_OT','TIPO','ID_EMBARCACION','EMBARCACION','ID_SISTEMA','SISTEMA',
     'ID_EQUIPO','EQUIPO','DESCRIPCION','ACCION','FECHA_REGISTRO','FECHA_PROGRAMADA',
     'FECHA_INICIO','FECHA_CIERRE','ESTADO','PRIORIDAD','TECNICO','HORAS_MANO_OBRA',
     'COSTO_REPUESTOS','COSTO_TOTAL','OBSERVACIONES','CREATED_BY','CREATED_AT']);
  cab(ss.getSheetByName(CONFIG.SHEETS.REGISTRO_FALLAS),
    ['ID_FALLA','USUARIO','ARMADOR','CODIGO_FALLA','CENTRO_MANT','UBICACION_TECNICA',
     'EQUIPO','SISTEMA','HORAS_OP','LISTA_REPUESTOS','PLAN_MANT','FECHA_REGISTRO',
     'NOTIFICACION','DESCRIPCION_FALLA','ORDEN_MANT','FECHA_CIERRE','EQUIPO_REPARADO',
     'MES','ANO','TIEMPO_FALLA_DIAS','SEVERIDAD','CAUSA_RAIZ','ACCION_TOMADA','COSTO_ESTIMADO']);
  cab(ss.getSheetByName(CONFIG.SHEETS.PARAMETROS),
    ['ID_PARAM','FECHA','ID_EMBARCACION','EMBARCACION','ID_EQUIPO','EQUIPO','CODIGO_GENERAL',
     'COMPONENTE','INSPECCION_VISUAL_SEMANAL','HORAS_OPERACION','OBSERVACION',
     'HOROMETRO_ULT_MANT','HOROMETRO_ACTUAL','HOROMETRO_PROX_MANT','HOROMETRO_REPORTADO',
     'HUMO','RPM','TEMP_AGUA_C','PRESION_LO_PSI','PRESION_FO_PSI','VOLTAJE',
     'OTRO_REPORTE','CONSUMO_PETROLEO','MILLAS_RECORRIDAS','HOR_INICIAL','HOR_FINAL','REGISTRADO_POR']);
  cab(ss.getSheetByName(CONFIG.SHEETS.KPIS),
    ['ID_KPI','ANO','MES','ID_EMBARCACION','EMBARCACION','ID_SISTEMA','SISTEMA','ID_EQUIPO',
     'EQUIPO','NUM_FALLAS','DIAS_TOTALES','DIAS_FOP','TIEMPO_OP','MTTR','MTBF',
     'TASA_FALLAS','DISPONIBILIDAD','PROB_FALLA_365','MANTENIBILIDAD','CRITICIDAD','FECHA_CALCULO']);
  cab(ss.getSheetByName(CONFIG.SHEETS.CERTIFICADOS),
    ['ID_CERT','ID_EMBARCACION','EMBARCACION','TIPO_CERTIFICADO','NUMERO','ORGANISMO_EMISOR',
     'FECHA_EMISION','FECHA_VENCIMIENTO','ESTADO','DIAS_ALERTA','OBSERVACIONES','ARCHIVO_URL']);
  cab(ss.getSheetByName(CONFIG.SHEETS.COMBUSTIBLE),
    ['ID_COMB','FECHA','ID_EMBARCACION','EMBARCACION','TIPO_COMBUSTIBLE','CANTIDAD_GALONES',
     'COSTO_POR_GALON','COSTO_TOTAL','MILLAS_RECORRIDAS','HORAS_OPERACION',
     'CONSUMO_POR_HORA','CONSUMO_POR_MILLA','PROVEEDOR','REGISTRO_POR','OBSERVACIONES']);
  cab(ss.getSheetByName(CONFIG.SHEETS.PLAN_MANT),
    ['ID_PLAN','ID_EMBARCACION','ID_EQUIPO','EQUIPO','FRECUENCIA','TIPO_FRECUENCIA',
     'ACCION','DESCRIPCION','HORAS_BASE','PROXIMA_EJECUCION','ULTIMA_EJECUCION',
     'ESTADO','RESPONSABLE','ACTIVO']);
  cab(ss.getSheetByName(CONFIG.SHEETS.AUDITORIA),
    ['TIMESTAMP','USUARIO','EMAIL','ACCION','ENTIDAD_ID','DETALLE','IP_APROX']);

  // USUARIOS — primer ADMIN = quien ejecuta
  const hU = ss.getSheetByName(CONFIG.SHEETS.USUARIOS);
  cab(hU, ['ID_USUARIO','NOMBRE','EMAIL','ROL','EMBARCACIONES_ASIGNADAS','ACTIVO','FECHA_REGISTRO']);
  if (hU.getLastRow() <= 1) {
    const em = Session.getActiveUser().getEmail();
    hU.appendRow(['USR001', em.split('@')[0], em, 'ADMIN', 'TODAS', true, new Date()]);
  }

  // CONFIGURACIÓN
  const hC = ss.getSheetByName(CONFIG.SHEETS.CONFIG);
  cab(hC, ['CLAVE','VALOR','DESCRIPCION']);
  if (hC.getLastRow() <= 1) {
    [['APP_VERSION','3.0','Versión del sistema'],
     ['EMPRESA','Agrorural','Nombre de la empresa'],
     ['CENTRO_MANT','CFMYT','Centro de mantenimiento'],
     ['ALERTAS_CERT_DIAS','30','Días anticipación alerta certificados'],
     ['MAX_DIAS_OT_ABIERTA','30','Días máximos OT abierta antes de alerta'],
     ['MONEDA','S/','Moneda local'],
     ['TIMEZONE','America/Lima','Zona horaria']
    ].forEach(r => hC.appendRow(r));
  }
}

// ════════════════════════════════════════════════════════════
//  UTILIDADES INTERNAS
// ════════════════════════════════════════════════════════════
function _ss() { return SpreadsheetApp.getActiveSpreadsheet(); }

function _leerHoja(nombreHoja) {
  const h = _ss().getSheetByName(nombreHoja);
  if (!h) return [];
  const data = h.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  const tz = Session.getScriptTimeZone();
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((col, i) => {
      obj[col] = row[i] instanceof Date
        ? Utilities.formatDate(row[i], tz, 'yyyy-MM-dd')
        : row[i];
    });
    return obj;
  }).filter(r => r[headers[0]] !== '' && r[headers[0]] !== null && r[headers[0]] !== undefined);
}

function _genId(prefix) {
  return prefix + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMddHHmmss')
         + Math.floor(Math.random()*100);
}

function _auditoria(accion, entidadId, detalle) {
  try {
    const u = getUsuarioActual();
    _ss().getSheetByName(CONFIG.SHEETS.AUDITORIA)
         .appendRow([new Date(), u.nombre, u.email, accion, entidadId, detalle, '']);
  } catch(e) { /* nunca bloquear por auditoría */ }
}

// ════════════════════════════════════════════════════════════
//  CACHÉ INTELIGENTE — evita golpear Sheets con 20+ usuarios
// ════════════════════════════════════════════════════════════
function _cacheGet(key) {
  try {
    const v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch(e) { return null; }
}
function _cacheSet(key, data) {
  try {
    const j = JSON.stringify(data);
    if (j.length < 100000) CacheService.getScriptCache().put(key, j, CONFIG.CACHE_TTL);
  } catch(e) {}
}
function _cacheInvalidar(key) {
  try { CacheService.getScriptCache().remove(key); } catch(e) {}
}

// ════════════════════════════════════════════════════════════
//  CATÁLOGOS (con caché de 2 minutos)
// ════════════════════════════════════════════════════════════
function getEmbarcaciones() {
  const k = 'emb_v1';
  return _cacheGet(k) || (_cacheSet(k, _leerHoja(CONFIG.SHEETS.EMBARCACIONES)), _cacheGet(k));
}
function getSistemas() {
  const k = 'sis_v1';
  return _cacheGet(k) || (_cacheSet(k, _leerHoja(CONFIG.SHEETS.SISTEMAS)), _cacheGet(k));
}
function getEquipos(idEmbarcacion) {
  const k = 'eq_' + (idEmbarcacion || 'all');
  let d = _cacheGet(k);
  if (!d) {
    const todos = _leerHoja(CONFIG.SHEETS.EQUIPOS);
    d = idEmbarcacion ? todos.filter(r => r.ID_EMBARCACION === idEmbarcacion) : todos;
    _cacheSet(k, d);
  }
  return d;
}
function getCatalogoAcciones() {
  const k = 'acc_v1';
  return _cacheGet(k) || (_cacheSet(k, _leerHoja(CONFIG.SHEETS.CATALOGO_ACCIONES)), _cacheGet(k));
}
function getConfiguracion() {
  const rows = _leerHoja(CONFIG.SHEETS.CONFIG);
  const cfg = {};
  rows.forEach(r => { cfg[r.CLAVE] = r.VALOR; });
  return cfg;
}

// ════════════════════════════════════════════════════════════
//  EQUIPOS
// ════════════════════════════════════════════════════════════
function guardarEquipo(datos) {
  _checkPermiso('equipos.crear');
  const h = _ss().getSheetByName(CONFIG.SHEETS.EQUIPOS);
  if (datos.ID_EQUIPO) {
    _checkPermiso('equipos.editar');
    const data = h.getDataRange().getValues();
    const idx  = data.findIndex(r => r[0] === datos.ID_EQUIPO);
    if (idx > 0) {
      data[0].forEach((col, i) => { if (datos[col] !== undefined) h.getRange(idx+1, i+1).setValue(datos[col]); });
      _cacheInvalidar('eq_all'); _cacheInvalidar('eq_'+datos.ID_EMBARCACION);
      _auditoria('EQUIPO_EDITADO', datos.ID_EQUIPO, datos.NOMBRE_EQUIPO);
      return { success: true, mensaje: 'Equipo actualizado' };
    }
  }
  const id = _genId('EQ');
  h.appendRow([id, datos.ID_EMBARCACION, datos.ID_SISTEMA, datos.CODIGO_EQUIPO||'',
    datos.NOMBRE_EQUIPO, datos.MARCA||'', datos.MODELO||'', datos.SERIE||'',
    datos.ANO_FABRICACION||'', datos.POTENCIA||'', datos.TIPO_INSPECCION||'IV',
    datos.HOROMETRO_ACTUAL||0, datos.HOROMETRO_ULT_MANT||0, datos.PROX_MANT_HORAS||0,
    'OPERATIVO', datos.CRITICIDAD||1, new Date()]);
  _cacheInvalidar('eq_all'); _cacheInvalidar('eq_'+datos.ID_EMBARCACION);
  _auditoria('EQUIPO_CREADO', id, datos.NOMBRE_EQUIPO);
  return { success: true, id, mensaje: 'Equipo registrado' };
}

// ════════════════════════════════════════════════════════════
//  ÓRDENES DE TRABAJO
// ════════════════════════════════════════════════════════════
function getOrdenesTrabajoResumen() {
  return _leerHoja(CONFIG.SHEETS.ORDENES); // Sin caché — siempre actualizado
}

function crearOrdenTrabajo(datos) {
  _checkPermiso('ordenes.crear');
  const h     = _ss().getSheetByName(CONFIG.SHEETS.ORDENES);
  const ahora = new Date();
  const numOT = 'OT-' + Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'yyyy')
                + '-' + String(h.getLastRow()).padStart(5,'0');
  const id    = _genId('OT');
  const user  = getUsuarioActual();
  h.appendRow([id, numOT, datos.TIPO, datos.ID_EMBARCACION, datos.EMBARCACION,
    datos.ID_SISTEMA, datos.SISTEMA, datos.ID_EQUIPO||'', datos.EQUIPO||'',
    datos.DESCRIPCION, datos.ACCION||'', ahora, datos.FECHA_PROGRAMADA||'',
    '', '', 'ABIERTA', datos.PRIORIDAD||'NORMAL', datos.TECNICO||'',
    0, 0, 0, datos.OBSERVACIONES||'', user.email, ahora]);
  _auditoria('OT_CREADA', id, numOT+' | '+datos.EMBARCACION);
  return { success: true, id, numOT, mensaje: 'OT creada: '+numOT };
}

function actualizarOrdenTrabajo(datos) {
  _checkPermiso('ordenes.cerrar');
  const h    = _ss().getSheetByName(CONFIG.SHEETS.ORDENES);
  const data = h.getDataRange().getValues();
  const idx  = data.findIndex(r => r[0] === datos.ID_OT);
  if (idx < 1) return { error: 'OT no encontrada' };
  data[0].forEach((col, i) => { if (datos[col] !== undefined) h.getRange(idx+1, i+1).setValue(datos[col]); });
  if (datos.ESTADO === 'CERRADA') h.getRange(idx+1, data[0].indexOf('FECHA_CIERRE')+1).setValue(new Date());
  _auditoria('OT_ACTUALIZADA', datos.ID_OT, 'Estado: '+datos.ESTADO);
  return { success: true, mensaje: 'OT actualizada' };
}

// ════════════════════════════════════════════════════════════
//  FALLAS
// ════════════════════════════════════════════════════════════
function getRegistroFallas(filtros) {
  let rows = _leerHoja(CONFIG.SHEETS.REGISTRO_FALLAS);
  if (filtros) {
    if (filtros.embarcacion) rows = rows.filter(r => r.UBICACION_TECNICA === filtros.embarcacion);
    if (filtros.ano)         rows = rows.filter(r => String(r.ANO) === String(filtros.ano));
    if (filtros.mes)         rows = rows.filter(r => String(r.MES) === String(filtros.mes));
  }
  return rows;
}

function registrarFalla(datos) {
  _checkPermiso('fallas.crear');
  const h    = _ss().getSheetByName(CONFIG.SHEETS.REGISTRO_FALLAS);
  const ahora = new Date();
  const user  = getUsuarioActual();
  const id    = _genId('FL');
  const codigo = 'F-'+Utilities.formatDate(ahora, Session.getScriptTimeZone(),'yyyyMMdd')
                 +'-'+String(h.getLastRow()).padStart(4,'0');
  h.appendRow([id, user.nombre, datos.ARMADOR||'Agrorural', codigo,
    datos.CENTRO_MANT||'CFMYT', datos.UBICACION_TECNICA, datos.EQUIPO, datos.SISTEMA,
    datos.HORAS_OP||0, datos.LISTA_REPUESTOS||'', datos.PLAN_MANT||'Correctivo', ahora,
    datos.NOTIFICACION||'', datos.DESCRIPCION_FALLA, datos.ORDEN_MANT||'', '',
    datos.EQUIPO_REPARADO||'NO', ahora.getMonth()+1, ahora.getFullYear(),
    datos.TIEMPO_FALLA_DIAS||0, datos.SEVERIDAD||'MEDIA',
    datos.CAUSA_RAIZ||'', datos.ACCION_TOMADA||'', datos.COSTO_ESTIMADO||0]);
  _auditoria('FALLA_REGISTRADA', id, codigo+' | '+datos.UBICACION_TECNICA+' | '+datos.EQUIPO);
  return { success: true, id, codigoFalla: codigo, mensaje: 'Falla registrada: '+codigo };
}

// ════════════════════════════════════════════════════════════
//  PARÁMETROS OPERATIVOS
// ════════════════════════════════════════════════════════════
function registrarParametros(datos) {
  _checkPermiso('parametros.crear');
  const h    = _ss().getSheetByName(CONFIG.SHEETS.PARAMETROS);
  const user = getUsuarioActual();
  const id   = _genId('PM');
  const ahora = new Date();
  h.appendRow([id, ahora, datos.ID_EMBARCACION, datos.EMBARCACION,
    datos.ID_EQUIPO, datos.EQUIPO, datos.CODIGO_GENERAL||'', datos.COMPONENTE||'',
    datos.INSPECCION_VISUAL||'', datos.HORAS_OP||0, datos.OBSERVACION||'',
    datos.HOROMETRO_ULT_MANT||0, datos.HOROMETRO_ACTUAL||0,
    datos.HOROMETRO_PROX_MANT||0, datos.HOROMETRO_REPORTADO||0,
    datos.HUMO||'NOR', datos.RPM||0, datos.TEMP_AGUA||0,
    datos.PRESION_LO||0, datos.PRESION_FO||0, datos.VOLTAJE||0,
    datos.OTRO_REPORTE||'', datos.CONSUMO_PETROLEO||0,
    datos.MILLAS_RECORRIDAS||0, datos.HOR_INICIAL||0, datos.HOR_FINAL||0, user.email]);
  // Actualizar horómetro del equipo automáticamente
  if (datos.ID_EQUIPO && parseFloat(datos.HOROMETRO_ACTUAL) > 0) {
    try {
      const hEq = _ss().getSheetByName(CONFIG.SHEETS.EQUIPOS);
      const eqData = hEq.getDataRange().getValues();
      const eqIdx  = eqData.findIndex(r => r[0] === datos.ID_EQUIPO);
      if (eqIdx > 0) {
        hEq.getRange(eqIdx+1, 12).setValue(parseFloat(datos.HOROMETRO_ACTUAL));
        _cacheInvalidar('eq_all'); _cacheInvalidar('eq_'+datos.ID_EMBARCACION);
      }
    } catch(e) {}
  }
  _auditoria('PARAMETROS_REGISTRADOS', id, datos.EMBARCACION+' | '+datos.EQUIPO);
  return { success: true, id, mensaje: 'Parámetros registrados correctamente' };
}

// ════════════════════════════════════════════════════════════
//  COMBUSTIBLE
// ════════════════════════════════════════════════════════════
function getCombustible(filtros) {
  let rows = _leerHoja(CONFIG.SHEETS.COMBUSTIBLE);
  if (filtros && filtros.embarcacion) rows = rows.filter(r => r.ID_EMBARCACION === filtros.embarcacion);
  return rows;
}

function registrarCombustible(datos) {
  _checkPermiso('combustible.crear');
  const h    = _ss().getSheetByName(CONFIG.SHEETS.COMBUSTIBLE);
  const user = getUsuarioActual();
  const id   = _genId('CB');
  const costo    = (datos.CANTIDAD_GALONES||0) * (datos.COSTO_POR_GALON||0);
  const cH       = datos.HORAS_OPERACION > 0 ? +(datos.CANTIDAD_GALONES/datos.HORAS_OPERACION).toFixed(2) : 0;
  const cM       = datos.MILLAS_RECORRIDAS > 0 ? +(datos.CANTIDAD_GALONES/datos.MILLAS_RECORRIDAS).toFixed(2) : 0;
  h.appendRow([id, datos.FECHA||new Date(), datos.ID_EMBARCACION, datos.EMBARCACION,
    datos.TIPO_COMBUSTIBLE||'DIESEL', datos.CANTIDAD_GALONES||0, datos.COSTO_POR_GALON||0,
    costo, datos.MILLAS_RECORRIDAS||0, datos.HORAS_OPERACION||0,
    cH, cM, datos.PROVEEDOR||'', user.email, datos.OBSERVACIONES||'']);
  _auditoria('COMBUSTIBLE_REGISTRADO', id, datos.EMBARCACION+' | '+(datos.CANTIDAD_GALONES||0)+' gal');
  return { success: true, id, mensaje: 'Consumo registrado' };
}

// ════════════════════════════════════════════════════════════
//  CERTIFICADOS
// ════════════════════════════════════════════════════════════
function getCertificados() {
  const rows = _leerHoja(CONFIG.SHEETS.CERTIFICADOS);
  const hoy  = new Date();
  return rows.map(c => {
    if (c.FECHA_VENCIMIENTO) {
      const venc = new Date(c.FECHA_VENCIMIENTO);
      const dias = Math.ceil((venc-hoy)/(1000*60*60*24));
      c.DIAS_RESTANTES = dias;
      c.ESTADO_ALERTA  = dias < 0 ? 'VENCIDO' : dias <= 30 ? 'CRITICO' : dias <= 90 ? 'ALERTA' : 'VIGENTE';
    }
    return c;
  });
}

function guardarCertificado(datos) {
  _checkPermiso('certificados.crear');
  const h  = _ss().getSheetByName(CONFIG.SHEETS.CERTIFICADOS);
  const id = _genId('CRT');
  h.appendRow([id, datos.ID_EMBARCACION, datos.EMBARCACION, datos.TIPO_CERTIFICADO,
    datos.NUMERO||'', datos.ORGANISMO_EMISOR||'', datos.FECHA_EMISION||'',
    datos.FECHA_VENCIMIENTO||'', 'VIGENTE', datos.DIAS_ALERTA||30,
    datos.OBSERVACIONES||'', datos.ARCHIVO_URL||'']);
  _auditoria('CERTIFICADO_CREADO', id, datos.TIPO_CERTIFICADO+' | '+datos.EMBARCACION);
  return { success: true, id, mensaje: 'Certificado registrado' };
}

// ════════════════════════════════════════════════════════════
//  PLAN DE MANTENIMIENTO
// ════════════════════════════════════════════════════════════
function getPlanesMant() { return _leerHoja(CONFIG.SHEETS.PLAN_MANT); }

function guardarPlanMant(datos) {
  _checkPermiso('plan.crear');
  const h  = _ss().getSheetByName(CONFIG.SHEETS.PLAN_MANT);
  const id = _genId('PL');
  h.appendRow([id, datos.ID_EMBARCACION, datos.ID_EQUIPO||'', datos.EQUIPO||'General',
    datos.FRECUENCIA, datos.TIPO_FRECUENCIA||'HORAS', datos.ACCION,
    datos.DESCRIPCION||'', datos.HORAS_BASE||0, datos.PROXIMA_EJECUCION||'',
    '', 'PENDIENTE', datos.RESPONSABLE||'', true]);
  _auditoria('PLAN_CREADO', id, datos.EQUIPO+' | '+datos.ACCION);
  return { success: true, id, mensaje: 'Plan creado' };
}

// ════════════════════════════════════════════════════════════
//  DASHBOARD CONSOLIDADO (1 llamada = todos los datos)
// ════════════════════════════════════════════════════════════
function getDashboardData() {
  try {
    const anoActual     = new Date().getFullYear();
    const hoy           = new Date();
    const embarcaciones = getEmbarcaciones();
    const fallasData    = _leerHoja(CONFIG.SHEETS.REGISTRO_FALLAS);
    const otData        = _leerHoja(CONFIG.SHEETS.ORDENES);
    const certData      = getCertificados();
    const combData      = _leerHoja(CONFIG.SHEETS.COMBUSTIBLE);

    const fallasAno    = fallasData.filter(r => String(r.ANO) === String(anoActual));
    const totalFallas  = fallasAno.length;
    const otAbiertas   = otData.filter(r => r.ESTADO === 'ABIERTA').length;
    const otCerradas   = otData.filter(r => r.ESTADO === 'CERRADA').length;
    const otVencidas   = otData.filter(o => o.ESTADO==='ABIERTA' && o.FECHA_REGISTRO &&
                           Math.ceil((hoy-new Date(o.FECHA_REGISTRO))/(1000*60*60*24)) > 30).length;
    const cVencidos    = certData.filter(c => c.ESTADO_ALERTA==='VENCIDO').length;
    const cCriticos    = certData.filter(c => c.ESTADO_ALERTA==='CRITICO').length;
    const totalGal     = combData.reduce((s,r) => s+(parseFloat(r.CANTIDAD_GALONES)||0), 0);
    const totalCosto   = combData.reduce((s,r) => s+(parseFloat(r.COSTO_TOTAL)||0), 0);

    // Fallas por embarcación
    const fallasPorEmb = {};
    fallasAno.forEach(r => { const k=r.UBICACION_TECNICA||'Sin asignar'; fallasPorEmb[k]=(fallasPorEmb[k]||0)+1; });

    // Fallas por mes
    const fallasPorMes = Array(12).fill(0);
    fallasAno.forEach(r => { if(r.MES) fallasPorMes[parseInt(r.MES)-1]++; });

    // Disponibilidad rápida
    const dispFlota = embarcaciones.map(emb => {
      const fe  = fallasAno.filter(r => r.UBICACION_TECNICA === emb.NOMBRE);
      const fop = fe.reduce((s,r) => s+(parseFloat(r.TIEMPO_FALLA_DIAS)||0), 0);
      const d   = (365-fop)/365;
      return { nombre: emb.NOMBRE, codigo: emb.CODIGO,
               disponibilidad: +(d*100).toFixed(1), diasFop: fop, numFallas: fe.length };
    });

    return {
      success: true,
      usuario:  getUserInfo(),
      resumen: {
        totalEmbarcaciones: embarcaciones.length,
        embarcacionesActivas: embarcaciones.filter(e=>e.ESTADO==='ACTIVO').length,
        totalFallas, totalOTs: otData.length, otAbiertas, otCerradas, otVencidas,
        certsVencidos: cVencidos, certsCriticos: cCriticos,
        totalCombustible: +totalGal.toFixed(0),
        totalCostoComb:   +totalCosto.toFixed(2),
        anoActual
      },
      fallasPorEmb, fallasPorMes, dispFlota, embarcaciones,
      ultimasFallas: fallasData.slice(-5).reverse(),
      ultimasOTs:    otData.filter(o=>o.ESTADO==='ABIERTA').slice(-8)
    };
  } catch(e) { return { error: e.message }; }
}

// ════════════════════════════════════════════════════════════
//  KPIs — ISO 14224 / SAP PM
// ════════════════════════════════════════════════════════════
function calcularKPIs(idEmbarcacion, ano) {
  try {
    _checkPermiso('kpis.ver');
    const fallas  = getRegistroFallas({ ano: String(ano) });
    const equipos = getEquipos(idEmbarcacion || null);
    const DIAS    = 365;
    const embNombre = idEmbarcacion ? (getEmbarcaciones().find(e=>e.ID===idEmbarcacion)||{}).NOMBRE : null;

    const kpis = equipos.map(equipo => {
      const fallasEq = fallas.filter(f =>
        f.EQUIPO === equipo.NOMBRE_EQUIPO &&
        (!embNombre || f.UBICACION_TECNICA === embNombre));
      const nf   = fallasEq.length;
      const fop  = fallasEq.reduce((s,f) => s+(parseFloat(f.TIEMPO_FALLA_DIAS)||0), 0);
      const top  = DIAS - fop;
      const mttr = nf > 0 ? fop/nf : 0;
      const mtbf = nf > 0 ? top/nf : top;
      const lam  = mtbf > 0 ? 1/mtbf : 0;
      const disp = top/DIAS;
      const rt   = lam > 0 ? 1-Math.exp(-lam*365) : 0;
      const mt   = mttr > 0 ? 1-Math.exp(-(1/mttr)*180) : 0;
      return {
        sistema: equipo.ID_SISTEMA, equipo: equipo.NOMBRE_EQUIPO,
        numFallas: nf, diasFop: fop, tiempoOp: top,
        mttr: +mttr.toFixed(2), mtbf: +mtbf.toFixed(2),
        tasaFallas: +lam.toFixed(6),
        disponibilidadPct: +(disp*100).toFixed(2),
        disponibilidadNum: disp,
        probFalla: +(rt*100).toFixed(2),
        mantenibilidad: +(mt*100).toFixed(2),
        criticidad: parseFloat(equipo.CRITICIDAD)||0
      };
    });
    return { success: true, kpis };
  } catch(e) { return { error: e.message }; }
}

// ════════════════════════════════════════════════════════════
//  DISPONIBILIDAD DE FLOTA
// ════════════════════════════════════════════════════════════
function getDisponibilidadFlota() {
  const embs = getEmbarcaciones();
  const fd   = _leerHoja(CONFIG.SHEETS.REGISTRO_FALLAS);
  const ano  = String(new Date().getFullYear());
  return embs.map(emb => {
    const fe  = fd.filter(r => r.UBICACION_TECNICA===emb.NOMBRE && String(r.ANO)===ano);
    const fop = fe.reduce((s,r) => s+(parseFloat(r.TIEMPO_FALLA_DIAS)||0), 0);
    const d   = (365-fop)/365;
    return { id: emb.ID, nombre: emb.NOMBRE, codigo: emb.CODIGO,
             numFallas: fe.length, diasFop: fop,
             disponibilidad: +(d*100).toFixed(1),
             estado: d>=0.85?'BUENA':d>=0.70?'REGULAR':'CRITICA' };
  });
}

// ════════════════════════════════════════════════════════════
//  MATRIZ DE CRITICIDAD
// ════════════════════════════════════════════════════════════
function getMatrizCriticidad() {
  return getEquipos().map(e => ({
    equipo: e.NOMBRE_EQUIPO, sistema: e.ID_SISTEMA, embarcacion: e.ID_EMBARCACION,
    criticidad: parseFloat(e.CRITICIDAD)||0,
    nivel: (parseFloat(e.CRITICIDAD)||0)>=32?'CRITICO'
          :(parseFloat(e.CRITICIDAD)||0)>=16?'ALTO'
          :(parseFloat(e.CRITICIDAD)||0)>=8 ?'MEDIO':'BAJO'
  }));
}

// ════════════════════════════════════════════════════════════
//  ALERTAS
// ════════════════════════════════════════════════════════════
function verificarAlertas() {
  try {
    const alertas = [];
    const hoy = new Date();

    getCertificados().forEach(c => {
      if (c.ESTADO_ALERTA==='VENCIDO')
        alertas.push({ tipo:'CRITICO', mensaje:`Certificado VENCIDO: ${c.TIPO_CERTIFICADO} — ${c.EMBARCACION}`, modulo:'certificados' });
      else if (c.ESTADO_ALERTA==='CRITICO')
        alertas.push({ tipo:'ALERTA', mensaje:`Vence en ${c.DIAS_RESTANTES} días: ${c.TIPO_CERTIFICADO} — ${c.EMBARCACION}`, modulo:'certificados' });
    });

    _leerHoja(CONFIG.SHEETS.ORDENES).forEach(ot => {
      if (ot.ESTADO==='ABIERTA' && ot.FECHA_REGISTRO) {
        const d = Math.ceil((hoy-new Date(ot.FECHA_REGISTRO))/(1000*60*60*24));
        if (d > 30) alertas.push({ tipo:'ALERTA', mensaje:`OT abierta hace ${d} días: ${ot.NUMERO_OT} — ${ot.EMBARCACION}`, modulo:'ordenes' });
      }
    });

    getEquipos().forEach(eq => {
      const act  = parseFloat(eq.HOROMETRO_ACTUAL)||0;
      const prox = parseFloat(eq.PROX_MANT_HORAS)||0;
      if (prox>0 && act>=prox*0.9 && act<prox)
        alertas.push({ tipo:'ALERTA', mensaje:`Próx. mant. en ${(prox-act).toFixed(0)}h: ${eq.NOMBRE_EQUIPO}`, modulo:'equipos' });
      if (prox>0 && act>=prox)
        alertas.push({ tipo:'CRITICO', mensaje:`MANT. VENCIDO por horómetro: ${eq.NOMBRE_EQUIPO} (${act}h / ${prox}h)`, modulo:'equipos' });
    });

    return { success: true, alertas };
  } catch(e) { return { error: e.message, alertas:[] }; }
}

// ════════════════════════════════════════════════════════════
//  REPORTE DIARIO POR EMAIL (Trigger automático)
// ════════════════════════════════════════════════════════════
function enviarReporteDiario() {
  try {
    const alts     = verificarAlertas().alertas || [];
    if (!alts.length) return;
    const usuarios = _leerHoja(CONFIG.SHEETS.USUARIOS);
    const destinos = usuarios.filter(u => (u.ROL==='ADMIN'||u.ROL==='SUPERVISOR') && u.ACTIVO===true).map(u => u.EMAIL);
    if (!destinos.length) return;
    const criticos = alts.filter(a=>a.tipo==='CRITICO');
    const normales = alts.filter(a=>a.tipo==='ALERTA');
    const html = `<div style="font-family:sans-serif;max-width:600px">
      <div style="background:#1a3a5c;padding:20px;border-radius:8px 8px 0 0">
        <h2 style="color:white;margin:0">🚢 CMMS Flota Marítima</h2>
        <p style="color:#90caf9;margin:4px 0 0">Reporte diario — ${Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm')}</p>
      </div>
      <div style="background:#fff;padding:20px;border:1px solid #ddd;border-radius:0 0 8px 8px">
        ${criticos.length?`<h3 style="color:#e53935">🔴 Crítico (${criticos.length})</h3><ul>${criticos.map(a=>`<li>${a.mensaje}</li>`).join('')}</ul>`:''}
        ${normales.length?`<h3 style="color:#fb8c00">🟡 Atención (${normales.length})</h3><ul>${normales.map(a=>`<li>${a.mensaje}</li>`).join('')}</ul>`:''}
        <p style="color:#999;font-size:12px;margin-top:20px">CMMS v${CONFIG.VERSION} — Este es un mensaje automático</p>
      </div></div>`;
    destinos.forEach(email => {
      GmailApp.sendEmail(email,
        `[CMMS] ${criticos.length} crítico(s), ${normales.length} alerta(s) — ${Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy')}`,
        'Ver en HTML', { htmlBody: html });
    });
  } catch(e) { Logger.log('Error reporte: '+e.message); }
}

// Ejecutar UNA vez manualmente para activar el reporte diario a las 7am:
function configurarTriggerDiario() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction()==='enviarReporteDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarReporteDiario').timeBased().atHour(7).everyDays(1).create();
  return '✅ Reporte diario activado — se enviará cada día a las 7:00 AM';
}


// ════════════════════════════════════════════════════════════
//  MÓDULO STOCK DE REPUESTOS Y CONSUMIBLES  v3.1
// ════════════════════════════════════════════════════════════

function _inicializarHojasStock(ss) {
  function cab(nombre, cols) {
    let h = ss.getSheetByName(nombre);
    if (!h) h = ss.insertSheet(nombre);
    if (h.getLastRow() === 0) {
      h.appendRow(cols);
      h.getRange(1,1,1,cols.length).setFontWeight('bold')
        .setBackground('#1a3a5c').setFontColor('#ffffff').setFontSize(11);
      h.setFrozenRows(1);
    }
    return h;
  }
  cab('Stock_Items', ['ID_ITEM','CLAVE','DESCRIPCION','STOCK_ACTUAL','STOCK_MINIMO',
    'UNIDAD','UBICACION','FAMILIA','FECHA_ULT_MOV','USUARIO_MOD']);
  // STOCK_ACTUAL: cantidad real en esa ubicación — se actualiza con cada movimiento
  cab('Stock_Movimientos', ['ID_MOV','FECHA','TIPO','ID_ITEM','CLAVE','DESCRIPCION',
    'CANTIDAD','UBICACION_ORIGEN','UBICACION_DESTINO','MOTIVO','OT_ASOCIADA','EMBARCACION','USUARIO']);
  // Combustible_Diario se inicializa en _inicializarHojaCombustibleDiario()
}

// ════════════════════════════════════════════════════════════
//  CARGA Y SINCRONIZACIÓN DEL CATÁLOGO DE STOCK  v3.2
//  El stock inicial viene del Excel (Existencia Actual).
//  Cada ítem se registra con un movimiento INVENTARIO_INICIAL
//  que establece la base. Los movimientos posteriores se suman.
//  Si el Excel se actualiza, ejecutar sincronizarCatalogoStock()
// ════════════════════════════════════════════════════════════
function cargarStockInicial() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hCat = ss.getSheetByName('Stock_Items');
  const hMov = ss.getSheetByName('Stock_Movimientos');
  if (!hCat || !hMov) {
    SpreadsheetApp.getUi().alert('Primero ejecuta inicializarSistema()');
    return;
  }

  // Verificar si ya fue cargado
  const yaExiste = hCat.getLastRow() > 1;
  if (yaExiste) {
    const r = SpreadsheetApp.getUi().alert(
      '⚠️ El catálogo ya tiene ' + (hCat.getLastRow()-1) + ' ítems.',
      '¿Deseas REEMPLAZARLO con los datos actuales del Excel?\n\n' +
      'ATENCIÓN: esto borrará el catálogo pero NO los movimientos previos.',
      SpreadsheetApp.getUi().ButtonSet.YES_NO
    );
    if (r !== SpreadsheetApp.getUi().Button.YES) return;
    hCat.getRange(2, 1, hCat.getLastRow(), 10).clearContent();
  }

  const items = _getStockItemsData();
  // [CLAVE, DESCRIPCION, EXISTENCIA_INICIAL, UNIDAD, UBICACION, FAMILIA]

  // Cargar catálogo (sin STOCK_ACTUAL — es calculado)
  const filasCat = items.map((it, i) => [
    'STK-' + String(i+1).padStart(4,'0'), // ID_ITEM
    it[0],        // CLAVE
    it[1],        // DESCRIPCION
    it[2],        // STOCK_ACTUAL = existencia del Excel (cantidad real inicial)
    0,            // STOCK_MINIMO (a definir por el usuario)
    it[3],        // UNIDAD
    it[4],        // UBICACION
    it[5],        // FAMILIA
    new Date(),   // FECHA_ULT_MOV
    'CARGA_INICIAL' // USUARIO_MOD
  ]);

  const LOTE = 100;
  for (let i = 0; i < filasCat.length; i += LOTE) {
    const lote = filasCat.slice(i, i + LOTE);
    hCat.getRange(hCat.getLastRow()+1, 1, lote.length, 10).setValues(lote);
    SpreadsheetApp.flush();
  }

  // Registrar movimiento INVENTARIO_INICIAL para cada ítem con stock > 0
  // Esto establece el saldo inicial desde el cual parten los movimientos
  const yaHayMovIniciales = hMov.getLastRow() > 1 &&
    hMov.getRange(2, 3).getValue() === 'INVENTARIO_INICIAL';

  if (!yaHayMovIniciales) {
    const movRows = filasCat
      .filter(r => Number(r[4]) > 0)  // solo items con existencia > 0
      .map((r, i) => [
        'MOV-INI-' + String(i+1).padStart(4,'0'),
        new Date(),
        'INVENTARIO_INICIAL',
        r[0],   // ID_ITEM
        r[1],   // CLAVE
        r[2],   // DESCRIPCION
        r[4],   // CANTIDAD = existencia base
        r[6],   // UBICACION_ORIGEN
        '',     // UBICACION_DESTINO
        'Inventario inicial cargado desde Excel',
        '', '', 'SISTEMA'
      ]);

    const LOTE2 = 100;
    for (let i = 0; i < movRows.length; i += LOTE2) {
      const lote = movRows.slice(i, i + LOTE2);
      hMov.getRange(hMov.getLastRow()+1, 1, lote.length, 13).setValues(lote);
      SpreadsheetApp.flush();
    }
    Logger.log('Movimientos iniciales registrados: ' + movRows.length);
  }

  _cacheInvalidar('stock_items');
  SpreadsheetApp.getUi().alert(
    '✅ Stock cargado correctamente\n\n' +
    '• ' + filasCat.length + ' ítems en catálogo\n' +
    '• ' + filasCat.filter(r=>Number(r[4])>0).length + ' con existencia inicial > 0\n\n' +
    'El stock es ahora calculado automáticamente\ndesde los movimientos.'
  );
}

// ── Sincronizar catálogo si el Excel cambió ──────────────────
// Ejecutar cuando se agreguen nuevos ítems al Excel de Drive
function sincronizarCatalogoStock() {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hCat = ss.getSheetByName('Stock_Items');
  const hMov = ss.getSheetByName('Stock_Movimientos');
  if (!hCat) { SpreadsheetApp.getUi().alert('Primero ejecuta inicializarSistema()'); return; }

  const items    = _getStockItemsData();
  const catData  = hCat.getDataRange().getValues();
  const hdr      = catData[0];
  const ci       = c => hdr.indexOf(c);
  const clavesExistentes = new Set(catData.slice(1).map(r => String(r[ci('CLAVE')])));

  // Solo agregar ítems NUEVOS (no actualizar existentes)
  const nuevos = items.filter(it => it[0] && !clavesExistentes.has(it[0]));

  if (!nuevos.length) {
    SpreadsheetApp.getUi().alert('✅ El catálogo ya está actualizado. No hay ítems nuevos en el Excel.');
    return;
  }

  const startIdx = hCat.getLastRow() - 1; // para numeración de IDs
  const filas = nuevos.map((it, i) => [
    'STK-' + String(startIdx + i + 1).padStart(4,'0'),
    it[0],    // CLAVE
    it[1],    // DESCRIPCION
    it[2],    // STOCK_ACTUAL = existencia del Excel
    0,        // STOCK_MINIMO
    it[3], it[4], it[5], new Date(), 'SINCRONIZACION'
  ]);
  hCat.getRange(hCat.getLastRow()+1, 1, filas.length, 10).setValues(filas);

  // Registrar inventario inicial de los nuevos
  if (hMov) {
    const movRows = filas
      .filter(r => Number(r[4]) > 0)
      .map((r, i) => [
        'MOV-SYNC-' + String(i+1).padStart(4,'0'),
        new Date(), 'INVENTARIO_INICIAL',
        r[0], r[1], r[2], r[4], r[6], '', 'Sincronización desde Excel', '', '', 'SISTEMA'
      ]);
    if (movRows.length) {
      hMov.getRange(hMov.getLastRow()+1, 1, movRows.length, 13).setValues(movRows);
    }
  }

  _cacheInvalidar('stock_items');
  SpreadsheetApp.getUi().alert(
    '✅ Sincronización completada\n' +
    '• ' + nuevos.length + ' ítems nuevos agregados al catálogo.'
  );
}

// ── Lectura de stock ─────────────────────────────────────────
// ════════════════════════════════════════════════════════════
//  STOCK — lógica basada en movimientos (v3.2)
//  El stock real = suma de INGRESOs − suma de SALIDAs/TRASLADOs
//  por cada (ID_ITEM + UBICACION). Stock_Items es solo catálogo.
// ════════════════════════════════════════════════════════════

// ── Calcula stock actual de cada ítem sumando sus movimientos ─
function _calcularStockDesdeMovimientos() {
  // Lee todos los movimientos y calcula el saldo neto por ítem y ubicación.
  // INVENTARIO_INICIAL = stock base del Excel (equivale a INGRESO)
  // INGRESO   → suma en ubicación origen
  // SALIDA    → resta en ubicación origen
  // TRASLADO  → resta en origen, suma en destino
  const movs = _leerHoja('Stock_Movimientos');
  const saldo = {}; // saldo[ID_ITEM][UBICACION] = cantidad neta
  movs.forEach(m => {
    const id   = m.ID_ITEM;
    const orig = m.UBICACION_ORIGEN  || '';
    const dest = m.UBICACION_DESTINO || '';
    const cant = Number(m.CANTIDAD || 0);
    const tipo = String(m.TIPO || '').toUpperCase();
    if (!id) return;
    if (!saldo[id]) saldo[id] = {};
    if (tipo === 'INVENTARIO_INICIAL' || tipo === 'INGRESO') {
      saldo[id][orig] = (saldo[id][orig] || 0) + cant;
    } else if (tipo === 'SALIDA') {
      saldo[id][orig] = (saldo[id][orig] || 0) - cant;
    } else if (tipo === 'TRASLADO') {
      saldo[id][orig] = (saldo[id][orig] || 0) - cant;
      if (dest) saldo[id][dest] = (saldo[id][dest] || 0) + cant;
    }
  });
  return saldo;
}

// ── getStockItems: lee STOCK_ACTUAL directamente de Stock_Items ─
// Cada fila = un ítem en una ubicación específica.
// El STOCK_ACTUAL se actualiza en tiempo real con cada movimiento.
function getStockItems(filtros) {
  _checkPermiso('stock.ver');
  const catalogo = _leerHoja('Stock_Items');
  if (!catalogo || !catalogo.length) return [];

  let items = catalogo.map(r => {
    const stockActual = Number(r.STOCK_ACTUAL || 0);
    const stockMini   = Number(r.STOCK_MINIMO || 0);
    return {
      id:          r.ID_ITEM,
      clave:       r.CLAVE,
      descripcion: r.DESCRIPCION,
      stockActual,
      stockTotal:  stockActual, // mismo campo — cada fila es una ubicación
      stockMinimo: stockMini,
      unidad:      r.UNIDAD,
      ubicacion:   r.UBICACION,
      familia:     r.FAMILIA,
      alerta:      stockActual <= stockMini && stockMini > 0
    };
  });

  if (filtros) {
    if (filtros.ubicacion)   items = items.filter(i => i.ubicacion === filtros.ubicacion);
    if (filtros.familia)     items = items.filter(i => i.familia   === filtros.familia);
    if (filtros.soloAlertas) items = items.filter(i => i.alerta);
    if (filtros.busqueda) {
      const q = filtros.busqueda.toLowerCase();
      items = items.filter(i =>
        i.descripcion.toLowerCase().includes(q) || i.clave.toLowerCase().includes(q));
    }
  }
  return items;
}

// ── Buscar ítems en catálogo (para el formulario de movimiento) ─
function buscarItemsCatalogo(texto) {
  _checkPermiso('stock.ver');
  const catalogo = _leerHoja('Stock_Items');
  if (!catalogo || !catalogo.length) return [];
  const q = (texto || '').toLowerCase();
  const saldo = _calcularStockDesdeMovimientos();
  // Agrupar por CLAVE para mostrar todas las ubicaciones del mismo ítem
  const porClave = {};
  catalogo
    .filter(r => r.DESCRIPCION && (
      r.DESCRIPCION.toLowerCase().includes(q) ||
      String(r.CLAVE || '').toLowerCase().includes(q)))
    .forEach(r => {
      const k = r.CLAVE || r.ID_ITEM;
      if (!porClave[k]) {
        porClave[k] = {
          id: r.ID_ITEM, clave: r.CLAVE,
          descripcion: r.DESCRIPCION,
          unidad: r.UNIDAD, familia: r.FAMILIA,
          ubicacion: r.UBICACION,
          stockTotal: 0,
          saldoPorUbic: {}
        };
      }
      const stock = Number(r.STOCK_ACTUAL || 0);
      porClave[k].stockTotal += stock;
      porClave[k].saldoPorUbic[r.UBICACION] = stock;
    });
  return Object.values(porClave).slice(0, 20);
}

// ── Historial de movimientos de un ítem ───────────────────────
function getMovimientosStock(idItem) {
  _checkPermiso('stock.ver');
  const data = _leerHoja('Stock_Movimientos');
  if (!data || !data.length) return [];
  return data
    .filter(r => !idItem || r.ID_ITEM === idItem)
    .map(r => ({
      id: r.ID_MOV, fecha: r.FECHA, tipo: r.TIPO,
      idItem: r.ID_ITEM, clave: r.CLAVE, descripcion: r.DESCRIPCION,
      cantidad: Number(r.CANTIDAD || 0),
      origen:   r.UBICACION_ORIGEN,
      destino:  r.UBICACION_DESTINO,
      motivo:   r.MOTIVO, otAsociada: r.OT_ASOCIADA,
      embarcacion: r.EMBARCACION, usuario: r.USUARIO
    })).reverse();
}

// ════════════════════════════════════════════════════════════
//  REGISTRAR MOVIMIENTO DE STOCK  v3.4
//  Actualiza STOCK_ACTUAL directamente en Stock_Items (Sheets)
//  Cada par (CLAVE + UBICACION) = una fila en Stock_Items
//  INGRESO  → suma en ubicación destino
//  SALIDA   → resta en ubicación origen
//  TRASLADO → resta en origen, suma en destino (crea fila si no existe)
// ════════════════════════════════════════════════════════════
function registrarMovimientoStock(datos) {
  _checkPermiso('stock.mover');

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const hItm = ss.getSheetByName('Stock_Items');
  const hMov = ss.getSheetByName('Stock_Movimientos');
  if (!hItm || !hMov) return { error: 'Hojas de stock no encontradas. Ejecuta inicializarSistema()' };

  const cantidad    = Number(datos.cantidad);
  const tipo        = String(datos.tipo || '').toUpperCase().trim();
  const ubicOrigen  = String(datos.ubicacionOrigen  || '').trim();
  const ubicDestino = String(datos.ubicacionDestino || '').trim();

  // ── Validaciones previas ─────────────────────────────────
  if (!datos.idItem)  return { error: 'Selecciona un repuesto de la lista' };
  if (cantidad <= 0)  return { error: 'La cantidad debe ser mayor a 0' };
  if (!['INGRESO','SALIDA','TRASLADO'].includes(tipo))
    return { error: 'Tipo de movimiento inválido: ' + tipo };
  if ((tipo === 'TRASLADO') && !ubicDestino)
    return { error: 'Selecciona la ubicación destino para el traslado' };
  if ((tipo === 'TRASLADO') && ubicOrigen === ubicDestino)
    return { error: 'La ubicación origen y destino no pueden ser iguales' };

  // ── Leer hoja LIVE (siempre datos frescos del Sheet) ─────
  // IMPORTANTE: leer con getValues() en cada operación para evitar
  // trabajar con copias en memoria desactualizadas
  function leerHojaLive() {
    const lastRow = hItm.getLastRow();
    if (lastRow < 2) return { hdr: [], rows: [], colIdx: () => -1 };
    const vals = hItm.getRange(1, 1, lastRow, hItm.getLastColumn()).getValues();
    const hdr  = vals[0];
    const ci   = c => hdr.indexOf(c);
    return { hdr, rows: vals, ci };
  }

  // ── Buscar fila por CLAVE + UBICACION (live) ─────────────
  function buscarFila(clave, ubicacion) {
    const { rows, ci } = leerHojaLive();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][ci('CLAVE')]).trim()    === String(clave).trim() &&
          String(rows[i][ci('UBICACION')]).trim() === String(ubicacion).trim()) {
        return i + 1; // número de fila 1-indexed en el Sheet
      }
    }
    return -1;
  }

  // ── Obtener info del ítem desde su ID ────────────────────
  function getInfoItem(idItem) {
    const { rows, ci } = leerHojaLive();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][ci('ID_ITEM')]).trim() === String(idItem).trim()) {
        return {
          fila:        i + 1,
          clave:       String(rows[i][ci('CLAVE')]),
          descripcion: String(rows[i][ci('DESCRIPCION')]),
          unidad:      String(rows[i][ci('UNIDAD')]),
          familia:     String(rows[i][ci('FAMILIA')]),
          stockMini:   Number(rows[i][ci('STOCK_MINIMO')] || 0),
          ubicacion:   String(rows[i][ci('UBICACION')])
        };
      }
    }
    return null;
  }

  // ── Leer STOCK_ACTUAL de una fila (live) ─────────────────
  function getStock(filaNum) {
    const { ci } = leerHojaLive();
    return Number(hItm.getRange(filaNum, ci('STOCK_ACTUAL') + 1).getValue() || 0);
  }

  // ── Escribir STOCK_ACTUAL en una fila ────────────────────
  function setStock(filaNum, nuevoStock, usuario) {
    const { ci } = leerHojaLive();
    hItm.getRange(filaNum, ci('STOCK_ACTUAL')   + 1).setValue(Math.max(0, nuevoStock));
    hItm.getRange(filaNum, ci('FECHA_ULT_MOV')  + 1).setValue(new Date());
    hItm.getRange(filaNum, ci('USUARIO_MOD')     + 1).setValue(usuario);
    SpreadsheetApp.flush(); // forzar escritura inmediata
  }

  // ── Crear nueva fila para ítem en nueva ubicación ────────
  function crearFila(info, ubicacion, stockInicial) {
    const nuevoId = _genId('STK');
    hItm.appendRow([
      nuevoId,
      info.clave,
      info.descripcion,
      stockInicial,
      0,            // STOCK_MINIMO
      info.unidad,
      ubicacion,
      info.familia,
      new Date(),
      getUsuarioActual().email
    ]);
    SpreadsheetApp.flush();
    return hItm.getLastRow();
  }

  // ─────────────────────────────────────────────────────────
  //  INICIO DEL PROCESO
  // ─────────────────────────────────────────────────────────
  const info = getInfoItem(datos.idItem);
  if (!info) return { error: 'Ítem no encontrado con ID: ' + datos.idItem };

  const clave       = info.clave;
  const descripcion = info.descripcion;
  const unidad      = info.unidad;
  const usuario     = getUsuarioActual().email;

  // Usar la ubicación de la fila del ítem si no se especificó origen
  const origenReal  = ubicOrigen || info.ubicacion;

  let stockOrigenFinal  = 0;
  let stockDestinoFinal = 0;
  let mensajeOp         = '';

  // ─────────────────────────────────────────────────────────
  if (tipo === 'INGRESO') {
    // Sumar stock en la ubicación de ingreso
    let filaIngreso = buscarFila(clave, origenReal);
    if (filaIngreso < 0) {
      // No existe fila para esta clave+ubicación → crear
      filaIngreso = crearFila(info, origenReal, 0);
    }
    const stockActual    = getStock(filaIngreso);
    stockOrigenFinal     = stockActual + cantidad;
    setStock(filaIngreso, stockOrigenFinal, usuario);
    mensajeOp = `Ingreso de ${cantidad} ${unidad} en ${origenReal}. Nuevo stock: ${stockOrigenFinal} ${unidad}`;

  // ─────────────────────────────────────────────────────────
  } else if (tipo === 'SALIDA') {
    const filaOrig = buscarFila(clave, origenReal);
    if (filaOrig < 0)
      return { error: `No hay registro de "${descripcion}" en ${origenReal}` };
    const stockActual = getStock(filaOrig);
    if (stockActual < cantidad)
      return { error: `Stock insuficiente en ${origenReal}.
Disponible: ${stockActual} ${unidad} — Solicitado: ${cantidad} ${unidad}` };
    stockOrigenFinal = stockActual - cantidad;
    setStock(filaOrig, stockOrigenFinal, usuario);
    mensajeOp = `Salida de ${cantidad} ${unidad} desde ${origenReal}. Stock restante: ${stockOrigenFinal} ${unidad}`;

  // ─────────────────────────────────────────────────────────
  } else if (tipo === 'TRASLADO') {
    // 1) Restar en origen
    const filaOrig = buscarFila(clave, origenReal);
    if (filaOrig < 0)
      return { error: `No hay registro de "${descripcion}" en ${origenReal}` };
    const stockOrigen = getStock(filaOrig);
    if (stockOrigen < cantidad)
      return { error: `Stock insuficiente en ${origenReal}.
Disponible: ${stockOrigen} ${unidad} — Solicitado: ${cantidad} ${unidad}` };
    stockOrigenFinal = stockOrigen - cantidad;
    setStock(filaOrig, stockOrigenFinal, usuario);

    // 2) Sumar en destino (crear fila si no existe)
    let filaDest = buscarFila(clave, ubicDestino);
    if (filaDest < 0) {
      filaDest = crearFila(info, ubicDestino, 0);
    }
    const stockDestino    = getStock(filaDest);
    stockDestinoFinal     = stockDestino + cantidad;
    setStock(filaDest, stockDestinoFinal, usuario);

    mensajeOp = `Traslado de ${cantidad} ${unidad}: ${origenReal} (quedan ${stockOrigenFinal}) → ${ubicDestino} (ahora ${stockDestinoFinal})`;
  }

  // ── Registrar en Stock_Movimientos ───────────────────────
  const idMov = _genId('MOV');
  hMov.appendRow([
    idMov,
    new Date(),
    tipo,
    datos.idItem,
    clave,
    descripcion,
    cantidad,
    origenReal,
    tipo === 'SALIDA'   ? (datos.embarcacion || datos.ubicacionDestino || 'USO INTERNO') :
    tipo === 'TRASLADO' ? ubicDestino : origenReal,
    datos.motivo     || '',
    datos.otAsociada || '',
    datos.embarcacion || '',
    usuario
  ]);
  SpreadsheetApp.flush();

  _cacheInvalidar('stock_items');
  _auditoria(
    'STOCK_' + tipo, idMov,
    `${tipo} | ${cantidad} ${unidad} | ${descripcion} | ${origenReal}${tipo==='TRASLADO'?' → '+ubicDestino:''}`
  );

  return {
    ok:                true,
    idMov,
    tipo,
    stockOrigenFinal,
    stockDestinoFinal,
    mensaje:           mensajeOp
  };
}


// ── Buscar ítems en catálogo (para el formulario de movimiento) ─
function buscarItemsCatalogo(texto) {
  _checkPermiso('stock.ver');
  const catalogo = _leerHoja('Stock_Items');
  if (!catalogo || !catalogo.length) return [];
  const q = (texto || '').toLowerCase();
  // Agrupar filas del mismo ítem por CLAVE — hay una fila por ubicación
  const porClave = {};
  catalogo
    .filter(r => r.DESCRIPCION && (
      r.DESCRIPCION.toLowerCase().includes(q) ||
      String(r.CLAVE || '').toLowerCase().includes(q)))
    .forEach(r => {
      const k = String(r.CLAVE || r.ID_ITEM);
      if (!porClave[k]) {
        porClave[k] = {
          id: r.ID_ITEM, clave: r.CLAVE,
          descripcion: r.DESCRIPCION,
          unidad: r.UNIDAD, familia: r.FAMILIA,
          ubicacion: r.UBICACION,
          stockTotal: 0,
          saldoPorUbic: {}
        };
      }
      const st = Number(r.STOCK_ACTUAL || 0);
      porClave[k].stockTotal += st;
      porClave[k].saldoPorUbic[r.UBICACION] = st;
    });
  return Object.values(porClave).slice(0, 20);
}

// ── Historial de movimientos de un ítem ───────────────────────
function getMovimientosStock(idItem) {
  _checkPermiso('stock.ver');
  const data = _leerHoja('Stock_Movimientos');
  if (!data || !data.length) return [];
  return data
    .filter(r => !idItem || r.ID_ITEM === idItem)
    .map(r => ({
      id: r.ID_MOV, fecha: r.FECHA, tipo: r.TIPO,
      idItem: r.ID_ITEM, clave: r.CLAVE, descripcion: r.DESCRIPCION,
      cantidad: Number(r.CANTIDAD || 0),
      origen:   r.UBICACION_ORIGEN,
      destino:  r.UBICACION_DESTINO,
      motivo:   r.MOTIVO, otAsociada: r.OT_ASOCIADA,
      embarcacion: r.EMBARCACION, usuario: r.USUARIO
    })).reverse();
}

// ── Resumen stock para dashboard ─────────────────────────────
function getResumenStock() {
  const items   = getStockItems();
  const alertas = items.filter(i => i.alerta);
  const conStock = items.filter(i => i.stockTotal > 0);
  const sinStock = items.filter(i => i.stockTotal <= 0);
  const porUbic  = {};
  items.forEach(i => { porUbic[i.ubicacion] = (porUbic[i.ubicacion]||0)+1; });
  return {
    totalItems:    items.length,
    conStock:      conStock.length,
    sinStock:      sinStock.length,
    totalAlertas:  alertas.length,
    porUbicacion:  porUbic,
    alertasDetalle: alertas.slice(0, 10)
  };
}


// ════════════════════════════════════════════════════════════
//  MÓDULO COMBUSTIBLE DIARIO POR EQUIPO  v3.1
// ════════════════════════════════════════════════════════════

// ════════════════════════════════════════════════════════════
//  MÓDULO COMBUSTIBLE DIARIO — Backend v3.5
//  Hoja: Combustible_Diario
//  Columnas: ID | FECHA | AÑO | MES | DIA | EMBARCACION |
//    EQUIPO | TIPO_EQUIPO | GALONES | HOROMETRO |
//    ACTIVIDAD | HORAS_NAV | USUARIO | OBSERVACION
// ════════════════════════════════════════════════════════════

const EQUIPOS_COMBUSTIBLE_V2 = {
  'GUANAY': [
    { codigo:'GUA-MP-BR', nombre:'Motor Principal Babor',       tipo:'MOTOR_PRINCIPAL',    kw:750  },
    { codigo:'GUA-MP-ER', nombre:'Motor Principal Estribor',    tipo:'MOTOR_PRINCIPAL',    kw:750  },
    { codigo:'GUA-GE-BR', nombre:'Grupo Electrógeno Babor',     tipo:'GRUPO_ELECTROGENO',  kw:76   },
    { codigo:'GUA-GE-ER', nombre:'Grupo Electrógeno Estribor',  tipo:'GRUPO_ELECTROGENO',  kw:76   },
    { codigo:'GUA-GE-PT', nombre:'Grupo Electrógeno Puerto',    tipo:'GRUPO_ELECTROGENO',  kw:28   }
  ],
  'PELICANO': [
    { codigo:'PEL-MP-BR', nombre:'Motor Principal Babor',       tipo:'MOTOR_PRINCIPAL',    kw:850  },
    { codigo:'PEL-MP-ER', nombre:'Motor Principal Estribor',    tipo:'MOTOR_PRINCIPAL',    kw:850  },
    { codigo:'PEL-GE-BR', nombre:'Grupo Electrógeno Babor',     tipo:'GRUPO_ELECTROGENO',  kw:99   },
    { codigo:'PEL-GE-ER', nombre:'Grupo Electrógeno Estribor',  tipo:'GRUPO_ELECTROGENO',  kw:99   },
    { codigo:'PEL-GE-PT', nombre:'Grupo Electrógeno Puerto',    tipo:'GRUPO_ELECTROGENO',  kw:28   }
  ],
  'DELFIN 11': [
    { codigo:'D11-MP-UN', nombre:'Motor Principal',             tipo:'MOTOR_PRINCIPAL',    kw:365  },
    { codigo:'D11-GE-PT', nombre:'Grupo Electrógeno Puerto',    tipo:'GRUPO_ELECTROGENO',  kw:8    }
  ],
  'DELFIN 12': [
    { codigo:'D12-MP-UN', nombre:'Motor Principal',             tipo:'MOTOR_PRINCIPAL',    kw:300  },
    { codigo:'D12-GE-PT', nombre:'Grupo Electrógeno Puerto',    tipo:'GRUPO_ELECTROGENO',  kw:8    }
  ],
  'ISLA CHINCHA': [
    { codigo:'ICH-GE-BR', nombre:'Grupo Electrógeno Babor',     tipo:'GRUPO_ELECTROGENO',  kw:65   },
    { codigo:'ICH-GE-ER', nombre:'Grupo Electrógeno Estribor',  tipo:'GRUPO_ELECTROGENO',  kw:99   }
  ]
};

const ACTIVIDADES_FLOTA = [
  'FONDEO DE EMBARCACION',
  'OPERACIONES FAENA DE AGUA',
  'OPERACIONES FAENA PETROLEO',
  'OPERACIONES FAENA VIVERES',
  'OPERACIONES MANTENIMIENTO',
  'OPERACIONES TRANSP CARGA',
  'OPERACIONES TRANSP PERSONAL',
  'TRANSPORTE DE GUANO',
  'EMBARQUE Y DESEMBARQUE DE GUANO',
  'PRUEBAS DE MAR',
  'OTRAS OPERACIONES'
];

function _inicializarHojaCombustibleDiario(ss) {
  // ── Hoja 1: Combustible_Diario ─────────────────────────
  // Una fila por equipo por día. El usuario puede editar manualmente.
  // El sistema lee siempre en tiempo real.
  let h1 = ss.getSheetByName('Combustible_Diario');
  if (!h1) {
    h1 = ss.insertSheet('Combustible_Diario');
    const cols1 = ['FECHA','ANO','MES','DIA','EMBARCACION','EQUIPO','TIPO_EQUIPO',
                   'GALONES','HOROMETRO','STOCK_ACTUAL','ACTIVIDAD','OBSERVACION'];
    h1.appendRow(cols1);
    const rHdr1 = h1.getRange(1, 1, 1, cols1.length);
    rHdr1.setFontBold(true).setBackground('#1a3a5c').setFontColor('#ffffff').setFontSize(11);
    h1.setFrozenRows(1);
    h1.setColumnWidth(1, 100); // FECHA
    h1.setColumnWidth(5, 110); // EMBARCACION
    h1.setColumnWidth(6, 200); // EQUIPO
    h1.setColumnWidth(10, 200); // ACTIVIDAD
    // Nota instructiva en fila 2 comentario
    h1.getRange('A2').setNote('Ingresa los datos directamente. FECHA formato: DD/MM/AAAA o AAAA-MM-DD. AÑO, MES y DIA deben ser números.');
  }

  // ── Hoja 2: Actividades_Flota ──────────────────────────
  // Una fila por actividad/embarcacion/mes. Editable manualmente.
  let h2 = ss.getSheetByName('Actividades_Flota');
  if (!h2) {
    h2 = ss.insertSheet('Actividades_Flota');
    const cols2 = ['ANO','MES','MES_NOMBRE','EMBARCACION','ACTIVIDAD',
                   'HORAS_NAV','GALONES','HORAS_ELEC','OBSERVACION'];
    h2.appendRow(cols2);
    const rHdr2 = h2.getRange(1, 1, 1, cols2.length);
    rHdr2.setFontBold(true).setBackground('#0d6efd').setFontColor('#ffffff').setFontSize(11);
    h2.setFrozenRows(1);
    h2.setColumnWidth(4, 110); // EMBARCACION
    h2.setColumnWidth(5, 230); // ACTIVIDAD
    h2.getRange('A2').setNote('Ingresa ANO y MES como números. MES_NOMBRE opcional (Ene, Feb...). ACTIVIDAD: tipo de operación.');
  }
  return h1;
}

// ── Registrar consumo diario por equipo ──────────────────────
function registrarCombustibleDiario(registros) {
  _checkPermiso('combustible.crear');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('Combustible_Diario') ||
             _inicializarHojaCombustibleDiario(ss);
  const filas = registros
    .filter(r => Number(r.galones||0) > 0 || Number(r.horometro||0) > 0 || Number(r.stockActual||0) > 0)
    .map(r => {
      const f = r.fecha ? new Date(r.fecha) : new Date();
      return [
        f,                             // FECHA
        f.getFullYear(),               // ANO
        f.getMonth() + 1,             // MES
        f.getDate(),                  // DIA
        r.embarcacion,                // EMBARCACION
        r.equipoNombre,               // EQUIPO
        r.tipoEquipo,                 // TIPO_EQUIPO
        Number(r.galones    || 0),    // GALONES
        Number(r.horometro  || 0),    // HOROMETRO
        Number(r.stockActual|| 0),    // STOCK_ACTUAL
        r.actividad    || '',         // ACTIVIDAD
        r.observacion  || ''          // OBSERVACION
      ];
    });
  if (!filas.length) return { ok: true, registrados: 0 };
  h.getRange(h.getLastRow()+1, 1, filas.length, 12).setValues(filas);
  SpreadsheetApp.flush();
  _auditoria('COMB_DIA', registros[0]?.embarcacion||'',
    `${filas.length} equipos — ${registros[0]?.embarcacion} — ${registros[0]?.fecha}`);
  return { ok: true, registrados: filas.length };
}

// ── Registrar actividad mensual de flota ─────────────────────
// Hoja separada: Actividades_Flota
// Columnas: ID | FECHA_REG | ANO | MES | EMBARCACION | ACTIVIDAD |
//           HORAS_NAV | GALONES | HORAS_NAV_COMP | HORAS_ELEC | OBSERVACION | USUARIO
function registrarActividadFlota(datos) {
  _checkPermiso('combustible.crear');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h    = ss.getSheetByName('Actividades_Flota');
  if (!h) {
    h = ss.insertSheet('Actividades_Flota');
    const cols = ['ID','FECHA_REG','ANO','MES','EMBARCACION','ACTIVIDAD',
                  'HORAS_NAV','GALONES','HORAS_NAV_COMP','HORAS_ELEC','OBSERVACION','USUARIO'];
    h.appendRow(cols);
    h.getRange(1,1,1,cols.length).setFontBold(true)
      .setBackground('#1a3a5c').setFontColor('#ffffff');
    h.setFrozenRows(1);
  }
  const ano     = Number(datos.ano  || new Date().getFullYear());
  const mes     = Number(datos.mes  || new Date().getMonth() + 1);
  const mesNom  = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'][mes-1]||'';
  const id = _genId('ACT');
  h.appendRow([
    ano, mes, mesNom,
    datos.embarcacion, datos.actividad,
    Number(datos.horasNav  || 0),
    Number(datos.galones   || 0),
    Number(datos.horasElec || 0),
    datos.observacion || ''
  ]);
  SpreadsheetApp.flush();
  _auditoria('ACTIVIDAD_FLOTA', id,
    `${datos.embarcacion} | ${datos.actividad} | ${mes}/${ano}`);
  return { ok: true, id };
}

// ── Obtener datos del dashboard de combustible ───────────────

// ── getCombustibleDiario — lectura en tiempo real ────────────
function getCombustibleDiario(filtros) {
  _checkPermiso('combustible.ver');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('Combustible_Diario');
  if (!h || h.getLastRow() < 2) return [];
  const lastCol = h.getLastColumn();
  const nCols   = Math.max(lastCol, 12);
  const vals = h.getRange(2, 1, h.getLastRow()-1, nCols).getValues();
  let regs = vals
    .filter(r => r[0] || r[4])
    .map(r => ({
      FECHA:        r[0],
      fecha:        r[0],
      ANO:          Number(r[1]||(r[0]?new Date(r[0]).getFullYear():0)),
      MES:          Number(r[2]||(r[0]?new Date(r[0]).getMonth()+1:0)),
      DIA:          Number(r[3]||(r[0]?new Date(r[0]).getDate():0)),
      EMBARCACION:  String(r[4]||'').trim(),
      EQUIPO:       String(r[5]||'').trim(),
      TIPO_EQUIPO:  String(r[6]||'').trim(),
      GALONES:      Number(r[7]||0),
      HOROMETRO:    Number(r[8]||0),
      STOCK_ACTUAL: Number(r[9]||0),   // col 10 nueva
      ACTIVIDAD:    String(r[10]||'').trim(),
      OBSERVACION:  String(r[11]||'').trim()
    }))
    .filter(r => r.EMBARCACION);
  if (filtros?.embarcacion) regs = regs.filter(r => r.EMBARCACION===filtros.embarcacion);
  if (filtros?.ano)         regs = regs.filter(r => r.ANO===Number(filtros.ano));
  if (filtros?.mes)         regs = regs.filter(r => r.MES===Number(filtros.mes));
  return regs.reverse();
}
function getDashboardCombustible(filtros) {
  _checkPermiso('combustible.ver');
  const anoSel = Number(filtros?.ano  || new Date().getFullYear());
  const mesSel = Number(filtros?.mes  || new Date().getMonth() + 1);
  const diaSel = filtros?.dia ? Number(filtros.dia) : null;

  // Leer SIEMPRE en tiempo real para reflejar ediciones manuales en el Sheet
  function leerCombustibleLive() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h  = ss.getSheetByName('Combustible_Diario');
    if (!h || h.getLastRow() < 2) return [];
    const lastColLive = h.getLastColumn();
    const nColsLive   = Math.max(lastColLive, 12);
    const vals = h.getRange(2, 1, h.getLastRow()-1, nColsLive).getValues();
    // Cols: FECHA|ANO|MES|DIA|EMBARCACION|EQUIPO|TIPO_EQUIPO|GALONES|HOROMETRO|STOCK_ACTUAL|ACTIVIDAD|OBSERVACION
    return vals.filter(r => r[0] || r[4]).map(r => ({
      FECHA:        r[0],
      ANO:          Number(r[1] || (r[0] ? new Date(r[0]).getFullYear() : 0)),
      MES:          Number(r[2] || (r[0] ? new Date(r[0]).getMonth()+1  : 0)),
      DIA:          Number(r[3] || (r[0] ? new Date(r[0]).getDate()     : 0)),
      EMBARCACION:  String(r[4] || '').trim(),
      EQUIPO:       String(r[5] || '').trim(),
      TIPO_EQUIPO:  String(r[6] || '').trim(),
      GALONES:      Number(r[7] || 0),
      HOROMETRO:    Number(r[8] || 0),
      STOCK_ACTUAL: Number(r[9] || 0),
      ACTIVIDAD:    String(r[10]|| '').trim(),
      OBSERVACION:  String(r[11]|| '').trim()
    })).filter(r => r.EMBARCACION && r.ANO > 0);
  }

  function leerActividadesLive() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const h  = ss.getSheetByName('Actividades_Flota');
    if (!h || h.getLastRow() < 2) return [];
    const vals = h.getRange(2, 1, h.getLastRow()-1, 9).getValues();
    // Cols: ANO|MES|MES_NOMBRE|EMBARCACION|ACTIVIDAD|HORAS_NAV|GALONES|HORAS_ELEC|OBSERVACION
    return vals.filter(r => r[3]).map(r => ({
      ANO:        Number(r[0] || 0),
      MES:        Number(r[1] || 0),
      MES_NOMBRE: String(r[2] || '').trim(),
      EMBARCACION:String(r[3] || '').trim(),
      ACTIVIDAD:  String(r[4] || '').trim(),
      HORAS_NAV:  Number(r[5] || 0),
      GALONES:    Number(r[6] || 0),
      HORAS_ELEC: Number(r[7] || 0),
      OBSERVACION:String(r[8] || '').trim()
    })).filter(r => r.EMBARCACION && r.ANO > 0);
  }

  const data = leerCombustibleLive();
  const embs = ['GUANAY','PELICANO','DELFIN 11','DELFIN 12','ISLA CHINCHA'];
  const meses = ['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];

  // ── Consumo del día seleccionado ─────────────────────────
  const diaData = diaSel
    ? data.filter(r => Number(r.ANO)===anoSel && Number(r.MES)===mesSel && Number(r.DIA)===diaSel)
    : data.filter(r => {
        const hoy = new Date(); hoy.setHours(0,0,0,0);
        const f   = new Date(r.FECHA); f.setHours(0,0,0,0);
        return f.getTime() === hoy.getTime();
      });

  const consumoDia = {};
  const stockDia   = {}; // último stock registrado del día por embarcación
  embs.forEach(e => {
    consumoDia[e] = { total:0, equipos:[], motorPrincipal:0, generadores:0 };
    stockDia[e]   = 0;
  });
  diaData.forEach(r => {
    const e = r.EMBARCACION;
    if (!consumoDia[e]) consumoDia[e] = { total:0, equipos:[], motorPrincipal:0, generadores:0 };
    const g   = Number(r.GALONES||0);
    const stk = Number(r.STOCK_ACTUAL||0);
    consumoDia[e].total += g;
    consumoDia[e].equipos.push({ eq:r.EQUIPO, tipo:r.TIPO_EQUIPO, gal:g, hor:Number(r.HOROMETRO||0), stk });
    if (r.TIPO_EQUIPO === 'MOTOR_PRINCIPAL')   consumoDia[e].motorPrincipal += g;
    if (r.TIPO_EQUIPO === 'GRUPO_ELECTROGENO') consumoDia[e].generadores    += g;
    if (stk > 0) stockDia[e] = stk; // tomar el último stock registrado
  });

  // ── Consumo del mes seleccionado ─────────────────────────
  const mesData  = data.filter(r => Number(r.ANO)===anoSel && Number(r.MES)===mesSel);
  const consumoMes = {};
  embs.forEach(e => consumoMes[e] = 0);
  mesData.forEach(r => { consumoMes[r.EMBARCACION] = (consumoMes[r.EMBARCACION]||0) + Number(r.GALONES||0); });

  // ── Consumo mensual del año seleccionado (12 meses) ─────
  const porMesAno = {};
  embs.forEach(e => { porMesAno[e] = Array(12).fill(0); });
  data.filter(r => Number(r.ANO)===anoSel).forEach(r => {
    const m = Number(r.MES) - 1;
    if (m >= 0 && m < 12 && porMesAno[r.EMBARCACION])
      porMesAno[r.EMBARCACION][m] += Number(r.GALONES||0);
  });

  // ── Comparativo año anterior ─────────────────────────────
  const anoAnt   = anoSel - 1;
  const porMesAnt = {};
  embs.forEach(e => { porMesAnt[e] = Array(12).fill(0); });
  data.filter(r => Number(r.ANO)===anoAnt).forEach(r => {
    const m = Number(r.MES) - 1;
    if (m >= 0 && m < 12 && porMesAnt[r.EMBARCACION])
      porMesAnt[r.EMBARCACION][m] += Number(r.GALONES||0);
  });

  // ── Total acumulado por embarcación ─────────────────────
  const acumulado = {};
  embs.forEach(e => acumulado[e] = 0);
  data.forEach(r => { acumulado[r.EMBARCACION] = (acumulado[r.EMBARCACION]||0) + Number(r.GALONES||0); });

  // ── Consumo acumulado por equipo ────────────────────────
  const porEquipo = {};
  data.forEach(r => {
    const k = r.EMBARCACION + '||' + r.EQUIPO;
    if (!porEquipo[k]) porEquipo[k] = { emb:r.EMBARCACION, eq:r.EQUIPO, tipo:r.TIPO_EQUIPO, gal:0, hor:0 };
    porEquipo[k].gal += Number(r.GALONES||0);
    porEquipo[k].hor += Number(r.HOROMETRO||0);
  });

  // ── Actividades de flota del mes ────────────────────────
  const actData  = leerActividadesLive();
  const actMes   = actData.filter(r => r.ANO===anoSel && r.MES===mesSel);

  // ── Verificar faltantes del día ─────────────────────────
  const faltantes = embs.filter(e => !consumoDia[e]?.total);

  return {
    ok: true,
    anoSel, mesSel, mesMesLabel: meses[mesSel-1],
    diaSel, meses, embs,
    consumoDia, consumoMes, porMesAno, porMesAnt, acumulado, stockDia,
    porEquipo: Object.values(porEquipo),
    actividadesMes: actMes,
    faltantes,
    totalDia:  Object.values(consumoDia).reduce((s,v)=>s+v.total, 0),
    totalMes:  Object.values(consumoMes).reduce((s,v)=>s+v, 0),
    totalAcum: Object.values(acumulado).reduce((s,v)=>s+v, 0)
  };
}

function getEquiposCombustibleV2(emb) {
  return EQUIPOS_COMBUSTIBLE_V2[emb] || [];
}
function getActividadesFlota() {
  return ACTIVIDADES_FLOTA;
}


// ════════════════════════════════════════════════════════════
//  FORMATOS G — Módulo de Parámetros y Mantenimiento v3.8
//  G-125: Seguimiento semanal al Plan de Mantenimiento
//  G-126: Actividades de Mantenimiento por Equipo
//  G-127: Registro de Parámetros Operativos (existente mejorado)
//  Nuevas hojas: G125_Seguimiento, G126_Actividades
// ════════════════════════════════════════════════════════════

function _inicializarFormatosG(ss) {
  function cab(nombre, cols, color) {
    let h = ss.getSheetByName(nombre);
    if (!h) {
      h = ss.insertSheet(nombre);
      h.appendRow(cols);
      h.getRange(1,1,1,cols.length)
        .setFontBold(true).setBackground(color||'#1a3a5c').setFontColor('#ffffff').setFontSize(10);
      h.setFrozenRows(1);
    }
    return h;
  }

  // G-125: Seguimiento semanal al plan de mantenimiento (una fila por componente)
  cab('G125_Seguimiento', [
    'ID','FECHA','SEMANA','AÑO','EMBARCACION','RESPONSABLE',
    'SISTEMA','EQUIPO','COMPONENTE','HORAS_REF','HORAS_REALIZADAS',
    'ESTADO_COMP','OBSERVACION_COMP',
    'ESTADO_GENERAL','OBS_GENERAL','REGISTRADO_POR'
  ], '#1565c0');

  // G-126: Actividades de mantenimiento por equipo (ficha imagen)
  cab('G126_Actividades', [
    'ID','FECHA','EMBARCACION','MATRICULA','LUGAR',
    'ID_EQUIPO','EQUIPO','MARCA','MODELO_SERIE','UBICACION_EQUIPO',
    'TIPO_MANTENIMIENTO',
    'HOROMETRO_ULT_CAMBIO_MP_ER','HOROMETRO_ACTUAL_MP_ER','HOROMETRO_PROX_MP_ER',
    'HOROMETRO_ULT_CAMBIO_MP_BR','HOROMETRO_ACTUAL_MP_BR','HOROMETRO_PROX_MP_BR',
    'HOROMETRO_ULT_CAMBIO_MP','HOROMETRO_ACTUAL_MP','HOROMETRO_PROX_MP',
    'HOROMETRO_ULT_CAMBIO_GE_ER','HOROMETRO_ACTUAL_GE_ER','HOROMETRO_PROX_GE_ER',
    'HOROMETRO_ULT_CAMBIO_GE_BR','HOROMETRO_ACTUAL_GE_BR','HOROMETRO_PROX_GE_BR',
    'HOROMETRO_ULT_CAMBIO_GE_AUX','HOROMETRO_ACTUAL_GE_AUX','HOROMETRO_PROX_GE_AUX',
    'OBSERVACION','REGISTRADO_POR'
  ], '#2e7d32');
}

// ════════════════════════════════════════════════════════════
//  G-125: Guardar seguimiento semanal
// ════════════════════════════════════════════════════════════
function guardarG125(datos) {
  _checkPermiso('parametros.crear');
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const h   = ss.getSheetByName('G125_Seguimiento') ||
              (() => { _inicializarFormatosG(ss); return ss.getSheetByName('G125_Seguimiento'); })();
  const usr = getUsuarioActual().email;
  const now = new Date();
  const fechaReg = datos.fecha ? new Date(datos.fecha) : now;

  // Calcular semana del año
  const start  = new Date(now.getFullYear(), 0, 1);
  const semana = Math.ceil(((now - start) / 86400000 + start.getDay() + 1) / 7);
  const idBase = _genId('G125');

  // Nuevo formato: guarda múltiples filas (una por componente)
  if (datos.registros && datos.registros.length > 0) {
    const filas = datos.registros.map((r, i) => [
      idBase + '-' + String(i+1).padStart(3,'0'), // ID único por fila
      fechaReg,
      semana,
      fechaReg.getFullYear(),
      datos.embarcacion,
      datos.responsable || '',
      r.sistema         || '',
      r.equipo          || '',
      r.componente      || '',
      r.horasRef        || 0,
      Number(r.horasRealizadas || 0),
      r.estado          || '',
      r.observacion     || '',
      datos.estadoGeneral       || '',
      datos.observacionGeneral  || '',
      usr
    ]);
    if (filas.length > 0) {
      h.getRange(h.getLastRow()+1, 1, filas.length, filas[0].length).setValues(filas);
      SpreadsheetApp.flush();
    }
    _auditoria('G125_GUARDADO', idBase,
      `${datos.embarcacion} | Semana ${semana} | ${filas.length} componentes`);
    return { ok: true, id: idBase, registros: filas.length,
             mensaje: `G-125 guardado: ${filas.length} componentes registrados` };
  }

  // Formato legacy (compatibilidad hacia atrás)
  h.appendRow([
    idBase, fechaReg, semana, fechaReg.getFullYear(),
    datos.embarcacion, datos.sistema||'', datos.equipo||'',
    datos.componente||'', datos.horasRef||0, datos.horasAplicadas||0,
    datos.estado||'', datos.observacion||'',
    datos.estadoGeneral||'', datos.observacionGeneral||'', usr
  ]);
  SpreadsheetApp.flush();
  _auditoria('G125_GUARDADO', idBase, `${datos.embarcacion} | ${datos.equipo}`);
  return { ok: true, id: idBase, mensaje: 'G-125 registrado correctamente' };
}

function getG125(filtros) {
  _checkPermiso('parametros.ver');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('G125_Seguimiento');
  if (!h || h.getLastRow() < 2) return [];
  const lastCol = Math.max(h.getLastColumn(), 15);
  const vals = h.getRange(2, 1, h.getLastRow()-1, lastCol).getValues();
  // Cols: ID|FECHA|SEMANA|AÑO|EMBARCACION|SISTEMA|EQUIPO|COMPONENTE|
  //       HORAS_REF|HORAS_REAL|ESTADO_COMP|OBS_COMP|ESTADO_GEN|OBS_GEN|USUARIO
  let rows = vals.filter(r => r[0]).map(r => ({
    ID: r[0], FECHA: r[1], SEMANA: r[2], ANO: r[3], EMBARCACION: r[4],
    SISTEMA: r[5], EQUIPO: r[6], COMPONENTE: r[7],
    HORAS_REF: r[8], HORAS_REAL: r[9],
    ESTADO_COMP: r[10], OBS_COMP: r[11],
    ESTADO: r[12], OBS_GENERAL: r[13], USUARIO: r[14]
  }));
  if (filtros?.embarcacion) rows = rows.filter(r => r.EMBARCACION === filtros.embarcacion);
  if (filtros?.ano)         rows = rows.filter(r => Number(r.ANO) === Number(filtros.ano));
  return rows.reverse();
}

// ════════════════════════════════════════════════════════════
//  G-126: Guardar ficha de actividades (imagen adjunta)
// ════════════════════════════════════════════════════════════
function guardarG126(datos) {
  _checkPermiso('parametros.crear');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('G126_Actividades') ||
             (() => { _inicializarFormatosG(ss); return ss.getSheetByName('G126_Actividades'); })();
  const usr= getUsuarioActual().email;
  const id = _genId('G126');

  // Obtener matrícula de la embarcación
  const embs = getEmbarcaciones();
  const emb  = embs.find(e => e.NOMBRE === datos.embarcacion) || {};

  h.appendRow([
    id,
    datos.fecha ? new Date(datos.fecha) : new Date(),
    datos.embarcacion,
    emb.CODIGO || '',
    datos.lugar || '',
    datos.idEquipo || '',
    datos.equipo   || '',
    datos.marca    || '',
    datos.modeloSerie || '',
    datos.ubicacionEquipo || '',
    datos.tipoMant || '',
    // M/P ER
    Number(datos.mpEr?.ultCambio || 0),
    Number(datos.mpEr?.actual    || 0),
    Number(datos.mpEr?.proxCambio|| 0),
    // M/P BR
    Number(datos.mpBr?.ultCambio || 0),
    Number(datos.mpBr?.actual    || 0),
    Number(datos.mpBr?.proxCambio|| 0),
    // M/P (único)
    Number(datos.mp?.ultCambio   || 0),
    Number(datos.mp?.actual      || 0),
    Number(datos.mp?.proxCambio  || 0),
    // GEN ER
    Number(datos.genEr?.ultCambio || 0),
    Number(datos.genEr?.actual    || 0),
    Number(datos.genEr?.proxCambio|| 0),
    // GEN BR
    Number(datos.genBr?.ultCambio || 0),
    Number(datos.genBr?.actual    || 0),
    Number(datos.genBr?.proxCambio|| 0),
    // GEN AUX
    Number(datos.genAux?.ultCambio || 0),
    Number(datos.genAux?.actual    || 0),
    Number(datos.genAux?.proxCambio|| 0),
    datos.observacion || '',
    usr
  ]);
  SpreadsheetApp.flush();
  _auditoria('G126_GUARDADO', id, `${datos.embarcacion} | ${datos.equipo} | ${datos.tipoMant}`);
  return { ok: true, id, mensaje: 'G-126 registrado correctamente' };
}

function getG126(filtros) {
  _checkPermiso('parametros.ver');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('G126_Actividades');
  if (!h || h.getLastRow() < 2) return [];
  const vals = h.getRange(2, 1, h.getLastRow()-1, 31).getValues();
  let rows = vals.filter(r => r[0]).map(r => ({
    id: r[0], fecha: r[1], embarcacion: r[2], matricula: r[3], lugar: r[4],
    idEquipo: r[5], equipo: r[6], marca: r[7], modeloSerie: r[8],
    ubicacion: r[9], tipoMant: r[10],
    mpEr:  { ult: r[11], act: r[12], prox: r[13] },
    mpBr:  { ult: r[14], act: r[15], prox: r[16] },
    mp:    { ult: r[17], act: r[18], prox: r[19] },
    genEr: { ult: r[20], act: r[21], prox: r[22] },
    genBr: { ult: r[23], act: r[24], prox: r[25] },
    genAux:{ ult: r[26], act: r[27], prox: r[28] },
    obs: r[29], usuario: r[30]
  }));
  if (filtros?.embarcacion) rows = rows.filter(r => r.embarcacion === filtros.embarcacion);
  return rows.reverse();
}

// ════════════════════════════════════════════════════════════
//  Datos auxiliares para los formularios
// ════════════════════════════════════════════════════════════
function getDatosFormulariosMant() {
  // Una sola llamada que devuelve todo lo necesario para los formularios
  return {
    embarcaciones: getEmbarcaciones(),
    equipos:       getEquipos(null),
    sistemas:      getSistemas(),
    acciones:      getCatalogoAcciones(),
    planes:        getPlanesMant()
  };
}


function _getStockItemsData() {
  // Fuente: STOCK__1_.xlsx — 909 ítems reales de Agro Rural
  // Columnas: [CLAVE, DESCRIPCION, EXISTENCIA_INICIAL, UNIDAD, UBICACION, FAMILIA]
  return [
    ["TPCH-001","ACEITE H68 CILINDRO X 55 GLNS - VISTONY (FLOTA)",1.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-002","ACEITE SAE 15W40 CILINDRO X 55 GLNS (FLOTA)",1.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-003","CABLE DE ACERO NEGRO 1/2' - ROLLO  (FLOTA)",5.0,"ROLLO","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-004","CABO DE NYLON TORCIDO 1' (FLOTA) ROLLO X 220M",1.0,"ROLLO","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-005","CABO DE NYLON TORCIDO 1  1/4' (FLOTA) ROLLO X 220M",1.0,"ROLLO","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-006","CABO DE NYLON TORCIDO 1  1/2' (FLOTA) ROLLO X 220M",0.0,"ROLLO","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-007","CABO DE NYLON TORCIDO 3/4' (FLOTA) ROLLO X 220M",1.0,"ROLLO","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-008","CABO DE NYLON TORCIDO 2' (FLOTA) ROLLO X 220M",3.0,"ROLLO","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-009","REFRIGERANTE ANTICONGELANTE VISTONY BALDE 5 GLN (FLOTA)",3.0,"BALDE","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["SIN-CLAVE-010","ACEITE LUBRICANTES MULTGRADO 15W40 X 55 GL CI-4 - PLUS - FENIX",7.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["SIN-CLAVE-011","ACEITE LUBRICANTES MULTGRADO 15W40 X 05 GL CK4 - EURO 9- FENIX",14.0,"BALDE","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-010","ACEITE LUBRICANTE MULTIGRADO SAE 15W-40 PARA MOTOR PETROLERO x 55 GAL. (FLOTA)",6.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-011","ACEITE LUBRICANTE SAE 40 MARCA LUBEXOL (FLOTA)",1.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-012","ACEITE DE SISTEMA HIDRAULICO 68 CST MARCA DC LUBE (FLOTA)",4.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-013","GRASA MULTIPROPOSITO NL GL2 MARCA LUBEXOL (FLOTA)",3.0,"BALDE","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["SIN-CLAVE-016","GRASA MULTIPROPOSITO NL GL2 MARCA VISTONY",6.0,"BALDE","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["SIN-CLAVE-017","EDREDON DE 1 PLAZA",34.0,"UNIDAD","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["SIN-CLAVE-018","SABANAS DE 1 PLAZA X 3 PIEZAS",102.0,"JUEGO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-014","LIQUIDO REFRIGERANTE X GLN MARCA DECLUBE (FLOTA)",72.0,"UNIDAD","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-015","BOTAS DE SEGURIDAD TALLA 38 - MARCA SANDDER (FLOTA)",0.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-016","BOTAS DE SEGURIDAD TALLA 39 - MARCA SANDDER (FLOTA)",0.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-017","BOTAS DE SEGURIDAD TALLA 40 - MARCA SANDDER (FLOTA)",2.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-018","BOTAS DE SEGURIDAD TALLA 41 - MARCA SANDDER (FLOTA)",14.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-019","BOTAS DE SEGURIDAD TALLA 42 - MARCA SANDDER (FLOTA)",1.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-020","BOTAS DE SEGURIDAD TALLA 43 - MARCA SANDDER (FLOTA)",7.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-021","BOTAS DE SEGURIDAD TALLA 44 - MARCA SANDDER (FLOTA)",1.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-022","CABLLE DE ACERO 1/2 PULG TIPO: LOMO DE CORVINA FLEXIBLE",100.0,"METROS","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-023","GANCHO DE ACERO  TIPO OCHO GIRATORIO DE 15 TON",2.0,"UNIDAD","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-024","GORRO DE LANA",2.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-025","GORRO LEGIONARIO",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-026","CAMISA DE TRABAJO MANGA LARGA CON CINTA REFLECTIVA TALLA 'M'",11.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-027","CAMISA DE TRABAJO MANGA LARGA CON CINTA REFLECTIVA TALLA 'L'",19.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-028","CAMISA DE TRABAJO MANGA LARGA CON CINTA REFLECTIVA TALLA 'XL'",4.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-029","PANTALON DRILL INDUSTRIAL CON CINTA REFLECTIVA TALLA 'M'",8.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-030","PANTALON DRILL INDUSTRIAL CON CINTA REFLECTIVA TALLA 'L'",19.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-031","PANTALON DRILL INDUSTRIAL CON CINTA REFLECTIVA TALLA 'XL'",5.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-032","POLO MANGA LARGA CON CUELLO CAMICERO Y CINTA REFELCTIVA TALLA 'M'",7.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-033","POLO MANGA LARGA CON CUELLO CAMICERO Y CINTA REFELCTIVA TALLA 'L'",2.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-034","POLO MANGA LARGA CON CUELLO CAMICERO Y CINTA REFELCTIVA TALLA 'XL'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-035","POLO MANGA CORTA CON CUELLO REDONDO TALLA 'M'",0.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-036","POLO MANGA CORTA CON CUELLO REDONDO TALLA 'L'",0.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-037","POLO MANGA CORTA CON CUELLO REDONDO TALLA 'XL'",0.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-038","CASACA IMPERMEABLE TALLA 'M'",1.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-039","CASACA IMPERMEABLE TALLA 'L'",2.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-040","CASACA IMPERMEABLE TALLA 'XL'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-041","CHOMPA DE LANA TALLA 'M'",1.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-042","CHOMPA DE LANA TALLA 'L'",20.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-043","CHOMPA DE LANA TALLA 'XL'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-044","CALENTADOR TIPO PANTALON TALLA 'ESTÁNDAR'",2.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-045","CALENTADOR TIPO POLERA TALLA 'ESTÁNDAR'",30.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-046","CHALECO DRILL CON CINTA REFLECTIVA TALLA 'M'",2.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-047","CHALECO DRILL CON CINTA REFLECTIVA TALLA 'L'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-048","CHALECO DRILL CON CINTA REFLECTIVA TALLA 'XL'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-049","CHALECO DRILL TIPO PERIODISTA TALLA 'M'",7.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-050","CHALECO DRILL TIPO PERIODISTA TALLA 'L'",20.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-051","CHALECO DRILL TIPO PERIODISTA TALLA 'XL'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-052","TRAJE DE PVC AMARILLO PARA LLUVIA CON CAPUCHA + OVERALL TALLA 'M'",7.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-053","TRAJE DE PVC AMARILLO PARA LLUVIA CON CAPUCHA + OVERALL TALLA 'L'",20.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-054","TRAJE DE PVC AMARILLO PARA LLUVIA CON CAPUCHA + OVERALL TALLA 'XL'",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-055","ALICATE DE CORTE 6 ¡n",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-056","ALICATE DE PUNTA LARGA 8 in",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-057","ALICATE PELACABLE 6 In",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-058","ALICATE UNIVERSAL 8 in",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-059","ARCO DESIERRA DE 12 In",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-060","CABLE DE BATERIA EXTERNA",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-061","CABLE VULCANIZADO 2X12 AWG X 100 m",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-062","CABLE VULCANIZADO 2X14 AWG",0.0,"METRO","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-063","CABLE VULCANIZADO 3X12 AWG",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-064","CABLE VULCANIZADO 3 X 14 AWG",0.0,"METRO","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-065","CINTA AISLANTE 19 mm X 18 m COLOR ROJO",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-066","CINTA AISLANTE VULCANIZANTE 19 mm X 15 m",15.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-067","CINTILLO DE SEGURIDAD DE PLASTICO X 100",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-068","DISCO DE DESBASTE 4 1/2 in X 1/4 in X 7/8 in",0.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-069","DISOLVENTE DE PINTURA EPOXICA",0.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-070","ENCHUFE INDUSTRIAL",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-071","FOCO DE SEÑALIZACION 6 V",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-072","FUSIBLE DE VIDRIO 10 A",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-073","HOJA DE SIERRA 12 in 18 DIENTES POR PULGADA",5.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-074","INTERRUPTOR CONMUTADOR SELECTOR 4 X 20 A",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-075","JUEGO DE BORNES DE COBRE DE BATERIA DE 12 V POSITIVO Y NEGATIVO",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-076","JUEGO DE BROCAS DE ACERO HSS X 27 PIEZAS",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-077","LED RGB",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-078","LINTERNA DE MANO CON LUZ LED RECARGABLE",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-079","LLAVE SACA FILTRO X 150 mm",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-080","LLAVE SACA FILTRO X 255 mm",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-081","LUMINARIA CON FOCO AHORRADOR 18 W",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-082","LUMINARIA CON LÁMPARA FLUORESCENTE TUBULAR LINEAL LED 18 W",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-083","LUMINARIA CON LUZ LED 18 W PARA ADOSAR",0.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-084","MASILLA MOLDEABLE EPÓXICA X 32 g",0.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-085","PASTA DETECTORA DE AGUA EN COMBUSTIBLE PARA VEHÍCULO DE TRANSPORTE AEREO",0.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-086","PILA ALCALINA RECARGABLE TIPO D",36.0,"UNIDAD","ALMACEN CHIMBOTE","MAT. ELECTRICIDAD"],
    ["TPCH-087","PINTURA EPOXICA",4.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-088","PULSADOR ELECTRICO 4 A 220 V",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["SIN-CLAVE-094","GRILLETE DE ACERO 7/8 in CON PASADOR DE SEGURIDAD",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-095","GRILLETE DE ACERO GIRATORIO DE 5/8 in",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-096","GRILLETE DE ACERO TIPO LIRA 1 in",30.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-097","GRILLETE DE ACERO TIPO LIRA 1/2 in",6.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-098","GRILLETE DE ACERO TIPO LIRA 5/8 in",3.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-099","GRILLETE DE ACERO TIPO LIRA CON PIN ROSCADO 1 1/2 in",12.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-100","GRILLETE DE FIERRO GALVANIZADO DE 3/4 in",30.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-101","MANGUERA DE NITRILO DE 2 1/2 in X 30 m",5.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-102","MANGUERA DE PVC TRANSPARENTE 1/2 in X 10 m",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-103","FOCO DE 220 V CON ROSCA E-14",60.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-104","FOCO DE SEÑALIZACION 6 V",24.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-105","FUSIBLE DE VIDRIO 10 A",20.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-106","INTERRRUPTOR 2 X 20 A",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-107","INTERRUPTOR MAGNETICO TIPO RIEL 2 X15 A",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-108","INTERRRUPTOR TERMOMAGNETICO BIPOLAR TIPO RIEL 10 A CURVA C4.5 KA",1.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-109","INTERRRUPTOR TERMOMAGNETICO BIPOLAR TIPO RIEL 16 A CURVA C4.5 KA",1.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-110","INTERRRUPTOR TERMOMAGNETICO BIPOLAR TIPO RIEL 6 A CHINT NXB-63",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-111","INTERRRUPTOR TERMOMAGNETICO BIPOLAR TIPO RIEL 6 A CURVA C4.5 KA",1.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-112","INTERRRUPTOR TERMOMAGNETICO RIEL DIN DE POLOS INT N 16A CAP CORTO CIRCUITO DE 1000 A TENSION AISLAMIENTO 500 TENSION 220 V",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-113","JUEGO DE BORNES DE COBRE DE BATERIA DE 12 V POSITIVO Y NEGATIVO",10.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-114","LAMPARA BULBO 5W 24 V",55.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-115","LAMPARA LED BULBO CLASICO 24W 220 V",80.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-116","LAMPARA LED REFLECTORA 18 W 24 VOLT",4.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-117","PULSADOR ELECTRICO 4 A 220 V",18.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-118","RELAY 24 VOL",4.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-119","RELE 230 V 50/60 HZ 10A 250 VAC",12.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-120","RELE DE 20 A BOSH",1.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-121","RELE DE 20 A SCHNEIDER",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-122","RELE DE 5 A",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-123","FILTRO DE AIRE CODIGO -P181099 - PARA GUANAY",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-124","FILTRO DE ACEITE - CODIGO LF3478 PARA GUANAY",14.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-125","FILTRO DE REFRIGERACION COD WF2076 - GUANAY",12.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-126","FILTRO DE AIRE CODIGO H19037 PARA ISLA CHINCHA",2.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["SIN-CLAVE-127","FILTRO DE ACEITE CODIGO LF16015 - GUANAY",4.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-089","RODILLO PARA PINTAR DE 4 in",9.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-090","SILICONA INCOLORA X 320 mL",1.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-091","VÁLVULA CHECK DE BRONCE 1 in",0.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-092","VÁLVULA CHECK DE BRONCE 1/2 in",0.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-093","VÁLVULA CHECK DE BRONCE 3/4 in",0.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-095","WINCHA DE METAL 5 m",0.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["TPCH-096","LIMPIADOR DE INODORO QUITA SARRO",50.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-097","LIMPIADOR MULTISUPERFICIES ANTIBACTERIAL BIODAGRADABE (DESINFECTANTE)",4.0,"GALON","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-098","ESCOBILLA PARA INODORO (HISOPO)",33.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-099","ESCOBILLÓN DE CERDA GRUESA PARA BARRIDO",9.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-100","LEJÍA CONCENTRADA X 1LT",14.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-139","KIT ANTIDERRAMES GRANDES (CILINDRO PLASTICO COLOR AZUL)CON SU RESPECTIVO MALETIN",3.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["SIN-CLAVE-140","KIT ANTIDERRAMES MEDIANOS(CILINDRO PLASTICO COLOR AZUL)",2.0,"UNIDAD","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["SIN-CLAVE-141","MANGUERA DE NITRILO DE 4 in X 30 m - VALVULA INGRESO - SALIDA",1.0,"ROLLO","ALMACEN CHIMBOTE","HERRAMIENTAS"],
    ["SIN-CLAVE-142","AMBIENTADOR LIQUIDO X 900 ML",50.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-143","DESINFECTANTE LIMPIADOR AROMATICO X 1 LT",50.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-144","DETERGENTE GRANULADO INDUSTRIAL X 15 KG",55.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-145","DETERGENTE GRANULADO X 250 G",341.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-146","DETERGENTE LIQUIDO LAVAVAJILLA X 500 ML",110.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-147","ESPONJA DE FIBRA SINTETICA",110.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-148","LEJIA AL 10 % X 1 LT",50.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-149","PASTA DENTIFRICA X 90 G",341.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["SIN-CLAVE-150","SERVILLETAS DE PAPEL DE DOBLE HOJA X 1000",20.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-101","JUEGO DE TRAPEADOR (CABEZAS DE MOPA + PALO)",45.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-102","RECOGEDOR DE PLASTICO",29.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-103","BALDES DE PLASTICO PARA 10 L",15.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-104","ESCOBA DE CERDA PLÁSTICA 40 CM",8.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-105","DESODORANTE DE AMBIENTE",24.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-106","ACIDO MURIATICO X 4L",3.0,"GALON","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-107","ESPONJA DE FIBRA DE METAL",116.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-108","PASTA DENTAL",83.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-109","JABON DE TOCADOR",106.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-110","ROLLO PAPEL TOALLA",57.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-111","ROLLOS PAPEL HIGIENICO",148.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-112","SERVILLETAS DE PAPEL DE DOBLE HOJA X 500",4.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-113","CUCHARAS PARA SOPA ACERO INOX",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-114","CUCHARAS PARA TE ACERO INOX",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-115","TENEDORES PARA MESA  ACERO INOX.",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-116","CUCHILLOS PARA MESA ACERO INOX.",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-117","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-118","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-119","CUCHARON PARA SOPA DE ACERO INOXIDABLE",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-120","TETERA DE ACERO INOXIDABLE 3.2 LT.",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-121","SET DE OLLA DE ACERO INOXIDABLE  6 PZ.",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-122","ESPUMADERA DE COCINA",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-123","CUCHARON DE MADERA",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-124","TABLA DE PICAR DE POLIPROPILENO DE 1/2 ' DE ESPESOR",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-125","PELADORES",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-126","JARRAS DE PLASTICO DE 4 LT",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-127","PINZA DE ACERO INOXIDABLE",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-128","CUCHARÓN DE ESPAGUETTI",0.0,"UNIDAD","ALMACEN CHIMBOTE","MENAJE"],
    ["TPCH-129","GUANTES DIELECTRICOS",35.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-130","GUANTES DE SEGURIDAD DE CUERO BADANA",35.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-131","CASCO DE SEGURIDAD BLANCO",10.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-132","CASCO DE SEGURIDAD ANARANJADO",10.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-133","OREJERAS PARA CASCO",4.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-134","OREJERAS (TAPONES ACÚSTICOS)",320.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-135","LENTES DE SEGURIDAD TRANSPARENTES",30.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-136","SOBRELENTES DE SEGURIDAD TRANSPARENTES",5.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-137","LENTES DE SEGURIDAD OSCUROS",30.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-138","SOBRELENTES DE SEGURIDAD OSCUROS",5.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-139","MASCARILLAS x 50",5.0,"CAJA","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-140","TRAPO INDUSTRIAL COSIDO DE COLORES",40.0,"KG","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-141","TRAPEADOR DE HILO",5.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-142","ATOMIZADOR (SIMPLE)  X 1 LITRO",13.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-143","ESCOBILLON DE CERDA DE PLASTICO 40 CM -",8.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-144","PAÑOS ABSORVENTES 38 X 38 cm",50.0,"UNIDAD","ALMACEN CHIMBOTE","ASEO Y LIMPIEZA"],
    ["TPCH-145","BOTAS DE JEBE CAÑA ALTA  TALLA - 41",7.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-146","BOTAS DE JEBE CAÑA ALTA  TALLA - 39",3.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-147","BOTAS DE JEBE CAÑA ALTA  TALLA - 40",40.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-148","CARETA FACIAL COMPLETO, TRANSPARENTE (ANSI Z87.1)",5.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-149","CASCO DE SEGURIDAD COLOR BLANCO",10.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-150","CASCO DE SEGURIDAD COLOR NARANJA",10.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-151","CASCO DE SEGURIDAD COLOR AMARILLO",1.0,"UNIDAD","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-152","BOTAS DE SEGURIDAD TALLA 43 - MARCA BATA (FLOTA)",1.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-153","BOTAS DE SEGURIDAD TALLA 39 - MARCA SAKERHET (FLOTA)",1.0,"PAR","ALMACEN CHIMBOTE","EPPS"],
    ["TPCH-154","CABO DE NYLON DE 1/2'",440.0,"METROS","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-155","CABO DE NYLON DE 3/4'",220.0,"METROS","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-156","GUARDACABO TIPO H DE 1/2 PULG",14.0,"UNIDAD","ALMACEN CHIMBOTE","FERRETERIA NAVAL"],
    ["TPCH-157","ACEITE SAE 40 API CF-4  CILINDRO X 55 GALONES",1.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-158","ACEITE SAE 68 - HIDRAULICO AW68 - CILINDRO X 55 GALONES",1.0,"CILINDRO","ALMACEN CHIMBOTE","GRASAS Y LUBRICANTES"],
    ["TPCH-159","ESCOBILLA DE ALAMBRE 4' PARA AMOLADORA TIPO COPA",36.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-160","LIJA NRO 50",22.0,"UNIDAD","ALMACEN CHIMBOTE","MANTENIMIENTO"],
    ["TPCH-161","PINTURA EPOXICA BONN EPOXI 720 MP NEGRO X 1 GALON",6.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-162","PINTURA EPOXICA BONN MASTIC 850 BLANCO RAL 9016",5.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-163","CATALIZADOR BONN MASTIC 850 (ENDURECEDOR)",11.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-164","PINTUTA EPOXICA BONN MASTIC 850 GRIS RAL 7040",1.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-165","CATALIZADOR BONN MASTIC 720MP (ENDURECEDOR)",7.0,"UNIDAD","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-166","PINTURA ANTICORROSIVA JET 62ZP X 1 GALON",4.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-167","CATALIZADOR PARA PINTURA ANTICORROSIVA JET 62ZP X 1 GALON",5.0,"GALON","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-168","CATALIZADOR PARA PINTURA ANTICORROSIVA JETDURAPOX X 1/4 GALON",5.0,"UNIDAD","ALMACEN CHIMBOTE","PINTURAS"],
    ["TPCH-169","JUEGO DE SEGURIDAD PARA RESP. DERRAMES DE HIDRO. HASTA 150 LT",3.0,"UNIDAD","ALMACEN CHIMBOTE","SEGURIDAD Y SALVAMENTO"],
    ["TPCH-170","JUEGO DE SEGURIDAD PARA RESP. DERRAMES DE HIDRO. HASTA 57 LT",3.0,"UNIDAD","ALMACEN CHIMBOTE","SEGURIDAD Y SALVAMENTO"],
    ["TPCH-171","JUEGO DE SEGURIDAD PARA RESP. DERRAMES DE HIDRO. HASTA 70 LT",2.0,"UNIDAD","ALMACEN CHIMBOTE","SEGURIDAD Y SALVAMENTO"],
    ["TPCH-172","PANTALON Y CAMISA INDUSTRIAL AZUL MARINO TALLA - M",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-173","PANTALON Y CAMISA INDUSTRIAL AZUL MARINO TALLA - S",3.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-174","PANTALON Y CAMISA INDUSTRIAL AZUL MARINO TALLA - XL",1.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-175","PANTALON Y CAMISA INDUSTRIAL NARANJA TALLA - L",8.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-176","PANTALON Y CAMISA INDUSTRIAL NARANJA TALLA - XXL",4.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-177","PANTALON Y CAMISA INDUSTRIAL NARANJA TALLA - XXXL",2.0,"UNIDAD","ALMACEN CHIMBOTE","VESTUARIO"],
    ["TPCH-178","PORTA CLIPS",3.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-179","PERFORADOR",0.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-180","CUADERNOS ANILLADOS TAMAÑO A4",37.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-181","PIONES BLANCO TAMAÑO A4",-2.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-182","CINTA DE EMBALAJE",6.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-183","LAPICERO COLOR NEGRO",17.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-184","LAPICERO COLOR AZUL",17.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-185","LAPICERO COLOR ROJO",17.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-186","PORTAFOLIOS GRANDES",9.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-187","HOJA BOND TAMAÑO A4 x 500 HOJAS",5.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-188","ENGRAPADORES",-1.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-189","GRAPAS x 5000 UNIDADES",2.0,"CAJITAS","ALMACEN CHIMBOTE","OTROS"],
    ["TPCH-190","REGLA DE 30 CM",1.0,"UNIDAD","ALMACEN CHIMBOTE","OTROS"],
    ["BZA-001","CASCO DE SEGURIDAD COLOR BLANCO (MARCA: SEGPRO)",3.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-002","CASCO DE SEGURIDAD COLOR NARANJA (MARCA: SEGPRO)",0.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-003","ARNES DE SEGURIDAD CON LINEA DE VIDA (NUEVO)",1.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-004","ARNES DE SEGURIDAD CON LINEA DE VIDA (USADO)",2.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-005","MASCARA DE SOLDAR",1.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-006","LENTES DE SEGURIDAD TRASNPARENTES (MARCA:LIBUS)",11.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-007","TAPONES DE OIDOS",22.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-008","BOTA DE JEBE TALLA: 40",1.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-009","BOTA DE JEBE TALLA: 41",3.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-010","BOTA DE JEBE TALLA: 42",2.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-011","BOTA DE JEBE TALLA: 43",0.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-012","GUANTES DE CUERO CROMO",12.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-013","GUANTES ANTICORTE CON PALMA RECUBIERTA",19.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-014","GUNATES RECUBIERTO DE NITRILO AZUL",14.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-015","GUANTES DE CUERO BADANA",19.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-016","GUANTES PARA SOLDAR",2.0,"PAR","ISLA CHINCHA","EPPS"],
    ["BZA-017","GANCHO TIPO PELICANO DE 5.3 TON",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-018","OCHO GIRATORIO DE 5/8 PULG",4.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-019","OCHO GIRATORIO DE 1 PULG",4.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-020","ESLABON DE 1 PULG (TRAMOS)",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-021","ESLABON DE 3/4 PULG (TRAMOS)",0.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-022","ESLABON DE 1/2 PULG (TRAMOS)",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-023","GUARDACABO TIPO H DE 1/2 PULG",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-024","GUARDACABO TIPO H DE 3/4 PULG",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-025","GUARDACABO TIPO H DE 1 1/2 PULG",1.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-026","GRAPA DE 1/2 PULG",4.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-027","GRAPA DE 5/8 PULG",9.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-028","TEMPLADORES DE 1/4 PULG",37.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-029","GRILLETE TIPO LIRA DE 1/4 PULG",25.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-030","GRILLETE TIPO LIRA DE 1 1/2 PULG",8.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-031","GRILLETE TIPO LIRA DE 1 1/4 PULG",4.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-032","GRILLETE TIPO LIRA DE 7/8 PULG",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-033","GRILLETE TIPO LIRA DE 1 PULG",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-034","GRILLETE TIPO LIRA DE 3/4 PULG",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-035","GRILLETE TIPO LIRA DE 1/2 PULG",3.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-036","GRILLETE TIPO LIRA DE 3/8 PULG",2.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-037","GRILLETE TIPO U DE 1 PULG",1.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-038","CABO DE NYLON DE 3/4 PULG",0.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-039","CABO DE NYLON DE 1/2 PULG",440.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-040","CABO DE NYLON DE 2 PULG",220.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-041","CABO DE NYLON DE 1 1/4 PULG",440.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-042","CABO DE NYLON DE 1 1/2 PULG",0.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-043","MOTONES DE FIERRO FUNDIDO",0.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-044","CABLE DE ACERO DE 3/4 PULG",0.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-045","BROCHA DE 2 1/2 PULG",11.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-046","BROCHA DE 2 PULG",20.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-047","BROCHA DE 1 PULG",5.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-048","SARTEN DE TEFLON N°26",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-049","BOLOS DE ACERO",1.0,"JUEGO","ISLA CHINCHA","MENAJE"],
    ["BZA-050","COLADOR DE PLASTICO DE 30 CM",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-051","SARTEN DE ALUMINIO GRANDE N°30",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-052","SARTEN DE ALUMINIO GRANDE N°24",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-053","AZUCARERA DE ACERO",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-054","ESPUMADERA DE COCINA",3.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-055","ESPATULA CON MANGO DE ACERO",3.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-056","CUCHARON DE MADERA",0.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-057","CUCHARÓN DE ESPAGUETTI",5.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-058","PELADOR DE PAPA",2.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-059","EXPRIMIDOR",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-060","PINZA DE ACERO INOXIDABLE",4.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-061","TENEDORES PARA MESA  ACERO INOX.",14.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-062","CUCHILLOS PARA MESA ACERO INOX.",14.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-063","CUCHARAS PARA TE ACERO INOX",14.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-064","RAYADOR DE ACERO",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-065","OLLA A PRESION DE 3 L",0.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-066","FUENTE DE ACERO DE 50CMx35CM",2.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-067","TETERA DE ACERO INOXIDABLE 3.2 LT.",2.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-068","PLATO PARA SOPA",5.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-069","TABLA DE PICAR DE POLIPROPILENO GRANDE",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-070","PLATO PARA ENTRADA",7.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-071","SEÑAL MIXTA DE HUMO",2.0,"UNIDAD","ISLA CHINCHA","SEGURIDAD Y SALVAMENTO"],
    ["BZA-072","FILTROS DE AGUA PARA TANQUE ROTOPLAS",6.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-073","SET DE OLLA DE ACERO INOXIDABLE  6 PZ.",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-074","SET DE OLLA DE ACERO INOXIDABLE  2PZ.",0.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-075","MANGUERA DE AGUA Y AIRE DE 1/4 PULG",50.0,"METROS","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-076","CABLLE DE ACERO 1/2 PULG TIPO: LOMO DE CORVINA FLEXIBLE",133.0,"METROS","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-077","GANCHO DE ACERO  TIPO OCHO GIRATORIO DE 15 TON",0.0,"UNIDAD","ISLA CHINCHA","FERRETERIA NAVAL"],
    ["BZA-078","PLATO PARA TE",11.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-079","VASO",0.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-080","CUCHARAS PARA SOPA ACERO INOX",14.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-081","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-082","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",2.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-083","CUCHARON PARA SOPA DE ACERO INOXIDABLE",3.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-084","JARRAS DE PLASTICO DE 4 LT",1.0,"UNIDAD","ISLA CHINCHA","MENAJE"],
    ["BZA-085","ACEITE LUBRICANTE (COD:15W40/MARCA: VISTONY)",22.0,"GALONES","ISLA CHINCHA","GRASAS Y LUBRICANTES"],
    ["BZA-086","ACEITE LUBRICANTE (COD:80W90/MARCA: VISTONY)",7.0,"BALDES","ISLA CHINCHA","GRASAS Y LUBRICANTES"],
    ["BZA-087","GRASA MULTIPROPOSITO (MARCA:: VISTONY)",2.0,"BALDES","ISLA CHINCHA","GRASAS Y LUBRICANTES"],
    ["BZA-088","ACEITE HIDRAULICO (COD: ISO68 / MARCA: VISTONY)",120.0,"GALONES","ISLA CHINCHA","GRASAS Y LUBRICANTES"],
    ["BZA-089","FOCO LED 12V DE 13 WATTS (MARCA: WAILEX)",10.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-090","FOCO LED 24V DE 13 WATTS (MARCA: WAILEX)",5.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-091","FOCO DE NAVECACION 24V CHICOS",3.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-092","FOCO DE NAVECACION 24V MEDIANO",16.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-093","SOQUETTE",0.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-094","FAJA DE DISTRIBUCION (COD: 8PK1535 - MARCA: BANDO)",1.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-095","ESCOBILLA TIPO COPA DE 3 PULG (MARCA: TRUPER)",16.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-096","LLAVE PARA LAVADERO PICO DE GANZO (MARCA G&A)",1.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-097","MANDIL PARA SOLDAR",1.0,"UNIDAD","ISLA CHINCHA","EPPS"],
    ["BZA-098","CEPILLO DE ALAMBRE CON MANGO (MARCA: TRUPER)",12.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-099","LLAVE DE PASO DE 1/2 PULG",1.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-100","LLAVE DE PASO DE 3/4 PULG",0.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-101","LLAVE DE PASO DE 1 PULG",0.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-102","CINTA AISLANTE COLOR ROJO",2.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-103","TEFLON 1/2 PULG x 12 M",15.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-104","RODILLO PARA PINTAR DE 3 PULG",14.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-105","RODILLO PARA PINTAR DE 9 PULG",4.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-106","LIMPIA CONTACTOS EN SPRAY",9.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-107","AFLOJATODO WD4O EN SPRAY",3.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-108","MASILLA EPOXICA (MARCA: DEVCON)",3.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-109","BORNES PARA BATERIA",4.0,"PAR","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-110","MENEKE",1.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-111","HOJA DE SIERRA DE 12 PULG (MARCA: TRUPER)",16.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-112","WINCHA DE 5M",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-113","DISCO DE DEBASTE DE 4 1/2 PULG (MARCA: DEWALT)",8.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-114","DISCO DE CORTE DE 4 1/2 PULG (MARCA: KRANENFLEX)",3.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-115","LIJA NRO 50",10.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-116","LIJA NRO 180",8.0,"UNIDAD","ISLA CHINCHA","MANTENIMIENTO"],
    ["BZA-117","PISTOLA DE SOLDAR ESTAÑO",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-118","REFLECTOR DE 50 WATTS (USADO)",1.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-119","REFLECTOR DE 500 WATTS (USADO)",2.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-120","FILTRO DE PETROLEO (COD: P551423/MARCA:DONALSON) PARA GE JHON DEERE",5.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-121","FILTRO DE PETROLEO (COD: P551425/MARCA:DONALSON) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-122","FILTRO DE ACEITE (COD:LF16243/MARCA:FLEERGUARD) PARA GE JHON DEERE",8.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-123","FILTRO PURIFICAOR DE ACEITE (COD:CV55017/MARCA: FLEETGUARD) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-124","FILTRO RACOR (COD: FS19763/MARCA: FLEETGUARD) PARA GE JHON DEERE",4.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-125","FILTRO DE PETROLEO (COD: P551423/MARCA:DONALSON) PARA GE JHON DEERE",6.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-126","FILTRO DE PETROLEO (COD: P551424/MARCA:DONALSON) PARA GE JHON DEERE",4.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-127","FILTRO DE PETROLEO (COD: P551425/MARCA:DONALSON) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-128","FILTRO DE PETROLEO (COD: P551435/MARCA:DONALSON) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-129","FILTRO DE PETROLEO (COD:PSC886/ MARCA: TECFIL) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-130","FILTRO DE ACEITE (COD:LF16243/MARCA:FLEERGUARD) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-131","FILTRO PURIFICAOR DE ACEITE (COD:CV55017/MARCA: FLEETGUARD) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-132","FILTRO PURIFICAOR DE ACEITE (COD:CV55017/SIN CODIGO RD) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-133","FILTRO RACOR (COD: FS19763/MARCA: FLEETGUARD) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-134","FILTRO RACOR (COD: FS19763/MARCA: DONALDSON) PARA GE JHON DEERE",0.0,"UNIDAD","ISLA CHINCHA","FILTROS"],
    ["BZA-135","TORQUIMETRO DE 1/4 PULG (MARCA:STANLEY) 73592",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-136","AMOLADOR DE MANO DE 700 WATTS (MARCA: DEWALT)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-137","AMOLADORA DE 700 WATTS (MARCA: GLUS)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-138","PISTOLA DE TEMPERATURA (MARCA: BOECO)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-139","MEGOMENTRO (MARCA: FLUKE)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-140","PIE DE REY DE 6 PULG",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-141","TALADRO PERCUTOR DE 800 WATTS (MARCA: BOSCH)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-142","MAQUINA DE SOLAR (MARCA: CROWN)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-143","PINZA AMPERIMETRICA DE 600V (MARCA: PRASEK)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-144","SET DE TARRAJA 1/8 - 1/2 PULG (MARCA TRUPER) NUEVO",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-145","SET DE DADOS (MARCA STANLEY) - FALTA LLAVE 14",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-146","SET DE DESARMADORES DE 20 PZAS (MARCA: STANLEY) INCOMPLETO",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-147","SET DE DADOS DE 25 PZAS (MARCA: TRUPER) NUEVO",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-148","PISTOLA DE AIRE DE 90 PSI (MARCA: TRUPER)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-149","EXTENSION DE DADO DE 1/2 PULLG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-150","ABRAZADERA DE ACERO INOX 50-70 (2 1/2)",10.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-151","ABRAZADERA DE ACERO INOX 40 -60 (2 1/2)",10.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-152","ABRAZADERA DE ACERO INOX 32 -50 (2 1/2)",15.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-153","ABRAZADERA DE ACERO INOX 70 -90 (2 1/2)",4.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-154","ABRAZADERA DE ACERO INOX 12 -22 (2 1/2)",5.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-155","ABRAZADERA DE ACERO INOX 64-67 (2 1/2)",5.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-156","ABRAZADERA DE ACERO INOX 48-51 (2 1/2)",3.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-157","ABRAZADERA DE ACERO INOX 52-57 (2 1/2)",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-158","ABRAZADERA DE ACERO INOX 74-79 (2 1/2)",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-159","EXTENSION DE 5M",1.0,"UNIDAD","ISLA CHINCHA","MAT. ELECTRICIDAD"],
    ["BZA-160","LLAVE MIXTA DE BOCA Y CORONA DE 6 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-161","LLAVE MIXTA DE BOCA Y CORONA DE 7 MM",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-162","LLAVE MIXTA DE BOCA Y CORONA DE 8 MM",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-163","LLAVE MIXTA DE BOCA Y CORONA DE 9 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-164","LLAVE MIXTA DE BOCA Y CORONA DE 10 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-165","LLAVE MIXTA DE BOCA Y CORONA DE 11 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-166","LLAVE MIXTA DE BOCA Y CORONA DE 12 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-167","LLAVE MIXTA DE BOCA Y CORONA DE 13 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-168","LLAVE MIXTA DE BOCA Y CORONA DE 14 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-169","LLAVE MIXTA DE BOCA Y CORONA DE 15 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-170","LLAVE MIXTA DE BOCA Y CORONA DE 17 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-171","LLAVE MIXTA DE BOCA Y CORONA DE 19 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-172","LLAVE MIXTA DE BOCA Y CORONA DE 22 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-173","LLAVE MIXTA DE BOCA Y CORONA DE 32 MM",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-174","LLAVE MIXTA DE BOCA Y CORONA DE 1 1/8 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-175","LLAVE MIXTA DE BOCA Y CORONA DE 1 1/16 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-176","LLAVE MIXTA DE BOCA Y CORONA DE 15/16 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-177","LLAVE MIXTA DE BOCA Y CORONA DE 13/16 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-178","LLAVE MIXTA DE BOCA Y CORONA DE 3/4 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-179","LLAVE MIXTA DE BOCA Y CORONA DE 11/16 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-180","LLAVE MIXTA DE BOCA Y CORONA DE 1 1/4 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-181","LLAVE MIXTA DE BOCA Y CORONA DE 1/4 PULG",0.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-182","SET DE DESARMADORES (MARCA: IMCO)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-183","MARTILLO CARPINTERO",0.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-184","ARCO SIERRA",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-185","ALICATE UNIVERSAL",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-186","ALICATE DE CORTE",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-187","ALICATE DE PRESION",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-188","LLAVE FRANCESA DE 8 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-189","LLAVE FRANCESA DE 12 PULG",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-190","LLAVE  STILLSON DE 12 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-191","LLAVE  STILLSON DE 18 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-192","SET DE LLAVES MIXTA DE 12 PZAS (06 MM A 22MM) MARCA:FERRAWANEY",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-193","SET PINZAS SACASEGUROS (3PZS)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-194","SET DE LIMAS (MARCA: KAMASA)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-195","LIMA MEDIA CAÑA DE 8 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-196","LIMA PLANA CAÑA DE 10 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-197","PISTOLA DE AIRE PARA PINTAR",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-198","SET EXTRACTOR DE PERNOS 1/8 - 3/4 (MARCA: TRUPER)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-199","JUEGO DE BROCAS 1MM - 100MM (MARCA: WURTH)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-200","JUEGO DE BROCAS (MARCA: BOSCH) USADO",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-201","BROCA DE COBALTO 5/32 PULG",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-202","SET DE LLAVES HEXAGONALES (COMPLETO)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-203","SET DE LLAVES ESTRELLADOS (COMPLETO)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-204","SET DE LLAVES HEXAGONALES (INCOMPLETO)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-205","JUEGO DE DESARMADORES PLANO Y ESTRELLA",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-206","BOMBA MANUAL TIPO RELOJ",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-207","ENCHUFE",0.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-208","SERRUCHO DE 20 PULG",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-209","GRASERAS (USADO)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-210","TECLE DE 2 TON (MARCA: YALE)",2.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-211","CINTAS ANTIOXIDANTES",12.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-212","SET SACABOCADOS 11PZS (REGULAR)",1.0,"UNIDAD","ISLA CHINCHA","HERRAMIENTAS"],
    ["BZA-213","MAGUERA CONTRTAINCENDIO DE NITRILO 1 1/2 PULG",2.0,"UNIDAD","ISLA CHINCHA","SEGURIDAD Y SALVAMENTO"],
    ["BZA-214","TORQUIMETRO (PROCEDENCIA: REMOLCADOR ALCATRAZ)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["BZA-215","TARJETA AVR STAMFOR PARA GE LISTE PETER (PROCEDENCIA: REMOLCADOR ALCATRAZ)",1.0,"UNIDAD","ISLA CHINCHA","MAQUINARIAS Y EQUIPOS"],
    ["PEL-001","GUANTES PARA SOLDAR",1.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-002","MASCARA DE SOLDAR",1.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-003","CALZADO DE SEGURIDAD TALLA:41",0.0,"PAR","PELICANO","EPPS"],
    ["PEL-004","ARNES DE SEGURIDAD MAS LINEA DE VIDA",2.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-005","LENTES DE SEGURIDAD TRASNPARENTES (MARCA:LIBUS)",0.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-006","LENTES DE SEGURIDAD OSCUROS (MARCA:LIBUS)",0.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-007","CASCO DE SEGURIDAD COLOR NARANJA (MARCA: CLUTE)",7.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-008","CASCO DE SEGURIDAD COLOR BLANCO (MARCA: CLUTE)",1.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-009","GUANTES DE CUERO BADANA COLOR NARANJA",0.0,"PAR","PELICANO","EPPS"],
    ["PEL-010","GUANTES DE CUERO CROMO",0.0,"PAR","PELICANO","EPPS"],
    ["PEL-011","GUANTES DE CUERO BADANA COLOR BLANCO",0.0,"PAR","PELICANO","EPPS"],
    ["PEL-012","GUANTES ANTICORTE CON PALMA RECUBIERTA",0.0,"PAR","PELICANO","EPPS"],
    ["PEL-013","GUNATES RECUBIERTO DE NITRILO AZUL",5.0,"PAR","PELICANO","EPPS"],
    ["PEL-014","MANDIL PARA SOLDAR",1.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-015","TAPONES AUDITIVOS",25.0,"UNIDAD","PELICANO","EPPS"],
    ["PEL-016","GRILLETE TIPO LIRA DE 1 1/2 PULG",7.0,"UNIDAD","PELICANO","FERRETERIA NAVAL"],
    ["PEL-017","GRILLETE TIPO LIRA DE 1 1/4 PULG",2.0,"UNIDAD","PELICANO","FERRETERIA NAVAL"],
    ["PEL-018","GRILLETE TIPO LIRA DE 1/4 PULG",0.0,"UNIDAD","PELICANO","FERRETERIA NAVAL"],
    ["PEL-019","CABO DE NYLON DE 1 1/2 PULG",110.0,"METROS","PELICANO","FERRETERIA NAVAL"],
    ["PEL-020","CABO DE NYLON DE 2 PULG",0.0,"METROS","PELICANO","FERRETERIA NAVAL"],
    ["PEL-021","GUARDACABO TIPO H DE 1 1/2 PULG",1.0,"UNIDAD","PELICANO","FERRETERIA NAVAL"],
    ["PEL-022","FILTRO DE ACEITE (COD: LF691A/MARCA:FLEETGUARD) PARA MOTOR PRINCIPAL",26.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-023","FILTRO DE PETROLEO (COD:FF5317/MARCA:FLEERGUARD) PARA MOTOR PRINCIPAL",10.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-024","FILTRO RACOR (COD: FSP 2040-3020202/MARCA: PUROLATOR) PARA MOTOR PRINCIPAL",0.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-025","FILTRO RACOR (COD: FS20202/MARCA: FLEETGUARD) PARA MOTOR PRINCIPAL",25.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-026","FILTRO DE AIRE (COD: 5865323/MARCA: CAT) PARA MOTOR PRINCIPAL (USADO/MUESTRA)",5.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-027","FILTRO DE PETROLEO (COD: FF185/MARCA: FLEETGUARD) PARA GE DOOSAN",14.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-028","FILTRO DE ACEITE (COD: LF3959/MARCA:FLEETGUARD) PARA GE DOOSAN",15.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-029","FILTRO DE ACEITE (COD: L20290/MARCA: PUROLATOR) PARA GE HYUNDAI",4.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-030","FILTRO DE PETROLEO (COD: F55175 / MARCA: PUROLATOR) PARA GE HYUNDAI",8.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-031","FILTRO DE AIRE (COD: E013412020/MARCA: HYUNDAI) PARA GE HYUNDAI",1.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-032","FILTRO PARA EQUIPO DE OSMOSIS",16.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-033","FILTRO ESPONJA DE AIRE PARA GE HYUNDAI",14.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-034","DISCO DE DEBASTE DE 4 1/2 PULG",34.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-035","DISCO DE CORTE DE 4 1/2 PULG",0.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-036","ESCOBILLA TIPO COPA DE 3 PULG (MARCA: TRUPER)",6.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-037","SET DE LIMAS",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-038","SET DE DESARMADORES DE 20 PZAS (MARCA STANLEY)",0.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-039","SET DE DADOS DE 29 PZAS (MARCA: STANLEY)",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-040","SET DE DADOS DE 25 PZAS (MARCA: STANLEY)",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-041","TECLE DE 2 TON (MARCA: YALE)",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-042","PISTOLA DE GRAVEDAD",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-043","PISTOLA DE AIRE",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-044","PIE DE REY DE 6 PULG",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-045","WINCHA DE 5 M",2.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-046","SET LLAVE ESTRELLADO TORX",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-047","PISTOLA PARA RIEGO",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-048","TIJERA CORTALATA",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-049","ALICATE UNVERSAL",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-050","ALICATE DE CORTE",1.0,"UNIDAD","PELICANO","HERRAMIENTAS"],
    ["PEL-051","RODILLO PARA PINTAR DE 9 PULG",0.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-052","MASILLA EPOXICA (MARCA: DEVCON)",1.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-053","DISCO DE PULIDO",0.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-054","CEPILLO DE ALAMBRE CON MANGO (MARCA: TRUPER)",1.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-055","RODILLO PARA PINTAR DE 4 PULG",6.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-056","RODILLO PARA PINTAR DE 9 PULG",4.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-057","CABLE ADAPTADOR PARA MOTOR CATERPILLAR",0.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-058","VALVULA DE PASO DE 1 PULG",4.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-059","VALVULA DE PASO DE 1/2 PULG",4.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-060","LLAVE DE PASO DE BRONCE 1/4 PULG",2.0,"UNIDAD","PELICANO","MANTENIMIENTO"],
    ["PEL-061","MULTITESTER (MARCA: AMPROBE)",1.0,"UNIDAD","PELICANO","MAQUINARIAS Y EQUIPOS"],
    ["PEL-062","TALADRO PERCUTOR DE 800 WATTS (MARCA: STANLEY)",1.0,"UNIDAD","PELICANO","MAQUINARIAS Y EQUIPOS"],
    ["PEL-063","MAQUINA DE SOLDAR (MARCA: BAUKER)",1.0,"UNIDAD","PELICANO","MAQUINARIAS Y EQUIPOS"],
    ["PEL-064","AMOLADORA DE 700 WATTS (MARCA: BOSCH)",0.0,"UNIDAD","PELICANO","MAQUINARIAS Y EQUIPOS"],
    ["PEL-065","TEFLON",4.0,"UNIDAD","PELICANO","MAT. ELECTRICIDAD"],
    ["PEL-066","CINTA AISLANTE COLOR ROJO",6.0,"UNIDAD","PELICANO","MAT. ELECTRICIDAD"],
    ["PEL-067","CINTA AISLANTE COLOR NEGRO",0.0,"UNIDAD","PELICANO","MAT. ELECTRICIDAD"],
    ["PEL-068","BORNES PARA BATERIA",6.0,"UNIDAD","PELICANO","MAT. ELECTRICIDAD"],
    ["PEL-069","LINTERNA DE MANO CON LUZ LED RECARGABLE",0.0,"UNIDAD","PELICANO","MAT. ELECTRICIDAD"],
    ["PEL-070","AZUCARERA DE ACERO",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-071","BOL DE ACERO",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-072","COLADOR DE ACERO CON MANGO",3.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-073","COLADOR DE PLASTICO",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-074","COLADOR DE PLASTICO CON MANGO",0.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-075","CUCHARAS PARA SOPA ACERO INOX",0.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-076","CUCHARAS PARA TE ACERO INOX",20.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-077","CUCHARÓN DE ESPAGUETTI",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-078","CUCHARON DE MADERA",4.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-079","CUCHARON PARA SOPA DE ACERO INOXIDABLE",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-080","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-081","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-082","CUCHILLOS PARA MESA ACERO INOX.",27.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-083","ESCURRIDOR DE PLASTICO",3.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-084","ESPATULA DE ACERO",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-085","ESPATULA DE TEFLON",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-086","ESPUMADERA DE COCINA",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-087","EXPRIMIDOR",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-088","FUENTES DE ACERO",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-089","JARRAS DE PLASTICO DE 4 LT",4.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-090","OLLA A PRESION",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-091","OLLAS DE ACERO QUIRURGICO DE 4 PZAS",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-092","PELADOR DE PAPA",0.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-093","PINZA DE ACERO INOXIDABLE",6.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-094","PLATO DE ENTRADA",12.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-095","PLATO SOPERO",12.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-096","PORTA CUBIERTOS DE PLASTICO",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-097","RAYADIR DE ACERO",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-098","SARTEN DE ALUMINIO NRO 30",5.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-099","SARTEN DE TEFLON NRO 26",0.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-100","SET DE OLLA DE ACERO INOXIDABLE  6 PZ.",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-101","TABLA DE PICAR DE POLIPROPILENO DE 1/2 ' DE ESPESOR",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-102","TAZA DE TE",13.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-103","TENEDORES PARA MESA  ACERO INOX.",17.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-104","TERMO DE 3L",1.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-105","TETERA DE ACERO INOXIDABLE 3.2 LT.",2.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-106","VASO DE VIDRIO",0.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-107","WOK DE TEFLON NRO 30",0.0,"UNIDAD","PELICANO","MENAJE"],
    ["PEL-108","BASE ZINCROMATO (MARCA: ANYPSA)",0.0,"GALON","PELICANO","PINTURAS"],
    ["PEL-109","CATALIZADOR (MARCA: ANYPSA)",0.0,"GALON","PELICANO","PINTURAS"],
    ["PEL-110","PINTURA EPOXICA COLOR NEGRO (MARCA: ANYPSA)",0.0,"GALON","PELICANO","PINTURAS"],
    ["PEL-111","PINTURA EPOXICA COLOR GRIS (MARCA:JET / CODIGO: PROTECTO2174:974)",0.0,"GALON","PELICANO","PINTURAS"],
    ["PEL-112","CATALIZADOR (MARCA: JET)",0.0,"GALON","PELICANO","PINTURAS"],
    ["PEL-113","DISOLVENTE EPOXICO (MARCA: ANYPSA)",2.0,"GALON","PELICANO","PINTURAS"],
    ["PEL-114","SEÑAL MIXTA DE HUMO",2.0,"UNIDAD","PELICANO","SEGURIDAD Y SALVAMENTO"],
    ["PEL-115","MANGUERA CONTRA INCENDIO DE NITRILO 1 1/2 PULG x 30M",2.0,"UNIDAD","PELICANO","SEGURIDAD Y SALVAMENTO"],
    ["PEL-116","MANGUERA CONTRA INCENDIO DE NITRILO 2 1/2 PULG x 30M",3.0,"UNIDAD","PELICANO","SEGURIDAD Y SALVAMENTO"],
    ["PEL-117","MANGUERA CONTRA INCENDIO DE JEBE Y LONA 1 1/2 PULG x 30M",4.0,"UNIDAD","PELICANO","SEGURIDAD Y SALVAMENTO"],
    ["PEL-118","PANTALOS DRIL DE SEGURIDAD COLOR AZUL TALLA: 30",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-119","PANTALOS DRIL DE SEGURIDAD COLOR AZUL TALLA: 32",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-120","PANTALOS DRIL DE SEGURIDAD COLOR AZUL TALLA: 34",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-121","CAMISA DRIL DE SEGURIDAD COLOR AZUL TALLA: M",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-122","CAMISA DRIL DE SEGURIDAD COLOR AZUL TALLA: L",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-123","CAMISA DRIL DE SEGURIDAD COLOR AZUL TALLA: XL",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-124","CAMISA DRIL DE SEGURIDAD COLOR AZUL TALLA: XXL",0.0,"UNIDAD","PELICANO","VESTUARIO"],
    ["PEL-125","FILTRO HIDRAULICO PARA CAJA REDUCTORA (HF6553) PARA MOTOR PRINCIPAL",8.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-126","FILTRO  DE PETROLEO (MARCA:FLEERGUARD / COD: FF5087) PARA GE HYUNDAI",6.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-127","FILTRO RACOR PARA GE HYUNDAI",11.0,"UNIDAD","PELICANO","FILTROS"],
    ["PEL-128","BASE GRI NIEBLA (MARCA: JET)",0.0,"GALON","PELICANO","PINTURAS"],
    ["D11-001","GRILLETE DE 1 1/4 PULG",4.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-002","GRILLETE DE 1 PULG",3.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-003","GRILLETE DE 3/4 PULG",3.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-004","GRILLETE DE 1/2 PULG",7.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-005","GRILLETE DE 1/4 PULG",3.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-006","GUARDACABO DE 1 PULG",2.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-007","GUARDACABO DE 5/8 PULG",3.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-008","GUARDACABO DE 1/2 PULG",4.0,"UNIDAD","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-009","FILTRO DE ACEITE (COD: LF691A/MARCA:FLEETGUARD) PARA MOTOR PRINCIPAL",1.0,"UNIDAD","DELFIN 11","FILTROS"],
    ["D11-010","FILTRO DE AIRE (COD:4N 0015) PARA MOTOR PRINCIPAL",2.0,"UNIDAD","DELFIN 11","FILTROS"],
    ["D11-011","FILTRO DE PETROLEO (COD: FF5319) PARA MOTOR PRINCIPAL",0.0,"UNIDAD","DELFIN 11","FILTROS"],
    ["D11-012","FILTRO RACOR (2020) PARA MOTOR PRINCIPAL",2.0,"UNIDAD","DELFIN 11","FILTROS"],
    ["D11-013","ACEITERA",1.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-014","BROCA DE 10 MM",2.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-015","BROCA DE COBALTO 5/32",2.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-016","ESCOBILLA TIPO COPA DE 3 PULG (MARCA: TRUPER)",0.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-017","ESTUCHE DE DADOS",1.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-018","GRASERAS",0.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-019","JUEGO DE DESARMADORES PLANO Y ESTRELLA",0.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-020","JUEGO DE EXTRACTOR DE PERNOS 05 PZAS",1.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-021","JUEGO DE SACABOCADO 12 PZAS",1.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-022","LIMA DE 08 PULG",4.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-023","PINZA DE CORTE",1.0,"UNIDAD","DELFIN 11","HERRAMIENTAS"],
    ["D11-024","BROCHA DE 02 PULG",0.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-025","AFLOJATODO WD4O EN SPRAY",0.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-026","ESCOBILLA DE ACERO CON MANGO DE PLASTICO",0.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-027","TEFLON",3.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-028","DISCO DE DEBASTE",7.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-029","DISCO DE CORTE",0.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-030","LIMPIA CONTACTOS EN SPRAY",0.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-031","RODILLO DE 6 PULG",0.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-032","RODILLO DE 4 PULG",4.0,"UNIDAD","DELFIN 11","MANTENIMIENTO"],
    ["D11-033","AMOLADORA DE 700 WATTS (MARCA: BOSCH)",1.0,"UNIDAD","DELFIN 11","MAQUINARIAS Y EQUIPOS"],
    ["D11-034","TALADRO PERCUTOR DE 800 WATTS (MARCA: BOSCH)",1.0,"UNIDAD","DELFIN 11","MAQUINARIAS Y EQUIPOS"],
    ["D11-035","BORNES PARA BATERIA",2.0,"UNIDAD","DELFIN 11","MAT. ELECTRICIDAD"],
    ["D11-036","CABLE VULCANIZADO 2X12 AWG X 100 m",0.0,"UNIDAD","DELFIN 11","MAT. ELECTRICIDAD"],
    ["D11-037","CINTA AISLANTE COLOR NEGRO",0.0,"UNIDAD","DELFIN 11","MAT. ELECTRICIDAD"],
    ["D11-038","CINTA AISLANTE COLOR ROJO",0.0,"UNIDAD","DELFIN 11","MAT. ELECTRICIDAD"],
    ["D11-039","LINTERNA DE MANO CON LUZ LED RECARGABLE",1.0,"UNIDAD","DELFIN 11","MAT. ELECTRICIDAD"],
    ["D11-040","PISTOLA DE SOLDAR ESTAÑO",1.0,"UNIDAD","DELFIN 11","MAT. ELECTRICIDAD"],
    ["D11-041","FUENTE DE ACERO DE 50CMx35CM",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-042","CUCHARAS PARA SOPA ACERO INOX",3.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-043","CUCHARAS PARA TE ACERO INOX",10.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-044","CUCHARÓN DE ESPAGUETTI",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-045","CUCHARON DE MADERA",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-046","CUCHARON PARA SOPA DE ACERO INOXIDABLE",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-047","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",2.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-048","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-049","CUCHILLOS PARA MESA ACERO INOX.",6.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-050","ESPATULA DE ACERO",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-051","ESPATULA DE TEFLON",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-052","ESPUMADERA DE COCINA",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-053","EXPRIMIDOR DE ACERO",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-054","JARRAS DE PLASTICO DE 4 LT",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-055","LICUADORA",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-056","PELADOR DE PAPA",2.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-057","PINZA DE ACERO INOXIDABLE",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-058","PINZAS DE ACERO",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-059","PLATO ENTRADA",4.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-060","PLATO SOPERO",3.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-061","PRENSAPAPA",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-062","RAYADOR DE ACERO",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-063","SET DE OLLA DE ACERO INOXIDABLE  6 PZ.",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-064","TABLA DE PICAR DE POLIPROPILENO DE 1/2 ' DE ESPESOR",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-065","TAZA DE TE",3.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-066","TENEDORES PARA MESA  ACERO INOX.",7.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-067","TERMO DE 3L",1.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-068","TETERA DE ACERO INOXIDABLE 3.2 LT.",0.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-069","VASO DE VIDRIO",6.0,"UNIDAD","DELFIN 11","MENAJE"],
    ["D11-070","CABO DE NYLON DE 1 1/2 PULG",0.0,"ROLLO","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-071","CABO DE NYLON DE 1 1/4 PULG",0.0,"ROLLO","DELFIN 11","FERRETERIA NAVAL"],
    ["D11-072","CABO DE NYLON DE 1 1/2 PULG (SEMI NUEVO)",1.0,"ROLLO","DELFIN 11","FERRETERIA NAVAL"],
    ["D12-001","MASCARA DE SOLDAR",2.0,"UNIDAD","DELFIN 12","EPPS"],
    ["D12-002","ARNES DE SEGURIDAD MAS LINEA DE VIDA (NUEVO)",2.0,"UNIDAD","DELFIN 12","EPPS"],
    ["D12-003","GUANTES DE CUERDO BADANA",1.0,"UNIDAD","DELFIN 12","EPPS"],
    ["D12-004","GUANTES DE SEGURIDAD ANTICORTE",3.0,"UNIDAD","DELFIN 12","EPPS"],
    ["D12-005","CABLE DE BATERIA MARCA: CATERPILLAR",1.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-006","CABO DE NYLON DE 1 1/2 PULG",380.0,"METROS","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-007","GRILLETE DE 1 1/4 PULG",160.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-008","GRILLETE DE 1 PULG",4.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-009","GRILLETE DE 1/2 PULG",11.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-010","GRILLETE DE 1/4 PULG",7.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-011","GRILLETE DE 3/4 PULG",7.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-012","GUARDACABO DE 1/2 PULG",23.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-013","GUARDACABO DE 1/4 PULG",7.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-014","GUARDACABO DE 3/4 PULG",1.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-015","GUARDACABO DE 5/8 PULG",3.0,"UNIDAD","DELFIN 12","FERRETERIA NAVAL"],
    ["D12-016","FILTRO DE ACEITE (COD: LF691A/MARCA:FLEETGUARD) PARA MOTOR PRINCIPAL",1.0,"UNIDAD","DELFIN 12","FILTROS"],
    ["D12-017","FILTRO DE AIRE (COD:4N 0015) PARA MOTOR PRINCIPAL",5.0,"UNIDAD","DELFIN 12","FILTROS"],
    ["D12-018","FILTRO DE PETROLEO (COD: FF5319) PARA MOTOR PRINCIPAL",3.0,"UNIDAD","DELFIN 12","FILTROS"],
    ["D12-019","FILTRO RACOR (2020) PARA MOTOR PRINCIPAL",3.0,"UNIDAD","DELFIN 12","FILTROS"],
    ["D12-020","ACEITE LUBRICANTE (COD:15W40/MARCA: VISTONY)",51.0,"GLN","DELFIN 12","GRASAS Y LUBRICANTES"],
    ["D12-021","GRASA MULTIPROPOSITO (MARCA:: VISTONY)",1.0,"BALDE","DELFIN 12","GRASAS Y LUBRICANTES"],
    ["D12-022","REFRIGERANTE (MARCA: VISTONY)",1.0,"GLN","DELFIN 12","GRASAS Y LUBRICANTES"],
    ["D12-023","ALICATE DE PRESION (MARCA: KAMASA)",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-024","ALICATE UNIVERSAL (USADO)",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-025","ARCO SIERRA",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-026","COMBA (USADO)",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-027","DESARMADORES ESTRELLA",2.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-028","DESARMADORES PLANO",3.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-029","ESCOBILLA DE ACERO CON MANGO DE PLASTICO",3.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-030","JUEGO DE DADOS (MARCA: STALEY)(NUEVO)",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-031","LLAVE MIXTA 09 MM",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-032","LLAVE MIXTA 10 MM",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-033","LLAVE MIXTA 11 MM",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-034","LLAVE MIXTA 12 MM",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-035","LLAVE MIXTA 14 MM",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-036","LLAVE MIXTA 15/16 PULG",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-037","LLAVE MIXTA 3/16 PULG",2.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-038","LLAVE MIXTA 3/4 PULG",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-039","LLAVE MIXTA 5/8 PULG",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-040","LLAVE MIXTA 7/8 PULG",2.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-041","MARTILLO CARPINTERO (USADO)",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-042","SERRUCHO DE 20 PULG",1.0,"UNIDAD","DELFIN 12","HERRAMIENTAS"],
    ["D12-043","AFLOJATODO WD4O EN SPRAY",3.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-044","ESCOBILLA TIPO COPA DE 3 PULG (MARCA: TRUPER)",5.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-045","HOJA DE SIERRA DE 12 PULG (MARCA: TRUPER)",6.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-046","LIJA NRO 180 Y NRO 50 (TRUPER)",2.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-047","LIMPIA CONTACTOS EN SPRAY",7.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-048","PICAPORTE DORADO",2.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-049","RODILLO DE 03 PULG",2.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-050","RODILLO DE 04 PULG",6.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-051","RODILLO DE 09 PULG",7.0,"UNIDAD","DELFIN 12","MANTENIMIENTO"],
    ["D12-052","MAQUINAS DE SOLDAR (MARCA: KUMAMOTO) (NUEVO)",1.0,"UNIDAD","DELFIN 12","MAQUINARIAS Y EQUIPOS"],
    ["D12-053","TALADRO PERCUTOR (MARCA: BOSCH)",1.0,"UNIDAD","DELFIN 12","MAQUINARIAS Y EQUIPOS"],
    ["D12-054","AMOLADORA DE 700 WATTS (MARCA: BLACK DECKER)",1.0,"UNIDAD","DELFIN 12","MAQUINARIAS Y EQUIPOS"],
    ["D12-055","TORNILLO DE BANCO DE 4 PULG",1.0,"UNIDAD","DELFIN 12","MAQUINARIAS Y EQUIPOS"],
    ["D12-056","TEFLON",2.0,"UNIDAD","DELFIN 12","MAT. ELECTRICIDAD"],
    ["D12-057","BORNES PARA BATERIA",2.0,"PAR","DELFIN 12","MAT. ELECTRICIDAD"],
    ["D12-058","CINTA AISLANTE COLOR ROJO",6.0,"UNIDAD","DELFIN 12","MAT. ELECTRICIDAD"],
    ["D12-059","CINTA AISLANTE COLOR NEGRO",2.0,"UNIDAD","DELFIN 12","MAT. ELECTRICIDAD"],
    ["D12-060","LINTERNA DE MANO CON LUZ LED RECARGABLE",1.0,"UNIDAD","DELFIN 12","MAT. ELECTRICIDAD"],
    ["D12-061","CUCHARAS PARA SOPA ACERO INOX",6.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-062","CUCHARAS PARA TE ACERO INOX",5.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-063","CUCHARÓN DE ESPAGUETTI",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-064","CUCHARON DE MADERA",2.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-065","CUCHARON PARA SOPA DE ACERO INOXIDABLE",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-066","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-067","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-068","CUCHILLOS PARA MESA ACERO INOX.",6.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-069","ESPUMADERA DE COCINA",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-070","JARRAS DE PLASTICO DE 4 LT",2.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-071","PELADOR DE PAPA",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-072","PINZA DE ACERO INOXIDABLE",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-073","SET DE OLLA DE ACERO INOXIDABLE  6 PZ.",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-074","TABLA DE PICAR DE POLIPROPILENO DE 1/2 ' DE ESPESOR",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-075","TENEDORES PARA MESA  ACERO INOX.",6.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["D12-076","TETERA DE ACERO INOXIDABLE 3.2 LT.",1.0,"UNIDAD","DELFIN 12","MENAJE"],
    ["GUA-001","ARNES DE SEGURIDAD",0.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-002","TAPONES AUDITIVOS",15.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-003","CARERA PARA SOLDAR",1.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-004","GUANTES PARA SOLDAR",1.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-005","LENTES DE SEGURIDAD TRASNPARENTES (MARCA:LIBUS)",6.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-006","LENTES DE SEGURIDAD OSCUROS",8.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-007","CASCO DE SEGURIDAD",8.0,"UNIDAD","GUANAY","EPPS"],
    ["GUA-008","GUANTES DE CUERO BADANA COLOR BLANCO",17.0,"PAR","GUANAY","EPPS"],
    ["GUA-009","GUANTES ANTICORTE CON PALMA RECUBIERTA",14.0,"PAR","GUANAY","EPPS"],
    ["GUA-010","GUNATES RECUBIERTO DE NITRILO AZUL",14.0,"PAR","GUANAY","EPPS"],
    ["GUA-011","GUNATES DE CUERO CROMO",24.0,"PAR","GUANAY","EPPS"],
    ["GUA-012","CABLE DE ACERO DE 3/4 PULG",0.0,"METROS","GUANAY","FERRETERIA NAVAL"],
    ["GUA-013","CABO DE NYLON DE 1 1/2 PULG",440.0,"METROS","GUANAY","FERRETERIA NAVAL"],
    ["GUA-014","CABO DE NYLON DE 2 PULG",220.0,"METROS","GUANAY","FERRETERIA NAVAL"],
    ["GUA-015","CABO DE NYLON DE 1 PULG",150.0,"METROS","GUANAY","FERRETERIA NAVAL"],
    ["GUA-016","CABO DE NYLON DE 1/2 PULG",80.0,"METROS","GUANAY","FERRETERIA NAVAL"],
    ["GUA-017","GUARDACABO DE 1/2 PULG",4.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-018","GUARDACABO DE 3/4 PULG",4.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-019","GUARDACABO DE 1 PULG",5.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-020","GUARDACABO DE 1 1/2 PULG",6.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-021","GUARDACABO DE 2 PULG",1.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-022","GRILLETE TIPO LIRA DE 3/4 PULG",2.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-023","GRILLETE TIPO LIRA DE 1 PULG",3.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-024","GRILLETE TIPO LIRA DE 1 1/4 PULG",10.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-025","GRILLETE TIPO LIRA DE 1 1/2 PULG",4.0,"UNIDAD","GUANAY","FERRETERIA NAVAL"],
    ["GUA-026","FILTRO DE PETROLEO (COD: FF5644/MARCA:FLEETGUARD) PARA MOTOR PRINCIPAL",20.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-027","FILTRO CENTRIFUGO (COD:CS41011/MARCA:FLEERGUARD) PARA MOTOR PRINCIPAL",6.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-028","FILTRO DE AGUA (COD:WF2076/MARCA: FLEETGUARD) PARA MOTOR PRINCIPAL",13.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-029","FILTRO RACOR (COD: FS53021/MARCA: FLEETGUARD) PARA MOTOR PRINCIPAL",17.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-030","FILTRO DE AIRE (COD: AF872M/MARCA: FLEETGUARD) PARA MOTOR PRINCIPAL",10.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-031","FILTRO DE ACEITE (COD: LF3715/MARCA: FLEETGUARD) PARA GE DOOSAN",8.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-032","FILTRO RACOR (COD: 2020N-10) PARA GE DOOSAN",29.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-033","FILTRO DE PETROLEO (COD: FF185/MARCA:FLEETGUARD) PARA GE DOOSAN",9.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-034","FILTRO DE ACEITE (COD: L20290/MARCA: PUROLATOR) PARA GE HYUNDAI",3.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-035","FILTRO DE PETROLEO (COD: F55175 / MARCA: PUROLATOR) PARA GE HYUNDAI",2.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-036","FILTRO RACOR (COD: FSP2040-30/MARCA:PUROLATOR) PARA GE HYUNDAI",4.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-037","FILTRO DE AIRE (COD: E013412020/MARCA: HYUNDAI) PARA GE HYUNDAI",2.0,"UNIDAD","GUANAY","FILTROS"],
    ["GUA-038","ACEITE LUBRICANTE (COD:15W40/MARCA: VISTONY)",89.0,"GALON","GUANAY","GRASAS Y LUBRICANTES"],
    ["GUA-039","ACEITE LUBRICANTE (COD:SAE40/MARCA: VISTONY)",5.0,"GALON","GUANAY","GRASAS Y LUBRICANTES"],
    ["GUA-040","ACEITE HIDRAULICO (COD: ISO68 / MARCA: VISTONY)",365.0,"GALON","GUANAY","GRASAS Y LUBRICANTES"],
    ["GUA-041","REFRIGERANTE (MARCA: VISTONY)",0.0,"GALON","GUANAY","GRASAS Y LUBRICANTES"],
    ["GUA-042","GRASA MULTIPROPOSITO (MARCA:: VISTONY)",1.0,"BALDES","GUANAY","GRASAS Y LUBRICANTES"],
    ["GUA-043","LLAVE MIXTA DE BOCA Y CORONA DE 9 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-044","LLAVE MIXTA DE BOCA Y CORONA DE 10 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-045","LLAVE MIXTA DE BOCA Y CORONA DE 11 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-046","LLAVE MIXTA DE BOCA Y CORONA DE 12 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-047","LLAVE MIXTA DE BOCA Y CORONA DE 13 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-048","LLAVE MIXTA DE BOCA Y CORONA DE 17 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-049","LLAVE MIXTA DE BOCA Y CORONA DE 19 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-050","LLAVE MIXTA DE BOCA Y CORONA DE 22 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-051","LLAVE MIXTA DE BOCA Y CORONA DE 24 MM",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-052","LLAVE MIXTA DE BOCA Y CORONA DE 1 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-053","LLAVE MIXTA DE BOCA Y CORONA DE 7/8 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-054","LLAVE MIXTA DE BOCA Y CORONA DE 15/16 PULG",2.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-055","LLAVE MIXTA DE BOCA Y CORONA DE 13/16 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-056","LLAVE MIXTA DE BOCA Y CORONA DE 11/16 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-057","LLAVE MIXTA DE BOCA Y CORONA DE 1 1/4 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-058","LLAVE MIXTA DE BOCA Y CORONA DE 1 1/8 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-059","LLAVE MIXTA DE BOCA Y CORONA DE 1 1/16 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-060","MARTILLO CARPINTERO",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-061","ARCO SIERRA",2.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-062","ALICATE UNIVERSAL",2.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-063","ALICATE DE CORTE",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-064","ALICATE DE PRESION",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-065","LLAVE FRANCESA DE 8 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-066","LLAVE FRANCESA DE 12 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-067","LLAVE  STILLSON DE 8 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-068","LLAVE  STILLSON DE 12 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-069","LLAVE  STILLSON DE 24 PULG",0.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-070","ENGRASADORA MANUAL (TRUPER)",2.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-071","ACEITERA",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-072","SET DE LIMAS",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-073","SET DE SACABOCADO DE 12 PZAS",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-074","SERRUCHO DE 20 PULG",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-075","JUEGO DE DESARMADORES",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-076","TIJERA CORTA LATA",2.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-077","COMBA 1LB MEDIANA",1.0,"UNIDAD","GUANAY","HERRAMIENTAS"],
    ["GUA-078","LIMPIA CONTACTOS EN SPRAY",10.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-079","AFLOJATODO WD4O EN SPRAY",1.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-080","DISCO DE DEBASTE (MARCA: DEWALT)",22.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-081","LIJA NRO 180 Y NRO 50 (TRUPER)",17.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-082","ESCOBILLA TIPO COPA DE 3 PULG (MARCA: TRUPER)",5.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-083","CEPILLO DE ALAMBRE CON MANGO (MARCA: TRUPER)",10.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-084","HOJA DE SIERRA DE 12 PULG (MARCA: TRUPER)",29.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-085","VALVULA CHECK DE 1/2 PULG DE BRONCE",5.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-086","LLAVE DE PASO DE 3/4 PULG (VALMAX)",1.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-087","LLAVE DE PASO DE 1 PULG (VALMAX)",1.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-088","AMOLADORA DE 700 WATTS (MARCA: BOSCH)",1.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-089","BROCHA PARA PINTAR DE 1 PULG",1.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-090","RODILLO PARA PINTAR DE 4 PULG",7.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-091","RODILLO PARA PINTAR DE 8 PULG",8.0,"UNIDAD","GUANAY","MANTENIMIENTO"],
    ["GUA-092","TALADRO PERCUTOR DE 800 WATTS (MARCA: BOSCH)",1.0,"UNIDAD","GUANAY","MAQUINARIAS Y EQUIPOS"],
    ["GUA-093","TECLE DE 3 TON",1.0,"UNIDAD","GUANAY","MAQUINARIAS Y EQUIPOS"],
    ["GUA-094","CUCHARAS PARA SOPA ACERO INOX",14.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-095","CUCHARAS PARA TE ACERO INOX",14.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-096","TENEDORES PARA MESA  ACERO INOX.",14.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-097","CUCHILLOS PARA MESA ACERO INOX.",14.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-098","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-099","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-100","CUCHARON PARA SOPA DE ACERO INOXIDABLE",2.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-101","TETERA DE ACERO INOXIDABLE 3.2 LT.",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-102","SET DE OLLA DE ACERO INOXIDABLE  6 PZ.",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-103","ESPUMADERA DE COCINA",1.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-104","CUCHARON DE MADERA",2.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-105","TABLA DE PICAR DE POLIPROPILENO DE 1/2 ' DE ESPESOR",1.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-106","PELADORES",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-107","JARRAS DE PLASTICO DE 4 LT",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-108","PINZA DE ACERO INOXIDABLE",1.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-109","CUCHARÓN DE ESPAGUETTI",0.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-110","LINTERNA DE MANO CON LUZ LED RECARGABLE",1.0,"UNIDAD","GUANAY","MENAJE"],
    ["GUA-111","DISOLVENTE EPOXICO (MARCA: ANYPSA)",8.0,"GALON","GUANAY","PINTURAS"],
    ["GUA-112","MAGUERA CONTRTAINCENDIO DE NITRILO 1 1/2 PULG",2.0,"UNIDAD","GUANAY","SEGURIDAD Y SALVAMENTO"],
    ["GUA-113","LUZ DE NAVEGACION (POY POY)",2.0,"UNIDAD","GUANAY","SEGURIDAD Y SALVAMENTO"],
    ["CIST-001","LEJIA X 1 LT",5.0,"UNIDAD","CISTERNA 1","ASEO Y LIMPIEZA"],
    ["CIST-002","LIMPIATODO DESINFECTANTE",3.0,"GALON","CISTERNA 1","ASEO Y LIMPIEZA"],
    ["CIST-003","LAVAVAJILLA EN PASTA DE 800 GR",6.0,"UNIDAD","CISTERNA 1","ASEO Y LIMPIEZA"],
    ["CIST-004","CASCO DE SEGURIDAD COLOR NARANJA",3.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-005","CEPILLO DE ALAMBRE CON MANGO (MARCA: TRUPER)",14.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-006","LENTES DE SEGURIDAD TRANSPARENTES",7.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-007","LENTES DE SEGURIDAD OSCUROS",6.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-008","TAPONES AUDITIVOS",5.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-009","GUANTES ANTICORTE CON PALMA RECUBIERTA",8.0,"PAR","CISTERNA 1","EPPS"],
    ["CIST-010","GUNATES RECUBIERTO DE NITRILO AZUL",17.0,"PAR","CISTERNA 1","EPPS"],
    ["CIST-011","GUANTES DE CUERO BADANA COLOR BLANCO",8.0,"PAR","CISTERNA 1","EPPS"],
    ["CIST-012","GUANTES DE CUERO CROMO",17.0,"PAR","CISTERNA 1","EPPS"],
    ["CIST-013","ARNES DE SEGURIDAD",1.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-014","OREJERA TIPO VINCHA",1.0,"UNIDAD","CISTERNA 1","EPPS"],
    ["CIST-015","GRILLETE TIPO LIRA DE 1/2 PULG",14.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-016","GRILLETE TIPO LIRA DE 3/4 PULG",8.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-017","GRILLETE TIPO LIRA DE 1 PULG",3.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-018","GRILLETE TIPO LIRA DE 1 1/4 PULG",2.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-019","GRILLETE TIPO LIRA DE 1 1/2 PULG",5.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-020","GUARDACABO DE 1 3/8 HASTA 1 1/2 PULG",6.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-021","GUARDACABO TIPO H DE 1 PULG",5.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-022","GUARDACABO TIPO L DE 1 PULG",5.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-023","GUARDACABO DE 3/4 PULG",6.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-024","PATESCA CON GANCHO (REGULAR)",2.0,"UNIDAD","CISTERNA 1","FERRETERIA NAVAL"],
    ["CIST-025","FILTRO DE ACEITE PARA GE LISTER PETER",7.0,"UNIDAD","CISTERNA 1","FILTROS"],
    ["CIST-026","FILTRO DE AIRE PARA GE LISTER PETER",29.0,"UNIDAD","CISTERNA 1","FILTROS"],
    ["CIST-027","WINCHA DE 5M",1.0,"UNIDAD","CISTERNA 1","HERRAMIENTAS"],
    ["CIST-028","SET LLAVE ESTRELLADO TORX",1.0,"UNIDAD","CISTERNA 1","HERRAMIENTAS"],
    ["CIST-029","RODILLO PARA PINTAR DE 03 PULG",6.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-030","RODILLO PARA PINTAR DE 04 PULG",11.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-031","RODILLO PARA PINTAR DE 09 PULG",2.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-032","BROCHA PARA PINTAR DE 2 PULG",3.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-033","BROCHA PARA PINTAR DE 1 PULG",4.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-034","LIMPIA CONTACTOS EN SPRAY",10.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-035","AFLOJATODO WD4O EN SPRAY",6.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-036","LIJA NRO 180",28.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-037","LIJA NRO 50",17.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-038","ELECTRODO PARA SOLDADURA",1.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-039","ESCOBILLA TIPO COPA DE 4 PULG (MARCA: KAMASA)",6.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-040","ESCOBILLA TIPO COPA DE 3 PULG (MARCA: TRUPER)",3.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-041","ESCOBILLA TIPO TRENZADO",2.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-042","TEFLON",4.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-043","LLAVE DE PASO DE 1 PULG",1.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-044","LLAVE DE PASO DE 3/4 PULG",2.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-045","FAJA DE DISTRIBUCION B53",2.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-046","ABRAZADERA PARA MANGUERA DE JEBE Y LONA NRO 94",4.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-047","ABRAZADERA DE ACERO INOX 80-85",2.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-048","HOJA DE SIERRA DE 12 PULG (MARCA: TRUPER)",4.0,"UNIDAD","CISTERNA 1","MANTENIMIENTO"],
    ["CIST-049","CINTA AISLANTE COLOR NEGRO",1.0,"UNIDAD","CISTERNA 1","MAT. ELECTRICIDAD"],
    ["CIST-050","CINTA AISLANTE COLOR ROJO",6.0,"UNIDAD","CISTERNA 1","MAT. ELECTRICIDAD"],
    ["CIST-051","TAZA DE TE",23.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-052","PLATO DE ENTRADA",14.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-053","PLATO SOPERO",1.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-054","VASO DE VIDRIO",6.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-055","CUCHILLOS PARA MESA ACERO INOX.",12.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-056","CUCHARAS PARA SOPA ACERO INOX",4.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-057","CUCHARAS PARA TE ACERO INOX",4.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-058","TENEDORES PARA MESA  ACERO INOX.",4.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-059","CUCHILLOS PARA MESA ACERO INOX.",4.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-060","CUCHILLO GRANDE PARA CARNE 10' MANGO POLIPROPILENO",1.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-061","CUCHILLO GRANDE PARA CARNE 8' MANGO POLIPROPILENO",1.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-062","CUCHARON PARA SOPA DE ACERO INOXIDABLE",1.0,"UNIDAD","CISTERNA 1","MENAJE"],
    ["CIST-063","DISOLVENTE EPOXICO (MARCA: ANYPSA)",73.0,"GALON","CISTERNA 1","PINTURAS"],
    ["CIST-064","PINTURA EPOXICA ZINCROMATO (MARCA: ANYPSA)",10.0,"GALON","CISTERNA 1","PINTURAS"],
    ["CIST-065","CATALIZADOR DE 1/4 (MARCA: ANYPSA)",10.0,"UNIDAD","CISTERNA 1","PINTURAS"]
  ];
}
