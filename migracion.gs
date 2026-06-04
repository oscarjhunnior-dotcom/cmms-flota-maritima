// ============================================================
//  CMMS FLOTA MARÍTIMA — Script de Migración Histórica
//  Carga 246 registros reales 2022-2026 desde Excel
//  Versión: 3.0-MIGRACION
//  INSTRUCCIONES:
//    1. Pegar este archivo como "migracion.gs" en Apps Script
//    2. Ejecutar: inicializarConHistorial()
//    3. Esperar ~30 segundos
//    4. ¡Listo! Ya tienes los datos reales y el sistema operativo
// ============================================================

// ════════════════════════════════════════════════════════════
//  FUNCIÓN PRINCIPAL — ejecutar UNA sola vez
// ════════════════════════════════════════════════════════════
function inicializarConHistorial() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    // Paso 1: Inicializar sistema base
    Logger.log('🚢 Paso 1/3 — Inicializando sistema...');
    inicializarSistema(); // Llama al Code.gs existente
    
    // Paso 2: Registrar técnicos reales como usuarios
    Logger.log('👥 Paso 2/3 — Registrando técnicos...');
    _migrarUsuarios();
    
    // Paso 3: Cargar historial de fallas
    Logger.log('📋 Paso 3/3 — Cargando 246 registros de fallas 2022-2026...');
    const resultado = _migrarFallas();
    
    const msg = `✅ MIGRACIÓN COMPLETADA\n\n` +
                `• Sistema inicializado con 14 hojas\n` +
                `• ${resultado.usuarios} técnicos registrados\n` +
                `• ${resultado.fallas} fallas históricas cargadas (2022-2026)\n` +
                `• 8 embarcaciones activas\n\n` +
                `El sistema está listo para operar.`;
    
    Logger.log(msg);
    ui.alert('✅ Migración exitosa', msg, ui.ButtonSet.OK);
    return { success: true, ...resultado };
    
  } catch(e) {
    const err = '❌ Error en migración: ' + e.message;
    Logger.log(err);
    ui.alert('Error', err, ui.ButtonSet.OK);
    return { error: e.message };
  }
}

// ════════════════════════════════════════════════════════════
//  TÉCNICOS REALES — extraídos del Excel
// ════════════════════════════════════════════════════════════
function _migrarUsuarios() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const h   = ss.getSheetByName('Usuarios');
  const existentes = h.getDataRange().getValues().slice(1).map(r => r[2]); // emails
  
  const tecnicos = [
    ['USR-M001', 'Jacob Medrano',     'jacob.medrano@agrorural.gob.pe',    'TECNICO',    'GUANAY,PANGA',             true],
    ['USR-M002', 'Manuel Calle',      'manuel.calle@agrorural.gob.pe',     'TECNICO',    'ISLA CHINCHA,PANGA',       true],
    ['USR-M003', 'Manuel Lopez',      'manuel.lopez@agrorural.gob.pe',     'TECNICO',    'ISLA CHINCHA,PELICANO',    true],
    ['USR-M004', 'Edison Moran',      'edison.moran@agrorural.gob.pe',     'TECNICO',    'DELFIN 11',                true],
    ['USR-M005', 'Lener Cisneros',    'lener.cisneros@agrorural.gob.pe',   'TECNICO',    'DELFIN 12',                true],
    ['USR-M006', 'Juan Cisneros',     'juan.cisneros@agrorural.gob.pe',    'TECNICO',    'ISLA CHINCHA',             true],
    ['USR-M007', 'Juan Cisnteros',    'juan.cisnteros@agrorural.gob.pe',   'TECNICO',    'ALCATRAZ',                 true],
    ['USR-M008', 'Victor Costilla',   'victor.costilla@agrorural.gob.pe',  'TECNICO',    'CISTERNA 1,PELICANO',      true],
    ['USR-M009', 'Johan Gutierrez',   'johan.gutierrez@agrorural.gob.pe',  'SUPERVISOR', 'TODAS',                    true],
    ['USR-M010', 'Juan Aguilar',      'juan.aguilar@agrorural.gob.pe',     'TECNICO',    'ISLA CHINCHA',             true],
    ['USR-M011', 'Martin Prado',      'martin.prado@agrorural.gob.pe',     'TECNICO',    'CISTERNA 1,PELICANO',      true],
    ['USR-M012', 'Jose Leon',         'jose.leon@agrorural.gob.pe',        'TECNICO',    'DELFIN 12,CISTERNA 1',     true],
    ['USR-M013', 'Jonathan Venegas',  'jonathan.venegas@agrorural.gob.pe', 'TECNICO',    'ISLA CHINCHA',             true],
  ];
  
  let agregados = 0;
  tecnicos.forEach(t => {
    if (!existentes.includes(t[2])) {
      h.appendRow([t[0], t[1], t[2], t[3], t[4], t[5], new Date()]);
      agregados++;
    }
  });
  
  Logger.log(`Técnicos agregados: ${agregados}`);
  return agregados;
}

// ════════════════════════════════════════════════════════════
//  FALLAS HISTÓRICAS — 246 registros reales 2022-2026
// ════════════════════════════════════════════════════════════
function _migrarFallas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('Registro_Fallas');
  
  // Verificar si ya fue migrado
  if (h.getLastRow() > 1) {
    const ui = SpreadsheetApp.getUi();
    const resp = ui.alert(
      '⚠️ Datos existentes',
      `La hoja Registro_Fallas ya tiene ${h.getLastRow()-1} registros.\n\n¿Deseas agregar el historial encima? (No se borrarán los existentes)`,
      ui.ButtonSet.YES_NO
    );
    if (resp !== ui.Button.YES) return { fallas: 0, mensaje: 'Migración cancelada por usuario' };
  }
  
  // ── 246 REGISTROS REALES DEL EXCEL ─────────────────────────
  // Columnas: ID_FALLA, USUARIO, ARMADOR, CODIGO_FALLA, CENTRO_MANT,
  //           UBICACION_TECNICA, EQUIPO, SISTEMA, HORAS_OP, LISTA_REPUESTOS,
  //           PLAN_MANT, FECHA_REGISTRO, NOTIFICACION, DESCRIPCION_FALLA,
  //           ORDEN_MANT, FECHA_CIERRE, EQUIPO_REPARADO, MES, ANO,
  //           TIEMPO_FALLA_DIAS, SEVERIDAD, CAUSA_RAIZ, ACCION_TOMADA, COSTO_ESTIMADO
  const FALLAS = [
["FL-HIST-001","Jacob Medrano","Agrorural","F-HIST-001","CFMYT","GUANAY","Electrobomba de Achique","Equipos Auxiliares",0,"","Correctivo","2022-01-04","","Obstrucción por sólidos o desgaste del impulsor por cavitación","Caja Chica","","OK",1,2022,2.0,"MEDIA","","",0],
["FL-HIST-002","Manuel Calle","Agrorural","F-HIST-002","CFMYT","ISLA CHINCHA","Grua 1-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2022-01-21","","Desalineación de Linea de Carga,Se necesita reparacion del carrete","Caja Chica","","OK",1,2022,10.0,"MEDIA","","",0],
["FL-HIST-003","Manuel Calle","Agrorural","F-HIST-003","CFMYT","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2022-01-23","","Falla en válvula solenoide del sistema hidráulico","tdr-8iut","","NO",1,2022,28.0,"MEDIA","","",0],
["FL-HIST-004","Manuel Calle","Agrorural","F-HIST-004","CFMYT","ISLA CHINCHA","Grua 1-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2022-02-04","","Falla en la tuerca de 1 1/16 de la cañeria.","tdr-8iut","","OK",2,2022,20.0,"MEDIA","","",0],
["FL-HIST-005","Manuel Lopez","Agrorural","F-HIST-005","CFMYT","ISLA CHINCHA","Motor Hidráulico 1","Sistema de Grua",0,"","Correctivo","2022-02-04","","Fuga de aceite por hermeto de cañeria","tdr-8iut","","OK",2,2022,20.0,"MEDIA","","",0],
["FL-HIST-006","Edison Moran","Agrorural","F-HIST-006","CFMYT","DELFIN 11","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-02-16","","Desprogramación del sistema de inyección","tdr-8iut","","OK",2,2022,60.0,"MEDIA","","",0],
["FL-HIST-007","Lener Cisneros","Agrorural","F-HIST-007","CFMYT","DELFIN 12","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-02-25","","Presencia de Humo Negro en Salida de escape","tdr-8iut","","OK",2,2022,30.0,"MEDIA","","",0],
["FL-HIST-008","Manuel Calle","Agrorural","F-HIST-008","CFMYT","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2022-02-24","","Desalineación de ejes en motor auxiliar","tdr-8iut","","NO",2,2022,28.0,"MEDIA","","",0],
["FL-HIST-009","Juan Cisneros","Agrorural","F-HIST-009","CFMYT","ISLA CHINCHA","Luces de Navegacion","Cubierta y Casco",0,"","Correctivo","2022-02-27","","Falla en la excitatriz del generador","Caja Chica","","OK",2,2022,2.0,"MEDIA","","",0],
["FL-HIST-010","Juan Aguilar","Agrorural","F-HIST-010","CFMYT","ISLA CHINCHA","Motor Eléctrico 1","Sistema de Grua",0,"","Correctivo","2022-03-02","","Falla por falso contacto del rele de supervision de corriente","Caja Chica","","OK",3,2022,5.0,"MEDIA","","",0],
["FL-HIST-011","Martin Prado","Agrorural","F-HIST-011","CFMYT","CISTERNA 1","Electrobomba de Contraincendio","Cubierta y Casco",0,"","Correctivo","2022-03-03","","Vibración excesiva en soporte de motor","Caja Chica","","OK",3,2022,2.0,"MEDIA","","",0],
["FL-HIST-012","Manuel Calle","Agrorural","F-HIST-012","CFMYT","ISLA CHINCHA","Grua 2-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2022-03-08","","Falla en los valvulas y pistones","tdr-8iut","","OK",3,2022,20.0,"MEDIA","","",0],
["FL-HIST-013","Manuel Lopez","Agrorural","F-HIST-013","CFMYT","ISLA CHINCHA","Motor Hidráulico 1","Sistema de Grua",0,"","Correctivo","2022-03-16","","Falla de corrosion en el tanque de expansion","tdr-8iut","","OK",3,2022,20.0,"MEDIA","","",0],
["FL-HIST-014","Manuel Calle","Agrorural","F-HIST-014","CFMYT","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2022-03-18","","Fuga de combustible en sistema de inyección","tdr-8iut","","NO",3,2022,28.0,"MEDIA","","",0],
["FL-HIST-015","Juan Aguilar","Agrorural","F-HIST-015","CFMYT","ISLA CHINCHA","Motor Eléctrico 1","Sistema de Grua",0,"","Correctivo","2022-03-27","","Corte en cableado del sistema de gobierno","Caja Chica","","OK",3,2022,5.0,"MEDIA","","",0],
["FL-HIST-016","Manuel Calle","Agrorural","F-HIST-016","CFMYT","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-03-15","","Falla en la inyeccion del motor cummins","tdr-8iut","","OK",3,2022,60.0,"MEDIA","","",0],
["FL-HIST-017","Edison Moran","Agrorural","F-HIST-017","CFMYT","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-04-11","","Temperatura elevada en el rodamiento del motor","tdr-8iut","","OK",4,2022,15.0,"MEDIA","","",0],
["FL-HIST-018","Juan Cisneros","Agrorural","F-HIST-018","CFMYT","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-05-11","","Falla en la cañeria de refrigeracion del motor","tdr-8iut","","",5,2022,35.0,"MEDIA","","",0],
["FL-HIST-019","Manuel Lopez","Agrorural","F-HIST-019","CFMYT","ISLA CHINCHA","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2022-04-12","","Las valvulas de los baños estan defectuosas","tdr-8iut","","OK",4,2022,90.0,"MEDIA","","",0],
["FL-HIST-020","Manuel Lopez","Agrorural","F-HIST-020","CFMYT","ISLA CHINCHA","Grua 2-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2022-04-18","","Falla en la valvula de seguridad, (relief)","tdr-8iut","","OK",4,2022,20.0,"MEDIA","","",0],
["FL-HIST-021","Manuel Calle","Agrorural","F-HIST-021","CFMYT","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-04-20","","Falla de bateria, se baja la bateria bornes oxidados","tdr-8iut","","OK",4,2022,15.0,"MEDIA","","",0],
["FL-HIST-022","Juan Cisnteros","Agrorural","F-HIST-022","CFMYT","ALCATRAZ","Casco Principal","Cubierta y Casco",0,"","Preventivo","2022-04-27","","oxidacion en la bita de popa,se necesita realizar el carenado","Caja Chica","","OK",4,2022,15.0,"MEDIA","","",0],
["FL-HIST-023","Juan Cisnteros","Agrorural","F-HIST-023","CFMYT","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-04-28","","Temperatura elevada en el aceite del motor","tdr-8iut","","OK",4,2022,25.0,"MEDIA","","",0],
["FL-HIST-024","Manuel Lopez","Agrorural","F-HIST-024","CFMYT","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-05-01","","Vibración excesiva en soporte de motor","tdr-8iut","","OK",5,2022,10.0,"MEDIA","","",0],
["FL-HIST-025","Manuel Calle","Agrorural","F-HIST-025","CFMYT","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-05-23","","Presencia de humo negro en la salida de escape. Falta revision de inyectores","tdr-8iut","","OK",5,2022,20.0,"MEDIA","","",0],
["FL-HIST-026","Edison Moran","Agrorural","F-HIST-026","CFMYT","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-05-05","","Falla en los inyectores de motores, presencia de humo negro","Caja Chica","","OK",5,2022,30.0,"MEDIA","","",0],
["FL-HIST-027","Lener Cisneros","Agrorural","F-HIST-027","CFMYT","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-06-07","","Falla en la salida de escape, posible fisura","tdr-8iut","","",6,2022,60.0,"MEDIA","","",0],
["FL-HIST-028","Victor Costilla","Agrorural","F-HIST-028","CFMYT","CISTERNA 1","ElectroBomba centrifuga aguas grises","Equipos Auxiliares",0,"","Correctivo","2022-05-08","","Vibración excesiva en soporte de motor","tdr-8iut","","OK",5,2022,125.0,"MEDIA","","",0],
["FL-HIST-029","Edison Moran","Agrorural","F-HIST-029","CFMYT","DELFIN 11","Casco Principal","Cubierta y Casco",0,"","Preventivo","2022-05-10","","casco de embarcacion lleno de inscrustaciones","tdr-8iut","","OK",5,2022,90.0,"MEDIA","","",0],
["FL-HIST-030","Edison Moran","Agrorural","F-HIST-030","CFMYT","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-05-16","","Presencia de humo negro, posible fallo de inyectores","tdr-8iut","","OK",5,2022,60.0,"MEDIA","","",0],
["FL-HIST-031","Edison Moran","Agrorural","F-HIST-031","CFMYT","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-05-16","","Falla en la tuberia de petroleo(cañeria)","Caja Chica","","OK",5,2022,90.0,"MEDIA","","",0],
["FL-HIST-032","Manuel Calle","Agrorural","F-HIST-032","CFMYT","PANGA","Casco Principal","Cubierta y Casco",0,"","Correctivo","2022-06-18","","la panga necesita con urgencia limpieza de casco, el avances tiene dificultad","Caja Chica","","OK",6,2022,12.0,"MEDIA","","",0],
["FL-HIST-033","Lener Cisneros","Agrorural","F-HIST-033","CFMYT","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-07-10","","Falla de cambio de filtro primario de petroleo","tdr-8iut","","OK",7,2022,15.0,"MEDIA","","",0],
["FL-HIST-034","Lener Cisneros","Agrorural","F-HIST-034","CFMYT","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-05-16","","Falla en cable de morce de arranque","tdr-8iut","","OK",5,2022,25.0,"MEDIA","","",0],
["FL-HIST-035","Jacob Medrano","Agrorural","F-HIST-035","CFMYT","GUANAY","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2022-05-25","","Error de sincronismo en grupo electrógeno","tdr-8iut","","NO",5,2022,2.0,"MEDIA","","",0],
["FL-HIST-036","Manuel Lopez","Agrorural","F-HIST-036","CFMYT","ISLA CHINCHA","Motor Eléctrico 2","Sistema de Grua",0,"","Correctivo","2022-05-31","","Transformador recalentado a 90°C","tdr-8iut","","OK",5,2022,20.0,"MEDIA","","",0],
["FL-HIST-037","Victor Costilla","Agrorural","F-HIST-037","CFMYT","CISTERNA 1","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-06-01","","Falla en válvula solenoide del sistema hidráulico","tdr-8iut","","OK",6,2022,80.0,"MEDIA","","",0],
["FL-HIST-038","Johan Gutierrez","Agrorural","F-HIST-038","CFMYT","CISTERNA 1","Cubierta Principal","Equipos Auxiliares",0,"","Correctivo","2022-06-04","","Sobrecalentamiento en tablero de distribución","Caja Chica","","OK",6,2022,5.0,"MEDIA","","",0],
["FL-HIST-039","Edison Moran","Agrorural","F-HIST-039","CFMYT","DELFIN 11","Cabre-estante","Sistema de Fondeo",0,"","Correctivo","2022-06-21","","Corte en cableado del sistema de gobierno","tdr-8iut","","OK",6,2022,90.0,"MEDIA","","",0],
["FL-HIST-040","Lener Cisneros","Agrorural","F-HIST-040","CFMYT","DELFIN 12","Bomba de Achique","Equipos Auxiliares",0,"","Correctivo","2022-08-02","","Falla en los impulsores de la bomba","tdr-8iut","","",8,2022,40.0,"MEDIA","","",0],
["FL-HIST-041","Juan Cisnteros","Agrorural","F-HIST-041","CFMYT","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-06-22","","Falla Electrica en alternador","tdr-8iut","","OK",6,2022,35.0,"MEDIA","","",0],
["FL-HIST-042","Manuel Lopez","Agrorural","F-HIST-042","CFMYT","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-07-04","","Falla SPN 29 FMI 4","tdr-8iut","","OK",7,2022,60.0,"MEDIA","","",0],
["FL-HIST-043","Manuel Calle","Agrorural","F-HIST-043","CFMYT","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-07-09","","Desgaste de escobillas en generador","tdr-8iut","","NO",7,2022,15.0,"MEDIA","","",0],
["FL-HIST-044","Manuel Lopez","Agrorural","F-HIST-044","CFMYT","ISLA CHINCHA","Motor Hidráulico 1","Sistema de Grua",0,"","Correctivo","2022-07-14","","Fuga de Aceite Hidraulico en tuberia de 1/2","tdr-8iut","","OK",7,2022,30.0,"MEDIA","","",0],
["FL-HIST-045","Manuel Calle","Agrorural","F-HIST-045","CFMYT","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2022-07-16","","Interferencia en sistema de navegación automatizado","tdr-8iut","","NO",7,2022,15.0,"MEDIA","","",0],
["FL-HIST-046","Manuel Calle","Agrorural","F-HIST-046","","PANGA","Bomba de Achique","Equipos Auxiliares",0,"","Correctivo","2022-07-22","","Falla en la bomba de achique, tiene dañado el embolo de bombeo","tdr-8iut","","OK",7,2022,30.0,"MEDIA","","",0],
["FL-HIST-047","Manuel Lopez","Agrorural","F-HIST-047","Flota","CISTERNA 1","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2022-07-25","","Cortocircuito en caja de distribución de puente","Caja Chica","","OK",7,2022,15.0,"MEDIA","","",0],
["FL-HIST-048","Jacob Medrano","Agrorural","F-HIST-048","Flota","GUANAY","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-07-26","","Vibracion del motor excesivo y inestabilidad de bomba de inyeccion","tdr-8iut","","OK",7,2022,2.0,"MEDIA","","",0],
["FL-HIST-049","Jacob Medrano","Agrorural","F-HIST-049","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-08-03","","Anomalía en sensor de velocidad de eje","tdr-8iut","","OK",8,2022,2.0,"MEDIA","","",0],
["FL-HIST-050","Johan Gutierrez","Agrorural","F-HIST-050","Flota","CISTERNA 1","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-08-12","","Pérdida de presión de aceite en el motor principal","Caja Chica","","OK",8,2022,2.0,"MEDIA","","",0],
["FL-HIST-051","Jacob Medrano","Agrorural","F-HIST-051","Flota","GUANAY","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2022-08-18","","Codigo de falla SD START FAIL, arrancador inoperativo","tdr-8iut","","OK",8,2022,2.0,"MEDIA","","",0],
["FL-HIST-052","Jacob Medrano","Agrorural","F-HIST-052","Flota","PANGA","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2022-09-19","","Exceo de oxido en la cubierta y la zona de espejo de popa","tdr-8iut","","OK",9,2022,10.0,"MEDIA","","",0],
["FL-HIST-053","Juan Cisneros","Agrorural","F-HIST-053","Flota","ISLA CHINCHA","Luces de Navegacion","Cubierta y Casco",0,"","Correctivo","2022-08-20","","Falla en la excitatriz del generador","Caja Chica","","OK",8,2022,2.0,"MEDIA","","",0],
["FL-HIST-054","Johan Gutierrez","Agrorural","F-HIST-054","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-09-12","","Desgaste de escobillas en generador","tdr-8iut","","OK",9,2022,2.0,"MEDIA","","",0],
["FL-HIST-055","Jacob Medrano","Agrorural","F-HIST-055","Flota","GUANAY","Unidad Hidralica","Sistema de Fondeo",0,"","Correctivo","2022-09-20","","Falla común operativa asociada al desgaste o ambiente marino","tdr-8iut","","OK",9,2022,2.0,"MEDIA","","",0],
["FL-HIST-056","Manuel Calle","Agrorural","F-HIST-056","Flota","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-09-24","","Falla de activo SPN 970 FMI 31","tdr-8iut","","OK",9,2022,60.0,"MEDIA","","",0],
["FL-HIST-057","Johan Gutierrez","Agrorural","F-HIST-057","Flota","CISTERNA 1","Eje de Propulsion","Sistema de Propulsion",0,"","Correctivo","2022-09-24","","Interferencia en sistema de navegación automatizado","Garantia","","NO",9,2022,20.0,"MEDIA","","",0],
["FL-HIST-058","Manuel Calle","Agrorural","F-HIST-058","Flota","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2022-09-24","","Falla en la faja de la bomba de enfriamiento y alternador","tdr-8iut","","OK",9,2022,20.0,"MEDIA","","",0],
["FL-HIST-059","Manuel Calle","Agrorural","F-HIST-059","Flota","ISLA CHINCHA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-09-27","","Incremento de temperatura de aceite del motor. Un 25%","tdr-8iut","","OK",9,2022,20.0,"MEDIA","","",0],
["FL-HIST-060","Juan Cisnteros","Agrorural","F-HIST-060","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-09-28","","Presencia de Humo negro en la salida de escape, posible daño de inyectores","tdr-8iut","","OK",9,2022,40.0,"MEDIA","","",0],
["FL-HIST-061","Manuel Calle","Agrorural","F-HIST-061","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-10-15","","Fallta del arrancador. La conexión de la llave esta operando mal","tdr-8iut","","OK",10,2022,25.0,"MEDIA","","",0],
["FL-HIST-062","Edison Moran","Agrorural","F-HIST-062","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-10-07","","Falla en el cable Morse, para el avance y el retroceso","tdr-8iut","","OK",10,2022,25.0,"MEDIA","","",0],
["FL-HIST-063","Lener Cisneros","Agrorural","F-HIST-063","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-08-20","","Fisura existente en la tubo de escape de silenciador","","","OK",8,2022,50.0,"MEDIA","","",0],
["FL-HIST-064","Martin Prado","Agrorural","F-HIST-064","Flota","CISTERNA 1","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2022-10-16","","Cortocircuito en caja de distribución de puente","Garantia","","OK",10,2022,120.0,"MEDIA","","",0],
["FL-HIST-065","Johan Gutierrez","Agrorural","F-HIST-065","Flota","CISTERNA 1","Tablero de Comunicaciones","Cubierta y Casco",0,"","Correctivo","2022-10-27","","Desgaste de escobillas en generador","Garantia","","OK",10,2022,120.0,"MEDIA","","",0],
["FL-HIST-066","Victor Costilla","Agrorural","F-HIST-066","Flota","CISTERNA 1","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2022-11-04","","Interferencia en sistema de navegación automatizado","Caja Chica","","OK",11,2022,15.0,"MEDIA","","",0],
["FL-HIST-067","Manuel Lopez","Agrorural","F-HIST-067","Flota","CISTERNA 1","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-11-09","","Sobrecalentamiento en tablero de distribución","Garantia","","OK",11,2022,20.0,"MEDIA","","",0],
["FL-HIST-068","Manuel Calle","Agrorural","F-HIST-068","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-11-18","","Fisura en la manguera de refrigeracion","Caja Chica","","OK",11,2022,15.0,"MEDIA","","",0],
["FL-HIST-069","Edison Moran","Agrorural","F-HIST-069","Flota","DELFIN 11","Cabre-estante","Sistema de Fondeo",0,"","Correctivo","2022-11-25","","Desgaste en el tubo de guiador de parte superior del winche","tdr-8iut","","OK",11,2022,90.0,"MEDIA","","",0],
["FL-HIST-070","Johan Gutierrez","Agrorural","F-HIST-070","Flota","CISTERNA 1","Electrobomba de Achique","Sistema de Achique /Lastre",0,"","Correctivo","2022-11-25","","Temperatura elevada en el rodamiento del motor","tdr-8iut","","OK",11,2022,125.0,"MEDIA","","",0],
["FL-HIST-071","Juan Cisnteros","Agrorural","F-HIST-071","Flota","ALCATRAZ","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-12-06","","Cortocircuito en caja del tablero principal","tdr-8iut","","OK",12,2022,35.0,"MEDIA","","",0],
["FL-HIST-072","Edison Moran","Agrorural","F-HIST-072","Flota","DELFIN 11","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-12-18","","Pérdida de presión de aceite en el motor principal","tdr-8iut","","OK",12,2022,120.0,"MEDIA","","",0],
["FL-HIST-073","Manuel Calle","Agrorural","F-HIST-073","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-12-20","","Inundacion de sala de maquinas, falla total de motor principal","Caja Chica","","OK",12,2022,10.0,"MEDIA","","",0],
["FL-HIST-074","Lener Cisneros","Agrorural","F-HIST-074","Flota","DELFIN 12","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-10-15","","Oxido permanente en las culata y en panel de control","tdr-8iut","","",10,2022,25.0,"MEDIA","","",0],
["FL-HIST-075","Manuel Calle","Agrorural","F-HIST-075","Flota","ISLA CHINCHA","Grua 1-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2022-12-21","","Cortocircuito en caja de distribución de puente","Caja Chica","","OK",12,2022,30.0,"MEDIA","","",0],
["FL-HIST-076","Manuel Calle","Agrorural","F-HIST-076","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-01-02","","Motor totalmente inoperativo","Caja Chica","","OK",1,2023,340.0,"MEDIA","","",0],
["FL-HIST-077","Martin Prado","Agrorural","F-HIST-077","Flota","CISTERNA 1","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2023-01-14","","Falla en la excitatriz del generador","Garantia","","OK",1,2023,120.0,"MEDIA","","",0],
["FL-HIST-078","Edison Moran","Agrorural","F-HIST-078","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-01-16","","Fisura de las valvulas en el resorte de calibracion","tdr-8iut","","OK",1,2023,90.0,"MEDIA","","",0],
["FL-HIST-079","Manuel Calle","Agrorural","F-HIST-079","Flota","ISLA CHINCHA","Grua 2-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2023-01-22","","Fuga de combustible en sistema de inyección","tdr-8iut","","OK",1,2023,20.0,"MEDIA","","",0],
["FL-HIST-080","Victor Costilla","Agrorural","F-HIST-080","Flota","CISTERNA 1","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-01-27","","Falla en válvula solenoide del sistema hidráulico","Caja Chica","","OK",1,2023,2.0,"MEDIA","","",0],
["FL-HIST-081","Johan Gutierrez","Agrorural","F-HIST-081","Flota","CISTERNA 1","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-02-02","","Vibración excesiva en soporte de motor","Caja Chica","","NO",2,2023,66.0,"MEDIA","","",0],
["FL-HIST-082","Lener Cisneros","Agrorural","F-HIST-082","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-02-22","","Desgaste de escobillas en generador","Caja Chica","","OK",2,2023,60.0,"MEDIA","","",0],
["FL-HIST-083","Jacob Medrano","Agrorural","F-HIST-083","Flota","GUANAY","Electrobomba de Achique","Equipos Auxiliares",0,"","Correctivo","2023-02-28","","Desbalanceo en alternador del grupo BR","Caja Chica","","OK",2,2023,3.0,"MEDIA","","",0],
["FL-HIST-084","Manuel Lopez","Agrorural","F-HIST-084","Flota","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2023-03-05","","Anomalía en sensor de velocidad de eje","tdr-8iut","","OK",3,2023,20.0,"MEDIA","","",0],
["FL-HIST-085","Manuel Calle","Agrorural","F-HIST-085","Flota","ISLA CHINCHA","Cargador de Baterias","Sistema de Energia",0,"","Correctivo","2023-03-10","","Desbalanceo en alternador del grupo BR","tdr-8iut","","NO",3,2023,10.0,"MEDIA","","",0],
["FL-HIST-086","Manuel Lopez","Agrorural","F-HIST-086","Flota","ISLA CHINCHA","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2023-03-13","","Falla en válvula solenoide del sistema hidráulico","tdr-8iut","","OK",3,2023,35.0,"MEDIA","","",0],
["FL-HIST-087","Johan Gutierrez","Agrorural","F-HIST-087","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2023-04-04","","Desgaste de escobillas en generador","tdr-8iut","","OK",4,2023,3.0,"MEDIA","","",0],
["FL-HIST-088","Juan Cisnteros","Agrorural","F-HIST-088","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-04-16","","Abertura de pernos de carcaza del arrancador","Caja Chica","","OK",4,2023,30.0,"MEDIA","","",0],
["FL-HIST-089","Manuel Calle","Agrorural","F-HIST-089","Flota","ALCATRAZ","Caja reductora Unico","Sistema de Propulsion",0,"","Correctivo","2022-07-15","","Cabeceo en el caja de engranajes","tdr-8iut","","OK",7,2022,90.0,"MEDIA","","",0],
["FL-HIST-090","Juan Cisnteros","Agrorural","F-HIST-090","Flota","ALCATRAZ","Caja reductora Unico","Sistema de Propulsion",0,"","Correctivo","2023-04-17","","Falla en la marcha retroceso y avance de la caja de transmision","tdr-8iut","","OK",4,2023,30.0,"MEDIA","","",0],
["FL-HIST-091","Lener Cisneros","Agrorural","F-HIST-091","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-04-20","","Pérdida de presión de aceite en el motor principal","Caja Chica","","OK",4,2023,60.0,"MEDIA","","",0],
["FL-HIST-092","Edison Moran","Agrorural","F-HIST-092","Flota","DELFIN 11","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2023-04-21","","Falla en el motor del congelador","Caja Chica","","OK",4,2023,120.0,"MEDIA","","",0],
["FL-HIST-093","Juan Cisnteros","Agrorural","F-HIST-093","Flota","ALCATRAZ","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2023-04-25","","Quemadura del estator","Caja Chica","","OK",4,2023,30.0,"MEDIA","","",0],
["FL-HIST-094","Edison Moran","Agrorural","F-HIST-094","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-04-29","","Vibración excesiva en soporte de motor","Caja Chica","","OK",4,2023,20.0,"MEDIA","","",0],
["FL-HIST-095","Juan Cisnteros","Agrorural","F-HIST-095","Flota","ALCATRAZ","Caja reductora Unico","Sistema de Propulsion",0,"","Correctivo","2023-05-02","","Corte en cableado del sistema de gobierno","tdr-8iut","","OK",5,2023,30.0,"MEDIA","","",0],
["FL-HIST-096","Jacob Medrano","Agrorural","F-HIST-096","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-05-10","","Presencia de humo negro en navegacion y cambio de sonido por escape","tdr-8iut","","OK",5,2023,3.0,"MEDIA","","",0],
["FL-HIST-097","Juan Cisnteros","Agrorural","F-HIST-097","Flota","ALCATRAZ","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2023-05-10","","Pérdida de presión de aceite en el motor principal","Caja Chica","","OK",5,2023,3.0,"MEDIA","","",0],
["FL-HIST-098","Jose Leon","Agrorural","F-HIST-098","Flota","DELFIN 12","Ancla","Sistema de Fondeo",0,"","Correctivo","2023-05-12","","Desbalanceo en alternador del grupo BR","Caja Chica","","OK",5,2023,25.0,"MEDIA","","",0],
["FL-HIST-099","Johan Gutierrez","Agrorural","F-HIST-099","Flota","PELICANO","Equipos de Comunicación","Cubierta y Casco",0,"","Correctivo","2023-06-20","","Falla en el equipo del clima","Caja Chica","","NO",6,2023,50.0,"MEDIA","","",0],
["FL-HIST-100","Edison Moran","Agrorural","F-HIST-100","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-06-21","","Falla en válvula solenoide del sistema hidráulico","tdr-8iut","","OK",6,2023,20.0,"MEDIA","","",0],
["FL-HIST-101","Juan Cisneros","Agrorural","F-HIST-101","Flota","ISLA CHINCHA","Winche de estribor","Sistema de Fondeo",0,"","Correctivo","2023-06-25","","Desprogramación del sistema de inyección","Caja Chica","","NO",6,2023,20.0,"MEDIA","","",0],
["FL-HIST-102","Juan Aguilar","Agrorural","F-HIST-102","Flota","ISLA CHINCHA","Grua 2-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2023-06-29","","Fallta en los terminales y niples de la valvulas","tdr-8iut","","OK",6,2023,15.0,"MEDIA","","",0],
["FL-HIST-103","Juan Aguilar","Agrorural","F-HIST-103","Flota","ISLA CHINCHA","Grua 1-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2023-06-29","","Fallta en las valvula contralbaalm","Caja Chica","","OK",6,2023,15.0,"MEDIA","","",0],
["FL-HIST-104","Jacob Medrano","Agrorural","F-HIST-104","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-07-02","","Error en módulo de control electrónico (ECU)","tdr-8iut","","NO",7,2023,3.0,"MEDIA","","",0],
["FL-HIST-105","Lener Cisneros","Agrorural","F-HIST-105","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-07-11","","Temperatura elevada en el rodamiento del motor","Caja Chica","","OK",7,2023,60.0,"MEDIA","","",0],
["FL-HIST-106","Victor Costilla","Agrorural","F-HIST-106","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2023-07-15","","Falla en válvula solenoide del sistema hidráulico","tdr-8iut","","OK",7,2023,35.0,"MEDIA","","",0],
["FL-HIST-107","Jacob Medrano","Agrorural","F-HIST-107","Flota","GUANAY","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2023-08-09","","Desbalanceo en alternador del grupo BR","tdr-8iut","","NO",8,2023,3.0,"MEDIA","","",0],
["FL-HIST-108","Jacob Medrano","Agrorural","F-HIST-108","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-08-11","","Presencia de humo negro y bajas revoluciones","tdr-8iut","","NO",8,2023,3.0,"MEDIA","","",0],
["FL-HIST-109","Johan Gutierrez","Agrorural","F-HIST-109","Flota","PELICANO","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2023-08-21","","Desprogramación del sistema de inyección","Caja Chica","","NO",8,2023,30.0,"MEDIA","","",0],
["FL-HIST-110","Manuel Calle","Agrorural","F-HIST-110","Flota","ISLA CHINCHA","Grua 2-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2023-08-27","","Anomalía en sensor de velocidad de eje","tdr-8iut","","OK",8,2023,30.0,"MEDIA","","",0],
["FL-HIST-111","Johan Gutierrez","Agrorural","F-HIST-111","Flota","PELICANO","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2023-09-14","","Error en módulo de control electrónico (ECU)","Caja Chica","","NO",9,2023,30.0,"MEDIA","","",0],
["FL-HIST-112","Juan Aguilar","Agrorural","F-HIST-112","Flota","ISLA CHINCHA","Grua 2-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2023-09-18","","Corte en cableado del sistema de gobierno","tdr-8iut","","OK",9,2023,30.0,"MEDIA","","",0],
["FL-HIST-113","Johan Gutierrez","Agrorural","F-HIST-113","Flota","PELICANO","Tablero de Comunicaciones","Cubierta y Casco",0,"","Correctivo","2023-09-19","","Desgaste de escobillas en generador","Garantia","","OK",9,2023,20.0,"MEDIA","","",0],
["FL-HIST-114","Jacob Medrano","Agrorural","F-HIST-114","Flota","GUANAY","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2023-09-21","","Falta de Estabilidad a la bomba de inyeccion","tdr-8iut","","OK",9,2023,3.0,"MEDIA","","",0],
["FL-HIST-115","Manuel Lopez","Agrorural","F-HIST-115","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2023-10-02","","Pérdida de presión de aceite en el motor principal","Caja Chica","","OK",10,2023,15.0,"MEDIA","","",0],
["FL-HIST-116","Victor Costilla","Agrorural","F-HIST-116","Flota","PELICANO","ElectroBomba centrifuga aguas grises","Equipos Auxiliares",0,"","Correctivo","2023-10-14","","Pérdida de presión de aceite en el motor principal","tdr-8iut","","OK",10,2023,25.0,"MEDIA","","",0],
["FL-HIST-117","Jacob Medrano","Agrorural","F-HIST-117","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-10-17","","falla en la cadena del saca filtro","Caja Chica","","OK",10,2023,3.0,"MEDIA","","",0],
["FL-HIST-118","Manuel Calle","Agrorural","F-HIST-118","Flota","ISLA CHINCHA","Electrobomba de Achique","Sistema de Achique /Lastre",0,"","Correctivo","2023-10-24","","Falla en el sonito en el interior de la bomba","tdr-8iut","","OK",10,2023,30.0,"MEDIA","","",0],
["FL-HIST-119","Edison Moran","Agrorural","F-HIST-119","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-10-27","","Error de sincronismo en grupo electrógeno","Caja Chica","","OK",10,2023,20.0,"MEDIA","","",0],
["FL-HIST-120","Jacob Medrano","Agrorural","F-HIST-120","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2023-11-12","","Sobrecalentamiento en tablero de distribución","tdr-8iut","","OK",11,2023,3.0,"MEDIA","","",0],
["FL-HIST-121","Johan Gutierrez","Agrorural","F-HIST-121","Flota","PELICANO","Equipos de Comunicación","Cubierta y Casco",0,"","Correctivo","2023-11-21","","television defectuoso","Caja Chica","","NO",11,2023,10.0,"MEDIA","","",0],
["FL-HIST-122","Juan Cisnteros","Agrorural","F-HIST-122","Flota","ALCATRAZ","Casco Principal","Cubierta y Casco",0,"","Preventivo","2023-11-30","","Limpieza de Casco por Incrustaciones Marinas","Caja Chica","","OK",11,2023,15.0,"MEDIA","","",0],
["FL-HIST-123","Juan Cisnteros","Agrorural","F-HIST-123","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-11-30","","Fuga de combustible en sistema de inyección","tdr-8iut","","OK",11,2023,30.0,"MEDIA","","",0],
["FL-HIST-124","Jacob Medrano","Agrorural","F-HIST-124","Flota","GUANAY","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-01-18","","Presencia de humo negro en navegacion","tdr-8iut","","OK",1,2024,60.0,"MEDIA","","",0],
["FL-HIST-125","Manuel Lopez","Agrorural","F-HIST-125","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Preventivo","2024-02-13","","Mantenimeinto 1734 horas","Caja Chica","","OK",2,2024,5.0,"MEDIA","","",0],
["FL-HIST-126","Jacob Medrano","Agrorural","F-HIST-126","Flota","GUANAY","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-02-20","","Faltar de estabilidad y perdidad de potencia","tdr-8iut","","OK",2,2024,90.0,"MEDIA","","",0],
["FL-HIST-127","Johan Gutierrez","Agrorural","F-HIST-127","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2024-03-20","","Fallta en el impeller de la bomba de agua salada","tdr-8iut","","OK",3,2024,10.0,"MEDIA","","",0],
["FL-HIST-128","Lener Cisneros","Agrorural","F-HIST-128","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-04-06","","Falla en la cañeria de agua de enfriamiento","Caja Chica","","OK",4,2024,30.0,"MEDIA","","",0],
["FL-HIST-129","Johan Gutierrez","Agrorural","F-HIST-129","Flota","GUANAY","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-04-09","","Falta de presion de aceite y movimiento erraticos equipo inoperativo","tdr-8iut","","OK",4,2024,90.0,"MEDIA","","",0],
["FL-HIST-130","Johan Gutierrez","Agrorural","F-HIST-130","Flota","GUANAY","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-04-23","","falla en la bomba de combustible, presencia de agua en los embolos","tdr-8iut","","OK",4,2024,252.0,"MEDIA","","",0],
["FL-HIST-131","Johan Gutierrez","Agrorural","F-HIST-131","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2024-04-27","","Alta temperatura de refrigeracion 195 °F. falla en el impeller","tdr-8iut","","OK",4,2024,90.0,"MEDIA","","",0],
["FL-HIST-132","Jacob Medrano","Agrorural","F-HIST-132","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-04-30","","Falla en arrancador, perno descolgado de la volante.","tdr-8iut","","OK",4,2024,245.0,"MEDIA","","",0],
["FL-HIST-133","Victor Costilla","Agrorural","F-HIST-133","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-05-03","","Falla en el gobernador de voltaje","tdr-8iut","","OK",5,2024,5.0,"MEDIA","","",0],
["FL-HIST-134","Manuel Lopez","Agrorural","F-HIST-134","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2024-05-06","","Fallta en el reductor de 1/2 a 2.5","Caja Chica","","OK",5,2024,5.0,"MEDIA","","",0],
["FL-HIST-135","Manuel Lopez","Agrorural","F-HIST-135","Flota","PELICANO","Motor Principal BR","Sistema de Energia",0,"","Correctivo","2024-05-08","","Ruptura de Anodos de Zinc de Enfriamiento","Caja Chica","","OK",5,2024,5.0,"MEDIA","","",0],
["FL-HIST-136","Jacob Medrano","Agrorural","F-HIST-136","Flota","GUANAY","Electrobomba de Achique","Equipos Auxiliares",0,"","Correctivo","2024-05-09","","falla en la bomba de descarga al mar, vastago roto","Caja Chica","","OK",5,2024,90.0,"MEDIA","","",0],
["FL-HIST-137","Johan Gutierrez","Agrorural","F-HIST-137","Flota","PELICANO","Equipos de Comunicación","Cubierta y Casco",0,"","Correctivo","2024-05-15","","Falla electrico y/o sonoro","Caja Chica","","NO",5,2024,5.0,"MEDIA","","",0],
["FL-HIST-138","Jacob Medrano","Agrorural","F-HIST-138","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2024-05-16","","Alta temperatura de refrigeracion. Impeller roto","tdr-8iut","","OK",5,2024,90.0,"MEDIA","","",0],
["FL-HIST-139","Johan Gutierrez","Agrorural","F-HIST-139","Flota","PELICANO","Tablero de Comunicaciones","Cubierta y Casco",0,"","Correctivo","2024-06-17","","Falla en la tarjeta de control","Garantia","","OK",6,2024,5.0,"MEDIA","","",0],
["FL-HIST-140","Victor Costilla","Agrorural","F-HIST-140","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Preventivo","2024-07-11","","Calidad de agua, se necesita set analizado de cloro","Caja Chica","","OK",7,2024,5.0,"MEDIA","","",0],
["FL-HIST-141","Lener Cisneros","Agrorural","F-HIST-141","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-07-15","","Falla en el alternador de motor","Caja Chica","","OK",7,2024,30.0,"MEDIA","","",0],
["FL-HIST-142","Edison Moran","Agrorural","F-HIST-142","Flota","DELFIN 11","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-07-20","","Falla en la volante del motor, no arranca por esta trancado el motor","tdr-8iut","","OK",7,2024,120.0,"MEDIA","","",0],
["FL-HIST-143","Victor Costilla","Agrorural","F-HIST-143","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2024-07-25","","Valvula check 3/4 de bronce","Caja Chica","","OK",7,2024,5.0,"MEDIA","","",0],
["FL-HIST-144","Victor Costilla","Agrorural","F-HIST-144","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2024-08-02","","Valvula check 1.0 de bronce","Caja Chica","","OK",8,2024,5.0,"MEDIA","","",0],
["FL-HIST-145","Victor Costilla","Agrorural","F-HIST-145","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-07","","Falla en el tapon de zinc en enfriamiento","Caja Chica","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-146","Victor Costilla","Agrorural","F-HIST-146","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-13","","hilo corto en el tapon de zinc","Caja Chica","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-147","Victor Costilla","Agrorural","F-HIST-147","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-13","","hilo corto en el tapon de zinc","Caja Chica","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-148","Victor Costilla","Agrorural","F-HIST-148","Flota","PELICANO","ElectroBomba centrifuga aguas grises","Equipos Auxiliares",0,"","Correctivo","2024-08-13","","Falla en el impulsor","tdr-8iut","","OK",8,2024,5.0,"MEDIA","","",0],
["FL-HIST-149","Johan Gutierrez","Agrorural","F-HIST-149","Flota","PELICANO","Electrobomba de Achique","Sistema de Achique /Lastre",0,"","Correctivo","2024-08-22","","Falla en la succion de la bomba","tdr-8iut","","OK",8,2024,5.0,"MEDIA","","",0],
["FL-HIST-150","Martin Prado","Agrorural","F-HIST-150","Flota","PELICANO","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2024-08-22","","acumulacion de agua en cubierta","Garantia","","OK",8,2024,5.0,"MEDIA","","",0],
["FL-HIST-151","Manuel Lopez","Agrorural","F-HIST-151","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-22","","llave termica , bornes de bateria","Garantia","","OK",8,2024,5.0,"MEDIA","","",0],
["FL-HIST-152","Manuel Lopez","Agrorural","F-HIST-152","Flota","PELICANO","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2024-08-22","","Presencia de agua en el eje de la bomba, falla en el sello","tdr-8iut","","NO",8,2024,5.0,"MEDIA","","",0],
["FL-HIST-153","Manuel Calle","Agrorural","F-HIST-153","Flota","ISLA CHINCHA","Grua 1-Esctructura-hidraulica","Sistema de Grua",0,"","Correctivo","2024-09-07","","Fallta en los pulsadores","Caja Chica","","OK",9,2024,90.0,"MEDIA","","",0],
["FL-HIST-154","Manuel Calle","Agrorural","F-HIST-154","Flota","ISLA CHINCHA","Motor Hidráulico 1","Sistema de Grua",0,"","Correctivo","2024-09-07","","Se percibe vibracion","tdr-8iut","","OK",9,2024,90.0,"MEDIA","","",0],
["FL-HIST-155","Manuel Calle","Agrorural","F-HIST-155","Flota","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2024-09-07","","Falla en el switch de operación","tdr-8iut","","NO",9,2024,20.0,"MEDIA","","",0],
["FL-HIST-156","Edison Moran","Agrorural","F-HIST-156","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-10-09","","Falla en morse en el retroceso","tdr-8iut","","OK",10,2024,20.0,"MEDIA","","",0],
["FL-HIST-157","Edison Moran","Agrorural","F-HIST-157","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-10-10","","falla en el morse por el desajuste de un perno","Caja Chica","","OK",10,2024,20.0,"MEDIA","","",0],
["FL-HIST-158","Edison Moran","Agrorural","F-HIST-158","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-10-15","","Fallas en la valvula check  del filtro racor","tdr-8iut","","OK",10,2024,20.0,"MEDIA","","",0],
["FL-HIST-159","Edison Moran","Agrorural","F-HIST-159","Flota","DELFIN 11","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-10-20","","Falla de bateria del motor","tdr-8iut","","OK",10,2024,60.0,"MEDIA","","",0],
["FL-HIST-160","Victor Costilla","Agrorural","F-HIST-160","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-11-08","","Excesiva vibracion por desgaste del amortiguador","tdr-8iut","","OK",11,2024,5.0,"MEDIA","","",0],
["FL-HIST-161","Victor Costilla","Agrorural","F-HIST-161","Flota","PELICANO","Unidad Hidralica","Sistema de Fondeo",0,"","Correctivo","2024-11-10","","Falla en el fusible marca Bussman 690v-80A","Caja Chica","","OK",11,2024,5.0,"MEDIA","","",0],
["FL-HIST-162","Johan Gutierrez","Agrorural","F-HIST-162","Flota","PELICANO","Eje de Propulsion","Sistema de Propulsion",0,"","Correctivo","2024-11-20","","Vibracion del sistema propulsion en travesia","Garantia","","NO",11,2024,5.0,"MEDIA","","",0],
["FL-HIST-163","Manuel Calle","Agrorural","F-HIST-163","Flota","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2024-11-23","","falla en switch de energia, no llega corriente","tdr-8iut","","NO",11,2024,38.0,"MEDIA","","",0],
["FL-HIST-164","Victor Costilla","Agrorural","F-HIST-164","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-12-09","","Falla en el regulador de voltaje y bomba de inyeccion","Garantia","","OK",12,2024,5.0,"MEDIA","","",0],
["FL-HIST-165","Manuel Calle","Agrorural","F-HIST-165","Flota","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2024-12-23","","Fuga de agua dulce en la bomba de refrigeracion","tdr-8iut","","NO",12,2024,8.0,"MEDIA","","",0],
["FL-HIST-166","Jacob Medrano","Agrorural","F-HIST-166","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2025-01-02","","Falla en arrancador, perno descolgado de la volante.","tdr-8iut","","OK",1,2025,20.0,"MEDIA","","",0],
["FL-HIST-167","Manuel Calle","Agrorural","F-HIST-167","Flota","ISLA CHINCHA","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2025-01-01","","Fuga de agua dulce en la bomba de refrigeracion","tdr-8iut","","NO",1,2025,25.0,"MEDIA","","",0],
["FL-HIST-168","Jacob Medrano","Agrorural","F-HIST-168","Flota","GUANAY","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2026-01-05","","Presencia de humo negro en navegacion","tdr-8iut","","NO",1,2026,20.0,"MEDIA","","",0],
["FL-HIST-169","Lener Cisneros","Agrorural","F-HIST-169","Flota","DELFIN 12","Inversor de Voltaje 24/220v","Sistema de Energia",0,"","Correctivo","2025-02-03","","falla de un circuito","Caja Chica","","OK",2,2025,30.0,"MEDIA","","",0],
["FL-HIST-170","Manuel Calle","Agrorural","F-HIST-170","Flota","ISLA CHINCHA","Cargador de Baterias","Sistema de Energia",0,"","Correctivo","2025-02-27","","falla en el cargador de bateria","tdr-8iut","","NO",2,2025,30.0,"MEDIA","","",0],
["FL-HIST-171","Edison Moran","Agrorural","F-HIST-171","Flota","DELFIN 11","Alternador de Baterias","Sistema de Energia",0,"","Correctivo","2025-03-21","","Fallta en la faja de eslabones","Caja Chica","","NO",3,2025,85.0,"MEDIA","","",0],
["FL-HIST-172","jonathan venega","Agrorural","F-HIST-172","Flota","ISLA CHINCHA","Tablero Electrico","Sistema de Fondeo",0,"","Correctivo","2025-03-27","","Falla en el rele y pulsado de parada de emergencia","Caja Chica","","NO",3,2025,5.0,"MEDIA","","",0],
["FL-HIST-173","Juan Cisneros","Agrorural","F-HIST-173","Flota","ISLA CHINCHA","Motor Hidráulico 1","Sistema de Grua",0,"","Correctivo","2025-04-09","","Falla en el manometro","Caja Chica","","OK",4,2025,15.0,"MEDIA","","",0],
["FL-HIST-174","Jose Leon","Agrorural","F-HIST-174","Flota","DELFIN 12","Habitabilidad","Cubierta y Casco",0,"","Correctivo","2025-04-19","","Falla en al cocina, Fatiga de los refuerzo","Caja Chica","","OK",4,2025,56.0,"MEDIA","","",0],
["FL-HIST-175","Juan Cisneros","Agrorural","F-HIST-175","Flota","ISLA CHINCHA","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2025-04-21","","Falla en la tapa de la congelora y no llega a congelar","Caja Chica","","OK",4,2025,5.0,"MEDIA","","",0],
["FL-HIST-176","Juan Cisneros","Agrorural","F-HIST-176","Flota","ISLA CHINCHA","Luces de Navegacion","Cubierta y Casco",0,"","Correctivo","2025-04-21","","quemado de foco de fondeo","Caja Chica","","OK",4,2025,20.0,"MEDIA","","",0],
["FL-HIST-177","Jacob Medrano","Agrorural","F-HIST-177","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2025-04-21","","Fuga de bomba de recirculacion de refrigerante, impeller roto","tdr-8iut","","OK",4,2025,54.0,"MEDIA","","",0],
["FL-HIST-178","Jose Leon","Agrorural","F-HIST-178","Flota","DELFIN 12","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2025-05-09","","Cocina malograda","Caja Chica","","OK",5,2025,11.0,"MEDIA","","",0],
["FL-HIST-179","Jose Leon","Agrorural","F-HIST-179","Flota","DELFIN 12","Ancla","Sistema de Fondeo",0,"","Correctivo","2025-05-20","","Caida y Ancla en el Bahia de Chimbote","Caja Chica","","OK",5,2025,25.0,"MEDIA","","",0],
["FL-HIST-180","Juan Cisneros","Agrorural","F-HIST-180","Flota","ISLA CHINCHA","Winche de babor","Sistema de Fondeo",0,"","Correctivo","2025-05-25","","Falla en winche de estribor. Sonido de engranajes malogrados","Caja Chica","","NO",5,2025,5.0,"MEDIA","","",0],
["FL-HIST-181","Jose Leon","Agrorural","F-HIST-181","Flota","DELFIN 11","Luces de Navegacion","Cubierta y Casco",0,"","Correctivo","2025-05-26","","Foco de faro pirata quemado","Caja Chica","","NO",5,2025,19.0,"MEDIA","","",0],
["FL-HIST-182","Manuel Calle","Agrorural","F-HIST-182","Flota","ISLA CHINCHA","Tablero Electrico","Sistema de Fondeo",0,"","Correctivo","2025-06-01","","Falla elecrtico en el panel de control de caseta en proa","Caja Chica","","NO",6,2025,5.0,"MEDIA","","",0],
["FL-HIST-183","Juan Cisnteros","Agrorural","F-HIST-183","Flota","ALCATRAZ","Casco Principal","Cubierta y Casco",0,"","Preventivo","2023-11-30","","Limpieza de Casco por Incrustaciones Marinas","Caja Chica","","OK",11,2023,15.0,"MEDIA","","",0],
["FL-HIST-184","Juan Cisnteros","Agrorural","F-HIST-184","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2022-06-22","","Falla en el arrancador","tdr-8iut","","OK",6,2022,30.0,"MEDIA","","",0],
["FL-HIST-185","Juan Cisnteros","Agrorural","F-HIST-185","Flota","ALCATRAZ","Caja reductora Unico","Sistema de Propulsion",0,"","Correctivo","2023-04-17","","Falla en la marcha retroceso y avance de la caja de transmision","tdr-8iut","","OK",4,2023,30.0,"MEDIA","","",0],
["FL-HIST-186","Juan Cisnteros","Agrorural","F-HIST-186","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-04-16","","Abertura de pernos de carcaza del arrancador","Caja Chica","","OK",4,2023,30.0,"MEDIA","","",0],
["FL-HIST-187","Juan Cisnteros","Agrorural","F-HIST-187","Flota","ALCATRAZ","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2023-04-25","","Quemadura del estator","Caja Chica","","OK",4,2023,30.0,"MEDIA","","",0],
["FL-HIST-188","Manuel Lopez","Agrorural","F-HIST-188","Flota","PELICANO","Motor Principal BR","Sistema de Energia",0,"","Correctivo","2024-05-08","","Ruptura de Anodos de Zinc de Enfriamiento","Caja Chica","","OK",5,2024,10.0,"MEDIA","","",0],
["FL-HIST-189","Manuel Lopez","Agrorural","F-HIST-189","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2024-05-06","","Fallta en el reductor de 1/2 a 2.5","Caja Chica","","OK",5,2024,15.0,"MEDIA","","",0],
["FL-HIST-190","Johan Gutierrez","Agrorural","F-HIST-190","Flota","PELICANO","Equipos de Comunicación","Cubierta y Casco",0,"","Correctivo","2024-05-15","","Falla electrico y/o sonoro","Caja Chica","","NO",5,2024,10.0,"MEDIA","","",0],
["FL-HIST-191","Manuel Lopez","Agrorural","F-HIST-191","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-22","","llave termica , bornes de bateria","Garantia","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-192","Victor Costilla","Agrorural","F-HIST-192","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2024-07-25","","Valvula check 3/4 de bronce","Caja Chica","","OK",7,2024,4.0,"MEDIA","","",0],
["FL-HIST-193","Victor Costilla","Agrorural","F-HIST-193","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2024-08-02","","Valvula check 1.0 de bronce","Caja Chica","","OK",8,2024,4.0,"MEDIA","","",0],
["FL-HIST-194","Victor Costilla","Agrorural","F-HIST-194","Flota","PELICANO","ElectroBomba centrifuga aguas grises","Equipos Auxiliares",0,"","Correctivo","2024-08-13","","Falla en el impulsor","tdr-8iut","","OK",8,2024,4.0,"MEDIA","","",0],
["FL-HIST-195","Victor Costilla","Agrorural","F-HIST-195","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-07","","Falla en el tapon de zinc en enfriamiento","Caja Chica","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-196","Victor Costilla","Agrorural","F-HIST-196","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-13","","hilo corto en el tapon de zinc","Caja Chica","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-197","Victor Costilla","Agrorural","F-HIST-197","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2024-08-13","","hilo corto en el tapon de zinc","Caja Chica","","OK",8,2024,2.0,"MEDIA","","",0],
["FL-HIST-198","Manuel Lopez","Agrorural","F-HIST-198","Flota","PELICANO","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2024-08-22","","Presencia de agua en el eje de la bomba, falla en el sello","tdr-8iut","","NO",8,2024,4.0,"MEDIA","","",0],
["FL-HIST-199","Victor Costilla","Agrorural","F-HIST-199","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Preventivo","2024-07-11","","Calidad de agua, se necesita set analizado de cloro","Caja Chica","","OK",7,2024,4.0,"MEDIA","","",0],
["FL-HIST-200","Johan Gutierrez","Agrorural","F-HIST-200","Flota","PELICANO","Electrobomba de Achique","Sistema de Achique /Lastre",0,"","Correctivo","2024-08-22","","Falla en la succion de la bomba","tdr-8iut","","OK",8,2024,4.0,"MEDIA","","",0],
["FL-HIST-201","Victor Costilla","Agrorural","F-HIST-201","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-11-08","","Excesiva vibracion por desgaste del amortiguador","tdr-8iut","","OK",11,2024,4.0,"MEDIA","","",0],
["FL-HIST-202","Victor Costilla","Agrorural","F-HIST-202","Flota","PELICANO","Unidad Hidralica","Sistema de Fondeo",0,"","Correctivo","2024-11-10","","Falla en el fusible marca Bussman 690v-80A","Caja Chica","","OK",11,2024,10.0,"MEDIA","","",0],
["FL-HIST-203","Johan Gutierrez","Agrorural","F-HIST-203","Flota","PELICANO","Eje de Propulsion","Sistema de Propulsion",0,"","Correctivo","2024-11-20","","Vibracion del sistema propulsion en travesia","Garantia","","NO",11,2024,10.0,"MEDIA","","",0],
["FL-HIST-204","Victor Costilla","Agrorural","F-HIST-204","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-12-09","","Falla en el regulador de voltaje y bomba de inyeccion","Garantia","","OK",12,2024,12.0,"MEDIA","","",0],
["FL-HIST-205","Jacob Medrano","Agrorural","F-HIST-205","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-10-17","","falla en la cadena del saca filtro","Caja Chica","","OK",10,2023,120.0,"MEDIA","","",0],
["FL-HIST-206","Jacob Medrano","Agrorural","F-HIST-206","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-08-11","","Presencia de humo negro y bajas revoluciones","tdr-8iut","","NO",8,2023,90.0,"MEDIA","","",0],
["FL-HIST-207","Jacob Medrano","Agrorural","F-HIST-207","Flota","GUANAY","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2023-09-21","","Falta de Estabilidad a la bomba de inyeccion","tdr-8iut","","OK",9,2023,90.0,"MEDIA","","",0],
["FL-HIST-208","Jacob Medrano","Agrorural","F-HIST-208","Flota","GUANAY","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2025-03-01","","Presencia de humo negro en navegacion, falla en los componentes de combustible","tdr-8iut","","NO",3,2025,90.0,"MEDIA","","",0],
["FL-HIST-209","Jacob Medrano","Agrorural","F-HIST-209","Flota","GUANAY","Motor Principal BR","Sistema de Propulsion",0,"","Correctivo","2023-05-10","","Presencia de humo negro en navegacion y cambio de sonido por escape","tdr-8iut","","OK",5,2023,90.0,"MEDIA","","",0],
["FL-HIST-210","Jacob Medrano","Agrorural","F-HIST-210","Flota","GUANAY","Unidad Hidralica","Sistema de Fondeo",0,"","Correctivo","2022-09-20","","Goteo en la empaquetuda del motor de caja","tdr-8iut","","OK",9,2022,90.0,"MEDIA","","",0],
["FL-HIST-211","Jacob Medrano","Agrorural","F-HIST-211","Flota","GUANAY","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2022-08-18","","Codigo de falla SD START FAIL, arrancador inoperativo","tdr-8iut","","OK",8,2022,90.0,"MEDIA","","",0],
["FL-HIST-212","Jacob Medrano","Agrorural","F-HIST-212","Flota","GUANAY","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2022-07-26","","Vibracion del motor excesivo y inestabilidad de bomba de inyeccion","tdr-8iut","","OK",7,2022,2.0,"MEDIA","","",0],
["FL-HIST-213","Edison Moran","Agrorural","F-HIST-213","Flota","DELFIN 11","Cubierta Principal","Cubierta y Casco",0,"","Correctivo","2023-04-21","","Falla en el motor del congelador","Caja Chica","","OK",4,2023,120.0,"MEDIA","","",0],
["FL-HIST-214","Edison Moran","Agrorural","F-HIST-214","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2023-01-16","","Fisura de las valvulas en el resorte de calibracion","tdr-8iut","","OK",1,2023,90.0,"MEDIA","","",0],
["FL-HIST-215","Edison Moran","Agrorural","F-HIST-215","Flota","DELFIN 11","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-07-20","","Falla en la volante del motor, no arranca por esta trancado el motor","tdr-8iut","","OK",7,2024,120.0,"MEDIA","","",0],
["FL-HIST-216","Edison Moran","Agrorural","F-HIST-216","Flota","DELFIN 11","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-10-20","","Falla de bateria del motor","tdr-8iut","","OK",10,2024,60.0,"MEDIA","","",0],
["FL-HIST-217","Edison Moran","Agrorural","F-HIST-217","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-10-09","","Falla en morse en el retroceso","tdr-8iut","","OK",10,2024,20.0,"MEDIA","","",0],
["FL-HIST-218","Edison Moran","Agrorural","F-HIST-218","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-10-15","","Fallas en la valvula check  del filtro racor","tdr-8iut","","OK",10,2024,20.0,"MEDIA","","",0],
["FL-HIST-219","Edison Moran","Agrorural","F-HIST-219","Flota","DELFIN 11","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-10-10","","falla en el morse por el desajuste de un perno","Caja Chica","","OK",10,2024,20.0,"MEDIA","","",0],
["FL-HIST-220","Lener Cisneros","Agrorural","F-HIST-220","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-04-06","","Falla en la cañeria de agua de enfriamiento","Caja Chica","","OK",4,2024,30.0,"MEDIA","","",0],
["FL-HIST-221","Lener Cisneros","Agrorural","F-HIST-221","Flota","DELFIN 12","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-07-15","","Falla en el alternador de motor","Caja Chica","","OK",7,2024,30.0,"MEDIA","","",0],
["FL-HIST-222","Lener Cisneros","Agrorural","F-HIST-222","Flota","DELFIN 12","Inversor de Voltaje 24/220v","Sistema de Energia",0,"","Correctivo","2025-02-03","","falla de un circuito","Caja Chica","","OK",2,2025,30.0,"MEDIA","","",0],
["FL-HIST-223","Johan Gutierrez","Agrorural","F-HIST-223","Flota","PELICANO","Tablero Electrico","Sistema de Fondeo",0,"","Correctivo","2025-03-21","","falla en el fusible de winche de proa","tdr-8iut","","OK",3,2025,2.0,"MEDIA","","",0],
["FL-HIST-224","Johan Gutierrez","Agrorural","F-HIST-224","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2025-03-26","","falla en el DISPLAY DE GRUPO ELECTROGENO DE PUERTO","Caja Chica","","OK",3,2025,2.0,"MEDIA","","",0],
["FL-HIST-225","Johan Gutierrez","Agrorural","F-HIST-225","Flota","PELICANO","Planta Osmosis","Equipos Auxiliares",0,"","Correctivo","2025-04-09","","Falla de baja presion en ingreso de la bomba Buster","tdr-8iut","","NO",4,2025,2.0,"MEDIA","","",0],
["FL-HIST-226","Johan Gutierrez","Agrorural","F-HIST-226","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2025-04-09","","Falla en el filtro de aire","Caja Chica","","NO",4,2025,2.0,"MEDIA","","",0],
["FL-HIST-227","Johan Gutierrez","Agrorural","F-HIST-227","Flota","PELICANO","Motor Principal ER","Sistema de Propulsion",0,"","Correctivo","2025-04-09","","Falla en filtro primario de combustible","Caja Chica","","OK",4,2025,2.0,"MEDIA","","",0],
["FL-HIST-228","Johan Gutierrez","Agrorural","F-HIST-228","Flota","PELICANO","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2025-04-09","","Falla en el AVR de Grupo Doosan BR","Caja Chica","","NO",4,2025,2.0,"MEDIA","","",0],
["FL-HIST-229","Martin Prado","Agrorural","F-HIST-229","Flota","PELICANO","Electrobomba de Contraincendio","Cubierta y Casco",0,"","Correctivo","2025-04-21","","Falla en acople de sistema de contraincedio","Caja Chica","","OK",4,2025,2.0,"MEDIA","","",0],
["FL-HIST-230","Jacob Medrano","Agrorural","F-HIST-230","Flota","GUANAY","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2025-04-21","","Fuga de bomba de recirculacion de refrigerante, impeller roto","tdr-8iut","","OK",4,2025,272.0,"MEDIA","","",0],
["FL-HIST-231","Johan Gutierrez","Agrorural","F-HIST-231","Flota","PELICANO","Grupo Electrogeno BR","Sistema de Energia",0,"","Correctivo","2025-05-28","","Falla en el AVR y falta programacion del display en modulo de Grupo Doosan","tdr-8iut","","NO",5,2025,2.0,"MEDIA","","",0],
["FL-HIST-232","Johan Gutierrez","Agrorural","F-HIST-232","Flota","PELICANO","Cubierta Principal","Equipos Auxiliares",0,"","Correctivo","2025-05-20","","Falla en la bomba manual de lubricacion","Caja Chica","","OK",5,2025,2.0,"MEDIA","","",0],
["FL-HIST-233","Johan Gutierrez","Agrorural","F-HIST-233","Flota","PELICANO","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2025-05-16","","Falta filtro de aceite para el mantenimiento","Caja Chica","","OK",5,2025,2.0,"MEDIA","","",0],
["FL-HIST-234","Johan Gutierrez","Agrorural","F-HIST-234","Flota","PELICANO","Motor Principal ER","Sistema de Gobierno",0,"","Correctivo","2025-05-20","","Falta filtro de aire para el motor principal","Caja Chica","","OK",5,2025,2.0,"MEDIA","","",0],
["FL-HIST-235","Jose Leon","Agrorural","F-HIST-235","Flota","DELFIN 12","Ancla","Sistema de Fondeo",0,"","Correctivo","2025-05-20","","Caida y Ancla en el Bahia de Chimbote","Caja Chica","","OK",5,2025,200.0,"MEDIA","","",0],
["FL-HIST-236","Johan Gutierrez","Agrorural","F-HIST-236","Flota","PELICANO","Casco Principal","Sistema de Energia",0,"","Correctivo","2025-06-03","","wincha en mal estado","Caja Chica","","NO",6,2025,2.0,"MEDIA","","",0],
["FL-HIST-237","Jacob Medrano","Agrorural","F-HIST-237","Flota","PANGA","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2025-01-31","","la volante no gira, con el arrancador operativo. Se paraliza el motor porque necesita desmontar","tdr-8iut","","NO",1,2025,352.0,"MEDIA","","",0],
["FL-HIST-238","Johan Gutierrez","Agrorural","F-HIST-238","Flota","PELICANO","Grupo Electrogeno ER","Sistema de Energia",0,"","Correctivo","2025-08-22","","Falla en el impulsor de la bomba de agua de mar de refrigeracion","Caja Chica","","OK",8,2025,2.0,"MEDIA","","",0],
["FL-HIST-239","Juan Cisnteros","Agrorural","F-HIST-239","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2024-03-02","","FALTA DE CARENADO","tdr-8iut","","NO",3,2024,360.0,"MEDIA","","",0],
["FL-HIST-240","Juan Cisnteros","Agrorural","F-HIST-240","Flota","ALCATRAZ","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2024-03-02","","FALTA DE CARENADO","tdr-8iut","","NO",3,2024,360.0,"MEDIA","","",0],
["FL-HIST-241","Juan Cisnteros","Agrorural","F-HIST-241","Flota","ALCATRAZ","Casco Principal","Cubierta y Casco",0,"","Correctivo","2024-03-03","","FALTA DE CARENADO","tdr-8iut","","NO",3,2024,360.0,"MEDIA","","",0],
["FL-HIST-242","Jose Leon","Agrorural","F-HIST-242","Flota","CISTERNA 1","Casco Principal","Cubierta y Casco",0,"","Correctivo","2024-01-02","","FALTA DE CARENADO","tdr-8iut","","NO",1,2024,360.0,"MEDIA","","",0],
["FL-HIST-243","Juan Cisnteros","Agrorural","F-HIST-243","Flota","CISTERNA 1","Casco Principal","Cubierta y Casco",0,"","Correctivo","2025-03-02","","FALTA DE CARENADO","tdr-8iut","","NO",3,2025,365.0,"MEDIA","","",0],
["FL-HIST-244","Juan Cisnteros","Agrorural","F-HIST-244","Flota","ALCATRAZ","Casco Principal","Cubierta y Casco",0,"","Correctivo","2025-03-02","","FALTA DE CARENADO","tdr-8iut","","NO",3,2025,365.0,"MEDIA","","",0],
["FL-HIST-245","Juan Cisnteros","Agrorural","F-HIST-245","Flota","ALCATRAZ","Motor Principal Unico","Sistema de Propulsion",0,"","Correctivo","2025-03-02","","FALTA DE CARENADO","tdr-8iut","","NO",3,2025,365.0,"MEDIA","","",0],
["FL-HIST-246","Juan Cisnteros","Agrorural","F-HIST-246","Flota","ALCATRAZ","Grupo Electrogeno de Puerto","Sistema de Energia",0,"","Correctivo","2025-03-02","","FALTA DE CARENADO","tdr-8iut","","NO",3,2025,365.0,"MEDIA","","",0]
  ];
  
  // Insertar en lotes de 50 para no agotar tiempo de ejecución
  const LOTE = 50;
  let total = 0;
  
  for (let i = 0; i < FALLAS.length; i += LOTE) {
    const lote = FALLAS.slice(i, i + LOTE);
    h.getRange(h.getLastRow() + 1, 1, lote.length, lote[0].length)
     .setValues(lote);
    total += lote.length;
    Logger.log(`  Insertadas ${total}/${FALLAS.length} fallas...`);
    SpreadsheetApp.flush(); // Forzar escritura en cada lote
  }
  
  Logger.log(`✅ ${total} fallas históricas cargadas`);
  return { fallas: total };
}

// ════════════════════════════════════════════════════════════
//  EQUIPOS REALES — extraídos del Excel
//  (los equipos únicos que aparecen en el historial)
// ════════════════════════════════════════════════════════════
function _migrarEquipos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h  = ss.getSheetByName('Equipos');
  if (h.getLastRow() > 1) { Logger.log('Equipos ya registrados'); return 0; }
  
  // Mapa embarcación → ID
  const embMap = {
    'ALCATRAZ':    'EMB001', 'GUANAY':       'EMB002',
    'ISLA CHINCHA':'EMB003', 'DELFIN 11':    'EMB004',
    'DELFIN 12':   'EMB005', 'PELICANO':     'EMB006',
    'CISTERNA 1':  'EMB007', 'PANGA':        'EMB008'
  };
  // Mapa sistema → ID
  const sisMap = {
    'Sistema de Propulsion':   'SIS001', 'Sistema de Gobierno':     'SIS002',
    'Sistema de Energia':      'SIS003', 'Sistema de Grua':         'SIS004',
    'Sistema de Fondeo':       'SIS005', 'Sistema de Achique /Lastre':'SIS006',
    'Sistema de Achique/Lastre':'SIS006','Equipos Auxiliares':      'SIS008',
    'Cubierta y Casco':        'SIS009'
  };
  
  // Equipos únicos por embarcación — extraídos del historial real
  const equipos = [
    // ID, EMB, SIS, COD, NOMBRE, MARCA, MODELO, SERIE, ANO, POT, TIPO_INSP, HOR_ACT, HOR_ULT, PROX, ESTADO, CRIT, FECHA
    ['EQ001','EMB002','SIS008','GUA-EA-001','Electrobomba de Achique','','','','','','IV',0,0,500,'OPERATIVO',3,new Date()],
    ['EQ002','EMB003','SIS004','ICH-SGR-001','Grua 1-Esctructura-hidraulica','','','','','','IM',0,0,1000,'OPERATIVO',4,new Date()],
    ['EQ003','EMB003','SIS003','ICH-SE-001','Grupo Electrogeno BR','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ004','EMB003','SIS004','ICH-SGR-002','Grua 2-Esctructura-hidraulica','','','','','','IM',0,0,1000,'OPERATIVO',4,new Date()],
    ['EQ005','EMB003','SIS004','ICH-SGR-003','Motor Hidráulico 1','','','','','','IV',0,0,500,'OPERATIVO',3,new Date()],
    ['EQ006','EMB004','SIS003','D11-SE-001','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ007','EMB005','SIS003','D12-SE-001','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ008','EMB003','SIS009','ICH-CC-001','Luces de Navegacion','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ009','EMB004','SIS001','D11-SP-001','Motor Principal Unico','Cummins','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ010','EMB005','SIS001','D12-SP-001','Motor Principal Unico','Cummins','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ011','EMB007','SIS009','CIS-CC-001','Electrobomba de Contraincendio','','','','','','IV',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ012','EMB003','SIS004','ICH-SGR-004','Motor Eléctrico 1','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ013','EMB008','SIS001','PAN-SP-001','Motor Principal Unico','Cummins','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ014','EMB001','SIS009','ALC-CC-001','Casco Principal','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ015','EMB001','SIS001','ALC-SP-001','Motor Principal Unico','','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ016','EMB003','SIS003','ICH-SE-002','Grupo Electrogeno ER','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ017','EMB003','SIS009','ICH-CC-002','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ018','EMB004','SIS005','D11-SF-001','Cabre-estante','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ019','EMB005','SIS008','D12-EA-001','Bomba de Achique','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ020','EMB001','SIS001','ALC-SP-002','Caja reductora Unico','','','','','','IM',0,0,1000,'OPERATIVO',4,new Date()],
    ['EQ021','EMB007','SIS003','CIS-SE-001','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ022','EMB007','SIS008','CIS-EA-001','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ023','EMB008','SIS009','PAN-CC-001','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ024','EMB002','SIS001','GUA-SP-001','Motor Principal BR','Cummins','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ025','EMB002','SIS001','GUA-SP-002','Motor Principal ER','Cummins','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ026','EMB002','SIS003','GUA-SE-001','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ027','EMB002','SIS003','GUA-SE-002','Grupo Electrogeno BR','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ028','EMB002','SIS003','GUA-SE-003','Grupo Electrogeno ER','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ029','EMB002','SIS005','GUA-SF-001','Unidad Hidralica','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ030','EMB007','SIS001','CIS-SP-001','Eje de Propulsion','','','','','','IM',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ031','EMB007','SIS001','CIS-SP-002','Motor Principal Unico','','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ032','EMB006','SIS009','PEL-CC-001','Equipos de Comunicación','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ033','EMB006','SIS003','PEL-SE-001','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ034','EMB006','SIS003','PEL-SE-002','Grupo Electrogeno BR','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ035','EMB006','SIS003','PEL-SE-003','Grupo Electrogeno ER','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ036','EMB006','SIS008','PEL-EA-001','Planta Osmosis','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ037','EMB006','SIS008','PEL-EA-002','ElectroBomba centrifuga aguas grises','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ038','EMB006','SIS009','PEL-CC-002','Tablero de Comunicaciones','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ039','EMB006','SIS001','PEL-SP-001','Motor Principal BR','','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ040','EMB006','SIS001','PEL-SP-002','Motor Principal ER','','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ041','EMB006','SIS005','PEL-SF-001','Unidad Hidralica','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ042','EMB006','SIS006','PEL-SAL-001','Electrobomba de Achique','','','','','','IV',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ043','EMB006','SIS009','PEL-CC-003','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ044','EMB006','SIS001','PEL-SP-003','Eje de Propulsion','','','','','','IM',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ045','EMB006','SIS009','PEL-CC-004','Electrobomba de Contraincendio','','','','','','IV',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ046','EMB005','SIS005','D12-SF-001','Ancla','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ047','EMB003','SIS005','ICH-SF-001','Winche de estribor','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ048','EMB003','SIS006','ICH-SAL-001','Electrobomba de Achique','','','','','','IV',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ049','EMB003','SIS003','ICH-SE-003','Cargador de Baterias','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ050','EMB007','SIS006','CIS-SAL-001','Electrobomba de Achique','','','','','','IV',0,0,0,'OPERATIVO',4,new Date()],
    ['EQ051','EMB007','SIS008','CIS-EA-002','ElectroBomba centrifuga aguas grises','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ052','EMB007','SIS009','CIS-CC-002','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ053','EMB007','SIS009','CIS-CC-003','Tablero de Comunicaciones','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ054','EMB004','SIS008','D11-EA-001','Cabre-estante','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ055','EMB004','SIS009','D11-CC-001','Casco Principal','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ056','EMB004','SIS008','D11-EA-002','Bomba de Achique','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ057','EMB004','SIS003','D11-SE-002','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ058','EMB004','SIS003','D11-SE-003','Alternador de Baterias','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ059','EMB003','SIS004','ICH-SGR-005','Motor Eléctrico 2','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ060','EMB003','SIS001','ICH-SP-001','Motor Principal Unico','','','','','','MP',0,0,500,'OPERATIVO',5,new Date()],
    ['EQ061','EMB003','SIS005','ICH-SF-002','Winche de babor','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ062','EMB003','SIS005','ICH-SF-003','Tablero Electrico','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ063','EMB005','SIS003','D12-SE-002','Inversor de Voltaje 24/220v','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ064','EMB006','SIS005','PEL-SF-002','Tablero Electrico','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ065','EMB001','SIS003','ALC-SE-001','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ066','EMB007','SIS001','CIS-SP-003','Casco Principal','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ067','EMB001','SIS009','ALC-CC-002','Grupo Electrogeno de Puerto','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ068','EMB003','SIS009','ICH-CC-003','Luces de Navegacion','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ069','EMB005','SIS009','D12-CC-001','Habitabilidad','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ070','EMB005','SIS009','D12-CC-002','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ071','EMB004','SIS009','D11-CC-002','Luces de Navegacion','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ072','EMB002','SIS008','GUA-EA-002','Electrobomba de Achique','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
    ['EQ073','EMB006','SIS003','PEL-SE-004','Motor Principal BR','','','','','','MP',0,0,500,'OPERATIVO',4,new Date()],
    ['EQ074','EMB006','SIS008','PEL-EA-003','Cubierta Principal','','','','','','IV',0,0,0,'OPERATIVO',2,new Date()],
    ['EQ075','EMB006','SIS003','PEL-SE-005','Casco Principal','','','','','','IV',0,0,0,'OPERATIVO',3,new Date()],
  ];
  
  h.getRange(h.getLastRow()+1, 1, equipos.length, equipos[0].length).setValues(equipos);
  Logger.log(`✅ ${equipos.length} equipos registrados`);
  return equipos.length;
}

// ════════════════════════════════════════════════════════════
//  MIGRACIÓN COMPLETA CON EQUIPOS
// ════════════════════════════════════════════════════════════
function migrarTodoIncluyendoEquipos() {
  inicializarConHistorial();
  const eq = _migrarEquipos();
  SpreadsheetApp.getUi().alert(
    '✅ Migración Total Completa',
    `Se cargaron además ${eq} equipos reales del historial.\n\nEl sistema tiene:\n• 246 fallas históricas 2022-2026\n• ${eq} equipos registrados\n• 13 técnicos del equipo\n• 8 embarcaciones\n\n¡Listo para operar!`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ════════════════════════════════════════════════════════════
//  ESTADÍSTICAS POST-MIGRACIÓN (verificar que todo quedó bien)
// ════════════════════════════════════════════════════════════
function verificarMigracion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stats = {};
  
  ['Embarcaciones','Sistemas','Equipos','Usuarios','Registro_Fallas',
   'Ordenes_Trabajo','Certificados','Combustible','Plan_Mantenimiento'].forEach(hoja => {
    const h = ss.getSheetByName(hoja);
    stats[hoja] = h ? Math.max(0, h.getLastRow()-1) : 'NO EXISTE';
  });
  
  // Fallas por año
  const hF = ss.getSheetByName('Registro_Fallas');
  const dataF = hF.getDataRange().getValues().slice(1);
  const porAno = {};
  dataF.forEach(r => {
    const ano = r[18]; // columna ANO
    if (ano) porAno[ano] = (porAno[ano]||0) + 1;
  });
  
  const resumen = Object.entries(stats)
    .map(([k,v]) => `  ${k}: ${v} registros`)
    .join('\n');
  
  const porAnoStr = Object.entries(porAno)
    .sort((a,b)=>a[0]-b[0])
    .map(([a,n]) => `  ${a}: ${n} fallas`)
    .join('\n');
    
  const msg = `📊 ESTADO DEL SISTEMA\n\n${resumen}\n\nFallas por año:\n${porAnoStr}`;
  Logger.log(msg);
  SpreadsheetApp.getUi().alert('Estado del Sistema', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  return { stats, porAno };
}
