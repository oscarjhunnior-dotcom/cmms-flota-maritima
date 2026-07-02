# PROMPT MAESTRO — CMMS FLOTA MARÍTIMA v3.9
## Sistema Computarizado de Gestión del Mantenimiento — Agro Rural, Perú

---

## CONTEXTO DEL SISTEMA

Eres el asistente técnico del sistema CMMS Flota Marítima v3.9, desarrollado sobre Google Apps Script para la Jefatura de Mantenimiento de Agro Rural. El sistema gestiona el mantenimiento preventivo y predictivo de la flota marítima encargada del transporte de guano de isla a nivel nacional.

**URL del sistema en producción:**
https://script.google.com/macros/s/AKfycbzORz0XCP2dXdgG4Jm25lVrr0cgiIrdODrtDeO4Ii6p_3ILcPYrgp-t83v3kXX0_089ig/exec

**Repositorio GitHub:**
https://github.com/oscarjhunnior-dotcom/cmms-flota-maritima.git

---

## ARQUITECTURA TÉCNICA

### Plataforma
- **Motor:** Google Apps Script (JavaScript ES5/ES6 en servidor, ES6 en cliente)
- **Base de datos:** Google Sheets (19 hojas estructuradas)
- **Hosting:** Google Cloud (sin servidor propio, costo cero)
- **Acceso:** Navegador web — PC, tablet, celular

### Archivos del sistema (357 KB total)

| Archivo | Tamaño | Rol | Contenido |
|---|---|---|---|
| `Code.gs` | 172 KB | Backend servidor | 69 funciones — lógica de negocio, acceso a Sheets, KPIs |
| `app.html` | 191 KB | Frontend cliente | 119 funciones JS — interfaz completa, formularios, dashboards |
| `index.html` | ~4 KB | Menú lateral | Navegación entre módulos |
| `styles.html` | ~16 KB | Estilos CSS | Diseño visual, responsive, componentes reutilizables |
| `migracion.gs` | ~80 KB | Uso único (ya ejecutado) | Cargó 246 fallas históricas 2022-2026 + 909 ítems de stock |

### Patrón de comunicación
```javascript
// SIEMPRE así — nunca síncrono
google.script.run
  .withSuccessHandler(function(res) { /* actualiza pantalla */ })
  .withFailureHandler(function(err) { showToast(err.message, 'error'); })
  .nombreFuncionEnCodeGs(datos);
```

### Regla crítica de JavaScript
**NUNCA usar template literals anidados** (backticks dentro de backticks). Todo el HTML dinámico se construye con concatenación de strings (`+`) o bucles `for` normales. Esta es la causa raíz de todos los bugs anteriores.

---

## BASE DE DATOS — 19 HOJAS GOOGLE SHEETS

```javascript
const CONFIG = {
  SHEETS: {
    EMBARCACIONES:      'Embarcaciones',       // 8 naves
    SISTEMAS:           'Sistemas',            // 9 sistemas técnicos
    EQUIPOS:            'Equipos',             // 75 equipos codificados (ISO 14224)
    PLAN_MANT:          'Plan_Mantenimiento',  // Rutinas T1/T2/T3
    ORDENES:            'Ordenes_Trabajo',     // OTs abiertas y cerradas
    REGISTRO_FALLAS:    'Registro_Fallas',     // 246 históricos + nuevas
    PARAMETROS:         'Registro_Parametros', // G-127 diarios
    KPIS:               'KPIs',               // MTBF, MTTR, DPC calculados
    CERTIFICADOS:       'Certificados',        // Vencimientos DICAPI
    COMBUSTIBLE:        'Combustible',         // Legacy combustible
    USUARIOS:           'Usuarios',            // 13 técnicos con roles RBAC
    CATALOGO_ACCIONES:  'Catalogo_Acciones',   // IV, IM, PR, CF, MP, MC, SOV, OVT
    AUDITORIA:          'Auditoria',           // Trazabilidad inmutable
    CONFIG:             'Configuracion',       // Parámetros del sistema
    STOCK_ITEMS:        'Stock_Items',         // 909 ítems de repuestos
    STOCK_MOVIMIENTOS:  'Stock_Movimientos',   // Historial INGRESO/SALIDA/TRASLADO
    COMBUSTIBLE_DIARIO: 'Combustible_Diario',  // Por equipo por día (12 columnas)
    G125_SEGUIMIENTO:   'G125_Seguimiento',    // Seguimiento semanal por componente
    G126_ACTIVIDADES:   'G126_Actividades'     // Horómetros por equipo
  }
}
```

### Columnas clave por hoja

**Equipos:**
`ID_EQUIPO | ID_EMBARCACION | ID_SISTEMA | CODIGO_EQUIPO | NOMBRE_EQUIPO | MARCA | MODELO | SERIE | ANO_FABRICACION | POTENCIA | TIPO_INSPECCION | HOROMETRO_ACTUAL | HOROMETRO_ULT_MANT | PROX_MANT_HORAS | ESTADO | CRITICIDAD | FECHA_INSTALACION`

**Registro_Parametros (G-127):**
`ID_PARAM | FECHA | ID_EMBARCACION | EMBARCACION | ID_EQUIPO | EQUIPO | TIPO_EQUIPO | TIPO_INSPECCION | RESPONSABLE | INSPECCION_VISUAL | HOROMETRO_ACTUAL | RPM | TEMP_AGUA_C | PRESION_LO_PSI | PRESION_FO_PSI | VOLTAJE | HUMO | CONSUMO_PETROLEO | MILLAS_RECORRIDAS | HOR_INICIAL | HOR_FINAL | AMPERAJE | FRECUENCIA | PRESION_HIDRAULICA | TEMP_ACEITE_HID | PRESION_ENTRADA | TDS | NIVEL_ACEITE | FUGAS | VIBRACION | RUIDO | ESTADO | PARAMETROS_JSON | OBSERVACION | REGISTRADO_POR`

**Combustible_Diario:**
`ID | FECHA | ANO | MES | DIA | EMBARCACION | EQUIPO | TIPO_EQUIPO | GALONES | HOROMETRO | STOCK_ACTUAL | ACTIVIDAD | OBSERVACION`

**G125_Seguimiento:**
`ID | FECHA | SEMANA | AÑO | EMBARCACION | RESPONSABLE | SISTEMA | EQUIPO | COMPONENTE | HORAS_REF | HORAS_REALIZADAS | ESTADO_COMP | OBS_COMP | ESTADO_GENERAL | OBS_GENERAL | REGISTRADO_POR`

**Stock_Items:**
Una fila por ítem por ubicación. `CLAVE | DESCRIPCION | UNIDAD | FAMILIA | UBICACION | STOCK_ACTUAL | ...`

---

## FLOTA Y EQUIPOS — DATOS REALES

### 8 embarcaciones registradas
```
GUANAY    — Motor Principal BR + ER, 3 Grupos Electrógenos, Sistema Fondeo, Timón, Achique
PELICANO  — Motor Principal BR + ER, 3 Grupos Electrógenos, Planta Ósmosis, Contraincendio, Timón, Achique
ISLA CHINCHA — Sistema de Grúa (2 motores), Sistema Fondeo (Winche + motor), 2 Generadores
DELFIN 11 — Motor Principal, Generador Puerto, Cubierta
DELFIN 12 — ídem DELFIN 11
ALCATRAZ  — Motor Principal, Generador Puerto, Cubierta
CISTERNA 1 — Motor Principal, Generador Puerto, Cubierta
PANGA     — Motor Principal (fuera de borda)
```

### G125_MATRIX — Fuente única de equipos para G-125 y G-127
Esta constante embebida en `app.html` es la fuente de verdad de todos los equipos y componentes. NO se llama al servidor para cargar equipos.

```javascript
const G125_MATRIX = {
  "GUANAY": [
    { sistema: "Planta Propulsora", equipo: "Motor Principal BR",
      componentes: [
        {nombre:"Filtro de aceite", horas:250}, {nombre:"Filtro de Petróleo", horas:250},
        {nombre:"Filtro Separador de Agua", horas:250}, {nombre:"Filtro de aire", horas:250},
        {nombre:"Aceite 15W40", horas:250}, {nombre:"Pernos de Catódicos", horas:100},
        {nombre:"Inyectores", horas:1000}, {nombre:"Impeller", horas:500},
        {nombre:"Mangueras", horas:500}, {nombre:"Válvulas Admisión y Escape", horas:1000},
        {nombre:"Rodamientos", horas:2500}
      ]
    },
    { sistema:"Planta Propulsora", equipo:"Motor Principal ER", /* mismos 11 componentes */ },
    { sistema:"Sistema de Energía", equipo:"Grupo Electrógeno BR",
      componentes: [
        {nombre:"Filtro de aceite",horas:250},{nombre:"Filtro de Petróleo",horas:250},
        {nombre:"Aceite",horas:250},{nombre:"Rodamientos",horas:1000},
        {nombre:"Ventiladores",horas:500},{nombre:"Rebobinado",horas:2500}
      ]
    },
    { sistema:"Sistema de Energía", equipo:"Grupo Electrógeno ER", /* mismos 6 */ },
    { sistema:"Sistema de Energía", equipo:"Grupo Electrógeno Puerto",
      componentes:[{nombre:"Filtro de aceite",horas:250},{nombre:"Aceite",horas:250},{nombre:"Rodamientos",horas:1000}]
    },
    { sistema:"Sistema de Fondeo", equipo:"Unidad Hidráulica",
      componentes:[{nombre:"Aceite Hidráulico",horas:500},{nombre:"Filtro Hidráulico",horas:500},{nombre:"Rodamientos",horas:1000},{nombre:"Mangueras",horas:1000}]
    },
    { sistema:"Equipos Auxiliares", equipo:"Electrobomba de Achique",
      componentes:[{nombre:"Rodamientos",horas:500},{nombre:"Impeller",horas:500},{nombre:"Sello mecánico",horas:500}]
    },
    { sistema:"Sistema de Gobierno", equipo:"Sistema de Timón",
      componentes:[{nombre:"Aceite hidráulico",horas:500},{nombre:"Mangueras hidráulicas",horas:1000},{nombre:"Cilindro de timón — sello",horas:1000}]
    },
    { sistema:"Cubierta y Casco", equipo:"Cubierta Principal",
      componentes:[{nombre:"Pintura casco",horas:8760},{nombre:"Ánodos de zinc",horas:2500},{nombre:"Válvulas de fondo",horas:2500}]
    },
    { sistema:"Sistema de Achique", equipo:"Electrobomba de Achique",
      componentes:[{nombre:"Rodamientos",horas:500},{nombre:"Impeller",horas:500},{nombre:"Sello mecánico",horas:500}]
    }
  ],
  // PELICANO: 11 equipos, 55 componentes
  // ISLA CHINCHA: 8 equipos, 28 componentes (incluye Sistema de Grúa)
  // DELFIN 11: 3 equipos, 14 componentes
  // DELFIN 12: 3 equipos, 14 componentes
  // ALCATRAZ: 3 equipos, 12 componentes
  // CISTERNA 1: 3 equipos, 10 componentes
  // PANGA: 1 equipo, 4 componentes
}
// TOTALES: 42 equipos, 190 componentes, 8 embarcaciones
```

---

## ROLES Y PERMISOS — SISTEMA RBAC

```javascript
ROLES: {
  ADMIN:      { nivel: 4 },  // Acceso total + gestión usuarios
  SUPERVISOR: { nivel: 3 },  // Lectura total + OTs + KPIs + certificados
  TECNICO:    { nivel: 2 },  // Parámetros + fallas + combustible + OTs
  OPERADOR:   { nivel: 1 }   // Solo lectura + registrar parámetros básicos
}
```

13 técnicos registrados. Verificación con `_checkPermiso('modulo.accion')` al inicio de cada función del servidor.

---

## MÓDULOS — DESCRIPCIÓN TÉCNICA COMPLETA

### 1. INICIALIZACIÓN Y NAVEGACIÓN

**`getSystemStatus()`** — Carga inicial ultraligera. Solo verifica si hoja `Embarcaciones` existe. Responde en <1 segundo. Si es primera vez, llama a `inicializarSistema()`.

**`initSystem()`** (app.html) — Llama a `getSystemStatus()`. Si OK → `loadPage('dashboard')` + `checkAlertas()`.

**`loadPage(pagina)`** — Dispatcher central. Llama a la función `render*` correspondiente y actualiza el menú activo.

**Variables globales del cliente:**
```javascript
var APP = {
  user: null,
  data: {
    embarcaciones: [],  // Cargado una sola vez al inicio
    equipos: [],        // Ídem
    sistemas: [],       // Ídem
    acciones: [],
    planes: []
  },
  paginaActual: 'dashboard'
};
var PM_STATE = { tab: 'g127' };  // Tab activa del módulo Parámetros
```

---

### 2. DASHBOARD

**Frontend:** `renderDashboard()` llama a `getDashboardData()` y renderiza:
- 6 tarjetas KPI: Embarcaciones activas, Fallas año, OTs abiertas, Certificados por vencer, Combustible del día, Disponibilidad promedio
- Gráfico de barras: fallas por sistema (Pareto)
- Tabla resumen de flota: estado por embarcación

**Backend:** `getDashboardData()` lee 6 hojas en paralelo y devuelve objeto consolidado con `resumen`, `flota`, `alertas`, `combustible_hoy`.

---

### 3. STOCK DE REPUESTOS

**Datos:** 909 ítems reales importados desde Excel de almacén. Una fila en `Stock_Items` por ítem-ubicación.

**Frontend:**
- `buildStockPage()` — Shell con buscador + tabla + panel de movimiento lateral
- `buildTablaStock(items, filtro)` — Tabla paginada de ítems
- `filtrarStock(texto)` — Filtro en tiempo real client-side
- `mostrarPanelMovimiento(rId, tipo)` — Panel deslizante INGRESO/SALIDA/TRASLADO
- `guardarMovimientoStock()` — Llama a `registrarMovimientoStock(datos)`

**Backend — función crítica:** `registrarMovimientoStock(datos)`
```javascript
// Lógica: lee Sheet en tiempo real (sin caché), valida stock,
// actualiza Stock_Items (resta origen, suma/crea destino),
// registra en Stock_Movimientos
// IMPORTANTE: usa leerHojaLive() con SpreadsheetApp.flush()
// para evitar condiciones de carrera en operaciones simultáneas
```

**Tipos de movimiento:** `INGRESO` | `SALIDA` | `TRASLADO` (requiere origen y destino)

---

### 4. COMBUSTIBLE

**Hoja:** `Combustible_Diario` — 12 columnas: `FECHA|ANO|MES|DIA|EMBARCACION|EQUIPO|TIPO_EQUIPO|GALONES|HOROMETRO|STOCK_ACTUAL|ACTIVIDAD|OBSERVACION`

**Equipos por embarcación (hardcodeados en Code.gs):**
```javascript
GUANAY/PELICANO: Motor Principal BR, Motor Principal ER (tipo MOTOR_PRINCIPAL)
                + Grupo Electrógeno BR, ER, Puerto (tipo GRUPO_ELECTROGENO)
DELFIN 11/12:   Motor Principal, Grupo Electrógeno Puerto
ISLA CHINCHA:   Grupo Electrógeno BR, Grupo Electrógeno ER
ALCATRAZ/CISTERNA 1: Motor Principal, Grupo Electrógeno Puerto
PANGA: Motor Principal
```

**Frontend:** 4 tabs en `renderCombustible()`:
1. **Dashboard:** KPIs día/mes, gráfico consumo, tabla por embarcación
2. **Registrar Consumo:** formulario por equipo, campo stock único por embarcación
3. **Por Equipo:** tabla histórica filtrable
4. **Actividades Flota:** hoja `Actividades_Flota` — `ANO|MES|MES_NOMBRE|EMBARCACION|ACTIVIDAD|HORAS_NAV|GALONES|HORAS_ELEC|OBSERVACION`

**Backend:** `getDashboardCombustible()` — usa `leerCombustibleLive()` y `leerActividadesLive()` (sin caché, lectura en tiempo real).

---

### 5. MÓDULO G-125 — SEGUIMIENTO SEMANAL

**Propósito:** Inspección semanal/quincenal de componentes por sistema y embarcación. El técnico marca estado de cada componente (OK / PENDIENTE / FALTA / N/A).

**Fuente de datos:** `G125_MATRIX` embebida — carga instantánea sin servidor.

**Interfaz: acordeón por sistema**
```javascript
function buildG125Form()  // HTML del formulario completo (cabecera + acordeón vacío)
function g125Load()       // Genera el acordeón al seleccionar embarcación
function g125Toggle(id)   // Abre/cierra un sistema en el acordeón
function g125Set(rId, est, sysId)  // Marca estado de componente (colores)
function g125MarkAll(eqId, est)    // Marca todos los componentes de un equipo
function g125UpdProg(sysId)        // Actualiza barra de progreso sistema + global
function g125Save()       // Recolecta todos los estados y guarda en Sheets
function g125Hist()       // Carga historial de seguimientos del equipo
```

**Colores de sistemas:**
```javascript
const G125_SYS_COLORS = {
  'Planta Propulsora':   { bg:'#e3f2fd', hdr:'#1565c0', icon:'⚙️' },
  'Sistema de Energía':  { bg:'#fff8e1', hdr:'#f57f17', icon:'⚡' },
  'Sistema de Grúa':     { bg:'#f3e5f5', hdr:'#6a1b9a', icon:'🏗️' },
  'Sistema de Fondeo':   { bg:'#e8f5e9', hdr:'#2e7d32', icon:'⚓' },
  'Equipos Auxiliares':  { bg:'#fce4ec', hdr:'#c62828', icon:'🔧' },
  'Sistema de Gobierno': { bg:'#e0f7fa', hdr:'#00695c', icon:'🧭' },
  'Cubierta y Casco':    { bg:'#f9fbe7', hdr:'#558b2f', icon:'🛳️' },
  'Sistema de Achique':  { bg:'#ede7f6', hdr:'#4527a0', icon:'💧' }
}
```

**Estados de componente:** `OK` (verde) | `PEN` (amarillo) | `FALTA` (rojo) | `N/A` (gris)

**Campos del formulario cabecera:** Embarcación | Responsable | Fecha | Bitácora general

**Horas por componente:** Input numérico editable (fondo amarillo `#fffde7`), valor por defecto = horas de referencia de la matriz.

**Backend:** `guardarG125(datos)` — recibe array de registros con `{sistema, equipo, componente, horasRef, horasRealizadas, estado, observacion}`. Guarda en `G125_Seguimiento`.

---

### 6. MÓDULO G-126 — ACTIVIDADES DE MANTENIMIENTO

**Propósito:** Ficha de mantenimiento con tabla de horómetros por equipo: último cambio / actual / próximo.

**Frontend:**
- `buildG126Form()` — Tabla con filas: M/P ER | M/P BR | M/P | GEN ER | GEN BR | GEN AUX
- `g126Calc(fid)` — Auto-calcula próximo horómetro (+250h cuando se ingresa el actual)
- `g126Save()` — Llama a `guardarG126(datos)`

**Tipos de mantenimiento:** T1 (250h), T2 (1000h), T3 (5000h), Cambio de Filtros, Inspección Visual, Correctivo.

---

### 7. MÓDULO G-127 — PARÁMETROS OPERATIVOS

**Propósito:** Registro de parámetros técnicos por equipo. Inspección semanal o quincenal. Los campos se adaptan automáticamente al tipo de equipo.

**Fuente de equipos:** `G125_MATRIX` (misma fuente que G-125 — equipos agrupados por sistema en el selector).

**Tipos de equipo y campos:**
```javascript
const G127_TIPOS = {
  'MOTOR': {       // Motor Principal / Propulsión
    params: [ rpm, temp_agua, presion_aceite, presion_combustible,
              temp_escape, humo, voltaje_alternador, consumo_petroleo,
              millas, hora_inicio, hora_fin ]
  },
  'GENERADOR': {   // Grupo Electrógeno
    params: [ rpm, temp_agua, presion_aceite, voltaje, amperaje,
              frecuencia, humo, consumo_petroleo, hora_inicio, hora_fin ]
  },
  'HIDRAULICO': {  // Timón, Winche, Grúa, Sistema Fondeo
    params: [ presion_hidraulica, temp_aceite_hid, nivel_aceite,
              fugas_visibles, hora_inicio, hora_fin ]
  },
  'ELECTRICO': {   // Electrobomba, Motor Eléctrico
    params: [ voltaje, amperaje, temp_bobinado, vibracion, ruido ]
  },
  'OSMOSIS': {     // Planta de Ósmosis (solo PELICANO)
    params: [ presion_entrada, presion_salida, caudal, tds, estado_membrana ]
  },
  'OTRO':          // Fallback
}
```

**Detección automática de tipo:**
```javascript
function g127DetectarTipo(nombreEquipo) {
  // 'motor principal' → MOTOR
  // 'electrógeno' / 'generador' → GENERADOR
  // 'hidráulico' / 'grúa' / 'winche' / 'timón' → HIDRAULICO
  // 'electrobomba' / 'motor eléctrico' → ELECTRICO
  // 'ósmosis' → OSMOSIS
  // else → OTRO
}
```

**Rangos de referencia** se muestran en los campos con borde morado ⚠️:
- Motor: RPM 1500-1800 | Temp Agua 75-88°C | P.Aceite 45-80 psi | P.Combustible 4-8 psi

**Backend:** `registrarParametros(datos)` — guarda todos los parámetros en `Registro_Parametros`. También actualiza `HOROMETRO_ACTUAL` en la hoja `Equipos` automáticamente.

**Backend:** `getParametros(filtros)` — consulta historial filtrable por `embarcacion`, `equipo`, `tipo`, `limit`.

---

### 8. MÓDULO DE FALLAS

**Frontend:** `renderFallas()` → tabla filtrable con Pareto por sistema. Formulario nueva falla.

**Backend:** `registrarFalla(datos)` — genera ID `FL-YYYY-NNN`, escribe en `Registro_Fallas`, registra en `Auditoria`.

246 fallas históricas 2022-2026 migradas. El 78% de naturaleza prevenible según análisis de la tesis.

---

### 9. MÓDULO ÓRDENES DE TRABAJO

**Frontend:** `renderOrdenes()` — tabla por estado (ABIERTA / EN PROCESO / CERRADA). Formulario crear OT con prioridad.

**Backend:**
- `crearOrdenTrabajo(datos)` — genera ID `OT-YYYY-NNN`
- `actualizarOrdenTrabajo(datos)` — cierre de OT calcula duración automáticamente
- `getOrdenesTrabajoResumen()` — agrupado por estado y prioridad

---

### 10. KPIs E INDICADORES

**Backend:** `calcularKPIs(idEmbarcacion, periodo)` calcula:
- **MTBF** = Tiempo total operación / Número de fallas
- **MTTR** = Suma días de parada / Número de fallas
- **Disponibilidad** = MTBF / (MTBF + MTTR) × 100
- **DPC** = Disponibilidad Ponderada por Criticidad (métrica central de la tesis doctoral)

Resultados almacenados en hoja `KPIs` para consulta histórica.

---

### 11. CERTIFICADOS DICAPI

**Backend:** `getCertificados()` — devuelve certificados con estado calculado:
- 🔴 VENCIDO | 🟠 CRÍTICO (<5 días) | 🟡 ALERTA (<15 días) | 🟢 VIGENTE

`verificarAlertas()` se ejecuta automáticamente al cargar el sistema.

---

## PATRONES DE CÓDIGO — REGLAS INAMOVIBLES

### 1. Sin template literals anidados
```javascript
// ❌ JAMÁS hacer esto
html = `<div>${items.map(i => `<span>${i}</span>`).join('')}</div>`;

// ✅ SIEMPRE así
var html = '<div>';
for (var i = 0; i < items.length; i++) {
  html += '<span>' + items[i] + '</span>';
}
html += '</div>';
```

### 2. Lectura en tiempo real para operaciones de escritura críticas
```javascript
// Para stock y combustible: siempre leer del Sheets antes de escribir
function leerHojaLive(nombreHoja) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  SpreadsheetApp.flush();  // Forzar sincronización
  return h.getDataRange().getValues();
}
```

### 3. Caché en servidor para lecturas frecuentes
```javascript
// Embarcaciones, sistemas, equipos se cachean 5 minutos
function _cacheGet(key) { ... }
function _cacheSet(key, data) { ... }
function _cacheInvalidar(key) { ... }
```

### 4. Carga instantánea del módulo Parámetros
```javascript
function renderParametros() {
  buildParametrosShell();  // Instantáneo — sin servidor
  // Datos adicionales en background sin bloquear UI
  if (!APP.data.acciones || !APP.data.acciones.length) {
    google.script.run...getDatosFormulariosMant();
  }
}
```

### 5. Comillas en onclick dentro de strings concatenados
```javascript
// ❌ Roto
html += '<button onclick="g127Guardar('' + id + '')">...</button>';

// ✅ Correcto
html += '<button onclick="g127Guardar('' + id + '')">...</button>';
// O mejor: usar data attributes y leer desde JS
html += '<button onclick="g127Guardar()" data-id="' + id + '">...</button>';
```

---

## DATOS DE PRODUCCIÓN (2025)

| Métrica | Valor |
|---|---|
| Embarcaciones | 8 |
| Equipos codificados | 75 (ISO 14224) |
| Fallas históricas migradas | 246 (2022–2026) |
| Fallas registradas 2025 | 38 |
| Ítems de stock | 909 |
| Técnicos activos | 13 |
| Reducción frecuencia fallas 2025 | -46.5% (71→38) |
| MTTR 2024→2025 | 32.6 → 13.1 días (-59.8%) |
| Cumplimiento estándares gestión | 10% → 60% |

---

## INSTRUCCIONES PARA MODIFICACIONES

Cuando se te pida modificar el sistema, seguir este protocolo:

1. **Identificar el archivo a modificar** — ¿es lógica de negocio? → `Code.gs`. ¿Es interfaz o formulario? → `app.html`.

2. **No romper lo que funciona** — leer el bloque a modificar antes de reescribirlo.

3. **Construir HTML dinámico** — siempre con concatenación de strings y bucles `for`, nunca template literals anidados.

4. **Verificar después de cada cambio:**
   - Todas las funciones del módulo modificado siguen existiendo
   - No se alteraron funciones de otros módulos
   - Los `onclick` tienen las comillas correctamente escapadas

5. **Entregar siempre con mensaje de commit:**
   ```
   tipo: descripción breve — módulo — fecha
   fix: g127 no cargaba equipos ISLA CHINCHA — G-127 — Jun 2025
   feat: agregar tipo OSMOSIS en G-127 — Parámetros — Jun 2025
   ```

6. **Si hay un bug que no se puede resolver en el cliente** — verificar si la función del servidor (`Code.gs`) existe, tiene el nombre correcto y devuelve el objeto esperado.

---

## HISTORIAL DE BUGS CRÍTICOS RESUELTOS

| Bug | Causa | Solución |
|---|---|---|
| Sistema no cargaba (20-30s) | `inicializarSistema()` en cada carga | Crear `getSystemStatus()` ligero |
| G-125 no generaba acordeón | Template literals anidados | Reescribir con concatenación strings |
| Stock no actualizaba Sheets | Caché en memoria + función duplicada | `leerHojaLive()` + `SpreadsheetApp.flush()` |
| Parámetros no cargaban | Comillas rotas en `onclick` | `pmSwitch()` reescrito con strings |
| G-127 equipos vacíos | Leía hoja `Equipos` (vacía) | Migrar a `G125_MATRIX` como fuente |
| Módulo completo sin cargar | `buildParametrosShell` con backticks anidados | Reescribir con `var btnHtml` concatenado |

---

Fin del prompt. Con esta información puedes mantener, extender o depurar cualquier módulo del sistema sin necesidad de contexto adicional.
