//CLIENT.GS
// ============================================================
// TEST CLAUDE API - Apps Script
// Configura tu API Key en Script Properties:
//   Archivo > Propiedades del proyecto > Propiedades de script
//   Clave: ANTHROPIC_API_KEY  Valor: sk-ant-...
// ============================================================

const CLAUDE_MODEL = 'claude-sonnet-4-5'; // o claude-opus-4-5, claude-haiku-4-5-20251001

/**
 * Llama a la API de Claude y devuelve el texto de respuesta.
 * @param {string} userMessage - El mensaje del usuario.
 * @param {string} [systemPrompt] - Prompt de sistema opcional.
 * @param {number} [maxTokens=1024] - Máximo de tokens en la respuesta.
 * @returns {string} Texto de respuesta de Claude.
 */
function llamarClaude(userMessage, systemPrompt = '', maxTokens = 1024) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('Falta ANTHROPIC_API_KEY en Script Properties.');

  const payload = {
    model: CLAUDE_MODEL,
    max_tokens: maxTokens,
    messages: [{ role: 'user', content: userMessage }]
  };

  if (systemPrompt) payload.system = systemPrompt;

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const code = response.getResponseCode();
  const body = JSON.parse(response.getContentText());

  if (code !== 200) {
    throw new Error(`Error ${code}: ${body.error?.message || JSON.stringify(body)}`);
  }

  return body.content[0].text;
}

// ============================================================
// FUNCIONES DE PRUEBA - Ejecútalas desde el menú Run
// ============================================================

function testBasico() {
  const respuesta = llamarClaude('¿Cuál es la capital de México? Responde en una sola oración.');
  Logger.log('✅ Respuesta: ' + respuesta);
}

function testConSystemPrompt() {
  const system = 'Eres un asistente notarial experto en derecho mexicano. Responde siempre de forma concisa y profesional.';
  const respuesta = llamarClaude('¿Qué es una escritura pública?', system);
  Logger.log('✅ Respuesta con system prompt:\n' + respuesta);
}

function testConversacionMultiturno() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('Falta ANTHROPIC_API_KEY en Script Properties.');

  const messages = [
    { role: 'user', content: 'Mi nombre es Luis.' },
    { role: 'assistant', content: 'Hola Luis, ¿en qué puedo ayudarte?' },
    { role: 'user', content: '¿Cómo me llamo?' }
  ];

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({ model: CLAUDE_MODEL, max_tokens: 256, messages }),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', options);
  const body = JSON.parse(response.getContentText());
  Logger.log('✅ Multi-turno: ' + body.content[0].text);
}

function testVerificarClave() {
  try {
    const resp = llamarClaude('Di solo: OK');
    Logger.log(resp === 'OK' || resp.includes('OK')
      ? '✅ API Key válida. Conexión exitosa.'
      : '✅ Conectado. Respuesta: ' + resp);
  } catch (e) {
    Logger.log('❌ Error: ' + e.message);
  }
}
function debugCalculoFestivos() {
  const TZ = Session.getScriptTimeZone();
  const hoy = new Date();
  
  Logger.log('=== DEBUG CÁLCULO FESTIVOS ===');
  Logger.log('Fecha actual: ' + Utilities.formatDate(hoy, TZ, 'dd/MM/yyyy HH:mm'));
  
  // 1. Verificar que lee los festivos correctamente
  const festivos = getFestivosActivos ? getFestivosActivos() : new Set();
  Logger.log('Total festivos cargados: ' + festivos.size);
  Logger.log('Festivos: ' + JSON.stringify([...festivos]));
  
  const hoyStr = Utilities.formatDate(hoy, TZ, 'yyyy-MM-dd');
  Logger.log('¿Hoy (' + hoyStr + ') es festivo? ' + festivos.has(hoyStr));
  
  // 2. Simular creación de ticket con SLA de 4h, 8h y 24h
  const slasPrueba = [4, 8, 24, 48];
  
  slasPrueba.forEach(function(slaHoras) {
    const venc = calcularFechaConHorarioLaboral(new Date(hoy), slaHoras);
    Logger.log(
      'SLA ' + slaHoras + 'h → Vencimiento: ' + 
      Utilities.formatDate(venc, TZ, 'dd/MM/yyyy HH:mm') +
      ' (día: ' + ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'][venc.getDay()] + ')'
    );
  });
  
  // 3. Verificar días siguientes hasta encontrar uno hábil
  Logger.log('--- Próximos 7 días ---');
  for (var i = 0; i <= 7; i++) {
    var d = new Date(hoy);
    d.setDate(hoy.getDate() + i);
    d.setHours(8, 0, 0, 0);
    var dStr   = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    var diaSem = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'][d.getDay()];
    var esFinde   = d.getDay() === 0 || d.getDay() === 6;
    var esFestivo = festivos.has(dStr);
    Logger.log(
      dStr + ' (' + diaSem + ')' +
      (esFinde    ? ' ← FIN DE SEMANA' : '') +
      (esFestivo  ? ' ← FESTIVO ✓'    : '') +
      (!esFinde && !esFestivo ? ' → HÁBIL' : '')
    );
  }
  
  Logger.log('=== FIN DEBUG ===');
  return '✅ Debug ejecutado — revisa el Logger (Ver > Registros)';
}

// ============================================================================
// BEXALTA HELP DESK — BACKEND V4.1
// Autodetección por usuario (Área/Ubicación) + catálogos por área
// Asignación automática, SLA, compresión de adjuntos, bitácora automática
// ============================================================================

// ============================================================================
// CONSTANTES Y CONFIGURACIÓN
// ============================================================================
const DB = {
  TICKETS: 'Tickets',
  COMMENTS: 'Comentarios',
  USERS: 'Usuarios',
  CATEGORIES: 'Categorias',
  CATEGORIES_TI: 'Categorias_TI',
  CATEGORIES_MTTO: 'Categorias_Mtto',
  AREAS: 'Areas',
  PRIORITIES: 'Prioridades',
  STATUSES: 'Estatus',
  CONFIG: 'Config',
  LOG: 'Bitacora',
  NOTIFS: 'Notificaciones',
  GERENTES: 'ConfigGerentes',
  ESCALAMIENTOS: 'EscalamientosLog'
};

const HEADERS = {
  Tickets: [
    'ID', 'Folio', 'Fecha', 'ReportaEmail', 'ReportaNombre', 'Área', 'Categoría',
    'Prioridad', 'AsignadoA', 'Título', 'Descripción', 'Estatus', 'SLA_Horas',
    'Vencimiento', 'ÚltimaActualización', 'Adjuntos', 'Origen', 'CarpetaAdjuntosId',
    'Ubicación', 'Presupuesto', 'TipoProveedor', 'TiempoCotizacion', 'AprobadorEmail',
    'StatusAutorizacion', 'FechaVisita', 'HoraVisita', 'NotasVisita','MotivoEscalamiento', 
    'FechaEscalamiento', 'SolicitanteEscalamiento', 'FueReabierto', 'ContadorReaperturas'
  ],
  Comentarios: ['ID', 'TicketID', 'Fecha', 'AutorEmail', 'AutorNombre', 'Comentario', 'Interno', 'FileID', 'FileURL', 'FileName'],
  // En la constante HEADERS, modificar Usuarios:
Usuarios: [
  'Email', 'Nombre', 'Rol', 'Área', 'Ubicación', 'Puesto', 
  'PasswordHash', 'ResetToken', 'ResetExp',
  'Disponible', 'MotivoAusencia', 'FechaInicioAusencia', 'FechaFinAusencia',
  'EmailNotificacion',
  'EstatusUsuario',
  'TelegramID' // <--- AGREGAR ESTO AL FINAL
],
  Categorias: ['Nombre', 'Área'],
  Areas: ['Nombre'],
  Prioridades: ['Nombre', 'SLA_Horas'],
  Estatus: ['Nombre', 'Orden'],
  Config: ['Clave', 'Valor'],
  Bitacora: ['Fecha', 'TicketID', 'Usuario', 'Acción', 'Detalle'],
  Notificaciones: ['ID', 'Fecha', 'Usuario', 'Tipo', 'Título', 'Mensaje', 'TicketID', 'Leido', 'Timestamp'],
  ConfigGerentes: ['Area', 'GerenteEmail', 'GerenteNombre'],
   EscalamientosLog: ['ID', 'TicketID', 'Folio', 'Area', 'Solicitante', 'Gerente','FechaSolicitud', 'FechaRespuesta', 'Estado', 'Motivo', 'NivelUrgencia', 'Respuesta']
};

// ============================================================
// CONFIGURACIÓN DE TELEGRAM POR GRUPOS
// Cada agente/gerente tiene su grupo privado con el bot
// ============================================================

const TELEGRAM_GRUPOS = {
  // Agentes - Grupos individuales
  'EJasso': '-5128054565',
  'ejasso': '-5128054565',
  'ejasso@bexalta.com': '-5128054565',
  'ejasso@bexalta.mx': '-5128054565',
  
  'RGNava': '-5235679022',
  'rgnava': '-5235679022',
  'rgnava@bexalta.com': '-5235679022',
  'rgnava@bexalta.mx': '-5235679022',
  
  'JLGonzalez': '-4997122821',
  'jlgonzalez': '-4997122821',
  'jlgonzalez@bexalta.com': '-4997122821',
  'jlgonzalez@bexalta.mx': '-4997122821',
  
  'DValdez': '-5023644227',
  'dvaldez': '-5023644227',
  'dvaldez@bexalta.com': '-5023644227',
  'dvaldez@bexalta.mx': '-5023644227',
  
  'JARamirez': '-5297882029',
  'jaramirez': '-5297882029',
  'jaramirez@bexalta.com': '-5297882029',
  'jaramirez@bexalta.mx': '-5297882029'
};

// Gerentes por área
const GERENTES_AREAS = {
  'sistemas': {
    email: 'rgnava@bexalta.com',
    nombre: 'RGNava',
    chatId: '-5235679022'
  },
  'mantenimiento': {
    email: 'jlgonzalez@bexalta.com',
    nombre: 'JLGonzalez',
    chatId: '-4997122821'
  }
};

// ============================================================================
// FUNCIONES CORE
// ============================================================================

/**
 * Entrada principal del Web App (GET)
 */
function doGet(e) {
  const action = e && e.parameter ? e.parameter.action : null;
  const ticketId = e && e.parameter ? e.parameter.id : null;
  const autoOpenTicket = e && e.parameter ? e.parameter.ticket : null;

  // 1) APROBACIÓN / RECHAZO DE COTIZACIONES
  if (action === 'approve_cot' || action === 'reject_cot') {
    try {
      let msg;
      if (action === 'reject_cot') {
        // Llama directamente a handleCotizacionApproval con 'reject'
        // que ahora usa 'Cotización Rechazada' en lugar de 'En Espera'
        msg = handleCotizacionApproval(ticketId, 'reject');
      } else {
        msg = handleCotizacionApproval(ticketId, 'approve');
      }

      const colorHeader = action === 'reject_cot' ? '#dc2626' : '#198754';
      const iconHeader  = action === 'reject_cot' ? '❌' : '✅';

      return HtmlService.createHtmlOutput(`
        <body style="font-family:Arial;text-align:center;background:#f7f9fc;">
          <div style="max-width:600px;margin:50px auto;padding:30px;border-radius:12px;background:white;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
            <h2 style="color:${colorHeader};margin-bottom:20px;">${iconHeader} Operación completada</h2>
            <p style="font-size:1.1em;color:#333;">${msg}</p>
            <p style="margin-top:30px;color:#6c757d;font-size:.9em;">Puede cerrar esta ventana.</p>
          </div>
        </body>
      `);
    } catch (err) {
      return HtmlService.createHtmlOutput(`
        <body style="font-family:Arial;text-align:center;background:#f7f9fc;">
          <div style="max-width:600px;margin:50px auto;padding:30px;border-radius:12px;background:white;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
            <h2 style="color:#dc3545;margin-bottom:20px;">Error</h2>
            <p style="font-size:1.1em;color:#333;">${err.message}</p>
          </div>
        </body>
      `);
    }
  }

  // 2) APROBACIÓN / RECHAZO DE ESCALAMIENTO
  if (action === 'approve_escalar' || action === 'reject_escalar') {
    try {
      const solicitante = e.parameter.by || '';
      const msg = handleEscalamientoApproval(ticketId, action.replace('_escalar', ''), solicitante);
      return HtmlService.createHtmlOutput(`
        <body style="font-family:Arial;text-align:center;background:#f7f9fc;">
          <div style="max-width:600px;margin:50px auto;padding:30px;border-radius:12px;background:white;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
            <h2 style="color:#16a34a;margin-bottom:20px;">Operación exitosa</h2>
            <p style="font-size:1.1em;color:#333;">${msg}</p>
            <p style="margin-top:30px;color:#6c757d;font-size:.9em;">Puede cerrar esta ventana.</p>
          </div>
        </body>
      `);
    } catch (err) {
      return HtmlService.createHtmlOutput(`
        <body style="font-family:Arial;text-align:center;background:#f7f9fc;">
          <div style="max-width:600px;margin:50px auto;padding:30px;border-radius:12px;background:white;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
            <h2 style="color:#dc3545;margin-bottom:20px;">Error</h2>
            <p style="font-size:1.1em;color:#333;">${err.message}</p>
          </div>
        </body>
      `);
    }
  }

  // 3) APROBACIÓN DE REPROGRAMACIÓN
  if (action === 'approve_reprog') {
    try {
      const fecha = e.parameter.fecha || '';
      const hora  = e.parameter.hora  || '';
      const msg   = aprobarReprogramacion(ticketId, fecha, hora);
      return HtmlService.createHtmlOutput(`
        <body style="font-family:Arial;text-align:center;background:#f7f9fc;">
          <div style="max-width:600px;margin:50px auto;padding:30px;border-radius:12px;background:white;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
            <h2 style="color:#16a34a;margin-bottom:20px;">✓ Reprogramación Aprobada</h2>
            <p style="font-size:1.1em;color:#333;">${msg}</p>
          </div>
        </body>
      `);
    } catch (err) {
      return HtmlService.createHtmlOutput(`<h3>Error: ${err.message}</h3>`);
    }
  }

  if (action === 'reject_reprog') {
    try {
      const msg = rechazarReprogramacion(ticketId);
      return HtmlService.createHtmlOutput(`
        <body style="font-family:Arial;text-align:center;background:#f7f9fc;">
          <div style="max-width:600px;margin:50px auto;padding:30px;border-radius:12px;background:white;box-shadow:0 4px 12px rgba(0,0,0,0.1);">
            <h2 style="color:#dc2626;margin-bottom:20px;">✗ Reprogramación Rechazada</h2>
            <p style="font-size:1.1em;color:#333;">${msg}</p>
          </div>
        </body>
      `);
    } catch (err) {
      return HtmlService.createHtmlOutput(`<h3>Error: ${err.message}</h3>`);
    }
  }

  // 4) RESET DE CONTRASEÑA POR TOKEN
  if (action === 'resetpass') {
    const token = e.parameter.token;
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);

    if (m.ResetToken == null) {
      return HtmlService.createHtmlOutput('<h3>Sistema no configurado para reset de contraseña.</h3>');
    }

    const idx = rows.findIndex(r => String(r[m.ResetToken]) === token);
    if (idx < 0) {
      return HtmlService.createHtmlOutput('<h3>Token inválido o expirado.</h3>');
    }

    const exp = Number(rows[idx][m.ResetExp] || 0);
    if (Date.now() > exp) {
      return HtmlService.createHtmlOutput('<h3>Este enlace ha expirado.</h3>');
    }

    const t = HtmlService.createTemplateFromFile('ResetPassword');
    t.token = token;
    return t.evaluate()
      .setTitle('Restablecer contraseña')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 5) VISTA DE REPORTES BI
  if (action === 'reportesBI') {
    return HtmlService.createHtmlOutputFromFile('ReportesBI')
      .setTitle('Reportes BI – Mesa de Ayuda BEXALTA')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 6) CARGA NORMAL DEL SISTEMA
  ensureSheets();
  const t = HtmlService.createTemplateFromFile('Index');
  t.brand = getConfig('nombreSistema') || 'Bexalta Helpdesk';
  t.autoOpenTicket = autoOpenTicket || '';

  return t.evaluate()
    .setTitle(t.brand)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


/**
 * Obtiene la URL pública de la Web App
 * @param {string} ticketId - (Opcional) ID del ticket para generar URL directa
 * @returns {string} URL del script, opcionalmente con parámetro de ticket
 */
function getScriptUrl(ticketId) {
  const baseUrl = ScriptApp.getService().getUrl();
  if (ticketId) {
    return baseUrl + '?ticket=' + encodeURIComponent(ticketId);
  }
  return baseUrl;
}

/**
 * Incluir archivos HTML parciales
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Obtener hoja por nombre (con validación)
 */
function getSheet(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Falta hoja: ' + name);
  return sh;
}

/**
 * Verificar si una hoja está vacía (solo encabezados)
 */
function isEmpty(sheetName) {
  const sh = getSheet(sheetName);
  return sh.getLastRow() <= 1;
}

/**
 * Insertar filas al final de una hoja
 */
function setValues(sheetName, rows) {
  const sh = getSheet(sheetName);
  if (!rows || !rows.length) return;
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

/**
 * Generar UUID único
 */
function genId() {
  return Utilities.getUuid();
}

/**
 * Convertir array de headers a mapa {header: índice}
 */
function _headerMap_(headers) {
  return Object.fromEntries(headers.map((h, i) => [h, i]));
}

// Variable global para almacenar datos en RAM durante el microsegundo que corre el script
const _GLOBAL_MEMO = {};

/**
 * Leer tabla completa por encabezados (Con Memoización extrema)
 */
function _readTableByHeader_(sheetName) {
  // Si ya leímos esta hoja en esta misma ejecución, la devolvemos al instante (0.001s)
  if (_GLOBAL_MEMO[sheetName]) {
    return _GLOBAL_MEMO[sheetName];
  }

  const sh = getSheet(sheetName);
  const lr = sh.getLastRow();
  const lc = sh.getLastColumn();
  if (lr < 1 || lc < 1) return { headers: [], rows: [] };
  
  const headers = sh.getRange(1, 1, 1, lc).getDisplayValues()[0];
  const rows = lr > 1 ? sh.getRange(2, 1, lr - 1, lc).getValues() : [];
  
  // Guardamos en la memoria volátil
  _GLOBAL_MEMO[sheetName] = { headers, rows };
  
  return _GLOBAL_MEMO[sheetName];
}

/**
 * Ejecutar función con bloqueo de documento (previene concurrencia)
 */
function withLock_(fn) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

/**
 * Wrapper seguro para funciones (captura errores)
 */
function safe(fn) {
  try {
    return fn();
  } catch (e) {
    Logger.log('ERROR: ' + e.message + '\n' + e.stack);
    return { ok: false, error: e.message };
  }
}

// ============================================================================
// INICIALIZACIÓN DE HOJAS
// ============================================================================

/**
 * Asegurar que todas las hojas existan con sus encabezados
 */
function ensureSheets() {
  const ss = SpreadsheetApp.getActive();

  Object.keys(DB).forEach(key => {
    const name = DB[key];
    let sh = ss.getSheetByName(name);

    if (!sh) {
      sh = ss.insertSheet(name);
      const hdr = HEADERS[name] || [];
      if (hdr.length) sh.getRange(1, 1, 1, hdr.length).setValues([hdr]);
    } else {
      // Verificar encabezados faltantes y agregarlos
      const want = HEADERS[name] || [];
      if (want.length) {
        const current = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getDisplayValues()[0];
        const missing = want.filter(h => !current.includes(h));
        if (missing.length) {
          sh.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
        }
      }
    }
  });

  // Semillas iniciales de catálogos
  if (isEmpty(DB.AREAS)) {
    setValues(DB.AREAS, [['Sistemas'], ['Mantenimiento'], ['Operaciones']]);
  }

  if (isEmpty(DB.PRIORITIES)) {
    setValues(DB.PRIORITIES, [
      ['Crítica', 2],
      ['Alta', 4],
      ['Media', 8],
      ['Baja', 12],
      ['Planeado', 72]
    ]);
  }

  if (isEmpty(DB.STATUSES)) {
    setValues(DB.STATUSES, [
      ['Nuevo', 1],
      ['Abierto', 2],
      ['En Proceso', 3],
      ['En Espera', 4],
      ['Resuelto', 5],
      ['Cerrado', 6]
    ]);
  }

  if (isEmpty(DB.CATEGORIES)) {
    setValues(DB.CATEGORIES, [
      ['Hardware', 'Sistemas'],
      ['Software', 'Sistemas'],
      ['Red', 'Sistemas'],
      ['Eléctrico', 'Mantenimiento'],
      ['Civil', 'Mantenimiento']
    ]);
  }

  if (isEmpty(DB.CONFIG)) {
    setValues(DB.CONFIG, [
      ['dominio', 'bexalta.com'],
      ['nombreSistema', 'Bexalta Helpdesk'],
      ['agente_sistemas', ''],
      ['agente_mantenimiento', ''],
      ['prioridad_default', 'Media'],
      ['telegram_token', ''],
      ['telegram_chat_id', ''],
      ['telegram_chat_admin', '']
    ]);
  }
    // Agregar después de isEmpty(DB.CONFIG)
  if (isEmpty(DB.GERENTES)) {
    setValues(DB.GERENTES, [
      ['Sistemas', 'rnava@bexalta.mx', 'Gerente de Sistemas'],
      ['Mantenimiento', 'JLGonzalez@bexalta.mx', 'Gerente de Mantenimiento'],
      ['Clientes', 'LMagana@bexalta.mx', 'Gerente de Clientes']
    ]);
  }

  // Inicializar hoja de días festivos con festivos de México
  try { inicializarHojaFestivos(); } catch(e) { Logger.log('Festivos init: ' + e.message); }
}

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

/**
 * Obtener configuración (una clave o todas)
 */
function getConfig(key) {
  const sh = getSheet(DB.CONFIG);
  const lr = sh.getLastRow();
  if (lr <= 1) return key ? '' : {};
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  const map = Object.fromEntries(data.filter(r => r[0]).map(r => [String(r[0]), r[1]]));
  return key ? map[key] : map;
}

/**
 * Actualizar configuración
 */
function updateConfig(data) {
  return withLock_(() => {
    const sh = getSheet(DB.CONFIG);
    const lr = sh.getLastRow();
    const existing = lr > 1 ? sh.getRange(2, 1, lr - 1, 2).getValues() : [];

    existing.forEach((r, i) => {
      const k = String(r[0] || '');
      if (k && k in data) sh.getRange(i + 2, 2).setValue(data[k]);
    });

    Object.keys(data).forEach(k => {
      const idx = existing.findIndex(r => String(r[0]) === k);
      if (idx < 0) sh.appendRow([k, data[k]]);
    });

    return { ok: true };
  });
}

// ============================================================================
// CACHÉ
// ============================================================================

function getCachedData(sheetName) {
  const cache = CacheService.getScriptCache();
  const metaKey = `sheet_meta_${sheetName}`;

  // 1. Intentar leer de caché
  try {
    const meta = cache.get(metaKey);
    if (meta) {
      const { chunks, ttl } = JSON.parse(meta);

      if (chunks === 1) {
        // Dataset pequeño: una sola clave
        const cached = cache.get(`sheet_${sheetName}`);
        if (cached) return JSON.parse(cached);
      } else {
        // Dataset grande: múltiples chunks
        const keys = Array.from({ length: chunks }, (_, i) => `sheet_${sheetName}_${i}`);
        const parts = cache.getAll(keys);

        // Verificar que TODOS los chunks existan
        const allPresent = keys.every(k => parts[k] != null);
        if (allPresent) {
          const combined = keys.map(k => JSON.parse(parts[k])).flat();
          return combined;
        }
      }
    }
  } catch (e) {
    Logger.log(`⚠️ Error leyendo caché de ${sheetName}: ${e.message}`);
  }

  // 2. Leer de hoja
  const sh = getSheet(sheetName);
  const lr = sh.getLastRow();
  if (lr <= 1) return [];
  const data = sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues();

  // 3. Guardar en caché
  try {
    const jsonStr = JSON.stringify(data);
    const CHUNK_SIZE = 90000; // ~90KB por chunk (límite es 100KB)

    // TTL diferenciado por hoja
    const ttlMap = {
      'Tickets': 180,          // 3 min (cambian frecuentemente)
      'Comentarios': 120,      // 2 min
      'Usuarios': 600,         // 10 min (cambian poco)
      'Notificaciones': 60,    // 1 min
      'Bitacora': 300,         // 5 min
      'Config': 1800,          // 30 min (casi nunca cambia)
      'EscalamientosLog': 120  // 2 min
    };
    const ttl = ttlMap[sheetName] || 300;

    if (jsonStr.length <= CHUNK_SIZE) {
      // Dataset pequeño: una sola clave
      cache.put(`sheet_${sheetName}`, jsonStr, ttl);
      cache.put(metaKey, JSON.stringify({ chunks: 1, ttl }), ttl);
    } else {
      // Dataset grande: dividir en chunks
      const numChunks = Math.ceil(jsonStr.length / CHUNK_SIZE);

      // Dividir el array de datos (no el JSON string) para mantener integridad
      const rowsPerChunk = Math.ceil(data.length / numChunks);
      const chunkEntries = {};

      for (let i = 0; i < numChunks; i++) {
        const start = i * rowsPerChunk;
        const end = Math.min(start + rowsPerChunk, data.length);
        const chunkData = data.slice(start, end);
        chunkEntries[`sheet_${sheetName}_${i}`] = JSON.stringify(chunkData);
      }

      // Guardar todos los chunks de golpe (más eficiente)
      cache.putAll(chunkEntries, ttl);
      cache.put(metaKey, JSON.stringify({ chunks: numChunks, ttl }), ttl);

      Logger.log(`✅ Caché ${sheetName}: ${data.length} filas en ${numChunks} chunks`);
    }
  } catch (e) {
    Logger.log(`⚠️ Error guardando caché de ${sheetName}: ${e.message}`);
  }

  return data;
}

function clearCache(sheetName) {
  try {
    const cache = CacheService.getScriptCache();

    const clearOne = (name) => {
      // Limpiar meta
      const metaKey = `sheet_meta_${name}`;
      const meta = cache.get(metaKey);

      if (meta) {
        try {
          const { chunks } = JSON.parse(meta);
          // Limpiar chunks
          const keys = [metaKey, `sheet_${name}`];
          for (let i = 0; i < (chunks || 1); i++) {
            keys.push(`sheet_${name}_${i}`);
          }
          cache.removeAll(keys);
        } catch (e) {
          cache.remove(metaKey);
          cache.remove(`sheet_${name}`);
        }
      } else {
        // Fallback por si no hay meta
        cache.remove(`sheet_${name}`);
      }
    };

    if (sheetName) {
      clearOne(sheetName);
    } else {
      ['Tickets', 'Usuarios', 'Comentarios', 'Bitacora', 'Notificaciones', 'EscalamientosLog'].forEach(clearOne);
    }
  } catch (e) {
    Logger.log('⚠️ Error limpiando caché: ' + e.message);
  }
}

// ============================================================================
// USUARIOS
// ============================================================================

function getUser(emailManual) {
  const email = (emailManual || '').trim().toLowerCase();
  if (!email) return { email: '', nombre: '', rol: 'usuario', area: '', ubicacion: '', puesto: '', estatusUsuario: 'Activo' };

  // BOOM: Usamos la Caché que dura 10 minutos, no el Sheet en vivo.
  const headers = HEADERS.Usuarios;
  const rows = getCachedData(DB.USERS); 
  const m = _headerMap_(headers);

  const row = rows.find(r => String(r[m.Email] || '').toLowerCase() === email);

  if (!row) {
    return { email, nombre: '', rol: 'usuario', area: '', ubicacion: '', puesto: '', estatusUsuario: 'Activo' };
  }

  const estatusUsuario = (m.EstatusUsuario != null ? row[m.EstatusUsuario] : 'Activo') || 'Activo';

  return {
    email,
    nombre: row[m.Nombre] || '',
    rol: row[m.Rol] || 'usuario',
    area: (m['Área'] != null ? row[m['Área']] : '') || '',
    ubicacion: (m['Ubicación'] != null ? row[m['Ubicación']] : '') || '',
    puesto: (m.Puesto != null ? row[m.Puesto] : '') || '',
    estatusUsuario: estatusUsuario
  };
}

function getUserInfo(email) {
  if (!email) return null;

  // Usar caché para máxima velocidad
  const headers = HEADERS.Usuarios;
  const rows = getCachedData(DB.USERS);
  const m = _headerMap_(headers);

  const row = rows.find(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
  if (!row) return null;

  const areaGerente = getGerenteArea(email); 

  return {
    email: row[m.Email],
    nombre: row[m.Nombre] || '',
    rol: (row[m.Rol] || 'usuario').toLowerCase().trim(),
    area: (m['Área'] != null ? row[m['Área']] : '') || '',
    ubicacion: (m['Ubicación'] != null ? row[m['Ubicación']] : '') || '',
    puesto: (m.Puesto != null ? row[m.Puesto] : '') || '',
    passwordHash: (m.PasswordHash != null ? row[m.PasswordHash] : '') || '',
    esGerente: !!areaGerente,
    areaGerente: areaGerente || null
  };
}

function listUsers() {
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);

  return rows.filter(r => r[m.Email]).map(r => ({
    email: r[m.Email],
    nombre: r[m.Nombre] || '',
    rol: r[m.Rol] || 'usuario',
    area: m['Área'] != null ? r[m['Área']] : '',
    ubicacion: m['Ubicación'] != null ? r[m['Ubicación']] : '',
    // AGREGAR ESTOS DOS:
    puesto: m.Puesto != null ? r[m.Puesto] : '',
    estatus: m.Estatus != null ? r[m.Estatus] : 'Activo'
  }));
}


/**
 * Crear nuevo usuario con envío de credenciales
 */
function createUser(email, nombre, rol, area, ubicacion, puesto, estatus) {
  if (!email) throw new Error('Email requerido');

  return withLock_(() => {
    const sh = getSheet(DB.USERS);
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);

    // Verificar si ya existe
    const existe = rows.find(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
    if (existe) throw new Error('El usuario ya existe');

    // Generar contraseña aleatoria
    const password = generarPasswordSeguro(); // [cite: 49]
    const hash = hashPassword(password);

    // Mapeo dinámico de columnas
    const newRow = headers.map(h => {
      switch (h) {
        case 'Email': return email.toLowerCase();
        case 'Nombre': return nombre || '';
        case 'Rol': return rol || 'usuario';
        case 'Área': return area || '';
        case 'Ubicación': return ubicacion || '';
        case 'Puesto': return puesto || '';
        case 'Estatus': return estatus || 'Activo'; // Nuevo campo
        case 'PasswordHash': return hash;
        case 'Disponible': return true;
        default: return '';
      }
    });

    sh.appendRow(newRow);
    clearCache(DB.USERS);

    // Enviar correo de bienvenida con credenciales
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Bienvenido al Help Desk - Credenciales de Acceso',
        htmlBody: `
          <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e0e0e0; border-radius: 5px;">
            <h2 style="color: #0b3ea1;">Bienvenido, ${nombre}</h2>
            <p>Se ha creado tu cuenta para acceder a la Mesa de Ayuda.</p>
            <p><strong>Tus credenciales son:</strong></p>
            <ul>
              <li><strong>Usuario:</strong> ${email}</li>
              <li><strong>Contraseña Temporal:</strong> ${password}</li>
            </ul>
            <p>Por favor, ingresa al sistema y cambia tu contraseña lo antes posible.</p>
            <p><a href="${getScriptUrl()}" style="background-color: #0b3ea1; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px;">Ir al Sistema</a></p>
          </div>
        `
      });
    } catch (e) {
      Logger.log('⚠️ Error enviando email: ' + e.message);
    }

    return { ok: true };
  });
}

/**
 * Actualizar usuario existente
 */
function updateUser(email, nombre, rol, area, ubicacion, puesto, estatus) {
  if (!email) throw new Error('Email requerido');

  return withLock_(() => {
    const sh = getSheet(DB.USERS);
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);

    const idx = rows.findIndex(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
    if (idx < 0) throw new Error('Usuario no encontrado');

    const r = rows[idx];
    // Actualizar campos
    if (m.Nombre != null) r[m.Nombre] = nombre;
    if (m.Rol != null) r[m.Rol] = rol;
    if (m['Área'] != null) r[m['Área']] = area;
    if (m['Ubicación'] != null) r[m['Ubicación']] = ubicacion;
    if (m.Puesto != null) r[m.Puesto] = puesto;
    
    // Manejo de Estatus (si la columna existe, si no, intenta crearla o ignorarla)
    if (m.Estatus != null) {
        r[m.Estatus] = estatus;
    } else {
        // Opcional: Crear columna si no existe (avanzado) o ignorar
    }

    // Si se da de baja, podríamos borrar el hash para impedir login
    if (estatus === 'Baja' && m.PasswordHash != null) {
        r[m.PasswordHash] = 'DISABLED'; 
    }

    sh.getRange(idx + 2, 1, 1, headers.length).setValues([r]);
    clearCache(DB.USERS);

    return { ok: true };
  });
}

/**
 * Eliminar usuario
 */
function deleteUser(email) {
  if (!email) throw new Error('Email requerido');

  return withLock_(() => {
    const sh = getSheet(DB.USERS);
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);

    const idx = rows.findIndex(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
    if (idx < 0) throw new Error('Usuario no encontrado');

    sh.deleteRow(idx + 2);
    clearCache(DB.USERS);

    return { ok: true };
  });
}

// ============================================================================
// AUTENTICACIÓN Y CONTRASEÑAS
// ============================================================================

/**
 * Generar contraseña segura de 12 caracteres
 */
function generarPasswordSeguro() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789!$%&*?';
  let pass = '';
  for (let i = 0; i < 12; i++) {
    pass += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return pass;
}



/**
 * Hash de contraseña con SHA-256
 */
function hashPassword(pwd) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    pwd,
    Utilities.Charset.UTF_8
  );
  return Utilities.base64Encode(bytes);
}

/**
 * Solicitar recuperación de contraseña
 */
function requestPasswordReset(email) {
  if (!email) throw new Error('Correo faltante');

  const sh = getSheet(DB.USERS);
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);

  const idx = rows.findIndex(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
  if (idx < 0) return { ok: true }; // No revelamos si el correo existe

  if (m.ResetToken == null || m.ResetExp == null) {
    throw new Error('Sistema no configurado para reset de contraseña');
  }

  const token = Utilities.getUuid();
  const exp = Date.now() + (1000 * 60 * 30); // 30 minutos

  rows[idx][m.ResetToken] = token;
  rows[idx][m.ResetExp] = exp;

  sh.getRange(idx + 2, 1, 1, headers.length).setValues([rows[idx]]);

  const appUrl = ScriptApp.getService().getUrl();
  const resetLink = `${appUrl}?action=resetpass&token=${token}`;

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Restablecer contraseña - Mesa de ayuda',
      htmlBody: `
        <p>Recibimos una solicitud para restablecer tu contraseña.</p>
        <p><a href="${resetLink}" style="font-size:16px;color:#0d6efd">Haz clic aquí para restablecerla</a></p>
        <p>Este enlace expirará en 30 minutos.</p>
      `
    });
  } catch (e) {
    Logger.log('⚠️ Error enviando email de reset: ' + e.message);
  }

  return { ok: true };
}

/**
 * Confirmar reset de contraseña con token
 */
function resetPasswordConfirm(token, newPass) {
  if (!token || !newPass) throw new Error('Datos incompletos');

  const sh = getSheet(DB.USERS);
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);

  if (m.ResetToken == null) throw new Error('Sistema no configurado');

  const idx = rows.findIndex(r => String(r[m.ResetToken]) === token);
  if (idx < 0) throw new Error('Token inválido');

  const exp = Number(rows[idx][m.ResetExp] || 0);
  if (Date.now() > exp) throw new Error('El enlace ha expirado.');

  const hash = hashPassword(newPass);

  rows[idx][m.PasswordHash] = hash;
  rows[idx][m.ResetToken] = '';
  rows[idx][m.ResetExp] = '';

  sh.getRange(idx + 2, 1, 1, headers.length).setValues([rows[idx]]);
  clearCache(DB.USERS);

  return { ok: true };
}

/**
 * Reset de contraseña por administrador
 */
function adminResetPassword(email) {
  if (!email) throw new Error('Email requerido');

  const sh = getSheet(DB.USERS);
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);

  const idx = rows.findIndex(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
  if (idx < 0) throw new Error('Usuario no encontrado');

  const newPass = generarPasswordSeguro();
  const hash = hashPassword(newPass);

  rows[idx][m.PasswordHash] = hash;
  if (m.ResetToken != null) rows[idx][m.ResetToken] = '';
  if (m.ResetExp != null) rows[idx][m.ResetExp] = '';

  sh.getRange(idx + 2, 1, 1, headers.length).setValues([rows[idx]]);
  clearCache(DB.USERS);

  try {
    MailApp.sendEmail({
      to: email,
      subject: 'Contraseña restablecida – Mesa de ayuda',
      htmlBody: `
        <p>Tu contraseña ha sido restablecida.</p>
        <p><b>${newPass}</b></p>
        <p>Puedes iniciar sesión inmediatamente.</p>
      `
    });
  } catch (e) {
    Logger.log('⚠️ Error enviando email: ' + e.message);
  }

  return { ok: true };
}


// ============================================================================
// CATÁLOGOS
// ============================================================================

/**
 * Obtener todos los catálogos
 */
function catalogs() {
  return {
    areas: listCol(DB.AREAS, 0),
    categories: getCategoriesByArea(),
    priorities: listPairs(DB.PRIORITIES),
    statuses: listCol(DB.STATUSES, 0)
  };
}

/**
 * Listar una columna de una hoja
 */
function listCol(sheetName, idx) {
  const sh = getSheet(sheetName);
  const lr = sh.getLastRow();
  if (lr <= 1) return [];
  const data = sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues();
  return data.filter(r => String(r[idx] || '') !== '').map(r => r[idx]);
}

/**
 * Listar pares nombre-valor
 */
function listPairs(sheetName) {
  const sh = getSheet(sheetName);
  const lr = sh.getLastRow();
  if (lr <= 1) return [];
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  return data.filter(r => String(r[0] || '') !== '').map(r => ({
    nombre: r[0],
    valor: Number(r[1]) || 0
  }));
}

/**
 * Obtener categorías agrupadas por área
 */
function getCategoriesByArea() {
  const sh = getSheet(DB.CATEGORIES);
  const lr = sh.getLastRow();
  if (lr <= 1) return {};
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  const map = {};
  data.forEach(([cat, area]) => {
    if (!cat || !area) return;
    (map[area] = map[area] || []).push(cat);
  });
  return map;
}

/**
 * Obtener catálogos completos desde hojas TI y Mtto
 */
function getCatalogosDesdeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Hojas principales
  const prioridadesSheet = ss.getSheetByName('Prioridades');
  const estatusSheet = ss.getSheetByName('Estatus');
  const areasSheet = ss.getSheetByName('Areas');

  // Categorías TI - AHORA LEE 7 COLUMNAS (incluyendo PalabrasClave)
  let tiData = [];
  const tiSheet = ss.getSheetByName('Categorias_TI');
  if (tiSheet && tiSheet.getLastRow() > 1) {
    const numCols = Math.min(tiSheet.getLastColumn(), 7); // Máximo 7 columnas
    tiData = tiSheet.getRange(2, 1, tiSheet.getLastRow() - 1, numCols).getValues();
  }

  // Categorías Mantenimiento - AHORA LEE 7 COLUMNAS
  let mttoData = [];
  const mttoSheet = ss.getSheetByName('Categorias_Mtto');
  if (mttoSheet && mttoSheet.getLastRow() > 1) {
    const numCols = Math.min(mttoSheet.getLastColumn(), 7);
    mttoData = mttoSheet.getRange(2, 1, mttoSheet.getLastRow() - 1, numCols).getValues();
  }

  const categorias = {};

  const procesarFila = (row) => {
    // Columnas: A=Nombre, B=Area, C=Ubicaciones, D=Prioridad, E=SLA, F=AgenteAsignado, G=PalabrasClave
    const [nombre, area, ubicRaw, prio, sla, agenteAsignado, palabrasClave] = row.map(x => String(x ?? '').trim());
    if (!nombre || !area) return;

    if (!categorias[area]) categorias[area] = [];

    const ubicaciones = (ubicRaw || '')
      .split(',')
      .map(u => u.trim().toLowerCase())
      .filter(Boolean);

    // ========== NUEVO: Procesar palabras clave de columna G ==========
    const keywords = (palabrasClave || '')
      .split(',')
      .map(k => k.trim().toLowerCase())
      .filter(Boolean);
    // =================================================================

    categorias[area].push({
      nombre,
      ubicaciones,
      prioridadDefault: prio || '',
      slaDefault: Number(sla) || 0,
      agenteAsignado: agenteAsignado || '',
      palabrasClave: keywords  // NUEVO: Array de palabras clave
    });
  };

  tiData.forEach(procesarFila);
  mttoData.forEach(procesarFila);

  // Prioridades
  let prioridades = [];
  if (prioridadesSheet && prioridadesSheet.getLastRow() > 1) {
    prioridades = prioridadesSheet.getRange(2, 1, prioridadesSheet.getLastRow() - 1, 2)
      .getValues()
      .map(r => ({ nombre: r[0], valor: Number(r[1]) }));
  }

  // Estatus
  let estatus = [];
  if (estatusSheet && estatusSheet.getLastRow() > 1) {
    estatus = estatusSheet.getRange(2, 1, estatusSheet.getLastRow() - 1, 1)
      .getValues()
      .map(r => String(r[0] ?? '').trim())
      .filter(Boolean);
  }

  // Áreas
  let areas = [];
  if (areasSheet && areasSheet.getLastRow() > 1) {
    areas = areasSheet.getRange(2, 1, areasSheet.getLastRow() - 1, 1)
      .getValues()
      .map(r => String(r[0] ?? '').trim())
      .filter(Boolean);
  }

  return { areas, categories: categorias, priorities: prioridades, statuses: estatus };
}

/**
 * Obtener categorías por área y ubicación del usuario
 */
function getCategoriasPorAreaYUbicacion(area, ubicacionesUsuario) {
  const cats = getCatalogosDesdeSheets().categories[area] || [];
  if (!ubicacionesUsuario) return cats;

  const userUbics = String(ubicacionesUsuario)
    .split(',')
    .map(u => u.trim().toLowerCase())
    .filter(Boolean);

  return cats.filter(cat => {
    if (!cat.ubicaciones || cat.ubicaciones.length === 0) return true;
    return cat.ubicaciones.some(u => userUbics.includes(u));
  });
}

/**
 * Obtener categorías por área y ubicación seleccionada
 */
function getCategoriasPorAreaYUbicacionSeleccionada(area, ubicacion) {
  if (!area || !ubicacion) return [];

  const cats = getCatalogosDesdeSheets().categories[area] || [];
  const u = ubicacion.toLowerCase().trim();

  return cats.filter(cat => {
    if (!cat.ubicaciones || cat.ubicaciones.length === 0) return true;
    return cat.ubicaciones.includes(u);
  });
}

/**
 * Alias para obtener categorías (usado desde frontend)
 */
function fetchCategorias(area, ubicacion) {
  return getCategoriasPorAreaYUbicacionSeleccionada(area, ubicacion);
}

// CRUD Áreas
function addArea(nombre, emailAdmin) {
  if (!nombre) throw new Error('Nombre requerido');

  // Verificar duplicados
  const sh = getSheet(DB.AREAS);
  const lr = sh.getLastRow();
  if (lr > 1) {
    const existing = sh.getRange(2, 1, lr - 1, 1).getValues().flat().map(s => String(s).toLowerCase().trim());
    if (existing.includes(nombre.toLowerCase().trim())) {
      throw new Error(`El área "${nombre}" ya existe`);
    }
  }

  setValues(DB.AREAS, [[nombre.trim()]]);
  registrarAuditoriaAdmin_('Área creada', `Nombre: ${nombre}`, emailAdmin);
  return { ok: true, message: `Área "${nombre}" creada exitosamente` };
}


function deleteArea(nombre, emailAdmin) {
  const sh = getSheet(DB.AREAS);
  const lr = sh.getLastRow();
  if (lr <= 1) throw new Error('Área no encontrada');
  const data = sh.getRange(2, 1, lr - 1, 1).getValues();
  const idx = data.findIndex(r => String(r[0]).trim().toLowerCase() === String(nombre).trim().toLowerCase());
  if (idx < 0) throw new Error('Área no encontrada');
  sh.deleteRow(idx + 2);
  registrarAuditoriaAdmin_('Área eliminada', `Nombre: ${nombre}`, emailAdmin);
  return { ok: true, message: `Área "${nombre}" eliminada` };
}



// CRUD Prioridades
function addPriority(nombre, slaHoras) {
  if (!nombre) throw new Error('Nombre requerido');
  setValues(DB.PRIORITIES, [[nombre, Number(slaHoras) || 0]]);
  return { ok: true };
}

function updatePriority(nombre, slaHoras) {
  const sh = getSheet(DB.PRIORITIES);
  const lr = sh.getLastRow();
  if (lr <= 1) throw new Error('Prioridad no encontrada');
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  const idx = data.findIndex(r => String(r[0]) === nombre);
  if (idx < 0) throw new Error('Prioridad no encontrada');
  sh.getRange(idx + 2, 2).setValue(Number(slaHoras) || 0);
  return { ok: true };
}

function deletePriority(nombre) {
  const sh = getSheet(DB.PRIORITIES);
  const lr = sh.getLastRow();
  if (lr <= 1) throw new Error('Prioridad no encontrada');
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  const idx = data.findIndex(r => String(r[0]) === nombre);
  if (idx < 0) throw new Error('Prioridad no encontrada');
  sh.deleteRow(idx + 2);
  return { ok: true };
}

// CRUD Estatus
function addStatus(nombre, orden) {
  if (!nombre) throw new Error('Nombre requerido');
  setValues(DB.STATUSES, [[nombre, Number(orden) || 0]]);
  return { ok: true };
}

function updateStatus(nombre, orden) {
  const sh = getSheet(DB.STATUSES);
  const lr = sh.getLastRow();
  if (lr <= 1) throw new Error('Estatus no encontrado');
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  const idx = data.findIndex(r => String(r[0]) === nombre);
  if (idx < 0) throw new Error('Estatus no encontrado');
  sh.getRange(idx + 2, 2).setValue(Number(orden) || 0);
  return { ok: true };
}

function deleteStatus(nombre) {
  const sh = getSheet(DB.STATUSES);
  const lr = sh.getLastRow();
  if (lr <= 1) throw new Error('Estatus no encontrado');
  const data = sh.getRange(2, 1, lr - 1, 2).getValues();
  const idx = data.findIndex(r => String(r[0]) === nombre);
  if (idx < 0) throw new Error('Estatus no encontrado');
  sh.deleteRow(idx + 2);
  return { ok: true };
}

/**
 * Actualizar catálogo desde panel de administración
 */
function updateCatalogFromPanel(row) {
  const sh = getSheet(DB.CATEGORIES);
  const data = sh.getDataRange().getDisplayValues();

  const idx = data.findIndex(r =>
    (r[0] || '').trim().toLowerCase() === (row.categoria || '').trim().toLowerCase() &&
    (r[1] || '').trim().toLowerCase() === (row.area || '').trim().toLowerCase()
  );

  if (idx >= 1) {
    sh.getRange(idx + 1, 1, 1, 3).setValues([[
      row.categoria || '',
      row.area || '',
      row.ubicacion || ''
    ]]);
  } else {
    sh.appendRow([
      row.categoria || '',
      row.area || '',
      row.ubicacion || ''
    ]);
  }

  return true;
}

// ============================================================================
// UTILIDADES DE SLA Y ASIGNACIÓN
// ============================================================================

/**
 * Ajustar SLA según puesto jerárquico del usuario
 */
function getAdjustedSLA(prioHoras, puesto) {
  const p = String(puesto || '').toLowerCase().trim();
  let adjustedSLA = prioHoras;

  // Si la prioridad es Crítica o Alta, no se ajusta
  if (prioHoras <= 4) return prioHoras;

  // Reducción para puestos jerárquicos
  if (p.includes('director')) {
    adjustedSLA = Math.min(4, prioHoras);
  } else if (p.includes('gerente')) {
    adjustedSLA = Math.min(8, prioHoras);
  }

  return Math.max(1, Math.ceil(adjustedSLA));
}

/**
 * Buscar SLA de una prioridad (con ajuste por puesto)
 */
function findPrioritySLA(nombrePrioridad, emailUsuario) {
  const prios = listPairs(DB.PRIORITIES);
  const row = prios.find(p => String(p.nombre).toLowerCase() === String(nombrePrioridad || '').toLowerCase());
  const baseSLA = row ? Number(row.valor) || 0 : 0;

  // Si se proporciona email, ajustar por puesto
  if (emailUsuario) {
    const userInfo = getUserInfo(emailUsuario);
    const puesto = userInfo ? userInfo.puesto : '';
    return getAdjustedSLA(baseSLA, puesto);
  }

  return baseSLA;
}

function computeDueDate(horasSLA, fechaInicio) {
  if (!horasSLA || horasSLA <= 0) return '';
  
  const inicio = fechaInicio ? new Date(fechaInicio) : new Date();
  
  // CORRECCIÓN: Usar la función que soporta horario partido (8-14, 16-18)
  // en lugar de sumarHorasLaborales que usa horario corrido.
  const vencimiento = calcularFechaConHorarioLaboral(inicio, Number(horasSLA));
  
  return vencimiento;
}

/**
 * Formatea las horas de SLA en formato legible
 * @param {number} horas - Horas de SLA
 * @returns {string} - Texto formateado
 */
function formatearSLA(horas) {
  if (!horas || horas <= 0) return '—';
  
  const diasLaborales = Math.floor(horas / HORARIO_LABORAL.horasPorDia);
  const horasRestantes = horas % HORARIO_LABORAL.horasPorDia;
  
  if (diasLaborales > 0 && horasRestantes > 0) {
    return `${diasLaborales} día(s) y ${horasRestantes} hora(s) laborales`;
  } else if (diasLaborales > 0) {
    return `${diasLaborales} día(s) laboral(es)`;
  } else {
    return `${horasRestantes} hora(s) laborales`;
  }
}

/**
 * Calcula el porcentaje de SLA consumido
 * @param {Date} fechaCreacion - Fecha de creación del ticket
 * @param {number} horasSLA - Horas totales de SLA
 * @returns {Object} - { porcentaje, horasTranscurridas, horasRestantes, vencido }
 */
function calcularPorcentajeSLA(fechaCreacion, horasSLA) {
  if (!fechaCreacion || !horasSLA) {
    return { porcentaje: 0, horasTranscurridas: 0, horasRestantes: horasSLA || 0, vencido: false };
  }
  
  const ahora = new Date();
  const horasTranscurridas = calcularHorasLaboralesTranscurridas(fechaCreacion, ahora);
  const horasRestantes = Math.max(0, horasSLA - horasTranscurridas);
  const porcentaje = Math.min(100, (horasTranscurridas / horasSLA) * 100);
  const vencido = horasTranscurridas >= horasSLA;
  
  return {
    porcentaje: Math.round(porcentaje * 10) / 10,
    horasTranscurridas: Math.round(horasTranscurridas * 10) / 10,
    horasRestantes: Math.round(horasRestantes * 10) / 10,
    vencido
  };
}

/**
 * Obtener información de SLA para un ticket
 * @param {string} ticketId - ID del ticket
 * @returns {Object} - Información de SLA
 */
function getSLAInfo(ticketId) {
  try {
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    
    const row = rows.find(r => String(r[m.ID]).trim() === String(ticketId).trim());
    if (!row) {
      return { ok: false, error: 'Ticket no encontrado' };
    }
    
    const fechaCreacion = new Date(row[m.Fecha]);
    const horasSLA = Number(row[m.SLA_Horas]) || 0;
    const vencimiento = row[m.Vencimiento] ? new Date(row[m.Vencimiento]) : null;
    const estatus = row[m.Estatus] || '';
    
    // No calcular SLA para tickets cerrados/resueltos
    if (['Cerrado', 'Resuelto'].includes(estatus)) {
      return {
        ok: true,
        horasSLA,
        vencimiento,
        estatus,
        activo: false,
        mensaje: 'Ticket cerrado'
      };
    }
    
    const slaInfo = calcularPorcentajeSLA(fechaCreacion, horasSLA);
    
    return {
      ok: true,
      horasSLA,
      horasSLAFormateado: formatearSLA(horasSLA),
      vencimiento,
      estatus,
      activo: true,
      ...slaInfo,
      mensaje: slaInfo.vencido 
        ? `⚠️ SLA vencido hace ${Math.abs(slaInfo.horasRestantes).toFixed(1)} horas laborales`
        : `${slaInfo.horasRestantes.toFixed(1)} horas laborales restantes (${slaInfo.porcentaje}% consumido)`
    };
    
  } catch (e) {
    Logger.log('Error en getSLAInfo: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Sugerir categoría por área
 */
function suggestCategory(area) {
  const cats = getCategoriesByArea();
  const arr = cats[area] || [];
  return arr.length ? arr[0] : '';
}



/**
 * Obtener agente por defecto según área
 */
function getDefaultAgentByArea(area) {
  const areaLower = (area || '').toLowerCase();

  if (areaLower === 'sistemas') {
    const cfg = getConfig('agente_sistemas');
    if (cfg) return cfg;
    return getDefaultAgent('agente_sistemas');
  }

  if (areaLower === 'mantenimiento') {
    const cfg = getConfig('agente_mantenimiento');
    if (cfg) return cfg;
    return getDefaultAgent('agente_mantenimiento');
  }

  return '';
}

/**
 * Obtener primer agente con un rol específico
 */
function getDefaultAgent(rol) {
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);

  const found = rows.find(r => String(r[m.Rol] || '').toLowerCase() === String(rol || '').toLowerCase());
  return found ? String(found[m.Email]) : '';
}


// ============================================================================
// ARCHIVOS Y CARPETAS
// ============================================================================

/**
 * Obtener o crear carpeta base de adjuntos
 */
function getOrCreateBaseFolder_() {
  const ROOT = 'Bexalta Helpdesk';
  const ADJ = 'Adjuntos';

  const it = DriveApp.getFoldersByName(ROOT);
  const root = it.hasNext() ? it.next() : DriveApp.createFolder(ROOT);

  const it2 = root.getFoldersByName(ADJ);
  const adj = it2.hasNext() ? it2.next() : root.createFolder(ADJ);

  return adj;
}

/**
 * Asegurar que existe carpeta para un ticket
 */
function ensureTicketFolder_(folio) {
  const base = getOrCreateBaseFolder_();
  const name = 'Ticket-' + folio;
  const it = base.getFoldersByName(name);
  return it.hasNext() ? it.next() : base.createFolder(name);
}

// ============================================================================
// TICKETS
// ============================================================================

/**
 * Generar siguiente número de folio (con lock)
 */
function nextFolio() {
  return withLock_(() => {
    const sh = getSheet(DB.TICKETS);
    const last = sh.getLastRow();
    const folio = last > 1 ? (Number(sh.getRange(last, 2).getDisplayValue()) || last - 1) + 1 : 1;
    return folio;
  });
}

function createTicket(payload) {
  return withLock_(() => {
    const reportaEmail = (payload.reportaEmail || '').trim();
    const userInfo = getUserInfo(reportaEmail);
    const id = genId();
    const folio = nextFolio();
    const fecha = new Date(); // Fecha de inicio (hoy)
    
    // 1. Limpieza de datos básicos
    const area = (payload.area || (userInfo ? userInfo.area : '') || '').trim();
    const categoria = (payload.categoria || suggestCategory(area) || '').trim();
    const ubicacion = (payload.ubicacion || '').trim();
    const reportaNombre = (payload.reportaNombre || (userInfo ? userInfo.nombre : '') || '').trim();
    
    // =================================================================================
    // 2. CÁLCULO INTELIGENTE DE SLA Y PRIORIDAD
    // =================================================================================
    const configSLA = obtenerConfiguracionSLA(area, categoria, ubicacion);
    const prioridad = (payload.prioridad && payload.prioridad !== '') 
                      ? payload.prioridad 
                      : (configSLA.prioridad || 'Media');
    const slaHoras = Number(configSLA.sla) || 24;
    const venc = calcularFechaConHorarioLaboral(fecha, slaHoras);
    
    // =================================================================================
    // 3. Asignación automática
    let asignadoA = payload.asignadoA ||
      asignarAgenteEquilibrado(area, ubicacion, categoria) ||
      getDefaultAgentByArea(area) ||
      '';
    
    // 4. Carpeta del ticket (Google Drive)
    const folder = ensureTicketFolder_(folio);
    
    // 5. Datos de Visita
    const requiereVisita = payload.requiereVisita === true;
    const visitaFecha = (payload.visitaFecha || '').trim();
    const visitaHora = (payload.visitaHora || '').trim();

    if (requiereVisita) {
      if (!visitaFecha || !visitaHora) {
        return { ok: false, error: 'Debe especificar fecha y hora para la visita programada.' };
      }
      
      const fechaIngresada = new Date(visitaFecha + 'T00:00:00');
      const manana = new Date();
      manana.setDate(manana.getDate() + 1);
      manana.setHours(0,0,0,0);
      if (fechaIngresada < manana) {
        return { ok: false, error: 'La visita no puede ser a quemarropa. Por favor, prográmela con al menos un día de anticipación.' };
      }
      
      const horaNum = parseInt(visitaHora.split(':')[0], 10);
      if (horaNum < 8 || horaNum >= 18) {
        return { ok: false, error: 'La hora de visita debe estar dentro del horario laboral (08:00 - 18:00).' };
      }
    }
    
    // =================================================================================
    // 6. CONSTRUCCIÓN BLINDADA DE LA FILA (Lee directamente de la hoja)
    // =================================================================================
    const sh = getSheet(DB.TICKETS);
    const ultimaColumna = sh.getLastColumn();
    // Leemos los encabezados reales directamente de la fila 1
    const headersReales = sh.getRange(1, 1, 1, ultimaColumna).getValues()[0];
    
    // Creamos un array del tamaño exacto de la hoja, lleno de espacios en blanco
    const row = new Array(headersReales.length).fill(''); 

    // Mapeamos los valores exactamente a la columna donde pertenecen
    for (let i = 0; i < headersReales.length; i++) {
      const h = String(headersReales[i]).trim();
      
      switch (h) {
        case 'ID': row[i] = id; break;
        case 'Folio': row[i] = folio; break;
        case 'Fecha': row[i] = fecha; break;
        case 'ReportaEmail': row[i] = reportaEmail; break;
        case 'ReportaNombre': row[i] = reportaNombre; break;
        case 'Área': 
        case 'Area': row[i] = area; break;
        case 'Categoría': 
        case 'Categoria': row[i] = categoria; break;
        case 'Prioridad': row[i] = prioridad; break;
        case 'AsignadoA': 
        case 'Asignado A': row[i] = asignadoA; break;
        case 'Título': row[i] = payload.titulo || ''; break;
        case 'Descripción': row[i] = payload.descripcion || ''; break;
        case 'Estatus': row[i] = requiereVisita ? 'Visita Programada' : 'Nuevo'; break;
        case 'SLA_Horas': row[i] = slaHoras; break;
        case 'Vencimiento': row[i] = venc; break;
        case 'ÚltimaActualización': row[i] = fecha; break;
        case 'Adjuntos': row[i] = payload.adjuntos || ''; break;
        case 'Origen': row[i] = payload.origen || 'Web'; break;
        case 'CarpetaAdjuntosId': row[i] = folder ? folder.getId() : ''; break;
        case 'Ubicación': 
        case 'Ubicacion': row[i] = ubicacion; break;
        case 'Presupuesto': row[i] = payload.presupuesto || ''; break;
        case 'TipoProveedor': row[i] = payload.tipoProveedor || ''; break;
        case 'TiempoCotizacion': row[i] = payload.tiempoCotizacion || ''; break;
        case 'AprobadorEmail': row[i] = payload.aprobadorEmail || ''; break;
        case 'StatusAutorizacion': row[i] = ''; break;
        case 'RequiereVisita': row[i] = requiereVisita ? 'Sí' : 'No'; break;
        case 'FechaVisita': 
        case 'Fecha Visita': row[i] = visitaFecha || ''; break;
        case 'HoraVisita': 
        case 'Hora Visita': row[i] = visitaHora ? "'" + visitaHora : ''; break; // Apóstrofe para forzar texto
        case 'NotasVisita': row[i] = payload.notasVisita || ''; break;
        case 'FueReabierto': row[i] = false; break;
        case 'ContadorReaperturas': row[i] = 0; break;
      }
    }
    
    // 7. Guardar ticket en Sheet (Alineación perfecta garantizada)
    sh.appendRow(row);
    SpreadsheetApp.flush(); 

    // Guardar en la hoja de visitas
    if (requiereVisita) {
      registrarEnHojaVisitas_({
        ticketId: id,
        folio: folio,
        agente: asignadoA || 'Sin asignar',
        accion: 'Programada',
        fechaVisita: visitaFecha,
        horaVisita: visitaHora,
        notas: payload.notasVisita || 'Visita solicitada en la creación del ticket'
      });
    }
    
    // 8. Registro en Bitácora
    registrarBitacora(
      id,
      'Creación',
      `Folio #${folio} · Área ${area} · Pri ${prioridad} · SLA ${slaHoras}h · Ubic ${ubicacion || 'N/A'}` +
      (requiereVisita ? ` · Visita ${visitaFecha} ${visitaHora}` : '')
    );
    
    if (asignadoA) {
      registrarBitacora(id, 'Asignación automática', `Asignado a ${asignadoA} (por ubicación: ${ubicacion || 'balanceo'})`);
    }
    
    clearCache(DB.TICKETS);
    
    // =================================================================================
    // 9. NOTIFICACIONES
    // =================================================================================
    
    if (asignadoA) {
      notifyUser(asignadoA, 'nuevo_ticket', 'Nuevo ticket asignado',
        `Se te ha asignado el ticket #${folio}: "${payload.titulo || 'Sin título'}"`,
        { ticketId: id, folio });
        
      try {
        notificarAgenteNuevoTicket(asignadoA, {
          folio,
          titulo: payload.titulo || 'Sin título',
          area,
          ubicacion,
          prioridad,
          reportaNombre,
          descripcion: payload.descripcion || '',
          visitaFecha: visitaFecha, 
          visitaHora: visitaHora
        });
      } catch (e) {
        Logger.log('⚠️ Error enviando email al agente: ' + e.message);
      }
      
      try {
        const ticketInfo = {
          folio, titulo: payload.titulo || 'Sin título', area, ubicacion, prioridad,
          reporta: reportaNombre, descripcion: payload.descripcion || ''
        };
        notificarAgenteTelegramGrupo(asignadoA, '🎫 <b>NUEVO TICKET ASIGNADO</b>', ticketInfo);
      } catch (e) {}
    }
    
    try {
      const gerente = getGerenteAreaConTelegram(area);
      if (gerente && gerente.chatId) {
        const ticketInfoGerente = { folio, titulo: payload.titulo || 'Sin título', area, prioridad };
        notificarGerenteTelegram(area, 'Nuevo ticket en tu área', ticketInfoGerente);
      }
    } catch (e) {}
    
    try {
      const cfg = getConfig();
      if (cfg.telegram_chat_admin && cfg.telegram_token) {
        let msg = `🎫 <b>Nuevo Ticket</b>\n<b>${payload.titulo || 'Sin título'}</b>\nÁrea: ${area}\nUbicación: ${ubicacion || 'No especificada'}\nPrioridad: ${prioridad}\nAsignado a: ${asignadoA || 'Sin asignar'}`;
        if (requiereVisita) msg += `\n📅 Visita: ${visitaFecha} ${visitaHora}`;
        telegramSend(msg, cfg.telegram_chat_admin);
      }
    } catch (e) {}
    
    // 10. Respuesta Final
    return {
      ok: true, id, folio,
      folderId: folder ? folder.getId() : '',
      asignadoA, requiereVisita, visitaFecha, visitaHora
    };
  });
}


/**
 * Subir archivos a la carpeta del ticket
 * CORREGIDO: Guarda nombre|URL para poder mostrar el nombre correcto
 */
function uploadFiles(ticketId, folio, files) {
  if (!files || !files.length) return { ok: true, uploaded: [] };

  return withLock_(() => {
    const sh = getSheet(DB.TICKETS);
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);

    const idx = rows.findIndex(r => r[m.ID] === ticketId);
    if (idx < 0) throw new Error('Ticket no encontrado');

    const folderId = rows[idx][m.CarpetaAdjuntosId];
    const folder = DriveApp.getFolderById(folderId || ensureTicketFolder_(folio).getId());

    const links = [];
    let fileIndex = 1;
    
    // Contar archivos existentes para numerar correctamente
    const existingFiles = String(rows[idx][m.Adjuntos] || '').split(';').filter(Boolean);
    fileIndex = existingFiles.length + 1;
    
    files.forEach((f, i) => {
      const base64 = String(f.dataUrl || '').split(',')[1] || '';
      if (!base64) return;

      const bin = Utilities.base64Decode(base64);
      
      // Obtener extensión del archivo original
      const originalName = f.name || 'archivo';
      const ext = originalName.includes('.') ? originalName.split('.').pop() : '';
      
      // Crear nombre legible: "Archivo 1.pdf", "Archivo 2.jpg", etc.
      const friendlyName = `Archivo ${fileIndex}${ext ? '.' + ext : ''}`;
      fileIndex++;
      
      let blob = Utilities.newBlob(bin, f.mimeType || MimeType.BINARY, friendlyName);

      // Comprimir si es muy grande
      if (blob.getBytes().length > 10 * 1024 * 1024) {
        blob = Utilities.zip([blob], friendlyName + '.zip');
      }

      const created = folder.createFile(blob);
      created.setDescription(`Original: ${originalName}`); // Guardar nombre original en descripción
      
      // Guardar formato: nombreVisible|URL
      links.push(`${friendlyName}|${created.getUrl()}`);
    });

    // Actualizar campo de adjuntos
    const prev = rows[idx][m.Adjuntos] || '';
    const merged = (prev ? prev + ';' : '') + links.join(';');
    sh.getRange(idx + 2, m.Adjuntos + 1).setValue(merged);

    clearCache(DB.TICKETS);
    touchTicket(ticketId);
    registrarBitacora(ticketId, 'Adjuntos', `Archivos subidos: ${links.length}`);

    return { ok: true, uploaded: links };
  });
}

// ╔═══════════════════════════════════════════════════════════════════════╗
// ║  listTickets — VERSIÓN CORREGIDA                                    ║
// ║  FIX: Eliminado bloque 3.6 "Filtrado temprano de cerrados"         ║
// ║       que removía tickets cerrados/cancelados antes del filtro      ║
// ║       de rol, causando que desaparecieran de Mis Tickets,          ║
// ║       Reportes, vista Agente y cálculos de KPIs.                   ║
// ║                                                                     ║
// ║  REEMPLAZAR función completa listTickets en backend (Code.gs)       ║
// ╚═══════════════════════════════════════════════════════════════════════╝

function listTickets(filter, emailManual) {
  const LOG_PREFIX = '[listTickets]';
  
  try {
    // ══════════════════════════════════════════════════════════════════
    // 1. RESOLVER USUARIO
    // ══════════════════════════════════════════════════════════════════
    const emailSesion = (Session.getActiveUser && Session.getActiveUser().getEmail && Session.getActiveUser().getEmail()) || '';
    const email = (emailManual && String(emailManual).trim()) ? String(emailManual).trim() : (emailSesion || '');
    
    const user = getUser(email);
    if (!user || !user.email) return { items: [], total: 0 };

    const uMail = String(user.email || '').trim().toLowerCase();
    
    // ══════════════════════════════════════════════════════════════════
    // 2. DETECCIÓN DE ROLES (JERARQUÍA)
    // ══════════════════════════════════════════════════════════════════
    const SUPERADMINS_HARDCODED = ['rgnava@bexalta.com', 'rgnava@bexalta.mx', 'admin', 'RCEsquivel'];
    const esSuperAdmin = SUPERADMINS_HARDCODED.some(sa => uMail.includes(sa.toLowerCase().split('@')[0]));
    const areaGerente = getGerenteArea(user.email);
    const esGerente = !!areaGerente;
    
    // ══════════════════════════════════════════════════════════════════
    // 3. OBTENER DATOS
    // ══════════════════════════════════════════════════════════════════
    const data = getCachedData(DB.TICKETS);
    if (!data || !Array.isArray(data) || !data.length) return { items: [], total: 0 };

    const hdr = HEADERS.Tickets;
    const tz = Session.getScriptTimeZone();
    const f = s => String(s || '').trim().toLowerCase();

    // 3.5 Variables de vista (necesarias para sección 5)
    const viewRole = filter ? (filter.role || 'usuario') : 'usuario';
    const esVistaGerencia = viewRole.startsWith('gerente') || viewRole === 'superadmin' || (filter && filter.view === 'gerencia');

    // ══════════════════════════════════════════════════════════════════
    // 4. NORMALIZAR FILAS
    // ══════════════════════════════════════════════════════════════════
    // NOTA: NO se hace filtrado previo de cerrados aquí.
    //       Todos los estatus pasan al filtro de rol (sección 5).
    let rows = data.map((r) => {
      const obj = Object.fromEntries(hdr.map((h, i) => [h, (r[i] === undefined || r[i] === null) ? '' : r[i]]));
      
      obj['Área'] = obj['Área'] || obj['Area'] || '';
      obj['Categoría'] = obj['Categoría'] || obj['Categoria'] || '';
      obj['Título'] = obj['Título'] || obj['Titulo'] || '';
      obj['Ubicación'] = obj['Ubicación'] || obj['Ubicacion'] || '';
      obj['ID'] = obj['ID'] || obj['Id'] || obj['id'] || '';
      
      obj._areaNorm = f(obj['Área']);
      obj._asignadoNorm = f(obj['AsignadoA']);
      obj._estatusNorm = f(obj['Estatus']);
      obj._prioridadNorm = f(obj['Prioridad']);
      obj._reportaEmailNorm = f(obj['ReportaEmail']);

      ['Fecha', 'Vencimiento', 'ÚltimaActualización'].forEach(k => {
        if (obj[k] instanceof Date) obj[k] = Utilities.formatDate(obj[k], tz, "yyyy-MM-dd'T'HH:mm:ss");
        else obj[k] = obj[k] ? String(obj[k]) : '';
      });
      return obj;
    });

    // ══════════════════════════════════════════════════════════════════
    // 5. LÓGICA DE FILTRADO SEGÚN "MODO DE VISTA"
    // ══════════════════════════════════════════════════════════════════

    Logger.log(`${LOG_PREFIX} Usuario: ${uMail} | EsSuperAdmin: ${esSuperAdmin} | Vista: ${viewRole} | Total bruto: ${rows.length}`);

    if (viewRole === 'usuario') {
      // --- CASO 1: "MIS TICKETS" ---
      // Solo tickets REPORTADOS por el usuario (incluye cerrados)
      rows = rows.filter(x => x._reportaEmailNorm === uMail);
      Logger.log(`   -> Mis Tickets Reportados: ${rows.length}`);
    } 
    else if (viewRole.includes('agente')) {
      // --- CASO 2: VISTA AGENTE ---
      // Solo tickets ASIGNADOS al usuario (incluye cerrados)
      rows = rows.filter(x => x._asignadoNorm === uMail);
      
      if (viewRole === 'agente_sistemas') {
        rows = rows.filter(x => x._areaNorm === 'sistemas');
      } else if (viewRole === 'agente_mantenimiento') {
        rows = rows.filter(x => x._areaNorm === 'mantenimiento');
      }
      Logger.log(`   -> Tickets Asignados (${viewRole}): ${rows.length}`);
    } 
    else if (esVistaGerencia) {
      // --- CASO 3: GERENCIA / ADMIN ---
      if (esSuperAdmin) {
        // SUPERADMIN: Ve TODO
        if (filter && filter.area && filter.area.toLowerCase() !== 'todas' && filter.area !== '') {
          rows = rows.filter(x => x._areaNorm === f(filter.area));
        }
        Logger.log(`   -> SuperAdmin Gerencia: ${rows.length}`);
      } 
      else if (esGerente) {
        const areaNorm = String(areaGerente || '').toLowerCase();
        if (areaNorm === 'clientes') {
          // Gerente Clientes: tickets de usuarios supervisados
          const supervisados = getUsuariosSupervisados().map(u => f(u));
          rows = rows.filter(x => supervisados.includes(x._reportaEmailNorm));
        } else {
          // Gerente de área: solo su área
          rows = rows.filter(x => x._areaNorm === areaNorm);
        }
        Logger.log(`   -> Gerente ${areaNorm}: ${rows.length}`);
      } 
      else {
        rows = [];
      }
    } 
    else {
      // Fallback
      rows = rows.filter(x => x._reportaEmailNorm === uMail);
    }

    // ══════════════════════════════════════════════════════════════════
    // 6. FILTROS DINÁMICOS (Buscador, Estatus, Fechas, etc.)
    // ══════════════════════════════════════════════════════════════════
    
    if (filter) {
      // --- Texto (Búsqueda) ---
      if (filter.q) {
        const q = f(filter.q);
        rows = rows.filter(x => {
          const folio = f(x['Folio']);
          const titulo = f(x['Título']);
          const desc = f(x['Descripción']);
          const rep = f(x['ReportaNombre']) + f(x['ReportaEmail']);
          const asig = f(x['AsignadoA']);
          const ubic = f(x['Ubicación']);
          return folio.includes(q) || titulo.includes(q) || desc.includes(q) || 
                 rep.includes(q) || asig.includes(q) || ubic.includes(q);
        });
      }

      // --- Estatus ---
      if (filter.estatus && filter.estatus.length) {
        const ev = filter.estatus.filter(e => e && f(e) !== 'todos' && f(e) !== '');
        if (ev.length > 0) rows = rows.filter(x => ev.some(e => x._estatusNorm === f(e)));
      }

      // --- Área (filtro dinámico adicional, útil para Reportes) ---
      if (filter.area && !esVistaGerencia) {
        // Solo aplicar si no es vista gerencia (ahí ya se aplicó en sección 5)
        rows = rows.filter(x => x._areaNorm === f(filter.area));
      }

      // --- Asignado A ---
      if (filter.asignadoA) rows = rows.filter(x => x._asignadoNorm === f(filter.asignadoA));
      
      // --- Prioridad ---
      if (filter.prioridad && filter.prioridad !== 'Todas') rows = rows.filter(x => x._prioridadNorm === f(filter.prioridad));
      
      // --- Ubicación ---
      if (filter.ubicacion) rows = rows.filter(x => f(x['Ubicación']).includes(f(filter.ubicacion)));

      // --- Fechas ---
      if (filter.dateFrom) {
        const dFrom = new Date(filter.dateFrom).getTime();
        rows = rows.filter(x => new Date(x['Fecha']).getTime() >= dFrom);
      }
      if (filter.dateTo) {
        const dTo = new Date(filter.dateTo);
        dTo.setHours(23, 59, 59);
        rows = rows.filter(x => new Date(x['Fecha']).getTime() <= dTo.getTime());
      }
      
      // --- Filtros Especiales (KPIs del Dashboard) ---
// --- Filtros Especiales (KPIs del Dashboard y stat cards) ---
      const now = Date.now();
      if (filter.special) {
        const isClosed = (st) => ['cerrado', 'resuelto', 'cancelado'].includes(st);
        
        switch (filter.special) {

          case 'activos':
            // Tickets NO cerrados que NO están vencidos (en tiempo)
            rows = rows.filter(x => {
              if (isClosed(x._estatusNorm)) return false;
              const v = x['Vencimiento'] ? new Date(String(x['Vencimiento']).replace(' ', 'T')).getTime() : null;
              return !v || isNaN(v) || v >= now;
            });
            break;

          case 'vencidos':
            // Tickets NO cerrados cuyo vencimiento ya pasó
            rows = rows.filter(x => {
              if (isClosed(x._estatusNorm)) return false;
              const v = x['Vencimiento'] ? new Date(String(x['Vencimiento']).replace(' ', 'T')).getTime() : null;
              return v && !isNaN(v) && v < now;
            });
            break;

          case 'porVencer':
            // NO cerrados con vencimiento en las próximas 4 horas
            rows = rows.filter(x => {
              if (isClosed(x._estatusNorm)) return false;
              const v = x['Vencimiento'] ? new Date(String(x['Vencimiento']).replace(' ', 'T')).getTime() : null;
              if (!v || isNaN(v)) return false;
              const h = (v - now) / 36e5;
              return h > 0 && h <= 4;
            });
            break;

          case 'resueltos':
          case 'resueltosTotales':
            // Todos los cerrados/resueltos/cancelados
            rows = rows.filter(x => isClosed(x._estatusNorm));
            break;

          case 'enProceso':
            rows = rows.filter(x => x._estatusNorm === 'en proceso');
            break;

          case 'sinAsignar':
            rows = rows.filter(x => {
              return (!x['AsignadoA'] || String(x['AsignadoA']).trim() === '') && !isClosed(x._estatusNorm);
            });
            break;

          case 'escalados':
            rows = rows.filter(x => x._estatusNorm === 'escalado');
            break;

          case 'visitasProgramadas':
          case 'visitas':
            rows = rows.filter(x => x._estatusNorm === 'visita programada');
            break;

          case 'pausados':
            rows = rows.filter(x => x._estatusNorm === 'en pausa' || x._estatusNorm === 'pausado');
            break;

          // Si llega un valor no reconocido, no filtrar (mostrar todo)
          default:
            Logger.log('[listTickets] ⚠️ filter.special no reconocido: ' + filter.special);
            break;
        }
      }
    }

// ══════════════════════════════════════════════════════════════════
    // 7. CALCULAR COUNTS GLOBALES (ANTES de paginar)
    // ══════════════════════════════════════════════════════════════════
      // 7. CALCULAR COUNTS GLOBALES (ANTES de paginar)
    const now7 = new Date();
    const counts = { 
      total: 0, activos: 0, vencidos: 0, resueltos: 0, 
      enProceso: 0, sinAsignar: 0, 
      escalados: 0, escaladosActivos: 0, escaladosVencidos: 0,
      nuevos: 0, visitaProgramada: 0 
    };

    rows.forEach(x => {
      counts.total++;
      const est = x._estatusNorm;
      const isClosed = est === 'cerrado' || est === 'resuelto' || est === 'cancelado';

      if (isClosed) {
        counts.resueltos++;
      } else {
        // ── Verificar vencimiento ANTES de clasificar ──
        const v = x['Vencimiento'] ? new Date(String(x['Vencimiento']).replace(' ', 'T')) : null;
        const estaVencido = v && !isNaN(v.getTime()) && v < now7;

        if (estaVencido) {
          counts.vencidos++;                              // SOLO vencidos
          if (est === 'escalado') counts.escaladosVencidos++;
        } else {
          counts.activos++;                               // SOLO activos (mutuamente exclusivo)
          if (est === 'escalado') counts.escaladosActivos++;
        }

        // Sub-conteos globales (siempre, independiente de vencido/activo)
        if (est === 'escalado') counts.escalados++;
        if (est === 'nuevo') counts.nuevos++;
        if (est === 'en proceso') counts.enProceso++;
        if (est === 'visita programada') counts.visitaProgramada++;
        if (!x['AsignadoA'] || String(x['AsignadoA']).trim() === '') counts.sinAsignar++;
      }
    });

    // ══════════════════════════════════════════════════════════════════
    // 8. ORDENAR
    // ══════════════════════════════════════════════════════════════════
    rows.sort((a, b) => {
      const fa = new Date(a['Fecha']).getTime() || 0;
      const fb = new Date(b['Fecha']).getTime() || 0;
      return fb - fa;
    });

    // ══════════════════════════════════════════════════════════════════
    // 9. PAGINAR
    // ══════════════════════════════════════════════════════════════════
    const total = rows.length;
    const page = (filter && filter.page) || 1;
    const pageSize = (filter && filter.pageSize) || 50;
    const start = (page - 1) * pageSize;

    Logger.log(`${LOG_PREFIX} ✅ ${total} tickets | counts: activos=${counts.activos} vencidos=${counts.vencidos} resueltos=${counts.resueltos}`);

    return {
      items: rows.slice(start, start + pageSize),
      total,
      page,
      pageSize,
      counts    // ← NUEVO: conteos globales SIEMPRE disponibles
    };

  } catch (e) {
    Logger.log(`${LOG_PREFIX} ❌ ERROR: ${e.message}\n${e.stack}`);
    return { items: [], total: 0, error: e.message };
  }
}



function esSuperAdmin(identificador) {
  const SUPERADMINS = [
    'rgnava',
    'RGNava',
    'rgnava@bexalta.com',
    'rgnava@bexalta.mx',
    'admin','RCEsquivel'
  ];
  
  const id = String(identificador || '').toLowerCase().trim();
  const idSinDominio = id.split('@')[0];
  
  return SUPERADMINS.some(sa => {
    const saLower = sa.toLowerCase();
    return id === saLower || idSinDominio === saLower;
  });
}


/**
 * Obtener detalle completo de un ticket
 */
function getTicket(id) {
  const safeId = String(id || '').trim();
  if (!safeId) return { ticket: null, comments: [] };

  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);

  const row = rows.find(r => String(r[m.ID]).trim() === safeId);
  if (!row) return { ticket: null, comments: [] };

  // Columnas que SÍ son fechas completas (con hora)
  const columnasFechaCompleta = [
    'Fecha', 'Vencimiento', 'ÚltimaActualización', 'FechaPrimeraRespuesta',
    'FechaResolucion', 'FechaCierreReal', 'FechaVisitaConfirmada'
  ];
  
  // Columnas que son SOLO fecha (sin hora)
  const columnasSoloFecha = ['FechaVisita'];
  
  // Columnas que son SOLO hora (NO convertir a Date, formatear como HH:mm)
  const columnasHora = ['HoraVisita'];

  const ticket = {};
  headers.forEach((h, i) => {
    let val = row[i];
    
    // Si el valor está vacío, null o undefined
    if (val === null || val === undefined || val === '') {
      ticket[h] = '';
      return;
    }
    
    // ============================================
    // COLUMNAS DE FECHA COMPLETA (con hora)
    // ============================================
    if (columnasFechaCompleta.includes(h)) {
      if (val instanceof Date && !isNaN(val.getTime())) {
        val = Utilities.formatDate(val, tz, 'dd/MM/yyyy HH:mm');
      } else if (typeof val === 'string' && val.trim()) {
        const d = new Date(val);
        if (!isNaN(d.getTime())) {
          val = Utilities.formatDate(d, tz, 'dd/MM/yyyy HH:mm');
        }
        // Si no es fecha válida, dejar el valor original
      }
    }
    // ============================================
    // COLUMNAS DE SOLO FECHA (sin hora)
    // ============================================
    else if (columnasSoloFecha.includes(h)) {
      if (val instanceof Date && !isNaN(val.getTime())) {
        val = Utilities.formatDate(val, tz, 'yyyy-MM-dd');
      } else if (typeof val === 'string' && val.trim()) {
        // Si ya viene como string, intentar normalizarla
        const d = new Date(val);
        if (!isNaN(d.getTime())) {
          val = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
        }
        // Si no es fecha válida, dejar el valor original
      }
    }
    // ============================================
    // COLUMNAS DE HORA - Formatear correctamente
    // Google Sheets guarda horas como Date con fecha 1899-12-30
    // ============================================
    else if (columnasHora.includes(h)) {
      if (val instanceof Date && !isNaN(val.getTime())) {
        // Corrección del bug de zona horaria 1899 de Google Sheets
        let hrs = val.getHours();
        let mins = val.getMinutes();
        if (mins >= 36) mins -= 36;
        else if (mins === 6) mins = 30;
        val = String(hrs).padStart(2, '0') + ':' + String(mins).padStart(2, '0');
      } else if (typeof val === 'string' && val.trim()) {
        val = val.replace(/'/g, '').trim(); // Quita el apóstrofe de texto
        const match = String(val).match(/(\d{1,2}):(\d{2})/);
        if (match) val = match[1].padStart(2, '0') + ':' + match[2];
        else val = '';
      } else {
        val = '';
      }
    }
    // ============================================
    // TODAS LAS DEMÁS COLUMNAS - NO tocar
    // ============================================
    
    ticket[h] = val;
  });

  ticket.ID = safeId;

  // Obtener comentarios
  let comments = [];
  try {
    comments = listComments(safeId);
  } catch (err) {
    Logger.log('⚠️ Error al cargar comentarios: ' + err.message);
  }

  return JSON.parse(JSON.stringify({ ticket, comments }));
}

/**
 * Actualizar ticket
 * 
 * CAMBIO: Escribe FechaResuelto (col Z) cuando el estatus pasa a
 * Resuelto / Cerrado / Completado / Cancelado.
 * Solo escribe si el campo está VACÍO (no sobreescribe si ya fue cerrado).
 * Limpia FechaResuelto si el ticket se REABRE.
 */
function updateTicket(id, fields) {
  return withLock_(() => {
    const sh = getSheet(DB.TICKETS);
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);

    const idx = rows.findIndex(r => String(r[m.ID]) === String(id));
    if (idx < 0) throw new Error('Ticket no encontrado');

    const row = rows[idx];
    const prevStatus     = row[m.Estatus];
    const prevAssign     = row[m.AsignadoA];
    const prevAuthStatus = m.StatusAutorizacion != null ? row[m.StatusAutorizacion] : null;
    const reportaEmail   = row[m.ReportaEmail];
    const folio          = row[m.Folio];

    // Aplicar cambios
    Object.keys(fields || {}).forEach(k => {
      if (m[k] != null && row[m[k]] !== fields[k]) {
        row[m[k]] = fields[k];
      }
    });

    // Actualizar timestamps y SLA
    row[m['ÚltimaActualización']] = new Date();

    // ─────────────────────────────────────────────────────────────────────
    // FechaResuelto: registrar el momento exacto en que el agente cierra
    // _readTableByHeader_ lee los headers REALES del Sheet, así que
    // m['FechaResuelto'] apunta directo a la columna Z (o donde esté).
    // ─────────────────────────────────────────────────────────────────────
    const ESTATUS_CIERRE  = ['resuelto', 'cerrado', 'completado', 'cancelado'];
    const ESTATUS_ACTIVOS = ['nuevo', 'abierto', 'en proceso', 'en cotización',
                             'visita programada', 'en espera', 'escalado'];

    const nuevoEstatusLow = String(fields['Estatus'] || row[m.Estatus] || '').toLowerCase();
    const idxFechaRes     = m['FechaResuelto'];  // detectado automáticamente

    if (idxFechaRes != null) {
      if (ESTATUS_CIERRE.includes(nuevoEstatusLow)) {
        // Solo escribir si aún está vacío (respetar la primera fecha de cierre)
        if (!row[idxFechaRes]) {
          row[idxFechaRes] = new Date();
          Logger.log(`[FechaResuelto] #${folio} → "${fields['Estatus']}" | ${row[idxFechaRes]}`);
        }
      } else if (ESTATUS_ACTIVOS.includes(nuevoEstatusLow) && 'Estatus' in fields) {
        // Ticket reabierto → limpiar para que el próximo cierre registre la fecha correcta
        if (row[idxFechaRes]) {
          Logger.log(`[FechaResuelto] #${folio} REABIERTO → limpiando: ${row[idxFechaRes]}`);
          row[idxFechaRes] = '';
        }
      }
    } else {
      Logger.log('⚠️ [FechaResuelto] Columna no encontrada en la hoja Tickets. Verifica que el header sea exactamente "FechaResuelto".');
    }
    // ─────────────────────────────────────────────────────────────────────

    if ('Prioridad' in fields) {
      const sla = findPrioritySLA(row[m.Prioridad], reportaEmail);
      row[m.SLA_Horas] = sla;
      // Recalcular vencimiento desde la fecha original de creación
      const fechaCreacion = row[m.Fecha] ? new Date(row[m.Fecha]) : new Date();
      row[m.Vencimiento] = computeDueDate(sla, fechaCreacion);
    }

    // Persistir
    sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
    clearCache(DB.TICKETS);

    // Bitácora y notificaciones
    if ('Estatus' in fields && fields.Estatus !== prevStatus) {
      registrarBitacora(id, 'Cambio de estatus', `${prevStatus || '—'} → ${fields.Estatus}`);
      addSystemComment(id, `Se cambió el estatus a "${fields.Estatus}".`, true);

      notifyUser(reportaEmail, 'status', 'Actualización de estatus',
        `Tu ticket #${folio} cambió a "${fields.Estatus}".`,
        { ticketId: id, folio, status: fields.Estatus });
    }

    if (m.StatusAutorizacion != null && row[m.StatusAutorizacion] !== prevAuthStatus) {
      registrarBitacora(id, 'Autorización Cotización', `${prevAuthStatus || '—'} → ${row[m.StatusAutorizacion]}`);
    }

    registrarTiempoAtencion(id, fields['Estatus']);

    if ('AsignadoA' in fields && fields.AsignadoA !== prevAssign) {
      registrarBitacora(id, 'Reasignación', `${prevAssign || '—'} → ${fields.AsignadoA || '—'}`);
    }

    // Enviar notificaciones si se resuelve
    if (String(row[m.Estatus]).toLowerCase() === 'resuelto') {
      const cfg  = getConfig();
      const url  = ScriptApp.getService().getUrl();
      const link = url + '?tid=' + encodeURIComponent(id) + '&action=close';

      // 1. EMAIL al usuario reportador
      try {
        const cuerpoHtml = `
          <div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;">
            <div style="background:#059669;color:white;padding:20px;text-align:center;border-radius:8px 8px 0 0;">
              <h2 style="margin:0;">✅ Ticket Resuelto</h2>
            </div>
            <div style="padding:25px;background:#f8fafc;border:1px solid #e2e8f0;">
              <p>Hola <strong>${reportaNombre || 'Usuario'}</strong>,</p>
              <p>Tu ticket <strong>#${folio}</strong> ha sido marcado como <strong style="color:#059669;">RESUELTO</strong>.</p>
              
              <div style="background:white;padding:15px;border-radius:8px;margin:20px 0;border-left:4px solid #059669;">
                <p style="margin:0 0 10px;"><strong>Título:</strong> ${row[m['Título']] || 'Sin título'}</p>
                <p style="margin:0;"><strong>Agente:</strong> ${row[m.AsignadoA] || 'N/A'}</p>
              </div>
              
              <p>Tienes <strong>48 horas</strong> para confirmar que el problema fue solucionado.</p>
              
              <div style="text-align:center;margin:25px 0;">
                <a href="${link}" style="display:inline-block;background:#059669;color:white;padding:12px 30px;text-decoration:none;border-radius:6px;font-weight:bold;">
                  ✓ Confirmar y Cerrar Ticket
                </a>
              </div>
              
              <p style="color:#64748b;font-size:13px;">Si no respondes en 48 horas, el ticket se cerrará automáticamente.</p>
            </div>
          </div>
        `;
        MailApp.sendEmail({
          to: reportaEmail,
          subject: `✅ Ticket #${folio} - Resuelto`,
          htmlBody: cuerpoHtml
        });
        Logger.log(`✅ Email de resolución enviado a ${reportaEmail}`);
      } catch (e) {
        Logger.log('⚠️ Error enviando email de resolución: ' + e.message);
      }

      // 2. NOTIFICACIÓN en sistema
      crearNotificacion(reportaEmail, 'resuelto', 'Ticket resuelto',
        `Tu ticket #${folio} fue marcado como resuelto. Por favor confirma.`, id);

      // 3. TELEGRAM al usuario (si tiene configurado)
      try {
        const keyTgUser  = 'tg_' + reportaEmail.split('@')[0].toLowerCase();
        const chatIdUser = getConfig(keyTgUser);
        if (chatIdUser) {
          const msgTg = `✅ *Ticket #${folio} Resuelto*\n\n` +
                        `Tu ticket ha sido marcado como resuelto.\n` +
                        `Por favor confirma que el problema fue solucionado.\n\n` +
                        `🔗 ${link}`;
          telegramSend(msgTg, chatIdUser);
        }
      } catch (e) {
        Logger.log('⚠️ Error Telegram usuario resolución: ' + e.message);
      }

      // 4. TELEGRAM al admin
      try {
        if (cfg.telegram_chat_admin && cfg.telegram_token) {
          const msgAdmin = `✅ *Ticket #${folio} RESUELTO*\n` +
                           `*Usuario:* ${reportaNombre || reportaEmail}\n` +
                           `*Agente:* ${row[m.AsignadoA] || 'N/A'}`;
          telegramSend(msgAdmin, cfg.telegram_chat_admin);
        }
      } catch (e) {
        Logger.log('⚠️ Error Telegram admin: ' + e.message);
      }
    }

    return { ok: true };
  });
}

/**
 * Obtener configuración de gerentes ultrarrápido
 */
function getGerentes() {
  try {
    const headers = HEADERS.ConfigGerentes;
    const rows = getCachedData(DB.GERENTES); // Leemos de caché, no del sheet
    const m = _headerMap_(headers);
    
    if(!rows || rows.length === 0) return [];

    return rows.map(r => ({
      area: r[m.Area] || '',
      email: r[m.GerenteEmail] || '',
      nombre: r[m.GerenteNombre] || ''
    })).filter(g => g.email);
  } catch (e) {
    Logger.log('⚠️ Error obteniendo gerentes: ' + e.message);
    return [];
  }
}

function getGerenteArea(email) {
  if (!email) return null;
  
  const gerentes = getGerentes();
  const emailLower = email.toLowerCase();
  const emailSinDominio = emailLower.split('@')[0]; // Obtener solo la parte antes del @
  
  const found = gerentes.find(g => {
    const gerenteEmail = (g.email || '').toLowerCase();
    // Comparar email completo O solo el usuario
    return gerenteEmail === emailLower || 
           gerenteEmail === emailSinDominio ||
           gerenteEmail.split('@')[0] === emailSinDominio;
  });
  
  return found ? found.area : null;
}

/**
 * Usuarios que la Gerencia de Clientes puede monitorear.
 */
function getUsuariosSupervisados() {
  // Retornamos los nombres de usuario o correos base que la gerente puede ver
  return ['maguzman', 'milopez', 'ralopez'].map(u => u.toLowerCase());
}

/**
 * Obtener rol efectivo incluyendo gerentes
 * MODIFICAR getUser o crear función auxiliar
 */
function getRolEfectivo(email) {
  const user = getUserByEmail(email);
  if (!user) return 'usuario';
  
  const rolBase = String(user.Rol || 'usuario').toLowerCase();
  
  // Si ya tiene rol de admin/agente, mantenerlo
  if (['admin', 'agente_sistemas', 'agente_mantenimiento'].includes(rolBase)) {
    return rolBase;
  }
  
  // Verificar si es gerente
  const areaGerente = getGerenteArea(email);
  if (areaGerente) {
    return 'gerente_' + areaGerente.toLowerCase().replace(/\s+/g, '_');
  }
  
  return rolBase;
}


/**
 * Actualizar timestamp de última modificación
 */
function touchTicket(id) {
  if (!id) return;

  const sh = getSheet(DB.TICKETS);
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);

  const idx = rows.findIndex(r => String(r[m.ID]).trim() === String(id).trim());
  if (idx < 0) return;

  const colUltimaAct = m['ÚltimaActualización'] + 1;
  if (colUltimaAct <= 0) return;

  sh.getRange(idx + 2, colUltimaAct).setValue(new Date());
}

function agentTouchesTicket(id, userEmail) {
  const u = getUser(userEmail);
  const role = String(u.rol || '').toLowerCase();

  // Solo agentes y admin pueden iniciar atención
  if (!['agente_sistemas', 'agente_mantenimiento', 'admin'].includes(role)) {
    return { ok: false, reason: 'no-agent', message: 'Solo agentes pueden iniciar atención' };
  }

  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);

  const idx = rows.findIndex(r => String(r[m.ID]) === String(id));
  if (idx < 0) return { ok: false, reason: 'not-found' };

  const row = rows[idx];
  const asignadoA = (row[m.AsignadoA] || '').toLowerCase().trim();
  const userEmailLower = (u.email || '').toLowerCase().trim();
  const folio = row[m.Folio];
  const titulo = row[m['Título']] || row[m.Titulo] || '';
  const reportaEmail = row[m.ReportaEmail] || '';
  const reportaNombre = row[m.ReportaNombre] || '';
  
  // ========== VALIDAR: Solo el agente asignado puede iniciar ==========
  if (role !== 'admin' && asignadoA && asignadoA !== userEmailLower) {
    return { 
      ok: false, 
      reason: 'not-assigned', 
      message: `Este ticket está asignado a otro agente (${row[m.AsignadoA]})` 
    };
  }
  
  // CORRECCIÓN: Validar que el agente pertenece al área del ticket
  const ticketArea = (row[m['Área']] || '').toLowerCase();
  if (role === 'agente_sistemas' && ticketArea !== 'sistemas' && !asignadoA) {
    return { ok: false, reason: 'wrong-area', message: 'Este ticket pertenece al área de Mantenimiento' };
  }
  if (role === 'agente_mantenimiento' && ticketArea !== 'mantenimiento' && !asignadoA) {
    return { ok: false, reason: 'wrong-area', message: 'Este ticket pertenece al área de Sistemas' };
  }
  // ====================================================================

  if (String(row[m.Estatus]) === 'Nuevo') {
    updateTicket(id, { Estatus: 'En Proceso' });
    addSystemComment(id, 'El agente abrió el ticket. Estatus cambiado a "En Proceso".', true);
    
    const agente = u.nombre || u.email;
    
    if (reportaEmail) {
      notifyUser(reportaEmail, 'ticket_en_proceso', 
        `Tu ticket #${folio} está siendo atendido`,
        `El agente ${agente} ha iniciado la atención de tu solicitud: "${titulo}"`,
        { ticketId: id, folio, agente }
      );
      
      try {
        const cuerpoHtml = `
          <h2>🔧 Tu Ticket está siendo Atendido</h2>
          <p>Hola <strong>${reportaNombre || 'Usuario'}</strong>,</p>
          <p>Tu solicitud está siendo atendida.</p>
          
          <div style="background:#f3f4f6;padding:15px;border-radius:8px;margin:15px 0;">
            <h3 style="margin:0 0 10px;">Ticket #${folio}</h3>
            <p style="margin:0;"><strong>${titulo}</strong></p>
          </div>
          
          <p><strong>Atendido por:</strong> ${agente}</p>
          <p><strong>Estado:</strong> <span style="color:#2563eb;font-weight:bold;">EN PROCESO</span></p>
        `;
        
        //enviarEmailNotificacion(reportaEmail, `🔧 Ticket #${folio} - En Proceso`, cuerpoHtml);
      } catch (e) {
        Logger.log('Error enviando email de inicio: ' + e.message);
      }
    }
    
    return { ok: true, changed: true };
  }

  return { ok: true, changed: false };
}


/**
 * Reasignar tickets sin movimiento en 48h
 */
function reasignarTicketsInactivos() {
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);

  const now = new Date();

  rows.forEach(r => {
    const est = String(r[m.Estatus]);
    if (['Nuevo', 'Abierto', 'En Proceso'].includes(est)) {
      const ult = new Date(r[m['ÚltimaActualización']]);
      const horas = (now - ult) / 36e5;

      if (horas > 48) {
        const id = r[m.ID];
        const area = r[m['Área']];
        const nuevo = asignarAgenteEquilibrado(area);

        if (nuevo) {
          updateTicket(id, { AsignadoA: nuevo });
          registrarBitacora(id, 'Reasignación automática', `Sin movimiento por 48h. Asignado a ${nuevo}`);
        }
      }
    }
  });
}

// ============================================================================
// COMENTARIOS
// ============================================================================

/**
 * Limpia el HTML del editor para que Telegram no arroje error "Unsupported start tag"
 */
function limpiarHtmlParaTelegram(htmlRaw) {
  if (!htmlRaw) return '';
  
  // 1. Convertir saltos de línea HTML a saltos de texto reales
  let texto = htmlRaw.replace(/<br\s*[\/]?>/gi, '\n');
  texto = texto.replace(/<\/p>/gi, '\n');
  texto = texto.replace(/<\/div>/gi, '\n');
  texto = texto.replace(/<\/li>/gi, '\n');
  
  // 2. Convertir etiquetas de Quill a etiquetas que Telegram SÍ acepta
  texto = texto.replace(/<strong>/gi, '<b>').replace(/<\/strong>/gi, '</b>');
  texto = texto.replace(/<em>/gi, '<i>').replace(/<\/em>/gi, '</i>');
  
  // 3. Eliminar absolutamente TODO el resto del HTML no soportado
  texto = texto.replace(/<(?!\/?(b|i|u|s|a|code|pre)(?=>|\s.*>))\/?.*?>/gi, '');
  
  // 4. Limpiar entidades HTML comunes
  texto = texto.replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&');
  
  // 5. Quitar excesos de saltos de línea
  return texto.replace(/\n{3,}/g, '\n\n').trim();
}

/**
 * Función principal para agregar comentarios a un ticket con soporte para adjuntos reales en Telegram
 */
function addComment(ticketId, comentario, emailUsuario, interno, fileId) {
  Logger.log(`\n===========================================`);
  Logger.log(`=== INICIO addComment ===`);
  
  if (!ticketId || (!comentario && !fileId)) {
    throw new Error('Datos insuficientes para comentar');
  }

  return withLock_(() => {
    // 1. OBTENER USUARIO QUE COMENTA
    const u = getUser(emailUsuario || '');
    
    // 2. VALIDACIÓN DE ESTATUS CERRADO
    const { headers: tHeaders, rows: tRows } = _readTableByHeader_(DB.TICKETS);
    const tm = _headerMap_(tHeaders);
    const tRow = tRows.find(r => String(r[tm.ID]) === String(ticketId));
    
    if (tRow && String(tRow[tm.Estatus] || '').toLowerCase() === 'cerrado') {
      throw new Error('⛔ El ticket está CERRADO y no admite más comentarios.');
    }

    // 3. OBTENER DATOS DEL TICKET PARA PERMISOS Y NOTIF
    const ticketData = getTicketById(ticketId);
    const AsignadoA = ticketData?.AsignadoA || '';
    const ReportaEmail = ticketData?.ReportaEmail || '';
    const Folio = ticketData?.Folio || '';

    let fileUrl = '';
    let fileName = '';
    let fileIdRetorno = fileId || ''; 

    // 4. PROCESAR ARCHIVO (Si existe)
    if (fileId && fileId !== '') {
      try {
        const file = DriveApp.getFileById(fileId);
        fileName = file.getName();
        fileUrl = file.getUrl();
        // Dar permisos de lectura a los involucrados
        if (AsignadoA) try { file.addViewer(AsignadoA); } catch(e){}
        if (ReportaEmail && ReportaEmail !== AsignadoA) try { file.addViewer(ReportaEmail); } catch(e){}
      } catch (e) {
        Logger.log('⚠️ Error procesando archivo: ' + e.message);
        fileIdRetorno = ''; fileUrl = ''; fileName = '';
      }
    }

    // 5. GUARDAR EN HOJA DE COMENTARIOS
    const sh = getSheet(DB.COMMENTS);
    const fecha = new Date();
    sh.appendRow([
      genId(), 
      ticketId, 
      fecha, 
      u.email, 
      u.nombre || '', 
      comentario || '', 
      !!interno, 
      fileIdRetorno || '', 
      fileUrl || '', 
      fileName || ''
    ]);

    // 6. ACTUALIZAR TIMESTAMP DEL TICKET (Touch)
    touchTicket(ticketId);
    
    // 7. BITÁCORA (Texto plano para la celda)
    const comentarioPlano = (comentario || '').replace(/<[^>]*>?/gm, '');
    registrarBitacora(ticketId, (interno ? 'Comentario interno' : 'Comentario'), (fileIdRetorno ? '📎 ' : '') + comentarioPlano.slice(0, 180));
    
    clearCache(DB.COMMENTS);

    // 8. NOTIFICACIONES (Solo si no es interno)
    if (!interno && u.email && ticketData) {
      
      const norm = e => String(e || '').trim().toLowerCase().split('@')[0];
      const autorBase = norm(u.email);
      const reportaBase = norm(ReportaEmail);
      const asignadoBase = norm(AsignadoA);
      
      // Limpieza de HTML para el cuerpo del mensaje de Telegram
      const comentarioLimpioTg = limpiarHtmlParaTelegram(comentario || '');

      // --- CASO A: NOTIFICAR AL AGENTE ASIGNADO ---
      if (asignadoBase && asignadoBase !== autorBase) {
        try {
          // Notif en App
          crearNotificacion(AsignadoA, 'comentario', 'Nuevo comentario en tu ticket', (u.nombre || u.email) + ' comentó el #' + Folio, ticketId);
          
          // Email
          const agenteData = getUserByEmail(AsignadoA);
          if (agenteData?.Email) {
            notificarNuevoComentario(agenteData.Email, ticketData, u.nombre || u.email, comentario, false, fileUrl);
          }
          
          // Telegram (Envío inteligente con archivo real si existe)
          const msgTg = `💬 <b>Comentario en Ticket #${Folio}</b>\nDe: ${u.nombre || u.email}\n\n${comentarioLimpioTg}`;
          telegramSendToGrupo(getTelegramChatIdGrupo(AsignadoA), msgTg, 'HTML', fileIdRetorno);
          
        } catch (e) { Logger.log('Error notif agente: ' + e.message); }
      }

      // --- CASO B: NOTIFICAR AL USUARIO QUE REPORTÓ ---
      if (reportaBase && reportaBase !== autorBase) {
        try {
          // Notif en App
          crearNotificacion(ReportaEmail, 'comentario', 'Nuevo comentario', 'Respondieron tu ticket #' + Folio, ticketId);
          
          // Email
          notificarNuevoComentario(ReportaEmail, ticketData, u.nombre || u.email, comentario, false, fileUrl);
          
          // Telegram (Envío inteligente con archivo real si existe)
          const msgTgUs = `💬 <b>Respuesta en tu Ticket #${Folio}</b>\nDe: ${u.nombre || u.email}\n\n${comentarioLimpioTg}`;
          telegramSendToGrupo(getTelegramChatIdGrupo(ReportaEmail), msgTgUs, 'HTML', fileIdRetorno);
          
        } catch (e) { Logger.log('Error notif usuario: ' + e.message); }
      }
    }

    Logger.log(`=== FIN addComment ===\n===========================================`);

    return { 
      ok: true, 
      fileUrl: fileUrl, 
      fileName: fileName, 
      timestamp: fecha.getTime() 
    };
  });
}


/**
 * ============================================================
 * FUNCIÓN AUXILIAR: Obtener comentarios con archivos adjuntos
 * Reemplaza la anterior listComments()
 * ============================================================
 */

function listComments(ticketId) {
  const sh = getSheet(DB.COMMENTS);
  const lr = sh.getLastRow();
  if (lr <= 1) return [];

  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const data = sh.getRange(2, 1, lr - 1, sh.getLastColumn()).getValues();

  return data
    .filter(r => String(r[1]).trim() === String(ticketId).trim())
    .map(r => {
      const fecha = r[2] instanceof Date ? r[2] : new Date(r[2]);
      const fechaFmt = isNaN(fecha) ? '' : Utilities.formatDate(fecha, tz, 'dd/MM/yyyy HH:mm');
      
      // ✅ AHORA RETORNA TODOS LOS CAMPOS INCLUYENDO ARCHIVOS
      return {
        id: r[0],
        ticketId: r[1],
        fecha: fechaFmt,
        autorEmail: r[3],
        autorNombre: r[4],
        comentario: r[5],
        interno: r[6] === true || r[6] === 'TRUE' || r[6] === true,
        fileId: r[7] || '',        // ✅ ID del archivo en Drive
        fileUrl: r[8] || '',       // ✅ URL del archivo (lo que el frontend necesita)
        fileName: r[9] || ''       // ✅ Nombre del archivo
      };
    });
}


/**
 * ============================================================
 * FUNCIÓN PARA SUBIR ARCHIVO EN COMENTARIO
 * Sube el archivo a la carpeta del ticket
 * ============================================================
 */

/**
 * ============================================================
 * FUNCIÓN: uploadCommentFile
 * Sube archivo adjunto a comentario
 * ============================================================
 * 
 * REQUISITOS:
 * - Las funciones getOrCreateBaseFolder_() y ensureTicketFolder_()
 *   ya existen en tu código
 * - Los archivos se guardan en Drive
 * - Se retorna el fileId y fileUrl
 */

function uploadCommentFile(ticketId, fileData, emailUsuario) {
  // ========== VALIDACIÓN ==========
  if (!fileData || !fileData.data) {
    return {
      ok: false,
      error: 'Sin datos de archivo'
    };
  }

  if (!ticketId) {
    return {
      ok: false,
      error: 'No hay ticket ID'
    };
  }

  try {
    // ========== 1. OBTENER DATOS DEL TICKET ==========
    // Necesitamos el Folio para crear la carpeta del ticket
    const ticketData = getTicketById(ticketId);
    if (!ticketData) {
      return {
        ok: false,
        error: 'Ticket no encontrado'
      };
    }

    const folio = ticketData.Folio || ticketId;
    Logger.log(`[uploadCommentFile] Procesando archivo para Ticket #${folio}`);

    // ========== 2. OBTENER CARPETA DEL TICKET ==========
    // Usa la función que ya existe en tu código
    const ticketFolder = ensureTicketFolder_(folio);
    if (!ticketFolder) {
      return {
        ok: false,
        error: 'No se pudo crear/obtener carpeta del ticket'
      };
    }

    Logger.log(`[uploadCommentFile] ✅ Carpeta del ticket: ${ticketFolder.getName()}`);

    // ========== 3. CONVERTIR BASE64 A BLOB ==========
    let bytes;
    try {
      bytes = Utilities.base64Decode(fileData.data);
    } catch (decodeErr) {
      return {
        ok: false,
        error: 'Error decodificando archivo: ' + decodeErr.message
      };
    }

    const blob = Utilities.newBlob(bytes, fileData.type || 'application/octet-stream', fileData.name);

    Logger.log(`[uploadCommentFile] Blob creado: ${fileData.name} (${bytes.length} bytes)`);

    // ========== 4. GUARDAR ARCHIVO EN DRIVE ==========
    let uploadedFile;
    try {
      uploadedFile = ticketFolder.createFile(blob);
    } catch (createErr) {
      return {
        ok: false,
        error: 'Error creando archivo en Drive: ' + createErr.message
      };
    }

    // ========== 5. CONFIGURAR PERMISOS ==========
    // Por defecto, solo el propietario (Bexalta helpdesk) puede ver
    // Los permisos específicos se otorgan en addComment()
    try {
      uploadedFile.setSharing(
        DriveApp.Access.ANYONE_WITH_LINK,
        DriveApp.Permission.VIEW
      );
      Logger.log(`[uploadCommentFile] ✅ Permisos configurados: ANYONE_WITH_LINK VIEW`);
    } catch (permErr) {
      Logger.log(`⚠️ Error configurando permisos: ${permErr.message}`);
      // Continuar sin romper el flujo
    }

    // ========== 6. OBTENER DATOS DEL ARCHIVO ==========
    const fileUrl = uploadedFile.getUrl();
    const fileId = uploadedFile.getId();
    const fileName = uploadedFile.getName();

    Logger.log(`[uploadCommentFile] ✅ Archivo guardado:`);
    Logger.log(`  - ID: ${fileId}`);
    Logger.log(`  - Nombre: ${fileName}`);
    Logger.log(`  - URL: ${fileUrl}`);

    // ========== 7. RETORNAR ÉXITO ==========
    return {
      ok: true,
      fileId: fileId,
      fileUrl: fileUrl,
      fileName: fileName,
      size: bytes.length,
      uploadedAt: new Date().getTime()
    };

  } catch (e) {
    Logger.log(`[uploadCommentFile] ❌ Error general: ${e.message}`);
    Logger.log(`  Stack: ${e.stack}`);
    return {
      ok: false,
      error: e.message
    };
  }
}



function getTicketById(ticketId) {
  if (!ticketId) return null;
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    
    const row = rows.find(r => String(r[m.ID] || '').trim() === String(ticketId).trim());
    if (!row) return null;
    
    // Retornar objeto con AMBAS versiones (mayúsculas y minúsculas)
    // para compatibilidad con todas las funciones
    return {
      // Versión con mayúsculas (para destructuring existente)
      ID: row[m.ID],
      Folio: row[m.Folio],
      Fecha: row[m.Fecha],
      ReportaEmail: row[m.ReportaEmail],
      ReportaNombre: row[m.ReportaNombre],
      Área: row[m['Área']],
      Categoría: row[m['Categoría']],
      Prioridad: row[m.Prioridad],
      AsignadoA: row[m.AsignadoA],
      Título: row[m['Título']],
      Descripción: row[m['Descripción']],
      Estatus: row[m.Estatus],
      Ubicación: row[m['Ubicación']],
      Vencimiento: row[m.Vencimiento],
      
      // Versión con minúsculas (para funciones de email)
      id: row[m.ID],
      folio: row[m.Folio],
      fecha: row[m.Fecha],
      reportaEmail: row[m.ReportaEmail],
      reportaNombre: row[m.ReportaNombre],
      area: row[m['Área']],
      categoria: row[m['Categoría']],
      prioridad: row[m.Prioridad],
      asignadoA: row[m.AsignadoA],
      titulo: row[m['Título']],
      descripcion: row[m['Descripción']],
      estatus: row[m.Estatus],
      ubicacion: row[m['Ubicación']],
      vencimiento: row[m.Vencimiento]
    };
  } catch (e) {
    Logger.log('⚠️ Error en getTicketById: ' + e.message);
    return null;
  }
}


// ============================================================================
// 2. getUserByEmail - Obtener usuario por email
// AGREGAR DESPUÉS DE getUserInfo (~línea 615)
// ============================================================================

/**
 * Obtener datos de un usuario por email
 * @param {string} email - Email del usuario
 * @returns {Object|null} - Datos del usuario o null
 */
function getUserByEmail(email) {
  if (!email) return null;
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    
    const emailLower = String(email).toLowerCase().trim();
    const row = rows.find(r => String(r[m.Email] || '').toLowerCase().trim() === emailLower);
    
    if (!row) return null;
    
    return {
      Email: row[m.Email],
      Nombre: row[m.Nombre],
      Rol: row[m.Rol],
      Área: row[m['Área']],
      Ubicación: row[m['Ubicación']],
      Puesto: row[m.Puesto],
      Disponible: row[m.Disponible],
      EmailNotificacion: row[m.EmailNotificacion]
    };
  } catch (e) {
    Logger.log('⚠️ Error en getUserByEmail: ' + e.message);
    return null;
  }
}
/**
 * Agregar comentario del sistema
 */
function addSystemComment(ticketId, texto, interno) {
  try {
    const u = getUser();
    const sh = getSheet(DB.COMMENTS);
    sh.appendRow([genId(), ticketId, new Date(), u.email || 'sistema', u.nombre || 'Sistema', texto, !!interno]);
    registrarBitacora(ticketId, interno ? 'Comentario interno' : 'Comentario', texto.slice(0, 180));
  } catch (e) {
    Logger.log('⚠️ Error en addSystemComment: ' + e.message);
  }
}

// ============================================================================
// BITÁCORA
// ============================================================================

/**
 * Registrar entrada en bitácora
 */
function registrarBitacora(ticketId, accion, detalle) {
  const sh = getSheet(DB.LOG);

  // Asegurar encabezados
  if (sh.getLastRow() < 1) {
    sh.getRange(1, 1, 1, HEADERS.Bitacora.length).setValues([HEADERS.Bitacora]);
  }

  const u = getUser();
  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';

  sh.appendRow([
    Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd HH:mm:ss'),
    ticketId,
    u.email || '',
    accion,
    detalle || ''
  ]);
}

/**
 * Obtener bitácora de un ticket
 */
function getBitacora(ticketId) {
  const sh = getSheet(DB.LOG);
  const lr = sh.getLastRow();
  if (lr <= 1) return [];

  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';
  const data = sh.getRange(2, 1, lr - 1, HEADERS.Bitacora.length).getValues();

  return data
    .filter(r => String(r[1]).trim() === String(ticketId).trim())
    .map(r => {
      const fecha = r[0] instanceof Date ? r[0] : new Date(r[0]);
      const fechaFmt = isNaN(fecha) ? '' : Utilities.formatDate(fecha, tz, 'dd/MM/yyyy HH:mm');
      return [fechaFmt, r[1], r[2], r[3], r[4]];
    });
}

/**
 * Registrar tiempo de atención al resolver/cerrar
 */
function registrarTiempoAtencion(ticketId, nuevoEstatus) {
  if (!nuevoEstatus || !['Resuelto', 'Cerrado'].includes(nuevoEstatus)) return;

  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);

  const idx = rows.findIndex(r => String(r[m.ID]) === ticketId);
  if (idx < 0) return;

  const row = rows[idx];
  const fechaCreacion = new Date(row[m.Fecha]);
  const fechaCierre = new Date();
  const horas = (fechaCierre - fechaCreacion) / 36e5;

  const duracion = horas.toFixed(2);
  registrarBitacora(ticketId, 'Tiempo de atención', `${duracion} h desde creación`);
}

/**
 * Marcar ticket como visto
 */
function marcarTicketVisto(ticketId, email) {
  registrarBitacora(ticketId, 'Visualización', `Ticket abierto por ${email}`);
}

// ============================================================================
// NOTIFICACIONES
// ============================================================================


function notifyUser(email, tipo, titulo, mensaje, meta = {}) {
  try {
    if (!email) return;
    const sh = getSheet(DB.NOTIFS);
    
    if (sh.getLastRow() < 1) {
      sh.getRange(1, 1, 1, HEADERS.Notificaciones.length).setValues([HEADERS.Notificaciones]);
    }

    const id = genId();
    const now = new Date();
    const ticketId = meta.ticketId || meta.TicketID || '';

    // ORDEN CORRECTO
    const row = [
      id,              // ID
      now,             // Fecha
      email,           // Usuario
      tipo || 'info',  // Tipo
      titulo || '',    // Título
      mensaje || '',   // Mensaje
      ticketId,        // TicketID ← IMPORTANTE
      false,           // Leido
      now.getTime()    // Timestamp
    ];
    
    sh.appendRow(row);
    clearCache(DB.NOTIFS);
  } catch (err) {
    Logger.log('⚠️ Error en notifyUser: ' + err.message);
  }
}

/**
 * Registrar notificación (alias)
 */
function registrarNotificacion(email, titulo, mensaje) {
  notifyUser(email, 'general', titulo, mensaje, {});
}

// ============================================================================
// MÉTRICAS Y REPORTES
// ============================================================================

/**
 * Estadísticas de un agente
 */
function getAgenteStats(email) {
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);

  const stats = { total: 0, abiertos: 0, proceso: 0, espera: 0, resueltos: 0, cerrados: 0 };

  rows.forEach(r => {
    const asignado = String(r[m.AsignadoA] || '').toLowerCase();
    const est = String(r[m.Estatus] || '').toLowerCase();

    if (asignado === email.toLowerCase()) {
      stats.total++;
      if (est.includes('nuevo') || est.includes('abierto')) stats.abiertos++;
      else if (est.includes('proceso')) stats.proceso++;
      else if (est.includes('espera')) stats.espera++;
      else if (est.includes('resuelto')) stats.resueltos++;
      else if (est.includes('cerrado')) stats.cerrados++;
    }
  });

  return stats;
}

/**
 * Rendimiento de un agente (promedio de tiempos)
 */
function getRendimientoAgente(email) {
  const sh = getSheet(DB.LOG);
  const lr = sh.getLastRow();
  if (lr <= 1) return { total: 0, promedio: 0 };

  const data = sh.getRange(2, 1, lr - 1, HEADERS.Bitacora.length).getValues();
  const m = _headerMap_(HEADERS.Bitacora);

  const registros = data.filter(r => String(r[m.Usuario]).toLowerCase() === String(email).toLowerCase());

  const tiempos = registros
    .filter(r => r[m['Acción']] === 'Tiempo de atención')
    .map(r => Number((r[m.Detalle] || '').replace(/[^\d.]/g, '')));

  const promedio = tiempos.length ? (tiempos.reduce((a, b) => a + b, 0) / tiempos.length).toFixed(2) : 0;

  return { total: tiempos.length, promedio };
}

/**
 * Resumen completo de agente
 */
function getResumenAgente(email) {
  const stats = getAgenteStats(email);
  const rendimiento = getRendimientoAgente(email);

  return {
    ...stats,
    promedioHoras: rendimiento.promedio,
    ticketsResueltos: rendimiento.total
  };
}

/**
 * Tickets vencidos de un agente
 */
function getTicketsVencidosAgente(email) {
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);
  const now = new Date();
  const tz = Session.getScriptTimeZone() || 'America/Mexico_City';

  return rows.filter(r => {
    const asig = String(r[m.AsignadoA] || '').toLowerCase();
    const venc = new Date(r[m.Vencimiento]);
    const est = String(r[m.Estatus] || '').toLowerCase();
    return asig === email.toLowerCase() && venc < now && !['resuelto', 'cerrado'].includes(est);
  }).map(r => ({
    folio: r[m.Folio],
    titulo: r[m['Título']],
    vencimiento: Utilities.formatDate(new Date(r[m.Vencimiento]), tz, 'dd/MM/yyyy HH:mm')
  }));
}

function getReportMetrics(filter, userEmail) {
  try {
    const user = getUser(userEmail);
    const role = (user.rol || '').toLowerCase();

    let userArea = '';
    if (role === 'agente_sistemas') userArea = 'Sistemas';
    else if (role === 'agente_mantenimiento') userArea = 'Mantenimiento';
    else if (role === 'gerente_sistemas' || (role === 'gerente' && (user.area || '').toLowerCase() === 'sistemas')) userArea = 'Sistemas';
    else if (role === 'gerente_mantenimiento' || (role === 'gerente' && (user.area || '').toLowerCase() === 'mantenimiento')) userArea = 'Mantenimiento';

    const isAdmin = ['admin', 'superadmin', 'director'].includes(role);
    if (!isAdmin && userArea) filter.area = userArea;
    if (!filter.area && userArea) filter.area = userArea;

    // OPTIMIZACIÓN: Usar caché en vez de getDataRange directo
    const hdr = HEADERS.Tickets;
    const data = getCachedData(DB.TICKETS);
    if (!data || !data.length) return { total: 0, cerrados: 0, vencidos: 0, aTiempo: 0, tarde: 0, promedioHoras: 0, items: [] };

    const m = {};
    hdr.forEach((h, i) => m[h] = i);

    // Filtros
    const fDesde = filter.dateFrom ? new Date(filter.dateFrom).setHours(0, 0, 0, 0) : null;
    const fHasta = filter.dateTo ? new Date(filter.dateTo).setHours(23, 59, 59, 999) : null;
    const filterArea = (filter.area || '').toLowerCase().trim();
    const filterPrio = (filter.prioridad || '').toLowerCase().trim();
    const filterUbic = (filter.ubicacion || '').toLowerCase().trim();
    const filterStatus = (filter.estatus && filter.estatus.length > 0)
      ? filter.estatus.map(s => s.toLowerCase()) : [];

    const now = new Date();
    let total = 0, cerrados = 0, vencidos = 0, aTiempo = 0, tarde = 0;
    let sumaTiempoResolucion = 0, countResueltos = 0;
    const items = [];

    const f = s => String(s || '').trim().toLowerCase();

    for (let i = 0; i < data.length; i++) {
      const r = data[i];
      const rArea = f(r[m['Área']]);
      const rStatus = f(r[m.Estatus]);
      const rPrio = f(r[m.Prioridad]);
      const rUbic = f(r[m['Ubicación']]);

      // Aplicar filtros
      if (filterArea && rArea !== filterArea) continue;

      const rFecha = r[m.Fecha] instanceof Date ? r[m.Fecha] : new Date(r[m.Fecha]);
      if (fDesde && rFecha < fDesde) continue;
      if (fHasta && rFecha > fHasta) continue;
      if (filterPrio && rPrio !== filterPrio) continue;
      if (filterStatus.length > 0 && !filterStatus.includes(rStatus)) continue;
      if (filterUbic && !rUbic.includes(filterUbic)) continue;

      total++;

      const isClosed = ['cerrado', 'resuelto', 'cancelado'].includes(rStatus);
      if (isClosed) cerrados++;

      // SLA Compliance
      const rVenc = r[m.Vencimiento] instanceof Date ? r[m.Vencimiento] : (r[m.Vencimiento] ? new Date(r[m.Vencimiento]) : null);

      if (!isClosed && rVenc && rVenc < now) {
        vencidos++;
      }

      if (isClosed && rVenc) {
        const ultAct = r[m['ÚltimaActualización']] instanceof Date ? r[m['ÚltimaActualización']] : new Date(r[m['ÚltimaActualización']]);
        if (!isNaN(ultAct.getTime()) && !isNaN(rVenc.getTime())) {
          if (ultAct <= rVenc) {
            aTiempo++;
          } else {
            tarde++;
          }
        }
        // Calcular tiempo de resolución
        if (!isNaN(rFecha.getTime()) && !isNaN(ultAct.getTime())) {
          sumaTiempoResolucion += (ultAct - rFecha) / (1000 * 60 * 60); // en horas
          countResueltos++;
        }
      }

      items.push({
        ID: r[m.ID],
        Estatus: String(r[m.Estatus] || ''),
        Prioridad: String(r[m.Prioridad] || ''),
        Area: String(r[m['Área']] || ''),
        Ubicacion: String(r[m['Ubicación']] || '')
      });
    }

    const promedioHoras = countResueltos > 0 ? Math.round(sumaTiempoResolucion / countResueltos) : 0;

    return {
      total,
      cerrados,
      vencidos,
      aTiempo,
      tarde,
      promedioHoras,
      slaCompliance: (aTiempo + tarde) > 0 ? Math.round((aTiempo / (aTiempo + tarde)) * 100) : 0,
      items
    };

  } catch (e) {
    Logger.log('Error en getReportMetrics: ' + e.message);
    return { error: e.message };
  }
}



/**
 * Exportar tickets a CSV
 */
function exportTicketsCSV(filtro) {
  const all = listTickets({ ...filtro, page: 1, pageSize: 100000 }).items;
  const cols = ['Folio', 'Fecha', 'ReportaEmail', 'ReportaNombre', 'Área', 'Categoría', 'Prioridad', 'AsignadoA', 'Título', 'Estatus', 'SLA_Horas', 'Vencimiento', 'ÚltimaActualización', 'Adjuntos'];
  const esc = v => `"${String(v == null ? '' : v).replace(/"/g, '""')}"`;
  const rows = [cols.map(esc).join(',')].concat(
    all.map(r => cols.map(c => esc(r[c])).join(','))
  );
  const csv = rows.join('\r\n');
  const filename = `tickets_${Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'America/Mexico_City', 'yyyyMMdd_HHmmss')}.csv`;

  return { ok: true, filename, mimeType: 'text/csv', csv };
}

// ============================================================================
// SLA: VENCIMIENTOS Y RECORDATORIOS
// ============================================================================

/**
 * Revisar tickets próximos a vencer (para trigger horario)
 * Usa cálculo de horas laborales
 */
function revisarVencimientos() {
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);
  
  const ahora = new Date();
  
  // Solo ejecutar durante horas laborales
  const horaActual = ahora.getHours();
  if (horaActual < HORARIO_LABORAL.horaInicio || horaActual >= HORARIO_LABORAL.horaFin) {
    Logger.log('Fuera de horario laboral, omitiendo revisión de vencimientos');
    return;
  }
  
  if (!esDiaLaboral(ahora)) {
    Logger.log('No es día laboral, omitiendo revisión de vencimientos');
    return;
  }

  rows.forEach(r => {
    const est = String(r[m.Estatus]);
    if (['Cerrado', 'Resuelto'].includes(est)) return;

    const fechaCreacion = r[m.Fecha] ? new Date(r[m.Fecha]) : null;
    const horasSLA = Number(r[m.SLA_Horas]) || 0;
    
    if (!fechaCreacion || !horasSLA) return;
    
    // Calcular SLA restante
    const slaInfo = calcularPorcentajeSLA(fechaCreacion, horasSLA);
    
    // Avisar cuando quede 1 hora laboral o menos (y no esté vencido aún)
    if (!slaInfo.vencido && slaInfo.horasRestantes <= 1 && slaInfo.horasRestantes > 0) {
      const folio = r[m.Folio];
      const asign = r[m.AsignadoA] || '';
      const ticketId = r[m.ID];

      if (asign) {
        try {
          MailApp.sendEmail({
            to: asign,
            subject: `⚠️ Ticket #${folio} próximo a vencer`,
            htmlBody: `
              <div style="font-family: Arial, sans-serif;">
                <h3 style="color: #f59e0b;">⚠️ Alerta de SLA</h3>
                <p>El ticket <strong>#${folio}</strong> vence en menos de <strong>1 hora laboral</strong>.</p>
                <p>Horas restantes: ${slaInfo.horasRestantes.toFixed(1)}</p>
                <p>Por favor, atiéndelo con prioridad.</p>
              </div>
            `
          });
        } catch (e) {
          Logger.log('⚠️ Error enviando recordatorio: ' + e.message);
        }
        
        // Notificación interna
        notifyUser(asign, 'sla_warning', 'SLA próximo a vencer',
          `El ticket #${folio} vence en ${slaInfo.horasRestantes.toFixed(1)} horas laborales.`,
          { ticketId, folio });
      }

      registrarBitacora(ticketId, 'Aviso SLA', `Quedan ${slaInfo.horasRestantes.toFixed(1)} horas laborales`);
    }
    
    // Marcar como vencido si corresponde
    if (slaInfo.vencido && !['Vencido'].includes(est)) {
      // Opcional: cambiar estado o solo registrar
      registrarBitacora(r[m.ID], 'SLA Vencido', 
        `Excedido por ${Math.abs(slaInfo.horasRestantes).toFixed(1)} horas laborales`);
    }
  });
}

// ============================================================================
// IA: SUGERENCIA DE CATEGORÍA
// ============================================================================

function normalizarTexto(t) {
  const STOP = new Set([
    'el','la','los','las','un','una','de','del','al','a','con','para','por',
    'que','me','mi','su','tu','se','es','esto','esta','estas','estos',
    'cuando','como','donde','hacer','necesito','favor','problema','ayuda',
    'puede','hay','otro','otros','requiere','requieren','tengo','no'
  ]);

  return String(t || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length > 2 && !STOP.has(w));
}

const DICCIONARIO = {
  // SISTEMAS
  wifi: ['wifi', 'inalambrico', 'inalámbrico', 'red', 'internet', 'conexion', 'conexión', 'señal', 'router', 'modem'],
  impresora: ['impresora', 'multifuncional', 'toner', 'tóner', 'escaneo', 'escaner', 'escáner', 'scanner', 'imprimir', 'impresion', 'impresión', 'copia', 'copias'],
  telefono: ['telefono', 'teléfono', 'extension', 'extensión', 'llamada', 'voz', 'conmutador', 'voip'],
  computadora: ['computadora', 'laptop', 'pc', 'equipo', 'monitor', 'pantalla', 'teclado', 'mouse', 'raton', 'ratón', 'cpu', 'desktop'],
  software: ['programa', 'software', 'aplicacion', 'aplicación', 'app', 'sistema', 'windows', 'office', 'excel', 'word', 'outlook', 'instalacion', 'instalación', 'actualizar', 'licencia'],
  correo: ['correo', 'email', 'mail', 'outlook', 'gmail', 'enviar', 'recibir', 'bandeja'],
  acceso: ['acceso', 'permiso', 'permisos', 'usuario', 'contraseña', 'password', 'cuenta', 'bloqueado', 'carpeta', 'servidor'],
  
  // MANTENIMIENTO
  cctv: ['cctv', 'camara', 'cámara', 'camaras', 'cámaras', 'grabacion', 'grabación', 'videograbacion', 'videograbación', 'vigilancia', 'nvr', 'dvr'],
  control_acceso: ['tarjeta', 'lectora', 'control', 'torniquete', 'biometrico', 'biométrico', 'huella'],
  aire: ['aire', 'acondicionado', 'clima', 'frio', 'frío', 'calor', 'termostato', 'minisplit', 'ventilacion', 'ventilación', 'temperatura'],
  agua: ['agua', 'fuga', 'cisterna', 'sanitario', 'lavabo', 'baño', 'wc', 'inodoro', 'tuberia', 'tubería', 'drenaje', 'goteo', 'humedad', 'plomeria', 'plomería'],
  elevador: ['elevador', 'ascensor', 'atorado', 'subir', 'bajar'],
  energia: ['luz', 'energia', 'energía', 'electrica', 'eléctrica', 'electricidad', 'planta', 'foco', 'apagon', 'apagón', 'contacto', 'enchufe', 'voltaje', 'corto', 'fusible'],
  incendio: ['incendio', 'alarma', 'detector', 'extintor', 'extinguidor', 'humo', 'fuego'],
  mobiliario: ['silla', 'escritorio', 'mesa', 'mueble', 'cajon', 'cajón', 'puerta', 'ventana', 'cerradura', 'llave', 'roto', 'dañado'],
  limpieza: ['limpieza', 'basura', 'sucio', 'derrame', 'olor', 'sanitizar']
};

function similitud(tokensA, tokensB) {
  const A = new Set(tokensA);
  const B = new Set(tokensB);
  const inter = [...A].filter(x => B.has(x)).length;
  return inter / Math.max(A.size, B.size, 1);
}

function scoreCategoria(textTokens, categoria) {
  const partes = categoria.nombre.split('-').map(p => normalizarTexto(p));
  let score = 0;

  // Nivel 1 (ej. Internet / Aire Acondicionado)
  score += similitud(textTokens, partes[0]) * 40;

  // Nivel 2
  if (partes[1]) score += similitud(textTokens, partes[1]) * 30;

  // Nivel 3
  if (partes[2]) score += similitud(textTokens, partes[2]) * 20;

  // Boost semántico por diccionario interno
  for (const grupo in DICCIONARIO) {
    if (DICCIONARIO[grupo].some(w => textTokens.includes(w))) {
      if (categoria.nombre.toLowerCase().includes(grupo)) {
        score += 15;
      }
    }
  }

  // ========== NUEVO: Boost por palabras clave de la hoja (columna G) ==========
  if (categoria.palabrasClave && categoria.palabrasClave.length > 0) {
    let matchesKeywords = 0;
    
    for (const keyword of categoria.palabrasClave) {
      // Normalizar la palabra clave
      const keywordTokens = normalizarTexto(keyword);
      
      // Verificar si algún token del texto coincide con la palabra clave
      for (const token of textTokens) {
        if (keyword.includes(token) || token.includes(keyword)) {
          matchesKeywords++;
          break;
        }
        // También verificar tokens normalizados
        if (keywordTokens.some(kt => kt === token || token.includes(kt) || kt.includes(token))) {
          matchesKeywords++;
          break;
        }
      }
    }
    
    // Boost proporcional al número de matches (máximo 50 puntos)
    if (matchesKeywords > 0) {
      const boostKeywords = Math.min(50, matchesKeywords * 20);
      score += boostKeywords;
    }
  }
  // ===========================================================================

  // Ubicación
  if (categoria.ubicaciones?.length) {
    const u = normalizarTexto(categoria.ubicaciones.join(' '));
    score += similitud(textTokens, u) * 10;
  }

  return Math.round(score);
}

// ============================================================
// PARTE 3: sugerirCategoriaIA MEJORADA
// Con mejor cálculo de confianza y soporte para múltiples sugerencias
// REEMPLAZAR la función existente (línea ~14308)
// ============================================================

function sugerirCategoriaIA(descripcion, ubicacionSeleccionada) {
  if (!descripcion || descripcion.trim().length < 3) {
    return { area: '', categoria: '', score: 0, confianza: 0, sugerencias: [] };
  }

  const { categories } = getCatalogosDesdeSheets();
  const textTokens = normalizarTexto(descripcion + ' ' + (ubicacionSeleccionada || ''));

  // Si no hay tokens útiles, retornar vacío
  if (textTokens.length === 0) {
    return { area: '', categoria: '', score: 0, confianza: 0, sugerencias: [] };
  }

  let ranking = [];

  for (const area in categories) {
    for (const cat of categories[area]) {
      const score = scoreCategoria(textTokens, cat);
      if (score > 0) {
        ranking.push({ 
          area, 
          categoria: cat.nombre, 
          score,
          palabrasClave: cat.palabrasClave || []
        });
      }
    }
  }

  ranking.sort((a, b) => b.score - a.score);

  if (ranking.length === 0) {
    return { area: '', categoria: '', score: 0, confianza: 0, sugerencias: [] };
  }

  const mejor = ranking[0];
  const segundo = ranking[1];
  const tercero = ranking[2];

  // ========== CÁLCULO DE CONFIANZA MEJORADO ==========
  let confianza = 50;
  
  // Base según score absoluto
  if (mejor.score >= 80) confianza = 90;
  else if (mejor.score >= 60) confianza = 80;
  else if (mejor.score >= 40) confianza = 70;
  else if (mejor.score >= 20) confianza = 60;
  
  // Ajustar por diferencia con el segundo
  if (!segundo) {
    confianza = Math.min(98, confianza + 15); // Única opción
  } else {
    const delta = mejor.score - segundo.score;
    if (delta > 30) confianza = Math.min(95, confianza + 10);
    else if (delta > 15) confianza = Math.min(90, confianza + 5);
    else if (delta < 5) confianza = Math.max(40, confianza - 15); // Muy parecidos
  }
  // ===================================================

  // Preparar sugerencias alternativas (top 3)
  const sugerencias = ranking.slice(0, 3).map(r => ({
    area: r.area,
    categoria: r.categoria,
    score: r.score
  }));

  return {
    area: mejor.area,
    categoria: mejor.categoria,
    score: mejor.score,
    confianza,
    sugerencias
  };
}




// ============================================================================
// COTIZACIONES Y APROBACIONES
// ============================================================================

/**
 * Manejar aprobación/rechazo de cotización
 * VERSIÓN MODIFICADA: Agrega notificación por email al usuario
 */
function handleCotizacionApproval(ticketId, action) {
  const safeId     = String(ticketId).trim();
  const safeAction = String(action).trim().toLowerCase();

  if (!safeId || (safeAction !== 'approve' && safeAction !== 'reject')) {
    return 'Acción o ID de ticket no válidos.';
  }

  // Leer info del ticket una sola vez
  let ticketInfo = null;
  let agenteEmail = '';
  try {
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    const ticketRow = rows.find(r => String(r[m.ID]) === safeId);
    if (ticketRow) {
      ticketInfo = {
        folio:        ticketRow[m.Folio],
        titulo:       ticketRow[m['Título']] || 'Sin título',
        reportaEmail: ticketRow[m.ReportaEmail],
        presupuesto:  ticketRow[m.Presupuesto]
      };
      agenteEmail = String(ticketRow[m.AsignadoA] || '').trim();
      // Si AsignadoA es nombre y no email, resolver
      if (agenteEmail && !agenteEmail.includes('@') && typeof getEmailNotificacion === 'function') {
        agenteEmail = getEmailNotificacion(agenteEmail) || agenteEmail;
      }
    }
  } catch (e) {
    Logger.log('Error obteniendo info de ticket: ' + e.message);
  }

  // ─── APROBAR ────────────────────────────────────────────────────────────
  if (safeAction === 'approve') {
    updateTicket(safeId, {
      StatusAutorizacion: 'Aprobado',
      Estatus: 'En Proceso'
    });
    addSystemComment(safeId, 'Cotización APROBADA vía enlace de email.', true);

    if (ticketInfo && ticketInfo.reportaEmail) {
      try {
        notificarUsuarioAprobacion(ticketInfo.reportaEmail, {
          folio:       ticketInfo.folio,
          titulo:      ticketInfo.titulo,
          presupuesto: ticketInfo.presupuesto
        }, true, 'La cotización ha sido aprobada y el trabajo procederá según lo acordado.');
      } catch (e) { Logger.log('⚠️ Error email aprobación usuario: ' + e.message); }
    }

    // Notificar al agente que fue aprobada
    if (agenteEmail && agenteEmail.includes('@')) {
      try {
        crearNotificacion(agenteEmail, 'info',
          `Cotización aprobada: #${ticketInfo && ticketInfo.folio}`,
          'La cotización fue aprobada. El ticket vuelve a En Proceso.', safeId);
        const chatId = getTelegramChatIdGrupo(agenteEmail);
        if (chatId) telegramSendToGrupo(chatId,
          `✅ <b>Cotización APROBADA - Ticket #${ticketInfo && ticketInfo.folio}</b>\n` +
          `El aprobador autorizó la cotización. El ticket está en <b>En Proceso</b>.`);
      } catch (e) { Logger.log('⚠️ Error notif agente aprobación: ' + e.message); }
    }

    return 'Cotización Aprobada. El ticket ha regresado al estado "En Proceso".';
  }

  // ─── RECHAZAR ───────────────────────────────────────────────────────────
  if (safeAction === 'reject') {
    // Nuevo estatus dedicado — el SLA sigue corriendo
    updateTicket(safeId, {
      StatusAutorizacion: 'Rechazado',
      Estatus: 'Cotización Rechazada'          // ← antes era 'En Espera'
    });
    addSystemComment(safeId,
      'Cotización RECHAZADA vía enlace de email. El agente debe replantear o escalar. El SLA sigue corriendo.',
      true);

    const folio  = ticketInfo ? ticketInfo.folio  : safeId;
    const titulo = ticketInfo ? ticketInfo.titulo : '';

    // 1. Notificar al AGENTE (omnicanal)
    if (agenteEmail && agenteEmail.includes('@')) {
      try {
        crearNotificacion(agenteEmail, 'alerta',
          `Cotización rechazada: #${folio}`,
          'Debes replantear la cotización. El SLA sigue corriendo.', safeId);

        const chatAgente = getTelegramChatIdGrupo(agenteEmail);
        if (chatAgente) telegramSendToGrupo(chatAgente,
          `⚠️ <b>Cotización RECHAZADA - Ticket #${folio}</b>\n\n` +
          `📝 <b>${titulo}</b>\n` +
          `❌ Rechazada por el aprobador vía correo.\n\n` +
          `⏰ <b>El SLA sigue corriendo.</b> Replantea la cotización o escala el ticket.`);

        const bodyEmailAgente =
          `<div style="font-family:sans-serif;max-width:600px;">` +
          `<h2 style="color:#dc2626;">⚠️ Cotización Rechazada</h2>` +
          `<p>La cotización del ticket <strong>#${folio}</strong> fue rechazada por el aprobador.</p>` +
          `<div style="background:#fef2f2;border-left:4px solid #dc2626;padding:15px;margin:15px 0;">` +
          `<p style="margin:0;"><strong>Motivo:</strong> Rechazado vía enlace de correo electrónico.</p>` +
          `</div>` +
          `<p style="color:#dc2626;"><strong>⏰ El SLA sigue corriendo.</strong> ` +
          `Revisa y replantea la cotización a la brevedad.</p></div>`;
        enviarEmailNotificacion(agenteEmail,
          `⚠️ Cotización Rechazada - Ticket #${folio}`, bodyEmailAgente);
      } catch (e) { Logger.log('⚠️ Error notif agente rechazo: ' + e.message); }
    }

    // 2. Notificar al USUARIO que reportó
    if (ticketInfo && ticketInfo.reportaEmail) {
      try {
        notificarUsuarioAprobacion(ticketInfo.reportaEmail, {
          folio:       folio,
          titulo:      titulo,
          presupuesto: ticketInfo.presupuesto
        }, false,
          'La cotización ha sido rechazada. El equipo técnico preparará una nueva propuesta y te notificará.');

        crearNotificacion(ticketInfo.reportaEmail, 'info',
          `Cotización en revisión: #${folio}`,
          'La cotización fue rechazada. El equipo preparará una nueva propuesta.', safeId);
      } catch (e) { Logger.log('⚠️ Error notif usuario rechazo: ' + e.message); }
    }

    return 'Cotización Rechazada. El ticket cambió a "Cotización Rechazada" — el SLA sigue activo y el agente fue notificado.';
  }

  return 'Acción o ID de ticket no válidos.';
}

/**
 * Envía notificaciones de aprobación OMNICANAL (Email, Telegram y Sistema)
 * Aclara los tiempos de ejecución y costos.
 */
function notifyApproval(ticketId, listaCorreosStr, presupuesto) {
  let ticketInfo = { folio: ticketId, titulo: 'Sin título', area: '', reportaNombre: '', ubicacion: '', tiempoCotizacion: '' };
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    const ticketRow = rows.find(r => String(r[m.ID]) === String(ticketId));
    
    if (ticketRow) {
      ticketInfo = {
        folio: ticketRow[m.Folio] || ticketId,
        titulo: ticketRow[m['Título']] || 'Sin título',
        area: ticketRow[m['Área']] || '',
        ubicacion: ticketRow[m['Ubicación']] || '',
        tiempoCotizacion: ticketRow[m.TiempoCotizacion] || 'No especificado',
        reportaNombre: ticketRow[m.ReportaNombre] || ticketRow[m.ReportaEmail] || 'Usuario'
      };
    }
  } catch (e) {
    Logger.log('Error obteniendo info de ticket: ' + e.message);
  }
  
  // Separar correos: El primero es el Aprobador Real (Gerente)
  const correos = listaCorreosStr.split(',').map(e => e.trim()).filter(e => e !== '');
  if (correos.length === 0) return;

  const aprobadorReal = correos[0]; 
  const notificados = correos.slice(1);

  const urlBase = ScriptApp.getService().getUrl();
  const approvalLink = `${urlBase}?action=approve_cot&id=${ticketId}`;
  const rejectLink = `${urlBase}?action=reject_cot&id=${ticketId}`;

  // --- 1. PLANTILLA EMAIL PARA EL APROBADOR ---
  const cuerpoAprobador = `
    <div style="font-family: sans-serif; max-width: 600px; border: 1px solid #e2e8f0; padding: 20px; border-radius: 8px;">
      <h2 style="color: #0d9488;">🚨 Aprobación de Cotización Requerida</h2>
      <p>Se requiere su autorización formal para proceder con el Ticket <b>#${ticketInfo.folio}</b></p>
      
      <div style="background: #f8fafc; padding: 15px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #0d9488;">
        <p><b>📝 Título:</b> ${ticketInfo.titulo}</p>
        <p><b>📍 Área/Ubicación:</b> ${ticketInfo.area} - ${ticketInfo.ubicacion}</p>
        <p><b>⏱️ Tiempo estimado de ejecución:</b> ${ticketInfo.tiempoCotizacion} días hábiles (tras aprobación)</p>
        <p style="font-size: 22px; color: #059669; margin-top: 15px;"><b>💰 Presupuesto: $${presupuesto} MXN</b></p>
      </div>
      
      <div style="text-align: center; margin-top: 30px;">
        <a href="${approvalLink}" style="background: #059669; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; margin-right: 10px; display: inline-block;">✓ APROBAR</a>
        <a href="${rejectLink}" style="background: #dc2626; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">✗ RECHAZAR</a>
      </div>
      <p style="font-size: 12px; color: #94a3b8; text-align: center; margin-top: 20px;">*Los días de ejecución son hábiles y comienzan a correr tras su aprobación.</p>
    </div>`;

  try {
    // -------------------------------------------------------------
    // A) NOTIFICAR AL APROBADOR (OMNICANAL)
    // -------------------------------------------------------------
    
    // 1. Enviar Email
    enviarEmailNotificacion(aprobadorReal, `🚨 ACCIÓN REQUERIDA: Aprobación Ticket #${ticketInfo.folio} - $${presupuesto}`, cuerpoAprobador);
    
    // 2. Notificación en Sistema (Campanita)
    crearNotificacion(aprobadorReal, 'aprobacion', `Cotización pendiente: #${ticketInfo.folio}`, `Se requiere tu aprobación por $${presupuesto} MXN. Ejecución: ${ticketInfo.tiempoCotizacion} días hábiles.`, ticketId);
    
    // 3. Enviar Telegram
    const chatId = getTelegramChatIdGrupo(aprobadorReal);
    if (chatId) {
      const msgTg = `🚨 <b>APROBACIÓN DE COTIZACIÓN</b>\n\n` +
                    `Ticket: <code>#${ticketInfo.folio}</code>\n` +
                    `<b>${ticketInfo.titulo}</b>\n` +
                    `📍 ${ticketInfo.ubicacion}\n\n` +
                    `⏱️ <b>Tiempo de ejecución:</b> ${ticketInfo.tiempoCotizacion} días hábiles\n` +
                    `💰 <b>Presupuesto: $${presupuesto} MXN</b>\n\n` +
                    `Revisa tu correo para aprobar o rechazar con un solo clic.`;
      telegramSendToGrupo(chatId, msgTg);
    }

    // -------------------------------------------------------------
    // B) ENVIAR COPIAS A NOTIFICADOS (Solo Email Informativo)
    // -------------------------------------------------------------
    if (notificados.length > 0) {
      const cuerpoNotificado = `
        <div style="font-family: sans-serif; max-width: 600px; border: 1px solid #e2e8f0; padding: 20px;">
          <h2 style="color: #64748b;">📢 Notificación de Cotización</h2>
          <p>Se ha generado una solicitud de presupuesto para el Ticket <b>#${ticketInfo.folio}</b></p>
          <p style="color: #ef4444; font-size: 13px;"><i>Nota: Usted recibe este correo como copia informativa. La aprobación será realizada por: ${aprobadorReal}</i></p>
          <div style="background: #f8fafc; padding: 15px; border-radius: 8px;">
            <p><b>Título:</b> ${ticketInfo.titulo}</p>
            <p><b>Tiempo de ejecución:</b> ${ticketInfo.tiempoCotizacion} días hábiles</p>
            <p><b>Presupuesto: $${presupuesto} MXN</b></p>
          </div>
        </div>`;
      enviarEmailNotificacion(notificados.join(','), `📢 COPIA: Solicitud de Cotización Ticket #${ticketInfo.folio}`, cuerpoNotificado);
    }
    
    Logger.log(`✅ Flujo Omnicanal de Cotización completado para ticket #${ticketInfo.folio}`);
  } catch (e) {
    Logger.log('⚠️ Error enviando notificaciones de cotización: ' + e.message);
  }
}

// ============================================================================
// TELEGRAM
// ============================================================================

/**
 * Enviar mensaje a Telegram
 */
function telegramSend(text, chatIdOptional) {
  const cfg = getConfig();
  const token = cfg.telegram_token;
  const defaultChat = cfg.telegram_chat_id;

  if (!token) {
    Logger.log('⚠️ Falta telegram_token en Config');
    return;
  }

  const chatId = chatIdOptional || defaultChat;
  if (!chatId) {
    Logger.log('⚠️ Falta chat_id para Telegram');
    return;
  }

  const url = `https://api.telegram.org/bot${token}/sendMessage`;

  const payload = {
    chat_id: chatId,
    text: String(text),
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };

  const params = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  try {
    UrlFetchApp.fetch(url, params);
  } catch (err) {
    Logger.log('⚠️ Error enviando Telegram: ' + err.message);
  }
}


/**
 * Configurar webhook de Telegram
 */
function telegramSetWebhook() {
  const cfg = getConfig();
  const token = cfg.telegram_token;

  const webAppUrl = ScriptApp.getService().getUrl();

  const url = `https://api.telegram.org/bot${token}/setWebhook?url=${encodeURIComponent(webAppUrl)}`;
  const res = UrlFetchApp.fetch(url).getContentText();
  console.log('Webhook response:', res);
}


// ============================================================================
// BOOTSTRAP / INICIO DE SESIÓN
// ============================================================================

/**
 * Inicializar sesión de usuario
 */
function bootstrap(email, pass) {
  if (!email || !pass) {
    return { error: 'Correo y contraseña requeridos' };
  }
  
  const userInfo = getUserInfo(email);
  if (!userInfo) {
    return { error: 'Usuario no encontrado' };
  }
  
  // Validar contraseña
  const hash = hashPassword(pass);
  if (userInfo.passwordHash !== hash) {
    return { error: 'Credenciales incorrectas' };
  }
  
  // NUEVO: Verificar si usuario está activo
  if (userInfo.estatusUsuario === 'Baja') {
    return { 
      ok: false, 
      error: 'Tu cuenta ha sido desactivada. Contacta al administrador.' 
    };
  }
  
  const cats = getCatalogosDesdeSheets();
  const config = getConfig();
  
  // =====================================================
  // NUEVO: Obtener información de disponibilidad si es agente
  // =====================================================
  let disponibilidadInfo = { disponible: true, motivo: '', fechaFin: '' };
  const rolLower = (userInfo.rol || '').toLowerCase();
  
  if (rolLower.includes('agente') || rolLower === 'admin') {
    disponibilidadInfo = getDisponibilidadAgente(email);
  }
  
  return {
    user: {
      email: userInfo.email,
      nombre: userInfo.nombre,
      rol: userInfo.rol,
      area: userInfo.area,
      ubicacion: userInfo.ubicacion,
      puesto: userInfo.puesto
    },
    userInfo: {
      email: userInfo.email,
      nombre: userInfo.nombre,
      rol: userInfo.rol,
      area: userInfo.area,
      ubicacion: userInfo.ubicacion,
      puesto: userInfo.puesto,
      // Campos de disponibilidad
      disponible: disponibilidadInfo.disponible,
      motivoAusencia: disponibilidadInfo.motivo || '',
      fechaFinAusencia: disponibilidadInfo.fechaFin || '',
      // Campos de gerente
      esGerente: userInfo.esGerente || false,
      areaGerente: userInfo.areaGerente || null
    },
    areas: cats.areas,
    ubicaciones: getTodasLasUbicaciones(),
    categories: cats.categories,
    priorities: cats.priorities,
    statuses: cats.statuses,
    config
  };
}

/**
 * Ping para verificar conectividad
 */
function pingServer() {
  return { ok: true, ts: new Date() };
}


// ============================================================================
// CORRECCIONES PARA Client.gs - MÁQUINA DE ESTADOS Y FLUJO DE TICKETS
// ============================================================================
// INSTRUCCIONES: Agregar este código al final de Client.gs
// ============================================================================

// ============================================================================
// MÁQUINA DE ESTADOS - DEFINICIÓN DE FLUJO
// ============================================================================

/**
 * Configuración visual de estados para el flujo
 */
const FLOW_CONFIG = {
  statusColors: {
    'Nuevo': { bg: '#1e3a8a', text: '#fff', icon: 'bi-plus-circle' },
    'Abierto': { bg: '#0ea5e9', text: '#fff', icon: 'bi-folder2-open' },
    'En Proceso': { bg: '#2563eb', text: '#fff', icon: 'bi-gear-fill' },
    'En Cotización': { bg: '#8b5cf6', text: '#fff', icon: 'bi-currency-dollar' },
    'Visita Programada': { bg: '#0d9488', text: '#fff', icon: 'bi-calendar-event' },
    'Escalado': { bg: '#dc2626', text: '#fff', icon: 'bi-exclamation-triangle' },
    'Resuelto': { bg: '#059669', text: '#fff', icon: 'bi-check-circle' },
    'Cerrado': { bg: '#111827', text: '#fff', icon: 'bi-lock-fill' },
    'Reabierto': { bg: '#f59e0b', text: '#000', icon: 'bi-arrow-counterclockwise' }
  }
};

// ============================================================================
// CONFIGURACIÓN DE HORARIO LABORAL PARA SLA
// ============================================================================

const HORARIO_LABORAL = {
  horaInicio: 8,    // 8:00 AM
  horaFin: 18,      // 6:00 PM (18:00)
  horasPorDia: 10,  // 10 horas laborales por día
  diasLaborales: [1, 2, 3, 4, 5], // Lunes=1 a Viernes=5 (domingo=0, sábado=6)
  zonaHoraria: 'America/Mexico_City'
};

// Días festivos opcionales (formato 'MM-DD')
// Puedes agregar más según necesites
const DIAS_FESTIVOS = [
  '01-01', // Año Nuevo
  '02-05', // Constitución
  '03-21', // Benito Juárez
  '05-01', // Día del Trabajo
  '09-16', // Independencia
  '11-20', // Revolución
  '12-25', // Navidad
];

// ============================================================================
// FUNCIONES DE CÁLCULO DE HORAS LABORALES
// ============================================================================

/**
 * Verifica si una fecha es día laboral (lunes a viernes, no festivo)
 * @param {Date} fecha - Fecha a verificar
 * @returns {boolean} - true si es día laboral
 */
function esDiaLaboral(fecha) {
  if (!fecha || isNaN(fecha.getTime())) return false;
  
  const diaSemana = fecha.getDay(); // 0=domingo, 6=sábado
  
  // Verificar si es día de semana laboral
  if (!HORARIO_LABORAL.diasLaborales.includes(diaSemana)) {
    return false;
  }
  
  // Verificar si es día festivo
  const mes = String(fecha.getMonth() + 1).padStart(2, '0');
  const dia = String(fecha.getDate()).padStart(2, '0');
  const fechaStr = `${mes}-${dia}`;
  
  if (DIAS_FESTIVOS.includes(fechaStr)) {
    return false;
  }
  
  return true;
}

/**
 * Verifica si una hora está dentro del horario laboral
 * @param {number} hora - Hora del día (0-23)
 * @returns {boolean} - true si está en horario laboral
 */
function esHoraLaboral(hora) {
  return hora >= HORARIO_LABORAL.horaInicio && hora < HORARIO_LABORAL.horaFin;
}

/**
 * Obtiene las horas laborales restantes en un día desde una hora específica
 * @param {number} horaActual - Hora actual (0-23)
 * @returns {number} - Horas laborales restantes
 */
function horasLaboralesRestantesEnDia(horaActual) {
  if (horaActual >= HORARIO_LABORAL.horaFin) {
    return 0;
  }
  
  if (horaActual < HORARIO_LABORAL.horaInicio) {
    return HORARIO_LABORAL.horasPorDia;
  }
  
  return HORARIO_LABORAL.horaFin - horaActual;
}

/**
 * Avanza una fecha al siguiente momento laboral válido
 * @param {Date} fecha - Fecha a ajustar
 * @returns {Date} - Fecha ajustada al próximo momento laboral
 */
function ajustarAHoraLaboral(fecha) {
  const resultado = new Date(fecha);
  
  // Si es fin de semana o festivo, avanzar al siguiente día laboral
  while (!esDiaLaboral(resultado)) {
    resultado.setDate(resultado.getDate() + 1);
    resultado.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
  }
  
  // Si es antes del horario laboral, ajustar a inicio
  if (resultado.getHours() < HORARIO_LABORAL.horaInicio) {
    resultado.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
  }
  
  // Si es después del horario laboral, avanzar al siguiente día laboral
  if (resultado.getHours() >= HORARIO_LABORAL.horaFin) {
    resultado.setDate(resultado.getDate() + 1);
    resultado.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
    
    // Verificar que el nuevo día sea laboral
    while (!esDiaLaboral(resultado)) {
      resultado.setDate(resultado.getDate() + 1);
    }
  }
  
  return resultado;
}

/**
 * Suma horas laborales a una fecha, respetando horario y días laborales
 * @param {Date} fechaInicio - Fecha de inicio
 * @param {number} horasASumar - Horas laborales a sumar
 * @returns {Date} - Nueva fecha con las horas sumadas
 */
function sumarHorasLaborales(fechaInicio, horasASumar) {
  if (!horasASumar || horasASumar <= 0) {
    return fechaInicio;
  }
  
  let fecha = ajustarAHoraLaboral(new Date(fechaInicio));
  let horasRestantes = horasASumar;
  
  // Límite de seguridad para evitar bucles infinitos
  const maxIteraciones = 365;
  let iteraciones = 0;
  
  while (horasRestantes > 0 && iteraciones < maxIteraciones) {
    iteraciones++;
    
    // Asegurar que estamos en un día laboral
    fecha = ajustarAHoraLaboral(fecha);
    
    // Calcular horas disponibles en el día actual
    const horaActual = fecha.getHours() + (fecha.getMinutes() / 60);
    const horasDisponiblesHoy = HORARIO_LABORAL.horaFin - horaActual;
    
    if (horasRestantes <= horasDisponiblesHoy) {
      // Las horas caben en el día actual
      const minutosASumar = horasRestantes * 60;
      fecha.setMinutes(fecha.getMinutes() + minutosASumar);
      horasRestantes = 0;
    } else {
      // Consumir el resto del día y pasar al siguiente
      horasRestantes -= horasDisponiblesHoy;
      fecha.setDate(fecha.getDate() + 1);
      fecha.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
    }
  }
  
  return fecha;
}

/**
 * Calcula las horas laborales transcurridas entre dos fechas
 * @param {Date} fechaInicio - Fecha de inicio
 * @param {Date} fechaFin - Fecha de fin
 * @returns {number} - Horas laborales transcurridas
 */
function calcularHorasLaboralesTranscurridas(fechaInicio, fechaFin) {
  if (!fechaInicio || !fechaFin) return 0;
  
  const inicio = new Date(fechaInicio);
  const fin = new Date(fechaFin);
  
  if (fin <= inicio) return 0;
  
  let horasTotales = 0;
  let fechaActual = ajustarAHoraLaboral(new Date(inicio));
  
  // Límite de seguridad
  const maxIteraciones = 365;
  let iteraciones = 0;
  
  while (fechaActual < fin && iteraciones < maxIteraciones) {
    iteraciones++;
    
    if (!esDiaLaboral(fechaActual)) {
      fechaActual.setDate(fechaActual.getDate() + 1);
      fechaActual.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
      continue;
    }
    
    const horaActual = fechaActual.getHours();
    
    // Si estamos fuera del horario laboral, ajustar
    if (horaActual < HORARIO_LABORAL.horaInicio) {
      fechaActual.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
      continue;
    }
    
    if (horaActual >= HORARIO_LABORAL.horaFin) {
      fechaActual.setDate(fechaActual.getDate() + 1);
      fechaActual.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
      continue;
    }
    
    // Calcular horas en este día
    const finDelDiaLaboral = new Date(fechaActual);
    finDelDiaLaboral.setHours(HORARIO_LABORAL.horaFin, 0, 0, 0);
    
    const finEfectivo = fin < finDelDiaLaboral ? fin : finDelDiaLaboral;
    const horasEnEsteDia = (finEfectivo - fechaActual) / (1000 * 60 * 60);
    
    if (horasEnEsteDia > 0) {
      horasTotales += horasEnEsteDia;
    }
    
    // Avanzar al siguiente día
    fechaActual.setDate(fechaActual.getDate() + 1);
    fechaActual.setHours(HORARIO_LABORAL.horaInicio, 0, 0, 0);
  }
  
  return Math.max(0, horasTotales);
}


/**
 * Matriz de transiciones válidas entre estados
 * ACTUALIZADA: Sin estado "En Espera"
 */
const STATE_TRANSITIONS = {

  // Ticket nuevo creado por usuario
  'Nuevo': ['En Proceso', 'Escalado'],
  
  // Ticket abierto (por si se usa)
  'Abierto': ['En Proceso', 'Escalado'],
  
  // En proceso - el agente está trabajando
  'En Proceso': [
    'En Cotización',      // Requiere presupuesto
    'Visita Programada',  // Se programó visita
    'Resuelto',           // Problema solucionado
    'Escalado'            // Requiere nivel superior
  ],
  
  // Esperando aprobación de cotización
  // Dentro del mapa de transiciones válidas:
'En Cotización': ['En Proceso', 'Escalado', 'Cotización Rechazada'],
'Cotización Rechazada': ['En Cotización', 'En Proceso', 'Escalado', 'Cerrado'],
  
  // Visita programada para atención en sitio
  'Visita Programada': ['En Proceso', 'Resuelto', 'Escalado'],
  
  // Escalado a nivel superior
  'Escalado': ['En Proceso', 'Resuelto'],
  
  // Resuelto - esperando confirmación del usuario
  'Resuelto': ['Cerrado', 'En Proceso'],  // En Proceso si rechaza (no Reabierto)
  
  // Cerrado definitivamente
  'Cerrado': ['En Proceso'],  // Solo reabrir dentro de X días
  
  // Reabierto (legacy - redirigir a En Proceso)
  'Reabierto': ['En Proceso', 'Resuelto']
};

/**
 * Configuración de acciones por estado y rol
 * ACTUALIZADO: Sin pausa ni "En Espera" + Permiso a Usuarios para Escalar
 */
const STATE_ACTIONS = {

  // ==================== NUEVO ====================
  'Nuevo': {
    agente: [
      { id: 'atender', label: 'Atender Ticket', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'solicitar_escalar', label: 'Solicitar Escalamiento', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ],
    admin: [
      { id: 'atender', label: 'Atender Ticket', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'escalar', label: 'Escalar', icon: 'bi-arrow-up-circle', nextState: 'Escalado', btnClass: 'btn-danger', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [
      { id: 'solicitar_escalar', label: 'Escalar por falta de atención', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ]  
  },

  // ==================== ABIERTO ====================
  'Abierto': {
    agente: [
      { id: 'atender', label: 'Atender Ticket', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'solicitar_escalar', label: 'Solicitar Escalamiento', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ],
    admin: [
      { id: 'atender', label: 'Atender Ticket', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'escalar', label: 'Escalar', icon: 'bi-arrow-up-circle', nextState: 'Escalado', btnClass: 'btn-danger', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [
      { id: 'solicitar_escalar', label: 'Escalar a Gerencia', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ]
  },

  // ==================== EN PROCESO ====================
  'En Proceso': {
    agente: [
      { id: 'cotizacion', label: 'Solicitar Cotización', icon: 'bi-currency-dollar', nextState: 'En Cotización', btnClass: 'btn-info', requiresCotizacion: true },
      { id: 'programar', label: 'Programar Visita', icon: 'bi-calendar-event', nextState: 'Visita Programada', btnClass: 'btn-outline-secondary', requiresVisita: true },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true },
      { id: 'solicitar_escalar', label: 'Solicitar Escalamiento', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ],
    admin: [
      { id: 'cotizacion', label: 'Solicitar Cotización', icon: 'bi-currency-dollar', nextState: 'En Cotización', btnClass: 'btn-info', requiresCotizacion: true },
      { id: 'programar', label: 'Programar Visita', icon: 'bi-calendar-event', nextState: 'Visita Programada', btnClass: 'btn-outline-secondary', requiresVisita: true },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true },
      { id: 'escalar', label: 'Escalar', icon: 'bi-arrow-up-circle', nextState: 'Escalado', btnClass: 'btn-danger', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [
      { id: 'solicitar_escalar', label: 'Escalar a Gerencia', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ]
  },

  // ==================== EN COTIZACIÓN ====================
  'En Cotización': {
    agente: [
      { id: 'retomar', label: 'Retomar', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true }
    ],
    admin: [
      { id: 'retomar', label: 'Retomar', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true },
      { id: 'escalar', label: 'Escalar', icon: 'bi-arrow-up-circle', nextState: 'Escalado', btnClass: 'btn-danger', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [
      { id: 'solicitar_escalar', label: 'Escalar a Gerencia', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ]
  },

  // ==================== VISITA PROGRAMADA ====================
  'Visita Programada': {
    agente: [
      { id: 'iniciar', label: 'Iniciar Atención', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'reprogramar', label: 'Reprogramar Visita', icon: 'bi-calendar-plus', action: 'reprogramar_visita', btnClass: 'btn-outline-secondary', requiresVisita: true },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true }
    ],
    admin: [
      { id: 'iniciar', label: 'Iniciar Atención', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'reprogramar', label: 'Reprogramar Visita', icon: 'bi-calendar-plus', action: 'reprogramar_visita', btnClass: 'btn-outline-secondary', requiresVisita: true },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true },
      { id: 'escalar', label: 'Escalar', icon: 'bi-arrow-up-circle', nextState: 'Escalado', btnClass: 'btn-danger', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [
      { id: 'solicitar_escalar', label: 'Escalar a Gerencia', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ]
  },

  // ==================== ESCALADO ====================
  'Escalado': {
    agente: [
      { id: 'retomar', label: 'Retomar', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true }
    ],
    admin: [
      { id: 'retomar', label: 'Retomar', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [] // Ya está escalado
  },

  // ==================== RESUELTO ====================
  'Resuelto': {
    agente: [
      { id: 'reabrir', label: 'Reabrir', icon: 'bi-arrow-counterclockwise', action: 'reabrir_ticket', btnClass: 'btn-outline-warning' }
    ],
    admin: [
      { id: 'cerrar_forzado', label: 'Cerrar (forzado)', icon: 'bi-lock-fill', nextState: 'Cerrado', btnClass: 'btn-dark', requiresComment: true },
      { id: 'reabrir', label: 'Reabrir', icon: 'bi-arrow-counterclockwise', action: 'reabrir_ticket', btnClass: 'btn-outline-warning' }
    ],
    usuario: [
      { id: 'confirmar', label: '✓ Sí, está resuelto', icon: 'bi-hand-thumbs-up', action: 'confirmar_resolucion', btnClass: 'btn-success' },
      { id: 'rechazar', label: '✗ No quedó resuelto', icon: 'bi-hand-thumbs-down', action: 'rechazar_resolucion', btnClass: 'btn-danger' }
    ]
  },

  // ==================== CERRADO ====================
  'Cerrado': {
    agente: [],
    admin: [
      { id: 'reabrir', label: 'Reabrir', icon: 'bi-arrow-counterclockwise', action: 'reabrir_ticket', btnClass: 'btn-outline-warning' }
    ],
    usuario: [],
    special: {
      reabrir: { 
        id: 'reabrir', 
        label: 'Reabrir Ticket', 
        icon: 'bi-arrow-counterclockwise', 
        action: 'reabrir_ticket', 
        btnClass: 'btn-outline-warning', 
        maxDays: 7
      }
    }
  },

  // ==================== REABIERTO (legacy) ====================
  'Reabierto': {
    agente: [
      { id: 'atender', label: 'Atender', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true }
    ],
    admin: [
      { id: 'atender', label: 'Atender', icon: 'bi-play-fill', nextState: 'En Proceso', btnClass: 'btn-primary' },
      { id: 'resolver', label: 'Marcar Resuelto', icon: 'bi-check-lg', nextState: 'Resuelto', btnClass: 'btn-success', requiresComment: true },
      { id: 'reasignar', label: 'Reasignar', icon: 'bi-person-plus', action: 'reasignar', btnClass: 'btn-outline-primary' }
    ],
    usuario: [
      { id: 'solicitar_escalar', label: 'Escalar a Gerencia', icon: 'bi-arrow-up-circle', action: 'solicitar_escalar', btnClass: 'btn-outline-danger', requiresComment: true }
    ]
  }
};

/**
 * Obtiene las acciones disponibles combinando roles (Agente + Dueño)
 */
function getAvailableActions(ticketId, userEmail) {
  try {
    const user = getUser(userEmail);
    // Rol global del usuario en el sistema (ej. agente_mantenimiento)
    const globalRole = (user.rol || 'usuario').toLowerCase();
    const emailLower = userEmail.toLowerCase().trim();
    
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    
    const row = rows.find(r => String(r[m.ID]).trim() === String(ticketId).trim());
    if (!row) {
      return { ok: false, error: 'Ticket no encontrado', actions: [] };
    }
    
    const currentStatus = row[m.Estatus] || 'Nuevo';
    const ticketArea = (row[m['Área']] || '').toLowerCase();
    const reportaEmail = (row[m.ReportaEmail] || '').toLowerCase();
    const asignadoA = (row[m.AsignadoA] || '').toLowerCase();
    const fechaUltimaAct = row[m['ÚltimaActualización']];
    const presupuesto = row[m.Presupuesto] || '';
    const statusAutorizacion = row[m.StatusAutorizacion] || '';
    
    // --- 1. DETERMINAR PERMISOS ESPECÍFICOS ---
    const isOwner = reportaEmail === emailLower;
    const isAssigned = asignadoA === emailLower;
    
    // Verificar si es Gerente del Área del ticket
    const areaGerente = getGerenteArea(userEmail);
    const esGerenteDelArea = areaGerente && areaGerente.toLowerCase() === ticketArea;
    
    // Verificar si es Agente DEL ÁREA del ticket
    const esAgenteDelArea = (globalRole === 'agente_sistemas' && ticketArea === 'sistemas') ||
                            (globalRole === 'agente_mantenimiento' && ticketArea === 'mantenimiento');

    const isAdmin = globalRole === 'admin';

    // === NUEVO: REGLA PARA GERENCIA DE CLIENTES ===
    const isGerenteClientes = (areaGerente || '').toLowerCase() === 'clientes';
    const esTicketDeSupervisado = isGerenteClientes && getUsuariosSupervisados().some(u => reportaEmail.includes(u));

    // --- 2. DETERMINAR ROL EFECTIVO (Para permisos de edición/comentario) ---
    let effectiveRole = 'usuario';
    
    if (isAdmin) effectiveRole = 'admin';
    else if (esGerenteDelArea) effectiveRole = 'gerente';
    else if (isAssigned || esAgenteDelArea) effectiveRole = 'agente';
    else if (isOwner) effectiveRole = 'usuario';
    else if (esTicketDeSupervisado) effectiveRole = 'gerente_lectura'; // <-- ROL ESPECIAL
    else effectiveRole = 'viewer'; // Solo lectura

    // --- 3. CONSTRUIR ACCIONES (Lógica Aditiva) ---
    const stateConfig = STATE_ACTIONS[currentStatus] || {};
    let actions = [];
    
    // A) Acciones Jerárquicas (Admin/Gerente/Agente)
    if (effectiveRole === 'admin' || effectiveRole === 'gerente') {
      actions = [...(stateConfig.admin || stateConfig.agente || [])];
    } else if (effectiveRole === 'agente') {
      actions = [...(stateConfig.agente || [])];
    } else if (effectiveRole === 'gerente_lectura') {
      // IMPORTANTE: Gerencia de clientes NO TIENE BOTONES para cambiar el estado
      actions = []; 
    }
    
    // B) Acciones de Dueño (Usuario)
    if (isOwner) {
      const userActions = stateConfig.usuario || [];
      const existingIds = new Set(actions.map(a => a.id));
      userActions.forEach(act => {
        if (!existingIds.has(act.id)) {
          actions.push(act);
        }
      });
    }
    
    // C) Acciones Especiales (Reabrir ticket cerrado)
    if (stateConfig.special && currentStatus === 'Cerrado') {
      const reopenAction = stateConfig.special.reabrir;
      if (reopenAction && fechaUltimaAct) {
        const daysSinceClosed = (Date.now() - new Date(fechaUltimaAct).getTime()) / (1000 * 60 * 60 * 24);
        if (daysSinceClosed <= reopenAction.maxDays) {
          // Permitir reabrir si es dueño, admin, gerente del área o agente del área
          if (isOwner || isAdmin || esGerenteDelArea || esAgenteDelArea) {
             if (!actions.some(a => a.id === reopenAction.id)) {
               actions.push(reopenAction);
             }
          }
        }
      }
    }
    
    // --- 4. LÓGICA DE COTIZACIÓN ---
    const hasCotizacion = !!presupuesto && String(presupuesto).trim() !== '';
    const isInCotizacionStatus = currentStatus === 'En Cotización';
    
    // Puede ver cotización si es staff, dueño, o GERENTE DE CLIENTES (gerente_lectura)
    const canSeeCotizacion = (['agente', 'admin', 'gerente', 'gerente_lectura'].includes(effectiveRole) || isOwner) && 
                             (isInCotizacionStatus || hasCotizacion);
    
    const showCotizacionReadOnly = canSeeCotizacion && (isInCotizacionStatus || hasCotizacion);
    
    return {
      ok: true,
      currentStatus,
      actions, 
      canEdit: isAdmin || esGerenteDelArea, // Gerente Clientes NO puede editar campos del ticket
      canComment: true, // Todos los involucrados pueden comentar
      canSeeCotizacion,
      canEditCotizacion: false, 
      showCotizacionReadOnly,
      hasCotizacion,
      statusAutorizacion,
      isOwner,
      isAssigned,
      effectiveRole,
      esGerente: esGerenteDelArea,
      areaGerente: areaGerente || null
    };
  } catch (e) {
    Logger.log('Error en getAvailableActions: ' + e.message);
    return { ok: false, error: e.message, actions: [] };
  }
}

/**
 * Valida si una transición de estado es permitida
 * @param {string} currentState - Estado actual
 * @param {string} newState - Nuevo estado propuesto
 * @returns {Object} - { allowed: boolean, reason: string }
 */
function validateStateTransition(currentState, newState) {
  const allowedTransitions = STATE_TRANSITIONS[currentState] || [];
  
  if (!allowedTransitions.includes(newState)) {
    return {
      allowed: false,
      reason: `No se puede cambiar de "${currentState}" a "${newState}". Transiciones válidas: ${allowedTransitions.join(', ') || 'ninguna'}`
    };
  }
  
  return { allowed: true };
}


/**
 * Ejecuta un cambio de estado con validación completa, sincronización de visitas,
 * manejo de cotizaciones, y notificaciones nativas del sistema.
 */
function executeStateChange(ticketId, newStatus, comment, extraData, userEmail) {
  return withLock_(() => {
    try {
      console.log(`--- INICIO executeStateChange | Ticket: ${ticketId} | Nuevo: ${newStatus} ---`);
      
      const user = getUser(userEmail);
      const sh = getSheet(DB.TICKETS);
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);

      const idx = rows.findIndex(r => String(r[m.ID]).trim() === String(ticketId).trim());
      if (idx < 0) return { ok: false, error: 'Ticket no encontrado' };

      const row = rows[idx];
      const currentStatus = row[m.Estatus] || 'Nuevo';
      const folio = row[m.Folio];
      const titulo = row[m['Título']] || 'Sin título';
      
      // === RESOLUCIÓN DE EMAILS USANDO CONSTANTES CORRECTAS (DB.USERS) ===
      const resolverEmail = (valor) => {
        if (!valor) return '';
        const v = String(valor).trim();
        if (v.includes('@')) return v;
        
        // Usar tu función nativa que ya hace esto perfecto
        if (typeof getEmailNotificacion === 'function') {
          const emailSys = getEmailNotificacion(v);
          if (emailSys && emailSys.includes('@')) return emailSys;
        }

        // Búsqueda de respaldo en DB.USERS
        try {
          const { headers: uH, rows: uR } = _readTableByHeader_(DB.USERS);
          const um = _headerMap_(uH);
          const uRow = uR.find(r => String(r[0]).toLowerCase() === v.toLowerCase() || String(r[um.Email] || '').toLowerCase() === v.toLowerCase());
          if (uRow && uRow[um.Email] && String(uRow[um.Email]).includes('@')) {
            return String(uRow[um.Email]).trim();
          }
        } catch(e) { console.error("Error búsqueda DB.USERS: " + e.message); }

        return v; 
      };

      const agenteEmail = resolverEmail(row[m.AsignadoA]);
      const reportaEmail = resolverEmail(row[m.ReportaEmail]);
      
      console.log(`Contexto Validado: Folio ${folio} | Agente: ${agenteEmail} | Usuario: ${reportaEmail}`);

      // 1. Idempotencia
      if (currentStatus === newStatus) return { ok: true };

      // 2. Validar transición y permisos
      const validation = validateStateTransition(currentStatus, newStatus);
      if (!validation.allowed) return { ok: false, error: validation.reason };

      // =====================================================================
      // 3. SINCRONIZACIÓN AUTOMÁTICA DE VISITA
      // =====================================================================
      if (currentStatus === 'Visita Programada' && newStatus !== 'Visita Programada') {
        registrarEnHojaVisitas_({
          ticketId: ticketId, folio: folio, agente: userEmail, accion: 'Atendida',
          fechaVisita: row[m.FechaVisita] || '', horaVisita: row[m.HoraVisita] || '',
          notas: `Visita finalizada por inicio de atención (${newStatus}). ` + (comment || '')
        });
        console.log("✅ Visita sincronizada automáticamente.");
      }

      // =====================================================================
      // 4. MANEJO DE COTIZACIONES
      // =====================================================================
      if (newStatus === 'En Cotización' && extraData && extraData.aprobadorEmail && extraData.presupuesto) {
        row[m.Presupuesto] = extraData.presupuesto;
        row[m.AprobadorEmail] = extraData.aprobadorEmail;
        row[m.TipoProveedor] = extraData.tipoProveedor || 'Externo';
        row[m.TiempoCotizacion] = extraData.tiempoCotizacion || '';
        row[m.StatusAutorizacion] = 'Pendiente';
        
        if (typeof notifyApproval === 'function') {
           notifyApproval(ticketId, extraData.aprobadorEmail, extraData.presupuesto);
        }
      }

      // 5. Actualizar estado y metadatos del Ticket
      row[m.Estatus] = newStatus;
      row[m['ÚltimaActualización']] = new Date();
      if (m.ActualizadoPor != null) row[m.ActualizadoPor] = userEmail;

      const ESTATUS_CIERRE = ['resuelto', 'cerrado', 'completado', 'cancelado'];
      if (ESTATUS_CIERRE.includes(newStatus.toLowerCase())) {
        if (m['FechaResuelto'] != null) row[m['FechaResuelto']] = new Date();
      }

      // 6. Persistir cambios en la hoja Tickets
      sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
      SpreadsheetApp.flush();
      clearCache(DB.TICKETS);

      // 7. Procesar Adjunto (Soporte para Botón Azul)
      let fileUrl = '';
      let fileName = '';
      const fileId = (extraData && extraData.fileId) ? extraData.fileId : '';

      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          fileUrl = file.getUrl();
          fileName = file.getName();
          if (agenteEmail && agenteEmail.includes('@')) try { file.addViewer(agenteEmail); } catch(e){}
          if (reportaEmail && reportaEmail.includes('@')) try { file.addViewer(reportaEmail); } catch(e){}
        } catch (e) { console.error('Error adjunto: ' + e.message); }
      }

      // 8. Registro en COMMENTS (Activación botón azul)
      const shComments = getSheet(DB.COMMENTS);
      const esRechazo = (currentStatus.toLowerCase() === 'resuelto' && newStatus.toLowerCase() === 'en proceso');
      
      const textoComentario = esRechazo 
        ? `❌ <b>Resolución RECHAZADA:</b>\n${comment || ''}`
        : `Estado cambiado a "${newStatus}":\n${comment || ''}`;

      shComments.appendRow([
        genId(), ticketId, new Date(), 'Sistema', 'Sistema', 
        textoComentario, false, fileId, fileUrl, fileName
      ]);

      registrarBitacora(ticketId, 'Cambio de estatus', `${currentStatus} → ${newStatus}`);
      clearCache(DB.COMMENTS);

      // =====================================================================
      // 9. BLOQUE DE NOTIFICACIONES (Aprovechando tus funciones nativas)
      // =====================================================================
      const appUrl = (typeof getScriptUrl === 'function') ? getScriptUrl(ticketId) : ScriptApp.getService().getUrl() + '?ticket=' + ticketId;

      // A) CASO: RECHAZO (Al Agente)
      if (esRechazo && agenteEmail && agenteEmail.includes('@')) {
        console.log("Enviando notificaciones de rechazo...");
        
        // 1. Telegram
        telegramSendToGrupo(getTelegramChatIdGrupo(agenteEmail), `❌ <b>Resolución RECHAZADA - Ticket #${folio}</b>\n<b>Motivo:</b> ${comment}`, 'HTML', fileId);

        // 2. Email (usando tu plantilla profesional)
        try {
          const bodyHtml = `<h2>Resolución Rechazada</h2>
                            <p>El ticket <strong>#${folio}</strong> fue reabierto.</p>
                            <div class="alert-box alert-danger"><strong>Motivo:</strong> ${comment}</div>
                            <p><a href="${appUrl}" class="btn btn-danger">Ver Ticket</a></p>`;
          enviarEmailNotificacion(agenteEmail, `❌ Rechazo: Ticket #${folio}`, bodyHtml);
          console.log(`✅ Email rechazo enviado a ${agenteEmail}`);
        } catch(e) { console.error("Error email rechazo: " + e.message); }

        // 3. Notificación App
        crearNotificacion(agenteEmail, 'rechazo', 'Resolución Rechazada', `Ticket #${folio} reabierto por el usuario`, ticketId);
      }

      // B) CASO: RESOLUCIÓN (Al Usuario)
      if (newStatus === 'Resuelto' && reportaEmail && reportaEmail.includes('@')) {
        console.log("Enviando notificaciones de resolución...");
        
        // 1. Notificación App
        crearNotificacion(reportaEmail, 'resuelto', 'Ticket Resuelto', `Tu ticket #${folio} ha sido resuelto.`, ticketId);
        
        // 2. Email (usando tu función nativa notificarUsuarioCambioEstado)
        try {
          if (typeof notificarUsuarioCambioEstado === 'function') {
            notificarUsuarioCambioEstado(reportaEmail, {folio: folio, titulo: titulo}, 'Resuelto', comment);
            console.log(`✅ Email resolución nativo enviado a ${reportaEmail}`);
          }
        } catch(e) { console.error("Error notif resolución nativa: " + e.message); }

        // 3. Telegram (si el usuario tiene Telegram)
        try {
           const keyTgUser = 'tg_' + reportaEmail.split('@')[0].toLowerCase();
           const chatIdUser = getConfig(keyTgUser);
           if (chatIdUser) {
             telegramSend(`✅ <b>Ticket #${folio} Resuelto</b>\nTu ticket ha sido marcado como resuelto.\n🔗 ${appUrl}`, chatIdUser);
           }
        } catch(e) { console.error("Error Telegram usuario: " + e.message); }
      }

      // C) CASO: COTIZACIÓN RECHAZADA (nuevo estatus dedicado)
if (newStatus === 'Cotización Rechazada') {
  // El tiempo sigue corriendo — NO se toca el Vencimiento
  // Limpiar datos de cotización pendiente
  if (m.StatusAutorizacion != null) row[m.StatusAutorizacion] = 'Rechazada';
  sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]); // re-guardar con StatusAutorizacion
  SpreadsheetApp.flush();

  // Notificar al AGENTE: debe replantear o escalar
  if (agenteEmail && agenteEmail.includes('@')) {
    const msgAgente = `⚠️ <b>Cotización RECHAZADA - Ticket #${folio}</b>\n\n` +
                      `📝 <b>${titulo}</b>\n` +
                      `❌ <b>Motivo:</b> ${comment || 'Sin comentario'}\n\n` +
                      `⏰ El SLA sigue corriendo. Por favor replantea la cotización o escala el ticket.`;
    try {
      const chatAgente = getTelegramChatIdGrupo(agenteEmail);
      if (chatAgente) telegramSendToGrupo(chatAgente, msgAgente);
      crearNotificacion(agenteEmail, 'alerta', `Cotización rechazada: #${folio}`,
        'Debes replantear la cotización. El SLA sigue corriendo.', ticketId);
      const bodyEmailAgente = `<h2 style="color:#dc2626;">⚠️ Cotización Rechazada</h2>
        <p>La cotización del ticket <strong>#${folio}</strong> fue rechazada.</p>
        <div style="background:#fef2f2;border-left:4px solid #dc2626;padding:15px;margin:15px 0;">
          <p style="margin:0;"><strong>Motivo:</strong> ${comment || 'No especificado'}</p>
        </div>
        <p style="color:#dc2626;"><strong>⏰ El SLA sigue corriendo.</strong> Revisa y replantea la cotización a la brevedad.</p>`;
      enviarEmailNotificacion(agenteEmail, `⚠️ Cotización Rechazada - Ticket #${folio}`, bodyEmailAgente);
    } catch(e) { Logger.log('Error notif rechazo cotización agente: ' + e.message); }
  }

  // Notificar al USUARIO que reportó: su solicitud requiere una nueva propuesta
  if (reportaEmail && reportaEmail.includes('@')) {
    crearNotificacion(reportaEmail, 'info', `Cotización en revisión: #${folio}`,
      'La cotización fue rechazada. El equipo preparará una nueva propuesta.', ticketId);
    try {
      const bodyEmailUser = `<h2 style="color:#f59e0b;">📋 Actualización en tu Ticket</h2>
        <p>La cotización enviada para tu ticket <strong>#${folio}</strong> no fue aprobada en esta ocasión.</p>
        <div style="background:#fffbeb;border-left:4px solid #f59e0b;padding:15px;margin:15px 0;">
          <p style="margin:0;"><strong>Estado:</strong> Cotización rechazada — en revisión</p>
          <p style="margin:5px 0 0;"><strong>Siguiente paso:</strong> El equipo técnico preparará una nueva propuesta.</p>
        </div>
        <p>Te notificaremos cuando haya una nueva cotización disponible.</p>`;
      enviarEmailNotificacion(reportaEmail, `📋 Actualización Ticket #${folio} - Cotización en revisión`, bodyEmailUser);
    } catch(e) { Logger.log('Error notif rechazo cotización usuario: ' + e.message); }
  }
}

      console.log(`--- FIN executeStateChange EXITOSO ---`);
      return { ok: true, newStatus, fileUrl };

    } catch (e) {
      console.error('Error crítico: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

/**
 * Obtener información completa del ticket incluyendo acciones disponibles
 * REEMPLAZA o complementa la función getTicket existente
 */
function getTicketWithActions(ticketId, userEmail) {
  try {
    const ticketData = getTicket(ticketId);
    
    if (!ticketData || !ticketData.ticket) {
      return { ok: false, error: 'Ticket no encontrado' };
    }
    
    const actionsInfo = getAvailableActions(ticketId, userEmail);
    
    return {
      ok: true,
      ticket: ticketData.ticket,
      comments: ticketData.comments || [],
      actions: actionsInfo.actions || [],
      canEdit: actionsInfo.canEdit || false,
      canSeeCotizacion: actionsInfo.canSeeCotizacion || false,
      canEditCotizacion: actionsInfo.canEditCotizacion || false,
      showCotizacionReadOnly: actionsInfo.showCotizacionReadOnly || false,
      hasCotizacion: actionsInfo.hasCotizacion || false,
      statusAutorizacion: actionsInfo.statusAutorizacion || '',
      isOwner: actionsInfo.isOwner || false,
      isAssigned: actionsInfo.isAssigned || false,
      effectiveRole: actionsInfo.effectiveRole || 'usuario',
      currentStatus: actionsInfo.currentStatus || ticketData.ticket['Estatus']
    };
    
  } catch (e) {
    Logger.log('Error en getTicketWithActions: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================================
// CORRECCIÓN: Auto-cierre de tickets resueltos sin respuesta en 48h
// ============================================================================

/**
 * Cerrar automáticamente tickets resueltos sin respuesta en 48h
 * Ejecutar con trigger cada hora
 */
function autocerrarResueltos48h() {
  try {
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    const sh = getSheet(DB.TICKETS);
    
    const ahora = new Date();
    let cerrados = 0;
    
    rows.forEach((row, idx) => {
      const estatus = row[m.Estatus];
      if (estatus !== 'Resuelto') return;
      
      const ultimaAct = new Date(row[m['ÚltimaActualización']]);
      if (isNaN(ultimaAct.getTime())) return;
      
      const horasTranscurridas = (ahora - ultimaAct) / (1000 * 60 * 60);
      
      if (horasTranscurridas >= 48) {
        row[m.Estatus] = 'Cerrado';
        row[m['ÚltimaActualización']] = ahora;
        
        sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
        
        const ticketId = row[m.ID];
        const folio = row[m.Folio];
        const reportaEmail = row[m.ReportaEmail];
        
        registrarBitacora(ticketId, 'Auto-cierre', 'Cerrado automáticamente tras 48h sin respuesta del usuario');
        addSystemComment(ticketId, 'Ticket cerrado automáticamente por falta de respuesta tras 48 horas.', true);
        
        notifyUser(reportaEmail, 'closed', 'Ticket cerrado automáticamente',
          `Tu ticket #${folio} fue cerrado automáticamente al no recibir confirmación en 48 horas.`,
          { ticketId, folio });
        
        cerrados++;
      }
    });
    
    if (cerrados > 0) {
      clearCache(DB.TICKETS);
      Logger.log(`Auto-cierre: ${cerrados} tickets cerrados`);
    }
    
    return { ok: true, cerrados };
    
  } catch (e) {
    Logger.log('Error en autocerrarResueltos48h: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================================
// VALIDACIONES Y CORRECCIONES DE SEGURIDAD
// ============================================================================

/**
 * Validar que el usuario tiene permiso para ver un ticket
 */
function canUserAccessTicket(ticketId, userEmail) {
  try {
    const user = getUser(userEmail);
    const role = (user.rol || 'usuario').toLowerCase();
    const emailLower = userEmail.toLowerCase().trim();
    
    // Admin puede ver todo
    if (role === 'admin') return true;
    
    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);
    
    const row = rows.find(r => String(r[m.ID]).trim() === String(ticketId).trim());
    if (!row) return false;
    
    const ticketArea = (row[m['Área']] || '').toLowerCase();
    const reportaEmail = (row[m.ReportaEmail] || '').toLowerCase();
    const asignadoA = (row[m.AsignadoA] || '').toLowerCase();
    
    // Usuario puede ver sus propios tickets
    if (reportaEmail === emailLower) return true;
    
    // Agente puede ver tickets de su área
    if (role === 'agente_sistemas' && ticketArea === 'sistemas') return true;
    if (role === 'agente_mantenimiento' && ticketArea === 'mantenimiento') return true;
    
    // Agente asignado puede ver el ticket
    if (asignadoA === emailLower) return true;
    
    // Gerente del área puede ver los tickets de su área
    const areaGerente = getGerenteArea(userEmail);
    if (areaGerente && areaGerente.toLowerCase() === ticketArea) return true;

    // === NUEVO: Gerente de Clientes puede ver tickets de sus supervisados ===
    if ((areaGerente || '').toLowerCase() === 'clientes') {
      const supervisados = getUsuariosSupervisados();
      if (supervisados.some(u => reportaEmail.includes(u))) return true;
    }
    
    return false;
    
  } catch (e) {
    Logger.log('Error en canUserAccessTicket: ' + e.message);
    return false;
  }
}

/**
 * Wrapper seguro para getTicketWithActions que valida acceso
 */
function getTicketWithActionsSafe(ticketId, userEmail) {
  if (!canUserAccessTicket(ticketId, userEmail)) {
    return { ok: false, error: 'No tienes permiso para ver este ticket' };
  }
  return getTicketWithActions(ticketId, userEmail);
}

function programarVisitaTicket(ticketId, fechaVisita, horaVisita, notasVisita, userEmail) {
  return withLock_(() => {
    try {
      const sh = getSheet(DB.TICKETS);
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      
      const idx = rows.findIndex(r => String(r[m.ID]).trim() === String(ticketId).trim());
      if (idx < 0) {
        return { ok: false, error: 'Ticket no encontrado' };
      }
      
      const row = rows[idx];
      const folio = row[m.Folio];
      const reportaEmail = row[m.ReportaEmail];

      // === NUEVO CANDADO DE HORARIO LABORAL ===
      const validacion = validarFechaHoraLaboral(fechaVisita, horaVisita);
      if (!validacion.valido) {
        return { ok: false, error: validacion.error };
      }
      // ========================================
      
      // Validar que la fecha sea futura
      const fechaHoraVisita = new Date(`${fechaVisita}T${horaVisita}:00`);
      if (isNaN(fechaHoraVisita.getTime())) {
        return { ok: false, error: 'Fecha/hora de visita inválida' };
      }
      
      if (fechaHoraVisita < new Date()) {
        return { ok: false, error: 'La fecha/hora de visita debe ser futura' };
      }
      
      // Actualizar campos de visita en TICKETS
      if (m.FechaVisita != null) row[m.FechaVisita] = fechaVisita;
      if (m.HoraVisita != null) row[m.HoraVisita] = "'" + horaVisita;
      if (m.NotasVisita != null) row[m.NotasVisita] = notasVisita || '';
      
      // Cambiar estado a "Visita Programada"
      row[m.Estatus] = 'Visita Programada';
      row[m['ÚltimaActualización']] = new Date();
      
      // Persistir cambios en TICKETS
      sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
      clearCache(DB.TICKETS);
      
      // Formatear fecha para mostrar
      const opciones = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
      const fechaFormateada = fechaHoraVisita.toLocaleDateString('es-MX', opciones);
      
      // Registrar en Bitácora
      registrarBitacora(ticketId, 'Visita programada', 
        `Fecha: ${fechaFormateada} a las ${horaVisita}. ${notasVisita ? 'Notas: ' + notasVisita : ''}`);
      
      // Agregar comentario
      addSystemComment(ticketId, 
        `📅 VISITA PROGRAMADA\nFecha: ${fechaFormateada}\nHora: ${horaVisita}\n${notasVisita ? 'Notas: ' + notasVisita : ''}\nProgramada por: ${userEmail}`, 
        false);

      // =========================================================
      // NUEVO: GUARDAR EN LA HOJA DE VISITAS PROGRAMADAS
      // =========================================================
      registrarEnHojaVisitas_({
        ticketId: ticketId,
        folio: folio,
        agente: userEmail,
        accion: 'Programada',
        fechaVisita: fechaVisita,
        horaVisita: horaVisita,
        notas: notasVisita || ''
      });
      // =========================================================
      
      // Notificaciones (Email / Telegram)
    
      if (reportaEmail) {
        notifyUser(reportaEmail, 'visita_programada', 'Visita programada',
          `Se ha programado una visita para atender tu ticket #${folio} el ${fechaFormateada} a las ${horaVisita}.`,
          { ticketId, folio, fecha: fechaVisita, hora: horaVisita });
          
        try {
          const htmlVisita = `
            <h2>📅 Visita Programada</h2>
            <p>Se ha agendado una visita en sitio para atender tu solicitud.</p>
            <div class="ticket-box">
              <h3>Ticket #${folio}</h3>
            </div>
            <div class="info-row"><span class="info-label">Fecha programada:</span><span class="info-value" style="color:#0d9488;">${fechaFormateada}</span></div>
            <div class="info-row"><span class="info-label">Hora aproximada:</span><span class="info-value" style="color:#0d9488;">${horaVisita}</span></div>
            ${notasVisita ? `<div class="alert-box alert-info"><p style="margin:0;"><strong>Notas del agente:</strong> ${notasVisita}</p></div>` : ''}
            <p style="color: #6b7280; font-size: 0.9em; margin-top: 15px;">Por favor asegúrate de estar disponible en el lugar y horario indicado.</p>
            <p style="margin-top:25px;"><a href="${getScriptUrl()}" class="btn btn-primary" style="background-color:#0d9488; border-color:#0d9488;">🔗 Ver Ticket</a></p>
          `;
          enviarEmailNotificacion(reportaEmail, `📅 Visita programada - Ticket #${folio}`, htmlVisita);
        } catch (e) { Logger.log('Error email visita: ' + e.message); }
      }
      
      const cfg = getConfig();
      if (cfg.telegram_chat_id && cfg.telegram_token) {
        telegramSend(`📅 <b>Visita Programada</b>\nTicket #${folio}\nFecha: ${fechaFormateada}\nHora: ${horaVisita}`, cfg.telegram_chat_id);
      }
      
      return { ok: true, fechaVisita, horaVisita };
      
    } catch (e) {
      Logger.log('Error en programarVisitaTicket: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

function getVisitasProgramadas(emailSolicitante, modo, areaFiltro) {
  try {
    const sheet = getSheet(DB.TICKETS);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const hdr = data[0];
    const idx = _headerMap_(hdr);

    // Búsqueda flexible de columnas
    const colEstatus = idx['Estatus'] !== undefined ? idx['Estatus'] : idx['Estado'];
    const colArea = idx['Área'] !== undefined ? idx['Área'] : idx['Area'];
    const colFecha = idx['FechaVisita'] !== undefined ? idx['FechaVisita'] : idx['Fecha Visita'];
    const colHora = idx['HoraVisita'] !== undefined ? idx['HoraVisita'] : idx['Hora Visita'];
    const colAsignado = idx['AsignadoA'] !== undefined ? idx['AsignadoA'] : idx['Asignado A'];

    const f = s => String(s || '').trim().toLowerCase();
    const emailNorm = f(emailSolicitante);
    let areaNorm = f(areaFiltro);
    
    // ==========================================================
    // FIX 1: SUPERADMIN VE TODO (Ignora el filtro de área)
    // ==========================================================
    const superAdmins = ['rgnava', 'admin', 'rcesquivel'];
    const esSuperAdmin = superAdmins.some(sa => emailNorm.includes(sa));
    if (esSuperAdmin) {
      areaNorm = ''; // Destruimos el filtro, queremos ver todo
    }

    const tz = Session.getScriptTimeZone();
    const ahora = new Date();
    const visitas = [];

    // Empezamos desde i=1 para saltar encabezados
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const estatusNorm = f(r[colEstatus]);
      
      // 1. Filtro de Estatus exacto
      if (estatusNorm !== 'visita programada') continue;

      // 2. Filtro por Rol/Área
      const areaTicket = f(r[colArea]);
      const asignadoTicket = f(r[colAsignado]);

      if (modo === 'agente') {
        if (asignadoTicket !== emailNorm) continue;
      } else if (modo === 'gerente' || modo === 'admin') {
        // Solo aplica si areaNorm tiene algo (si es superadmin, está vacío y lo ignora)
        if (areaNorm && areaNorm !== 'todas' && areaTicket !== areaNorm) continue;
      }

      // 3. Parseo de Fecha
      let fv = "";
      const fvRaw = r[colFecha];
      
      if (fvRaw instanceof Date) {
        fv = Utilities.formatDate(fvRaw, tz, "yyyy-MM-dd");
      } else if (fvRaw) {
        let strDate = String(fvRaw).split('T')[0].trim();
        if (strDate.includes('/')) {
          const partes = strDate.split('/');
          if (partes.length === 3) {
            if (partes[2].length === 4) {
               fv = `${partes[2]}-${partes[1].padStart(2,'0')}-${partes[0].padStart(2,'0')}`;
            } else {
               fv = strDate; 
            }
          }
        } else {
          fv = strDate;
        }
      }

      // 4. Parseo de Hora
      let hv = "";
      const hvRaw = r[colHora];
      if (hvRaw instanceof Date) {
        hv = Utilities.formatDate(hvRaw, tz, "HH:mm");
      } else {
        hv = String(hvRaw || "").replace(/'/g, "").trim();
        const match = hv.match(/(\d{1,2}:\d{2})/);
        hv = match ? match[1] : "08:00"; 
      }

      // ==========================================================
      // FIX 2: NO DESCARTAR TICKETS SIN FECHA (Mostrarlos como error)
      // ==========================================================
      let tiempoEstado = 'futura';
      let fechaFormateada = '';

      if (!fv) {
        // Si no hay fecha (Como el ticket #152)
        tiempoEstado = 'vencida'; // Lo forzamos a 'vencida' para que salga ROJO en pantalla
        fechaFormateada = '⚠️ SIN FECHA ASIGNADA';
        hv = '--:--';
      } else {
        // Lógica normal si sí hay fecha
        const fechaVisitaDt = new Date(fv + 'T' + hv + ':00');
        if (!isNaN(fechaVisitaDt.getTime())) {
          const diffHrs = (fechaVisitaDt - ahora) / 36e5;
          if (diffHrs < 0) tiempoEstado = 'vencida';
          else if (diffHrs <= 4) tiempoEstado = 'proxima';
          else if (diffHrs <= 24) tiempoEstado = 'hoy_manana';
          
          fechaFormateada = fechaVisitaDt.toLocaleDateString('es-MX', { weekday: 'long', day: 'numeric', month: 'short' });
        } else {
          fechaFormateada = fv; // Fallback
        }
      }

      visitas.push({
        ticketId: r[idx['ID']],
        folio: r[idx['Folio']],
        titulo: r[idx['Título']] || "Sin título",
        ubicacion: r[idx['Ubicación']] || "N/A",
        area: r[colArea],
        asignadoA: r[colAsignado],
        fechaVisita: fv,
        horaVisita: hv,
        tiempoEstado: tiempoEstado,
        fechaFormateada: fechaFormateada
      });
    }
    
    const orden = { 'vencida': 0, 'proxima': 1, 'hoy_manana': 2, 'futura': 3 };
    return visitas.sort((a, b) => orden[a.tiempoEstado] - orden[b.tiempoEstado]);

  } catch (e) {
    console.error('Error en getVisitasProgramadas:\n' + e.message);
    return [];
  }
}

function reprogramarVisitaTicket(ticketId, nuevaFecha, nuevaHora, motivo, notas, userEmail) {
  return withLock_(() => {
    try {
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      const sh = getSheet(DB.TICKETS);
      
      const idx = rows.findIndex(r => String(r[m.ID]).trim() === String(ticketId).trim());
      if (idx < 0) return { ok: false, error: 'Ticket no encontrado' };
      
      const row = rows[idx];
      const folio = row[m.Folio];
      const reportaEmail = row[m.ReportaEmail];
      const titulo = row[m['Título']] || '';
      const asignadoA = row[m.AsignadoA];
      
      // 1. Guardar en la hoja de auditoría de visitas
      registrarEnHojaVisitas_({
        ticketId: ticketId,
        folio: folio,
        agente: userEmail,
        accion: 'Reprogramada',
        fechaVisita: nuevaFecha,
        horaVisita: nuevaHora, // Pasamos la hora limpia a la bitácora
        notas: `Motivo: ${motivo}. ${notas || ''}`
      });

      // 2. Actualizar datos en la tabla principal 
      // EL APÓSTROFE (') EVITA EL BUG DEL 1899 (16:36)
      row[m.FechaVisita] = nuevaFecha;
      row[m.HoraVisita] = "'" + nuevaHora; 
      if (m.NotasVisita != null) row[m.NotasVisita] = notas || '';
      row[m['ÚltimaActualización']] = new Date();
      
      sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
      clearCache(DB.TICKETS);
      
      // 3. Registrar en bitácora y comentarios
      registrarBitacora(ticketId, 'Visita Reprogramada', `Nueva fecha: ${nuevaFecha} a las ${nuevaHora}. Motivo: ${motivo}`);
      addSystemComment(ticketId, `🔄 VISITA REPROGRAMADA\nNueva Fecha: ${nuevaFecha}\nHora: ${nuevaHora}\nMotivo: ${motivo}\nNotas: ${notas || 'Ninguna'}`, false);

      // 4. Notificar OMNICANAL
      if (asignadoA) {
        notificarAgenteTodosCanales(asignadoA, 'reprog_aprobada', {
          ticketId: ticketId,
          folio: folio,
          titulo: titulo,
          fecha: nuevaFecha,
          hora: nuevaHora,
          tituloNotif: '📅 Visita Reprogramada: #' + folio,
          mensajeNotif: `Visita reprogramada para el ${nuevaFecha} a las ${nuevaHora}`,
          htmlCuerpo: `<h2>🔄 Visita Reprogramada</h2><p>La visita para el ticket #${folio} ha sido actualizada.</p><p><strong>Nueva Fecha:</strong> ${nuevaFecha}</p><p><strong>Nueva Hora:</strong> ${nuevaHora}</p><p><strong>Motivo:</strong> ${motivo}</p>`
        });
      }

      return { ok: true, message: 'Visita reprogramada correctamente' };
    } catch (e) {
      Logger.log('Error reprogramando: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

function registrarEnHojaVisitas_(datos) {
  try {
    const NOMBRE_HOJA_VISITAS = 'VisitasProgramadas';
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(NOMBRE_HOJA_VISITAS);

    if (!sh) {
      sh = ss.insertSheet(NOMBRE_HOJA_VISITAS);
      sh.appendRow(['ID', 'FechaRegistro', 'TicketID', 'Folio', 'Agente', 'Accion', 'FechaVisita', 'HoraVisita', 'Notas']);
    }

    sh.appendRow([
      Utilities.getUuid(),
      new Date(),
      datos.ticketId,
      datos.folio,
      datos.agente,
      datos.accion,
      datos.fechaVisita,
      datos.horaVisita,
      datos.notas
    ]);
    
    // IMPORTANTE: Forzar el guardado y limpiar el caché de esta hoja específica
    SpreadsheetApp.flush(); 
    clearCache(NOMBRE_HOJA_VISITAS); 
    
    Logger.log(`✅ Visita registrada y caché limpio para ticket ${datos.folio}`);
  } catch (e) {
    console.error('Error guardando en VisitasProgramadas: ' + e.message);
  }
}


// =====================================================================
// FUNCIONES PARA REPORTES BI V2 - AGREGAR A Client.gs
// =====================================================================

function getBIData_V17(periodo, fInicioStr, fFinStr) {
  console.log(`--- INICIO BI V18 --- Periodo: ${periodo}`);

  try {
    const hdr  = HEADERS.Tickets;
    const data = getCachedData(DB.TICKETS);
    if (!data || !data.length) {
      return JSON.stringify({
        ok: true,
        dataset: [],
        resumen: { total: 0, activos: 0, cerrados: 0, vencidos: 0, aTiempo: 0, tarde: 0 },
        ticketsTarde: []
      });
    }

    const idx = {};
    hdr.forEach((h, i) => idx[h] = i);

    // ── Detectar índices reales desde el Sheet ──────────────────────────────
    let idxFechaRes    = -1;
    let idxFechaCierre = -1;
    let idxNotas       = -1;

    try {
      const sh           = getSheet(DB.TICKETS);
      const sheetHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

      idxFechaRes    = sheetHeaders.findIndex(h => String(h).trim() === 'FechaResolucion');
      idxFechaCierre = sheetHeaders.findIndex(h => String(h).trim() === 'FechaCierre');
      idxNotas = sheetHeaders.findIndex(h => {
  const colName = String(h).trim().toLowerCase();
  // Busca cualquiera de estas coincidencias en el nombre de la columna
  return colName === 'cantidadnotas' || 
         colName === 'notas' || 
         colName === 'comentarios' || 
         colName === 'cantidad de notas' || 
         colName === 'cantidad notas';
});

      Logger.log(`[BI] FechaResolucion: ${idxFechaRes} | FechaCierre: ${idxFechaCierre} | CantidadNotas: ${idxNotas}`);

      if (idxFechaRes === -1)
        Logger.log('⚠️ [BI] Columna FechaResolucion no encontrada.');
      if (idxFechaCierre === -1)
        Logger.log('⚠️ [BI] Columna FechaCierre no encontrada. Rendimiento de agentes usará FechaResolucion como fallback.');
      if (idxNotas === -1)
        Logger.log('⚠️ [BI] Columna CantidadNotas no encontrada. Tab Notas mostrará vacío.');

    } catch (eH) {
      Logger.log('⚠️ [BI] Error leyendo headers del Sheet: ' + eH.message);
    }
    // ────────────────────────────────────────────────────────────────────────

    // --- RANGO DE FECHAS ---
    const hoy = new Date();
    let fechaMin = new Date(hoy);
    let fechaMax = new Date(hoy);
    fechaMax.setHours(23, 59, 59, 999);
    fechaMin.setHours(0, 0, 0, 0);

    if (periodo === 'custom' && fInicioStr && fFinStr) {
      const pI = fInicioStr.split('-');
      const pF = fFinStr.split('-');
      fechaMin = new Date(pI[0], pI[1] - 1, pI[2], 0,  0,  0);
      fechaMax = new Date(pF[0], pF[1] - 1, pF[2], 23, 59, 59);
    } else if (periodo === 'semana') {
      fechaMin.setDate(hoy.getDate() - 7);
    } else if (periodo === 'mes') {
      fechaMin.setDate(hoy.getDate() - 30);
    } else if (periodo === 'trimestre') {
      fechaMin.setDate(hoy.getDate() - 90);
    } else if (periodo === 'todo') {
      fechaMin = new Date(2000, 0, 1);
    }

    const minTime = fechaMin.getTime();
    const maxTime = fechaMax.getTime();
    const nowTime = Date.now();

    // --- PARSER DE FECHAS ---
    const parseFecha = (valor) => {
      if (!valor) return null;
      if (valor instanceof Date) return valor.getTime();
      const str = String(valor).trim();

      const isoDate = new Date(str);
      if (!isNaN(isoDate.getTime())) return isoDate.getTime();

      const partesEspacio = str.split(' ');
      const partes        = partesEspacio[0].split('/');
      if (partes.length === 3) {
        const dia  = parseInt(partes[0], 10);
        const mes  = parseInt(partes[1], 10) - 1;
        const anio = parseInt(partes[2], 10);
        if (!isNaN(dia) && !isNaN(mes) && !isNaN(anio)) {
          return new Date(anio, mes, dia, 12, 0, 0).getTime();
        }
      }
      return null;
    };

    // --- FILTRADO Y MAPEO ---
    const dataset      = [];
    const ticketsTarde = [];

    data.forEach(row => {
      if (!row[idx['ID']]) return;

      const tTime = parseFecha(row[idx['Fecha']]);

      if (periodo !== 'todo') {
        if (!tTime) return;
        if (tTime < minTime || tTime > maxTime) return;
      }

      // Estatus
      const est       = String(row[idx['Estatus']] || '');
      const estLow    = est.toLowerCase();
      const esCerrado = ['resuelto', 'cerrado', 'completado', 'cancelado'].some(s => estLow.includes(s));

      // Vencimiento
      const vencRaw  = row[idx['Vencimiento']];
      const vencTime = parseFecha(vencRaw);

      const estaVencido = !esCerrado && vencTime != null && vencTime < nowTime;

      // ── SLA: solo con FechaResolucion ────────────────────────────────────
      let cerradoATiempo = null;
      let diasRetraso    = 0;

      if (esCerrado && vencTime != null && idxFechaRes >= 0) {
        const fResTime = parseFecha(row[idxFechaRes]);
        if (fResTime) {
          cerradoATiempo = fResTime <= vencTime;
          if (!cerradoATiempo) {
            diasRetraso = Math.ceil((fResTime - vencTime) / (1000 * 60 * 60 * 24));
          }
        }
      }

      if (estaVencido && vencTime) {
        diasRetraso = Math.ceil((nowTime - vencTime) / (1000 * 60 * 60 * 24));
      }
      // ────────────────────────────────────────────────────────────────────

      // Plan 72hrs
      const slaHoras = Number(row[idx['SLA_Horas']] || 0);
      const esPlan72 = slaHoras >= 72;

      const fIso = tTime ? new Date(tTime).toISOString() : null;

      // ── FechaCierre ──────────────────────────────────────────────────────
      // Prioridad: columna FechaCierre → fallback a FechaResolucion
      // Usado por el frontend para calcular horas de resolución por agente
      let fechaCierreIso = null;
      if (idxFechaCierre >= 0) {
        const ft = parseFecha(row[idxFechaCierre]);
        if (ft) fechaCierreIso = new Date(ft).toISOString();
      }
      if (!fechaCierreIso && idxFechaRes >= 0) {
        const ft = parseFecha(row[idxFechaRes]);
        if (ft) fechaCierreIso = new Date(ft).toISOString();
      }
      // ────────────────────────────────────────────────────────────────────

      // ── CantidadNotas ────────────────────────────────────────────────────
      const cantidadNotas = idxNotas >= 0 ? (Number(row[idxNotas]) || 0) : 0;
      // ────────────────────────────────────────────────────────────────────

      const ticket = {
        id:             row[idx['ID']],
        folio:          row[idx['Folio']],
        titulo:         row[idx['Título']]        || 'Sin Asunto',
        area:           row[idx['Área']]          || 'N/A',
        categoria:      row[idx['Categoría']]     || 'N/A',
        agente:         row[idx['AsignadoA']]     || 'Sin Asignar',
        estatus:        est,
        prioridad:      row[idx['Prioridad']]     || 'Normal',
        ubicacion:      row[idx['Ubicación']]     || 'N/A',
        cliente:        row[idx['ReportaNombre']] || 'N/A',
        fecha:          fIso,
        esCerrado:      esCerrado,
        cerradoATiempo: cerradoATiempo,
        diasRetraso:    diasRetraso,
        estaVencido:    estaVencido,
        esPlan72:       esPlan72,
        slaHoras:       slaHoras,
        vencimiento:    vencTime ? new Date(vencTime).toISOString() : null,
        fechaCierre:    fechaCierreIso,   // ← NUEVO: para horas resolución
        cantidadNotas:  cantidadNotas     // ← NUEVO: para tab Notas
      };

      dataset.push(ticket);

      if (cerradoATiempo === false || estaVencido) {
        ticketsTarde.push(ticket);
      }
    });

    // --- RESUMEN ---
    const cerradosSinFechaRes = dataset.filter(t =>
      t.esCerrado && t.cerradoATiempo === null
    ).length;

    const resumen = {
      total:         dataset.length,
      activos:       dataset.filter(t => !t.esCerrado).length,
      cerrados:      dataset.filter(t => t.esCerrado).length,
      vencidos:      dataset.filter(t => t.estaVencido).length,
      aTiempo:       dataset.filter(t => t.cerradoATiempo === true).length,
      tarde:         dataset.filter(t => t.cerradoATiempo === false).length,
      plan72:        dataset.filter(t => t.esPlan72).length,
      sinFechaRes:   cerradosSinFechaRes,
      slaCompliance: 0
    };

    const cerradosConSLA = resumen.aTiempo + resumen.tarde;
    if (cerradosConSLA > 0) {
      resumen.slaCompliance = Math.round((resumen.aTiempo / cerradosConSLA) * 100);
    }

    console.log(
      `BI V18: ${dataset.length} tickets | Vencidos: ${resumen.vencidos} | ` +
      `SLA: ${resumen.slaCompliance}% (${resumen.aTiempo}✅ / ${resumen.tarde}❌) | ` +
      `Sin FechaRes: ${resumen.sinFechaRes} | ` +
      `Con FechaCierre: ${dataset.filter(t => t.fechaCierre).length} | ` +
      `Con Notas: ${dataset.filter(t => t.cantidadNotas > 0).length}`
    );

    return JSON.stringify({ ok: true, dataset, resumen, ticketsTarde });

  } catch (e) {
    console.error(e);
    return JSON.stringify({ ok: false, error: e.message });
  }
}

/**
 * getTicketDetail - Obtiene detalle de un ticket para el modal de BI
 * @param {string} ticketId - ID del ticket
 * @returns {Object} Datos del ticket
 */
function getTicketDetail(ticketId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Tickets');
    
    if (!sheet) {
      return { ok: false, error: 'Hoja de tickets no encontrada' };
    }
    
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // Buscar columna ID
    var idCol = headers.indexOf('ID');
    if (idCol === -1) {
      return { ok: false, error: 'Columna ID no encontrada' };
    }
    
    // Buscar el ticket
    var row = null;
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][idCol]) === String(ticketId)) {
        row = data[i];
        break;
      }
    }
    
    if (!row) {
      return { ok: false, error: 'Ticket no encontrado' };
    }
    
    // Helper para obtener valor por nombre de columna
    function getVal(colName) {
      var ci = headers.indexOf(colName);
      return ci !== -1 ? (row[ci] || '') : '';
    }
    
    // Formatear fecha de vencimiento
    var vencimientoStr = 'No definido';
    var fechaVenc = getVal('Vencimiento');
    if (fechaVenc) {
      try {
        vencimientoStr = Utilities.formatDate(
          new Date(fechaVenc),
          Session.getScriptTimeZone(),
          'dd/MM/yyyy HH:mm'
        );
      } catch (e) {
        vencimientoStr = String(fechaVenc);
      }
    }
    
    return {
      ok: true,
      data: {
        id: getVal('ID'),
        folio: getVal('Folio'),
        titulo: getVal('Título') || getVal('Titulo') || 'Sin título',
        descripcion: getVal('Descripción') || getVal('Descripcion') || 'Sin descripción',
        solicitante: getVal('ReportaNombre') || getVal('Nombre') || getVal('Solicitante') || 'N/A',
        asignado: getVal('AsignadoA') || getVal('Asignado') || 'Sin asignar',
        area: getVal('Área') || getVal('Area') || 'N/A',
        estatus: getVal('Estatus') || 'Pendiente',
        prioridad: getVal('Prioridad') || 'Normal',
        vencimiento: vencimientoStr,
        ubicacion: getVal('Ubicación') || getVal('Ubicacion') || 'N/A'
      }
    };
  } catch (e) {
    console.error('Error en getTicketDetail: ' + e.message);
    return { ok: false, error: e.message };
  }
}


//Escalamientos

/**
 * NUEVA — Escapa HTML para prevenir XSS en emails generados por backend.
 * CONTEXTO: solicitarEscalamiento() (línea 6029) y handleEscalamientoApproval() 
 * llaman escapeHtml() que solo existía en el frontend.
 * UBICACIÓN: Agregar en la sección de FUNCIONES CORE (~después de línea 600)
 */
function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function notificarAdminsEscalamiento(ticketId, folio, titulo, area, motivo, nivelUrgencia, solicitanteEmail) {
  try {
    const admins = getAdminEmails(); 
    if (!admins || !admins.length) return;

    const urgenciaTexto = nivelUrgencia === 'critica' ? '🔴 CRÍTICA' : nivelUrgencia === 'alta' ? '🟠 ALTA' : '🔵 NORMAL';
    const appUrl = ScriptApp.getService().getUrl();
    const approveLink = `${appUrl}?action=approve_escalar&id=${ticketId}&by=${encodeURIComponent(solicitanteEmail)}`;
    const rejectLink = `${appUrl}?action=reject_escalar&id=${ticketId}&by=${encodeURIComponent(solicitanteEmail)}`;

    const htmlBody = `
        <h2>⚠️ Escalamiento SIN GERENTE</h2>
        <p>El siguiente ticket fue escalado, pero el área <strong>${area}</strong> no tiene un gerente asignado en el sistema.</p>
        <div class="ticket-box">
          <h3>Ticket #${folio}</h3>
          <p><strong>${titulo}</strong></p>
        </div>
        <div class="info-row"><span class="info-label">Urgencia:</span><span class="info-value">${urgenciaTexto}</span></div>
        <div class="info-row"><span class="info-label">Solicitante:</span><span class="info-value">${solicitanteEmail}</span></div>
        <div class="alert-box alert-danger">
          <p style="margin:0 0 10px;"><strong>Motivo del escalamiento:</strong></p>
          <p style="margin:0; background:white; padding:10px; border-radius:4px;">${escapeHtml(motivo)}</p>
        </div>
        <div style="margin-top:25px; text-align:center;">
          <a href="${approveLink}" style="display:inline-block; background:#16a34a; color:white; padding:12px 24px; text-decoration:none; border-radius:6px; margin: 0 10px; font-weight:bold;">✓ Aprobar</a>
          <a href="${rejectLink}" style="display:inline-block; background:#dc2626; color:white; padding:12px 24px; text-decoration:none; border-radius:6px; margin: 0 10px; font-weight:bold;">✗ Rechazar</a>
        </div>
    `;

    admins.forEach(adminEmail => {
      enviarEmailNotificacion(adminEmail, `⚠️ [${urgenciaTexto}] Escalamiento SIN GERENTE - #${folio} - ${area}`, htmlBody);
      crearNotificacion(adminEmail, 'escalamiento_sin_gerente', `Escalamiento sin gerente - #${folio}`, `Área: ${area}. Urgencia: ${urgenciaTexto}. Solicitante: ${solicitanteEmail}`, ticketId);
      const chatId = getTelegramChatIdGrupo(adminEmail);
      if (chatId) telegramSendToGrupo(chatId, `⚠️ <b>ESCALAMIENTO SIN GERENTE</b>\n📋 Ticket: #${folio}\n📍 Área: ${area}\n⚡ Urgencia: ${urgenciaTexto}\n👤 Solicitante: ${solicitanteEmail}\n💬 Motivo: ${motivo.substring(0, 100)}...`);
    });
  } catch (e) { Logger.log('Error en notificarAdminsEscalamiento: ' + e.message); }
}

function solicitarEscalamiento(ticketId, motivo, nivelUrgencia, solicitanteEmail) {
  return withLock_(() => {
    try {
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      const sh = getSheet(DB.TICKETS);

      const rowIndex = rows.findIndex(r => String(r[m.ID]) === String(ticketId));
      if (rowIndex < 0) return { ok: false, error: 'Ticket no encontrado' };

      const row = rows[rowIndex];
      const folio = row[m.Folio];
      const titulo = row[m['Título']] || '';
      const area = row[m['Área']] || '';
      const prioridad = row[m.Prioridad];

      // 1. Actualizar campos
      if (m.MotivoEscalamiento != null) row[m.MotivoEscalamiento] = motivo;
      if (m.FechaEscalamiento != null) row[m.FechaEscalamiento] = new Date();
      if (m.SolicitanteEscalamiento != null) row[m.SolicitanteEscalamiento] = solicitanteEmail;
      row[m.Estatus] = 'Escalado';
      row[m['ÚltimaActualización']] = new Date();

      sh.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
      SpreadsheetApp.flush(); // FORZAR GUARDADO INMEDIATO
      clearCache(DB.TICKETS);

      const gerenteDelArea = getGerenteDelArea(area);
      const gerenteEmail = gerenteDelArea ? gerenteDelArea.email : null;

      registrarEscalamiento_(ticketId, folio, area, solicitanteEmail, gerenteEmail, motivo, nivelUrgencia);

      // Si no hay gerente, enviar a Admins y terminar
      if (!gerenteDelArea || !gerenteEmail) {
        notificarAdminsEscalamiento(ticketId, folio, titulo, area, motivo, nivelUrgencia, solicitanteEmail);
        registrarBitacora(ticketId, 'Escalamiento solicitado', `Por: ${solicitanteEmail}. Urgencia: ${nivelUrgencia}. Sin gerente — notificado a admins.`);
        addSystemComment(ticketId, `⚠️ ESCALAMIENTO SOLICITADO\nUrgencia: ${nivelUrgencia}\nMotivo: ${motivo}\n(Se notificó a Administradores)`, false);
        return { ok: true, warning: 'No se encontró gerente, se notificó a administradores' };
      }

      // Notificar Gerente
      const urgenciaTexto = nivelUrgencia === 'critica' ? '🔴 CRÍTICA' : nivelUrgencia === 'alta' ? '🟠 ALTA' : '🔵 NORMAL';
      const appUrl = ScriptApp.getService().getUrl();
      const approveLink = `${appUrl}?action=approve_escalar&id=${ticketId}&by=${encodeURIComponent(solicitanteEmail)}`;
      const rejectLink = `${appUrl}?action=reject_escalar&id=${ticketId}&by=${encodeURIComponent(solicitanteEmail)}`;

      const htmlBody = `
        <h2>🔺 Solicitud de Escalamiento</h2>
        <div class="ticket-box">
          <h3>Ticket #${folio}</h3>
          <p><strong>${titulo}</strong></p>
        </div>
        <div class="info-row"><span class="info-label">Área:</span><span class="info-value">${area}</span></div>
        <div class="info-row"><span class="info-label">Prioridad:</span><span class="info-value">${prioridad}</span></div>
        <div class="info-row"><span class="info-label">Urgencia:</span><span class="info-value">${urgenciaTexto}</span></div>
        <div class="info-row"><span class="info-label">Solicitante:</span><span class="info-value">${solicitanteEmail}</span></div>
        <div class="alert-box alert-warning">
          <p style="margin:0 0 10px;"><strong>Motivo del escalamiento:</strong></p>
          <p style="margin:0; background:white; padding:10px; border-radius:4px;">${escapeHtml(motivo)}</p>
        </div>
        <div style="margin-top:25px; text-align:center;">
          <a href="${approveLink}" style="display:inline-block; background:#16a34a; color:white; padding:12px 24px; text-decoration:none; border-radius:6px; margin: 0 10px; font-weight:bold;">✓ Aprobar</a>
          <a href="${rejectLink}" style="display:inline-block; background:#dc2626; color:white; padding:12px 24px; text-decoration:none; border-radius:6px; margin: 0 10px; font-weight:bold;">✗ Rechazar</a>
        </div>
      `;

      enviarEmailNotificacion(gerenteEmail, `🔺 [${urgenciaTexto}] Escalamiento - Ticket #${folio}`, htmlBody);
      notificarGerenteTelegram(area, `🔺 ESCALAMIENTO PENDIENTE\nUrgencia: ${urgenciaTexto}`, { folio, titulo, area, solicitante: solicitanteEmail, motivo });
      crearNotificacion(gerenteEmail, 'escalamiento_pendiente', 'Escalamiento pendiente', `El ticket #${folio} requiere tu autorización para escalar. Urgencia: ${urgenciaTexto}`, ticketId);
      
      registrarBitacora(ticketId, 'Escalamiento solicitado', `Por: ${solicitanteEmail}. Urgencia: ${nivelUrgencia}. Gerente: ${gerenteEmail}`);
      addSystemComment(ticketId, `⚠️ ESCALAMIENTO SOLICITADO\nUrgencia: ${nivelUrgencia}\nMotivo: ${motivo}\nGerente notificado: ${gerenteEmail}`, false);

      return { ok: true, gerenteNotificado: gerenteEmail };

    } catch (e) {
      Logger.log('Error en solicitarEscalamiento: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

/**
 * NUEVA — Envía email con enlace directo al ticket.
 * CONTEXTO: handleEscalamientoApproval() (líneas 6178, 6209) la invoca pero nunca fue definida.
 * UBICACIÓN: Agregar en la sección de NOTIFICACIONES (~después de línea 8500)
 */
function enviarEmailConEnlaceDirecto(destinatario, asunto, ticketId, folio, contenidoHtml) {
  try {
    if (!destinatario) return;

    const appUrl = ScriptApp.getService().getUrl();
    const ticketUrl = `${appUrl}?ticket=${encodeURIComponent(ticketId)}`;

    MailApp.sendEmail({
      to: destinatario,
      subject: asunto,
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="padding: 25px; background: #f8fafc; border-radius: 12px; border: 1px solid #e2e8f0;">
            ${contenidoHtml}
            <div style="margin-top: 25px; padding-top: 20px; border-top: 1px solid #e2e8f0; text-align: center;">
              <a href="${ticketUrl}"
                 style="display: inline-block; background: #3b82f6; color: white; padding: 12px 30px;
                        text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 14px;">
                📋 Ver Ticket #${folio}
              </a>
              <p style="margin-top: 15px; color: #94a3b8; font-size: 0.85em;">
                Si el botón no funciona, copia este enlace:<br>
                <a href="${ticketUrl}" style="color: #3b82f6; word-break: break-all;">${ticketUrl}</a>
              </p>
            </div>
          </div>
        </div>`
    });
    Logger.log(`✅ Email con enlace directo enviado a ${destinatario}`);
  } catch (e) {
    Logger.log(`⚠️ Error en enviarEmailConEnlaceDirecto: ${e.message}`);
  }
}



function handleEscalamientoApproval(ticketId, action, solicitanteEmail) {
  const safeId = String(ticketId || '').trim();
  if (!safeId) throw new Error('ID de ticket no válido');

  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);
  const sh = getSheet(DB.TICKETS);

  const rowIndex = rows.findIndex(r => String(r[m.ID]).trim() === safeId);
  if (rowIndex < 0) throw new Error('Ticket no encontrado');

  const row = rows[rowIndex];
  const folio      = row[m.Folio];
  const titulo     = row[m['Título']]  || '';
  const area       = row[m['Área']]    || '';
  const ubicacion  = row[m['Ubicación']] || '';
  const prioridad  = row[m.Prioridad]  || '';
  const asignadoA  = row[m.AsignadoA]  || '';   // ← agente asignado

  if (action === 'approve') {
    row[m.Estatus]               = 'En Proceso';
    row[m['ÚltimaActualización']] = new Date();

    sh.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
    clearCache(DB.TICKETS);

    actualizarEscalamiento_(safeId, 'Aprobado', 'Aprobado por gerencia vía email');

    addSystemComment(safeId,
      '✅ ESCALAMIENTO APROBADO\nEl gerente ha aprobado el escalamiento. El ticket vuelve a "En Proceso" para atención prioritaria.',
      true);
    registrarBitacora(safeId, 'Escalamiento aprobado', 'Aprobado por gerencia');

    // ── Notificar al AGENTE asignado ──────────────────────────────────────
    if (asignadoA) {
      try {
        notificarAgenteTodosCanales(asignadoA, 'escalamiento_aprobado', {
          ticketId: safeId,
          folio,
          titulo,
          area,
          ubicacion,
          prioridad,
          tituloNotif: '✅ Escalamiento Aprobado: #' + folio,
          mensajeNotif: 'Gerencia ha aprobado el escalamiento. Atención prioritaria requerida.'
        });
      } catch (e) {
        Logger.log('⚠️ Error notificando agente (approve): ' + e.message);
      }
    }

    // ── Notificar al solicitante si es diferente al agente ─────────────────
    if (solicitanteEmail && solicitanteEmail !== asignadoA) {
      crearNotificacion(solicitanteEmail, 'escalamiento_aprobado', 'Escalamiento aprobado',
        `Tu solicitud de escalamiento para el ticket #${folio} ha sido APROBADA.`, safeId);
      enviarEmailConEnlaceDirecto(solicitanteEmail,
        `✅ Escalamiento APROBADO - Ticket #${folio}`, safeId, folio,
        `<h2 style="color:#16a34a;">✅ Escalamiento Aprobado</h2>
         <p>Tu solicitud de escalamiento para el ticket <b>#${folio}</b> ha sido <strong style="color:#16a34a;">APROBADA</strong>.</p>
         <p>El ticket ha sido marcado para atención prioritaria.</p>`);
    }

    return 'Escalamiento APROBADO. El ticket ha sido marcado para atención prioritaria.';

  } else if (action === 'reject') {
    row[m.Estatus]               = 'En Proceso';
    row[m['ÚltimaActualización']] = new Date();

    sh.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
    clearCache(DB.TICKETS);

    actualizarEscalamiento_(safeId, 'Rechazado', 'Rechazado por gerencia vía email');

    addSystemComment(safeId,
      '❌ ESCALAMIENTO RECHAZADO\nEl gerente ha rechazado la solicitud de escalamiento. El ticket continúa en proceso normal.',
      true);
    registrarBitacora(safeId, 'Escalamiento rechazado', 'Rechazado por gerencia');

    // ── Notificar al AGENTE asignado ──────────────────────────────────────
    if (asignadoA) {
      try {
        notificarAgenteTodosCanales(asignadoA, 'escalamiento_rechazado', {
          ticketId: safeId,
          folio,
          titulo,
          area,
          ubicacion,
          prioridad,
          motivo: 'Rechazado por gerencia vía email',
          tituloNotif: '❌ Escalamiento Rechazado: #' + folio,
          mensajeNotif: 'Gerencia ha rechazado el escalamiento. El ticket continúa en proceso normal.'
        });
      } catch (e) {
        Logger.log('⚠️ Error notificando agente (reject): ' + e.message);
      }
    }

    // ── Notificar al solicitante si es diferente al agente ─────────────────
    if (solicitanteEmail && solicitanteEmail !== asignadoA) {
      crearNotificacion(solicitanteEmail, 'escalamiento_rechazado', 'Escalamiento rechazado',
        `Tu solicitud de escalamiento para el ticket #${folio} ha sido RECHAZADA.`, safeId);
      enviarEmailConEnlaceDirecto(solicitanteEmail,
        `❌ Escalamiento RECHAZADO - Ticket #${folio}`, safeId, folio,
        `<h2 style="color:#dc2626;">❌ Escalamiento Rechazado</h2>
         <p>Tu solicitud de escalamiento para el ticket <b>#${folio}</b> ha sido <strong style="color:#dc2626;">RECHAZADA</strong>.</p>
         <p>Por favor, continúa con la atención normal del ticket.</p>`);
    }

    return 'Escalamiento RECHAZADO. El ticket continúa en proceso normal.';
  }

  return 'Acción no válida.';
}


function reasignarTicket(ticketId, nuevoAgenteEmail, motivo, emailSolicitante) {
  return withLock_(() => {
    try {
      const user = getUser(emailSolicitante || '');
      const userRole = (user.rol || '').toLowerCase();
      
      // 1. OBTENER EL TICKET
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      const sh = getSheet(DB.TICKETS);
      
      const rowIndex = rows.findIndex(r => String(r[m.ID]).trim() === String(ticketId).trim());
      if (rowIndex < 0) return { ok: false, error: 'Ticket no encontrado' };
      
      const ticketArea = (rows[rowIndex][m['Área']] || '').toLowerCase();
      const agenteAnterior = rows[rowIndex][m.AsignadoA] || '';
      const folio = rows[rowIndex][m.Folio] || ticketId;
      
      // 2. VERIFICACIÓN DE PERMISOS
      var SUPERADMINS = ['rgnava', 'rcesquivel'];
      var esSuperAdmin = SUPERADMINS.some(function(sa) { 
        return user.email.toLowerCase().includes(sa.toLowerCase()); 
      });
      
      // Admin: rol exacto 'admin' O rol que contenga 'admin'
      var esAdmin = userRole === 'admin' || userRole.includes('admin');
      
      // Gerente: cualquier gerente puede reasignar (no solo el de esa área)
      var esGerente = userRole.includes('gerente');
      
      // Gerente específico del área (desde ConfigGerentes)
      var esGerenteConfig = false;
      try {
        var areaGerente = getGerenteArea(user.email);
        esGerenteConfig = !!areaGerente; // Si está en ConfigGerentes, es gerente
      } catch (e) {}

      // REGLA: SuperAdmin, Admin, o cualquier Gerente pueden reasignar cualquier ticket
      if (!esSuperAdmin && !esAdmin && !esGerente && !esGerenteConfig) {
        return { ok: false, error: '⛔ Acceso denegado: Solo Gerentes y Administradores pueden reasignar tickets.' };
      }
      
      // 3. VALIDAR NUEVO AGENTE
      if (!nuevoAgenteEmail) return { ok: false, error: 'Debe especificar un agente' };
      
      // ... (resto de la lógica de guardado y notificación igual) ...
      const nuevoAgente = getUserByEmail(nuevoAgenteEmail);
      if (!nuevoAgente) return { ok: false, error: 'Agente no encontrado' };
      
      // Guardar cambios
      sh.getRange(rowIndex + 2, m.AsignadoA + 1).setValue(nuevoAgenteEmail);
      sh.getRange(rowIndex + 2, m['ÚltimaActualización'] + 1).setValue(new Date());
      
      registrarBitacora(ticketId, 'Reasignación', 
        `De "${agenteAnterior}" a "${nuevoAgenteEmail}". Motivo: ${motivo || 'Manual'}. Por: ${user.nombre}`);
      
      // Notificaciones...
      try {
        notificarAgenteTodosCanales(nuevoAgenteEmail, 'reasignacion', {
          ticketId: ticketId,
          folio: folio,
          titulo: rows[rowIndex][m['Título']] || '',
          area: rows[rowIndex][m['Área']] || '',
          ubicacion: rows[rowIndex][m['Ubicación']] || '',
          prioridad: rows[rowIndex][m.Prioridad] || '',
          reportaNombre: rows[rowIndex][m.ReportaNombre] || '',
          motivo: motivo || 'Reasignación por gerencia',
          tituloNotif: 'Ticket Reasignado',
          mensajeNotif: 'Se te ha reasignado el ticket #' + folio
        });
      } catch (e) { Logger.log('Error notificación reasignación omnicanal: ' + e.message); }
      
      clearCache(DB.TICKETS);
      return { ok: true, message: 'Ticket reasignado correctamente' };
      
    } catch (e) {
      Logger.log('Error reasignarTicket: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

// ============================================================
// PARTE 3: BACKEND - Nuevas funciones para gerentes
// ============================================================

/**
 * Modificar fecha de vencimiento (SLA) - Solo admin y gerentes
 */
function modificarVencimientoTicket(ticketId, nuevaFechaVencimiento, motivo, emailUsuario) {
  return withLock_(() => {
    try {
      const user = getUser(emailUsuario);
      const role = (user.rol || '').toLowerCase();
      
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      const sh = getSheet(DB.TICKETS);
      
      const rowIndex = rows.findIndex(r => String(r[m.ID]).trim() === String(ticketId).trim());
      if (rowIndex < 0) {
        return { ok: false, error: 'Ticket no encontrado' };
      }
      
      const ticketArea = (rows[rowIndex][m['Área']] || '').toLowerCase();
      const vencimientoAnterior = rows[rowIndex][m.Vencimiento] || '';
      
      // Validar permisos
      const areaGerente = getGerenteArea(emailUsuario);
      const esGerenteDelArea = areaGerente && areaGerente.toLowerCase() === ticketArea;
      
      if (role !== 'admin' && !esGerenteDelArea) {
        return { ok: false, error: 'No tienes permisos para modificar el vencimiento' };
      }
      
      // Actualizar vencimiento
      const vencimientoCol = m.Vencimiento;
      const updateCol = m['ÚltimaActualización'];
      
      sh.getRange(rowIndex + 2, vencimientoCol + 1).setValue(new Date(nuevaFechaVencimiento));
      sh.getRange(rowIndex + 2, updateCol + 1).setValue(new Date());
      
      // Bitácora
      registrarBitacora(ticketId, 'Vencimiento modificado', 
        `De "${vencimientoAnterior}" a "${nuevaFechaVencimiento}". Motivo: ${motivo || 'Sin especificar'}. Por: ${user.nombre || user.email}`);
      
      clearCache(DB.TICKETS);
      
      return { ok: true, message: 'Vencimiento actualizado correctamente' };
      
    } catch (e) {
      Logger.log('Error en modificarVencimientoTicket: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}


/**
 * Modificar categoría - Solo admin y gerentes
 */
function modificarCategoriaTicket(ticketId, nuevaCategoria, motivo, emailUsuario) {
  return withLock_(() => {
    try {
      const user = getUser(emailUsuario);
      const role = (user.rol || '').toLowerCase();
      
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      const sh = getSheet(DB.TICKETS);
      
      const rowIndex = rows.findIndex(r => String(r[m.ID]).trim() === String(ticketId).trim());
      if (rowIndex < 0) {
        return { ok: false, error: 'Ticket no encontrado' };
      }
      
      const ticketArea = (rows[rowIndex][m['Área']] || '').toLowerCase();
      const categoriaAnterior = rows[rowIndex][m['Categoría']] || '';
      
      // Validar permisos
      const areaGerente = getGerenteArea(emailUsuario);
      const esGerenteDelArea = areaGerente && areaGerente.toLowerCase() === ticketArea;
      
      if (role !== 'admin' && !esGerenteDelArea) {
        return { ok: false, error: 'No tienes permisos para modificar la categoría' };
      }
      
      // Actualizar categoría
      const categoriaCol = m['Categoría'];
      const updateCol = m['ÚltimaActualización'];
      
      sh.getRange(rowIndex + 2, categoriaCol + 1).setValue(nuevaCategoria);
      sh.getRange(rowIndex + 2, updateCol + 1).setValue(new Date());
      
      // Bitácora
      registrarBitacora(ticketId, 'Categoría modificada', 
        `De "${categoriaAnterior}" a "${nuevaCategoria}". Motivo: ${motivo || 'Sin especificar'}. Por: ${user.nombre || user.email}`);
      
      clearCache(DB.TICKETS);
      
      return { ok: true, message: 'Categoría actualizada correctamente' };
      
    } catch (e) {
      Logger.log('Error en modificarCategoriaTicket: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

/**
 * Obtener lista de agentes para reasignación
 */
function getAgentes() {
  try {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    
    return rows
      .filter(r => {
        const rol = String(r[m.Rol] || '').toLowerCase();
        return rol === 'agente' || rol === 'admin';
      })
      .map(r => ({
        email: r[m.Email],
        nombre: r[m.Nombre] || r[m.Email],
        rol: r[m.Rol],
        area: r[m.Área] || ''
      }))
      .sort((a, b) => (a.nombre || '').localeCompare(b.nombre || ''));
      
  } catch (e) {
    Logger.log('Error en getAgentes: ' + e.message);
    return [];
  }
}

function confirmarResolucion(ticketId, comentario, emailManual) {
  const user = getUser(emailManual || '');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(DB.TICKETS);

  const data = sh.getDataRange().getValues();
  const hdr = data[0];

  const idCol = hdr.indexOf('ID');
  const folioCol = hdr.indexOf('Folio');
  const tituloCol = hdr.indexOf('Título');
  const statusCol = hdr.indexOf('Estatus');
  const reportaCol = hdr.indexOf('ReportaEmail');
  const asignadoCol = hdr.indexOf('AsignadoA');
  const lastUpdateCol = hdr.indexOf('ÚltimaActualización');

  for (let i = 1; i < data.length; i++) {
    if (data[i][idCol] === ticketId) {
      const currentStatus = data[i][statusCol];
      const reportaEmail = String(data[i][reportaCol]).toLowerCase().trim();
      const userEmail = String(user.email).toLowerCase().trim();
      const asignadoEmail = String(data[i][asignadoCol] || '').trim();
      const folio = data[i][folioCol] || ticketId;
      const titulo = data[i][tituloCol] || 'Sin título';

      // Validación de permisos
      if (reportaEmail !== userEmail && user.rol !== 'admin') {
        return { ok: false, error: 'Solo el usuario que reportó puede confirmar la resolución' };
      }

      // Validación de estatus
      if (currentStatus !== 'Resuelto') {
        return { ok: false, error: 'Solo se pueden confirmar tickets en estado Resuelto' };
      }

      // Cerrar ticket
      sh.getRange(i + 1, statusCol + 1).setValue('Cerrado');
      sh.getRange(i + 1, lastUpdateCol + 1).setValue(new Date());

      // Bitácora
      registrarBitacora(
        ticketId,
        'Resolución confirmada',
        comentario || 'Usuario confirmó que el problema fue resuelto'
      );

      // ========== NUEVO: Notificar al agente asignado ==========
      if (asignadoEmail) {
        const cuerpoHtml = `<h2>✅ Resolución Confirmada</h2><p>El usuario ha confirmado que el problema fue resuelto satisfactoriamente.</p><div class="ticket-box"><h3>Ticket #${folio}</h3><p><strong>${titulo}</strong></p></div><div class="alert-box alert-success"><p style="margin:0;">🎉 <strong>¡Excelente trabajo!</strong> El ticket ha sido cerrado exitosamente.</p></div><p style="margin-top:25px;"><a href="${getScriptUrl()}" class="btn btn-success">🔗 Ver Ticket #${folio}</a></p>`;
        
        notificarAgenteTodosCanales(asignadoEmail, 'resolucion_confirmada', {
          ticketId: ticketId,
          folio: folio,
          titulo: titulo,
          tituloNotif: '✅ Resolución confirmada: #' + folio,
          mensajeNotif: 'El usuario confirmó la resolución del ticket. Ticket cerrado.',
          htmlCuerpo: cuerpoHtml
        });
      }
      // =========================================================

      clearCache(DB.TICKETS);

      return { ok: true, message: 'Ticket cerrado exitosamente' };
    }
  }

  return { ok: false, error: 'Ticket no encontrado' };
}



// ============================================================
// PARTE 2: BACKEND - reabrirTicket CORREGIDA
// ============================================================
// - Cambia a "En Proceso" (no "Reabierto")
// - Agrega marcador "FueReabierto = true"
// - Reinicia el SLA
// - Verifica 5 días hábiles desde el cierre
// REEMPLAZAR función completa (línea ~6396)

function reabrirTicket(ticketId, motivo, emailManual) {
  const user = getUser(emailManual || '');
  if (!motivo || motivo.trim().length < 10) return { ok: false, error: 'El motivo es obligatorio (mínimo 10 caracteres)' };

  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(DB.TICKETS);
    
    // Asegurarnos de que existan las columnas nuevas si el Sheet es viejo
    const maxCols = sh.getLastColumn();
    const currentHeaders = sh.getRange(1, 1, 1, maxCols).getValues()[0];
    if (!currentHeaders.includes('FueReabierto')) {
      sh.getRange(1, maxCols + 1).setValue('FueReabierto');
      sh.getRange(1, maxCols + 2).setValue('ContadorReaperturas');
    }

    const { headers, rows } = _readTableByHeader_(DB.TICKETS);
    const m = _headerMap_(headers);

    const rowIndex = rows.findIndex(r => String(r[m.ID]) === String(ticketId));
    if (rowIndex < 0) return { ok: false, error: 'Ticket no encontrado' };

    const row = rows[rowIndex];
    
    // 1. REINICIAR SLA (Desde AHORA mismo)
    const area = row[m['Área']];
    const categoria = row[m['Categoría']];
    const ubicacion = row[m['Ubicación']];
    const configSLA = obtenerConfiguracionSLA(area, categoria, ubicacion);
    const horasSLA = Number(configSLA.sla) || 24;
    const ahora = new Date();
    const nuevoVencimiento = calcularFechaConHorarioLaboral(ahora, horasSLA);

    // 2. ACTUALIZAR DATOS
    row[m.Estatus] = 'En Proceso'; // DEBE REGRESAR A "EN PROCESO"
    row[m['ÚltimaActualización']] = ahora;
    row[m.Vencimiento] = nuevoVencimiento;
    
    // MARCADORES DE REAPERTURA
    if (m.FueReabierto != null) row[m.FueReabierto] = true;
    if (m.ContadorReaperturas != null) {
      row[m.ContadorReaperturas] = Number(row[m.ContadorReaperturas] || 0) + 1;
    }
    
    // Limpiar Fechas de Cierre/Resolución para que no afecten los KPIs
    if (m.FechaResuelto != null) row[m.FechaResuelto] = '';
    if (m.FechaCierre != null) row[m.FechaCierre] = '';

    // Extender la fila si agregamos columnas nuevas
    while (row.length < headers.length) row.push('');

    sh.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
    clearCache(DB.TICKETS);

    registrarBitacora(ticketId, 'Reapertura', `Motivo: ${motivo}. SLA reiniciado a ${horasSLA}h.`);
    addSystemComment(ticketId, `🔄 TICKET REABIERTO\nMotivo: ${motivo}\nNuevo Vencimiento: ${nuevoVencimiento.toLocaleString('es-MX')}`, true);

    // 3. NOTIFICAR AL AGENTE ASIGNADO OMNICANAL
    const asignadoA = row[m.AsignadoA];
    const folio = row[m.Folio];
    const titulo = row[m['Título']];

    if (asignadoA) {
      notificarAgenteTodosCanales(asignadoA, 'reapertura', {
        ticketId: ticketId,
        folio: folio,
        titulo: titulo,
        area: area,
        ubicacion: ubicacion,
        prioridad: row[m.Prioridad],
        reportaNombre: row[m.ReportaNombre] || user.nombre || user.email,
        vencimiento: nuevoVencimiento,
        motivo: motivo,
        tituloNotif: '🔄 Ticket Reabierto #' + folio,
        mensajeNotif: 'El usuario reabrió el ticket. Motivo: ' + motivo + '. Nuevo SLA activado.'
      });
    }

    return { ok: true, message: 'Ticket reabierto y notificado al agente.' };
  });
}


// ============================================================
// PARTE 3: FUNCIÓN AUXILIAR - Calcular días hábiles
// ============================================================

function calcularDiasHabiles(fechaInicio, fechaFin) {
  let diasHabiles = 0;
  const fecha = new Date(fechaInicio);
  
  while (fecha < fechaFin) {
    const diaSemana = fecha.getDay();
    // 0 = Domingo, 6 = Sábado
    if (diaSemana !== 0 && diaSemana !== 6) {
      diasHabiles++;
    }
    fecha.setDate(fecha.getDate() + 1);
  }
  
  return diasHabiles;
}


// ============================================================
// PARTE 4: FUNCIÓN AUXILIAR - Calcular nuevo vencimiento SLA
// ============================================================
/**
 * Calcula fecha de vencimiento SLA
 */
function calcularVencimientoSLA(prioridad, fechaInicio) {
  try {
    // Obtener horas SLA según prioridad
    const { headers, rows } = _readTableByHeader_('Prioridades');
    const m = _headerMap_(headers);
    
    const prioRow = rows.find(r => 
      String(r[m.Nombre] || '').toLowerCase() === String(prioridad || '').toLowerCase()
    );
    
    const horasSLA = prioRow ? Number(prioRow[m.HorasSLA] || prioRow[m.Horas] || 24) : 24;
    
    // Calcular vencimiento considerando horario laboral
    const inicio = fechaInicio ? new Date(fechaInicio) : new Date();
    return calcularFechaConHorarioLaboral(inicio, horasSLA);
  } catch (e) {
    Logger.log('⚠️ Error calculando SLA: ' + e.message);
    // Fallback: sumar horas directamente
    const venc = new Date();
    venc.setHours(venc.getHours() + 24);
    return venc;
  }
}

/**
 * Obtener días festivos de la hoja "DiasFestivos"
 * Formato esperado: Columna A = Fecha (Date)
 */
function getDiasFestivos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('DiasFestivos');
    if (!sheet || sheet.getLastRow() < 2) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    return data
      .map(row => {
        const d = row[0];
        if (d instanceof Date) {
          return Utilities.formatDate(d, 'America/Mexico_City', 'yyyy-MM-dd');
        }
        return null;
      })
      .filter(d => d !== null);
  } catch (e) {
    Logger.log('Error obteniendo días festivos: ' + e.message);
    return [];
  }
}

/**
 * Verificar si una fecha es día festivo
 */
function esDiaFestivo(fecha) {
  const festivos = getDiasFestivos();
  const fechaStr = Utilities.formatDate(fecha, 'America/Mexico_City', 'yyyy-MM-dd');
  return festivos.includes(fechaStr);
}

/**
 * Calcula fecha final sumando horas laborales (Exactitud de minutos)
 * Horario: Lunes-Viernes | 8:00-14:00 y 16:00-18:00
 * Zona Horaria: America/Mexico_City
 */
function calcularFechaConHorarioLaboral(fechaInicio, horasNecesarias) {
  const TZ = 'America/Mexico_City';
  let fecha = new Date(fechaInicio);
  let tiempoRestanteMs = horasNecesarias * 60 * 60 * 1000;

  const HORA_INICIO_1 = 8;
  const HORA_FIN_1    = 14;
  const HORA_INICIO_2 = 16;
  const HORA_FIN_2    = 18;

  // ── FESTIVOS: leer una sola vez y cachear en esta ejecución ──────────────
  let _festivosSet = null;
  function _getFestivos() {
    if (_festivosSet) return _festivosSet;
    _festivosSet = new Set();
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName('DiasFestivos');
      if (sh) {
        const rows = sh.getDataRange().getValues().slice(1); // saltar header
        const tz   = Session.getScriptTimeZone();
        rows.forEach(function(r) {
          // Columna 3 (índice 3) = Activo
          const activo = r[3];
          if (activo === false || String(activo).toLowerCase() === 'false') return;
          const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
          if (!isNaN(d.getTime())) {
            _festivosSet.add(Utilities.formatDate(d, tz, 'yyyy-MM-dd'));
          }
        });
      }
    } catch(e) {
      Logger.log('calcularFechaConHorarioLaboral - error leyendo festivos: ' + e.message);
    }
    return _festivosSet;
  }

  // ── Helpers ──────────────────────────────────────────────────────────────
  function esDiaFestivo(d) {
    const festivos = _getFestivos();
    const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return festivos.has(key);
  }

  const esDiaInhabil = function(d) {
    const dia = d.getDay(); // 0 Dom, 6 Sab
    return dia === 0 || dia === 6 || esDiaFestivo(d);
  };

  function ajustarAHorarioLaboral(d) {
    // 1. Saltar días inhábiles
    let guard = 0;
    while (esDiaInhabil(d) && guard < 15) {
      d.setDate(d.getDate() + 1);
      d.setHours(HORA_INICIO_1, 0, 0, 0);
      guard++;
    }

    const h         = d.getHours();
    const m         = d.getMinutes();
    const timeFloat = h + m / 60;

    if (timeFloat < HORA_INICIO_1) {
      d.setHours(HORA_INICIO_1, 0, 0, 0);
    } else if (timeFloat >= HORA_FIN_1 && timeFloat < HORA_INICIO_2) {
      d.setHours(HORA_INICIO_2, 0, 0, 0);
    } else if (timeFloat >= HORA_FIN_2) {
      d.setDate(d.getDate() + 1);
      d.setHours(HORA_INICIO_1, 0, 0, 0);
      ajustarAHorarioLaboral(d); // recursión para saltar festivos encadenados
    }
    return d;
  }

  // ── Lógica principal (sin cambios respecto a tu versión original) ────────
  fecha = ajustarAHorarioLaboral(fecha);

  let iteraciones = 0;
  while (tiempoRestanteMs > 0 && iteraciones < 1000) {
    iteraciones++;

    const h = fecha.getHours();
    let fechaLimiteBloque = new Date(fecha);

    if (h < HORA_FIN_1) {
      fechaLimiteBloque.setHours(HORA_FIN_1, 0, 0, 0);
    } else {
      fechaLimiteBloque.setHours(HORA_FIN_2, 0, 0, 0);
    }

    const msDisponibles = fechaLimiteBloque.getTime() - fecha.getTime();

    if (tiempoRestanteMs <= msDisponibles) {
      fecha.setTime(fecha.getTime() + tiempoRestanteMs);
      tiempoRestanteMs = 0;
    } else {
      tiempoRestanteMs -= msDisponibles;
      fecha.setTime(fechaLimiteBloque.getTime());
      fecha.setSeconds(fecha.getSeconds() + 1);
      fecha = ajustarAHorarioLaboral(fecha);
      fecha.setSeconds(0);
    }
  }

  return fecha;
}

function getAdminEmails() {
  try {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    return rows
      .filter(r => String(r[m.Rol]).toLowerCase() === 'admin' && String(r[m.Estatus] || r[m.EstatusUsuario] || '').toLowerCase() !== 'baja')
      .map(r => r[m.Email]);
  } catch(e) {
    Logger.log('Error en getAdminEmails: ' + e.message);
    return [];
  }
}



/**
 * Obtener URL del ticket
 */
function getTicketUrl(ticketId) {
  const scriptUrl = ScriptApp.getService().getUrl();
  return `${scriptUrl}?ticket=${ticketId}`;
}

// ============================================================================
// PASO 2: AGREGAR ESTAS FUNCIONES PARA EL SISTEMA DE NOTIFICACIONES
// ============================================================================

/**
 * Obtener notificaciones de un usuario
 * CORREGIDO: Manejo robusto de Booleanos (TRUE/FALSE) y Nombres de Usuario
 */
function getNotificaciones(email, limite) {
  if (!email) return [];
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.NOTIFS);
    if (!headers || !rows) return [];
    
    const m = _headerMap_(headers);
    
    // Normalizamos el email que solicita (ej: rgnava@bexalta.com -> rgnava)
    const emailLower = String(email).toLowerCase().trim();
    const usuarioBase = emailLower.split('@')[0]; 

    return rows
      .filter(r => {
        // 1. LÓGICA DE USUARIO FLEXIBLE
        // Permite que "RGNava" (Hoja) coincida con "rgnava@bexalta.com" (Login)
        const rowUser = String(r[m.Usuario] || '').toLowerCase().trim();
        const rowUserBase = rowUser.split('@')[0];
        
        return rowUser === emailLower || rowUserBase === usuarioBase;
      })
      .map(r => {
        // 2. LÓGICA DE LEÍDO ROBUSTA
        // Convierte cualquier cosa que parezca "verdad" en true
        const valRaw = r[m.Leido];
        const valString = String(valRaw).toUpperCase().trim();
        
        const isLeido = (valRaw === true) || 
                        (valString === 'TRUE') || 
                        (valString === '1') || 
                        (valString === 'SI') ||
                        (valString === 'YES');

        return {
          id: String(r[m.ID] || genId()),
          fecha: r[m.Fecha] ? new Date(r[m.Fecha]).toISOString() : '',
          tipo: String(r[m.Tipo] || 'info'),
          titulo: String(r[m['Título']] || ''),
          mensaje: String(r[m.Mensaje] || ''),
          ticketId: String(r[m.TicketID] || ''),
          leido: isLeido,  // <--- Aquí asignamos el valor corregido
          ts: Number(r[m.Timestamp]) || 0
        };
      })
      .sort((a, b) => b.ts - a.ts) // Más recientes primero
      .slice(0, limite || 20);

  } catch (e) {
    Logger.log('Error en getNotificaciones: ' + e.message);
    return [];
  }
}

function marcarNotificacionLeida(notifId) {
  if (!notifId) return { ok: false, error: 'ID requerido' };
  
  Logger.log(`[ACTION] Intentando marcar LEÍDA la notificación: ${notifId}`);

  return withLock_(() => {
    try {
      const sh = getSheet(DB.NOTIFS);
      // Usamos getDataRange para leer datos frescos, no cacheados
      const data = sh.getDataRange().getValues();
      const headers = data[0];
      
      const idCol = headers.indexOf('ID');
      const leidoCol = headers.indexOf('Leido');
      
      if (idCol < 0 || leidoCol < 0) {
        Logger.log('[ERROR] No se encontraron columnas ID o Leido');
        return { ok: false, error: 'Estructura DB incorrecta' };
      }
      
      let found = false;
      for (let i = 1; i < data.length; i++) {
        // Convertimos ambos a String y trim por si acaso hay espacios invisibles
        if (String(data[i][idCol]).trim() === String(notifId).trim()) {
          
          Logger.log(`[MATCH] Fila ${i+1} encontrada. Valor anterior: ${data[i][leidoCol]}`);
          
          // Escritura directa a la celda
          sh.getRange(i + 1, leidoCol + 1).setValue(true); // TRUE booleano
          
          found = true;
          break;
        }
      }

      if (found) {
        SpreadsheetApp.flush(); // FORZAR GUARDADO EN DISCO
        clearCache(DB.NOTIFS);  // BORRAR CACHÉ VIEJA
        Logger.log('[SUCCESS] Notificación marcada y caché limpiado.');
        return { ok: true };
      } else {
        Logger.log('[WARN] ID no encontrado en la hoja.');
        return { ok: false, error: 'No encontrada' };
      }

    } catch (e) {
      Logger.log('[EXCEPTION] ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

function marcarTodasNotificacionesLeidas(email) {
  if (!email) return { ok: false, error: 'Email requerido' };
  
  Logger.log(`[ACTION] Marcar TODAS leídas para: ${email}`);

  return withLock_(() => {
    try {
      const sh = getSheet(DB.NOTIFS);
      const lastRow = sh.getLastRow();
      if (lastRow < 2) return { ok: true, updated: 0 };

      // Leer datos frescos
      const data = sh.getDataRange().getValues();
      const headers = data[0];
      const userColIdx = headers.indexOf('Usuario');
      const leidoColIdx = headers.indexOf('Leido');

      if (userColIdx === -1 || leidoColIdx === -1) {
        return { ok: false, error: 'Columnas no encontradas' };
      }

      const emailLower = email.toLowerCase().trim();
      let updatedCount = 0;
      
      // Obtener rango de la columna Leido para editar en bloque
      const leidoRange = sh.getRange(2, leidoColIdx + 1, lastRow - 1, 1);
      const leidoValues = leidoRange.getValues(); 

      for (let i = 0; i < leidoValues.length; i++) {
        // data[i+1] es la fila correspondiente (saltando header)
        const rowUser = String(data[i + 1][userColIdx] || '').toLowerCase().trim();
        
        // Verificación robusta de "falso"
        const val = leidoValues[i][0];
        const isLeido = val === true || String(val).toLowerCase() === 'true' || val === 1;

        if (rowUser === emailLower && !isLeido) {
          leidoValues[i][0] = true; // Actualizar en matriz
          updatedCount++;
        }
      }

      if (updatedCount > 0) {
        leidoRange.setValues(leidoValues); // Escribir cambios
        SpreadsheetApp.flush(); // FORZAR GUARDADO
        clearCache(DB.NOTIFS);  // LIMPIAR CACHÉ
        Logger.log(`[SUCCESS] ${updatedCount} notificaciones actualizadas.`);
      } else {
        Logger.log(`[INFO] No había notificaciones pendientes.`);
      }

      return { ok: true, updated: updatedCount };
    } catch (e) {
      Logger.log('[EXCEPTION] ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}


/**
 * Crear una nueva notificación
 * @param {string} email - Email del destinatario
 * @param {string} tipo - Tipo de notificación
 * @param {string} titulo - Título de la notificación
 * @param {string} mensaje - Mensaje de la notificación
 * @param {string} ticketId - ID del ticket relacionado (opcional)
 * @returns {Object} Resultado de la operación
 */
function crearNotificacion(email, tipo, titulo, mensaje, ticketId) {
  if (!email) return { ok: false, error: 'Email requerido' };
  
  try {
    const sh = getSheet(DB.NOTIFS);
    const now = new Date();
    const id = genId();
    
    const { headers } = _readTableByHeader_(DB.NOTIFS);
    const m = _headerMap_(headers);
    
    // Construir fila según headers actuales
    const newRow = headers.map(h => {
      switch(h) {
        case 'ID': return id;
        case 'Fecha': return now;
        case 'Usuario': return email;
        case 'Tipo': return tipo || 'info';
        case 'Título': return titulo || '';
        case 'Mensaje': return mensaje || '';
        case 'TicketID': return ticketId || '';
        case 'Leido': return false;
        case 'Timestamp': return now.getTime();
        default: return '';
      }
    });
    
    sh.appendRow(newRow);
    clearCache(DB.NOTIFS);
    
    return { ok: true, id };
  } catch (e) {
    Logger.log('Error creando notificación: ' + e.message);
    return { ok: false, error: e.message };
  }
}

function pollNotifications(email, lastTs) {
  if (!email) return [];
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.NOTIFS);
    if (!headers || !rows) return [];
    
    const m = _headerMap_(headers);
    const emailLower = email.toLowerCase();
    const ts = Number(lastTs) || 0;
    
    return rows
      .filter(r => {
        const isUser = String(r[m.Usuario] || '').toLowerCase() === emailLower;
        const isNew  = Number(r[m.Timestamp] || 0) > ts;
        return isUser && isNew;
      })
      .map(r => ({
        id:       String(r[m.ID] || ''),
        fecha:    r[m.Fecha] ? new Date(r[m.Fecha]).toISOString() : '',
        tipo:     String(r[m.Tipo] || 'info'),
        titulo:   String(r[m['Título']] || ''),
        mensaje:  String(r[m.Mensaje] || ''),
        ticketId: String(r[m.TicketID] || ''),
        leido:    String(r[m.Leido] || '').toLowerCase() === 'true', // ← leer campo real
        ts:       Number(r[m.Timestamp]) || 0
      }))
      .sort((a, b) => b.ts - a.ts);
  } catch (e) {
    Logger.log('Error en pollNotifications: ' + e.message);
    return [];
  }
}

// ============================================================================
// PASO 3: FUNCIONES PARA ENVIAR NOTIFICACIONES AUTOMÁTICAS
// Agregar estas llamadas donde corresponda en tu código existente
// ============================================================================

/**
 * Notificar al usuario cuando se crea un ticket
 * @param {Object} ticket - Datos del ticket
 */
function notificarTicketCreado(ticket) {
  crearNotificacion(
    ticket.reportaEmail,
    'nuevo_ticket',
    'Ticket creado',
    `Tu ticket #${ticket.folio} ha sido registrado correctamente.`,
    ticket.id
  );
}

/**
 * Notificar al agente cuando se le asigna un ticket
 * @param {Object} ticket - Datos del ticket
 * @param {string} agenteEmail - Email del agente
 */
function notificarAsignacion(ticket, agenteEmail) {
  crearNotificacion(
    agenteEmail,
    'asignacion',
    'Nuevo ticket asignado',
    `Se te ha asignado el ticket #${ticket.Folio}: ${ticket['Título'] || ''}`,
    ticket.ID || ticket.Id || ticket.id
  );
}

/**
 * Notificar al usuario cuando su ticket cambia de estado
 * @param {Object} ticket - Datos del ticket
 * @param {string} nuevoEstatus - Nuevo estado
 */
function notificarCambioEstatus(ticket, nuevoEstatus) {
  const ticketId = ticket.ID || ticket.Id || ticket.id;
  const folio = ticket.Folio;
  const userEmail = ticket.ReportaEmail;
  
  let tipo = 'info';
  let titulo = 'Actualización de ticket';
  let mensaje = `Tu ticket #${folio} ha sido actualizado a: ${nuevoEstatus}`;
  
  if (nuevoEstatus === 'Resuelto') {
    tipo = 'resuelto';
    titulo = '¡Ticket resuelto!';
    mensaje = `Tu ticket #${folio} ha sido marcado como resuelto. Por favor confirma si el problema fue solucionado.`;
  } else if (nuevoEstatus === 'Cerrado') {
    tipo = 'cerrado';
    titulo = 'Ticket cerrado';
    mensaje = `Tu ticket #${folio} ha sido cerrado.`;
  } else if (nuevoEstatus === 'En Cotización') {
    tipo = 'aprobacion';
    titulo = 'Cotización pendiente';
    mensaje = `Tu ticket #${folio} requiere aprobación de cotización.`;
  }
  
  crearNotificacion(userEmail, tipo, titulo, mensaje, ticketId);
}

/**
 * Notificar cuando hay un nuevo comentario
 * @param {Object} ticket - Datos del ticket
 * @param {string} autorEmail - Email del autor del comentario
 */
function notificarComentario(ticket, autorEmail) {
  const ticketId = ticket.ID || ticket.Id || ticket.id;
  const folio = ticket.Folio;
  
  // Notificar al reportero si el comentario es de otra persona
  if (ticket.ReportaEmail && ticket.ReportaEmail !== autorEmail) {
    crearNotificacion(
      ticket.ReportaEmail,
      'comentario',
      'Nuevo comentario',
      `Se ha agregado un comentario a tu ticket #${folio}.`,
      ticketId
    );
  }
  
  // Notificar al agente asignado si el comentario es del usuario
  if (ticket.AsignadoA && ticket.AsignadoA !== autorEmail) {
    crearNotificacion(
      ticket.AsignadoA,
      'comentario',
      'Nuevo comentario',
      `El usuario ha comentado en el ticket #${folio}.`,
      ticketId
    );
  }
}

/**
 * Notificar vencimiento de SLA
 * @param {Object} ticket - Datos del ticket
 */
function notificarVencimiento(ticket) {
  const ticketId = ticket.ID || ticket.Id || ticket.id;
  const folio = ticket.Folio;
  
  // Notificar al agente
  if (ticket.AsignadoA) {
    crearNotificacion(
      ticket.AsignadoA,
      'vencimiento',
      '⚠️ SLA vencido',
      `El ticket #${folio} ha excedido su tiempo de SLA. Requiere atención urgente.`,
      ticketId
    );
  }
  
  // Notificar a admins
  const admins = getAdmins();
  admins.forEach(admin => {
    crearNotificacion(
      admin.email,
      'vencimiento',
      '⚠️ SLA vencido',
      `El ticket #${folio} asignado a ${ticket.AsignadoA || 'nadie'} ha vencido.`,
      ticketId
    );
  });
}


function verificarVencimientos() {
  var LOG = '[verificarVencimientos]';
  var ahora = new Date();

  // =======================================================
  // NUEVO: VALIDACIÓN DE HORARIO LABORAL ESTRICTO
  // =======================================================
  // 1. Validar que sea día laboral (Lunes a Viernes, no festivo)
  if (typeof esDiaLaboral === 'function' && !esDiaLaboral(ahora)) {
    Logger.log(LOG + ' Día no laboral, abortando envío de alertas.');
    return;
  }

  // 2. Validar que la hora actual esté dentro del horario: 8-14 o 16-18
  var h = ahora.getHours();
  // Si son antes de las 8am, o entre 2pm y 3:59pm, o después de las 6pm... abortar.
  if (h < 8 || (h >= 14 && h < 16) || h >= 18) {
    Logger.log(LOG + ' Fuera de horario laboral (8-14, 16-18), abortando envío de alertas.');
    return;
  }
  // =======================================================
  
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(DB.TICKETS);
    var _data = _readTableByHeader_(DB.TICKETS);
    var headers = _data.headers;
    var rows = _data.rows;
    var m = _headerMap_(headers);
    var ahora = new Date();
    var notificados = 0;
    var yaVencidos = 0;

    // Necesitamos columna VencimientoNotificado. Si no existe, crearla.
    if (m.VencimientoNotificado == null) {
      var nuevaCol = headers.length + 1;
      sh.getRange(1, nuevaCol).setValue('VencimientoNotificado');
      headers.push('VencimientoNotificado');
      m = _headerMap_(headers);
      rows.forEach(function(r) { r.push(''); });
      Logger.log(LOG + ' Columna VencimientoNotificado creada');
    }

    rows.forEach(function(r, idx) {
      var estatus = String(r[m.Estatus] || '').toLowerCase().trim();
      
      // Solo tickets activos (no cerrados/resueltos/cancelados)
      if (['cerrado', 'resuelto', 'cancelado'].indexOf(estatus) >= 0) return;
      
      var vencimiento = r[m.Vencimiento];
      if (!vencimiento) return;
      
      var fechaVenc = new Date(vencimiento);
      if (isNaN(fechaVenc.getTime())) return;
      
      var diffHoras = (fechaVenc - ahora) / (1000 * 60 * 60);
      var yaNotificado = String(r[m.VencimientoNotificado] || '').toLowerCase();
      
      var folio = r[m.Folio] || '';
      var titulo = r[m['Título']] || 'Sin título';
      var area = r[m['Área']] || '';
      var ubicacion = r[m['Ubicación']] || '';
      var prioridad = r[m.Prioridad] || '';
      var asignadoA = r[m.AsignadoA] || '';
      var ticketId = r[m.ID] || '';
      var reportaNombre = r[m.ReportaNombre] || '';

      var datosNotif = {
        ticketId: ticketId,
        folio: folio,
        titulo: titulo,
        area: area,
        ubicacion: ubicacion,
        prioridad: prioridad,
        reportaNombre: reportaNombre,
        vencimiento: vencimiento,
        horasRestantes: Math.max(0, Math.round(diffHoras * 10) / 10),
        tituloNotif: '',
        mensajeNotif: ''
      };

      // ── CASO 1: Ticket YA VENCIÓ (diffHoras <= 0) ──
      if (diffHoras <= 0 && yaNotificado !== 'vencido') {
        yaVencidos++;
        datosNotif.tituloNotif = '🚨 SLA VENCIDO: Ticket #' + folio;
        datosNotif.mensajeNotif = 'El ticket ha excedido su fecha límite. Requiere atención URGENTE.';
        
        // Notificar al agente
        if (asignadoA) {
          notificarAgenteTodosCanales(asignadoA, 'vencido', datosNotif);
        }
        
        // Notificar al gerente del área
        try {
          var gerente = getGerenteDelArea(area);
          if (gerente && gerente.email) {
            notificarAgenteTodosCanales(gerente.email, 'vencido', datosNotif);
            notificarGerenteTelegram(area, 'SLA VENCIDO - Atención urgente', {
              folio: folio,
              titulo: titulo,
              area: area,
              prioridad: prioridad,
              motivo: 'El ticket ha superado su fecha límite de SLA'
            });
          }
        } catch (e) {
          Logger.log(LOG + ' Error notificando gerente: ' + e.message);
        }
        
        // Notificar a admins (in-system)
        try {
          var admins = getAdmins();
          admins.forEach(function(admin) {
            crearNotificacion(admin.email, 'vencimiento', '🚨 SLA Vencido', 
              'Ticket #' + folio + ' asignado a ' + (asignadoA || 'nadie') + ' ha vencido.', ticketId);
          });
        } catch (e) {}
        
        // Marcar como notificado para no repetir
        r[m.VencimientoNotificado] = 'vencido';
        sh.getRange(idx + 2, 1, 1, headers.length).setValues([r]);
        notificados++;
      }
      
      // ── CASO 2: Ticket POR VENCER (0 < diffHoras <= 4) ──
      else if (diffHoras > 0 && diffHoras <= 4 && yaNotificado !== 'porVencer' && yaNotificado !== 'vencido') {
        datosNotif.tituloNotif = '⏰ Próximo a vencer: Ticket #' + folio;
        datosNotif.mensajeNotif = 'Vence en ' + datosNotif.horasRestantes + ' horas. Prioriza su atención.';
        
        // Notificar al agente
        if (asignadoA) {
          notificarAgenteTodosCanales(asignadoA, 'vencimiento', datosNotif);
        }
        
        // Notificar al gerente solo por sistema (no saturar con pre-vencimientos)
        try {
          var gerente2 = getGerenteDelArea(area);
          if (gerente2 && gerente2.email) {
            crearNotificacion(gerente2.email, 'vencimiento', '⏰ Ticket por vencer',
              'Ticket #' + folio + ' vence en ' + datosNotif.horasRestantes + ' horas.', ticketId);
          }
        } catch (e) {}
        
        // Marcar
        r[m.VencimientoNotificado] = 'porVencer';
        sh.getRange(idx + 2, 1, 1, headers.length).setValues([r]);
        notificados++;
      }
    });
    
    clearCache(DB.TICKETS);
    Logger.log(LOG + ' Completado. Notificados: ' + notificados + ' | Ya vencidos: ' + yaVencidos);
    
  } catch (e) {
    Logger.log(LOG + ' ERROR: ' + e.message + '\n' + e.stack);
  }
}

// ============================================================================
// TRIGGER: VERIFICAR Y REACTIVAR AGENTES CUYA AUSENCIA HA TERMINADO
// Ejecutar diariamente
// ============================================================================

/**
 * Verificar agentes cuya fecha de fin de ausencia ha pasado
 * y reactivarlos automáticamente
 */
function verificarFinAusenciasAgentes() {
  try {
    const sh = getSheet(DB.USERS);
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    
    if (m.Disponible == null || m.FechaFinAusencia == null) return;
    
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    
    let reactivados = 0;
    
    rows.forEach((row, idx) => {
      const disponible = row[m.Disponible];
      const fechaFin = row[m.FechaFinAusencia];
      
      // Si no está disponible y tiene fecha de fin
      if ((disponible === false || String(disponible).toLowerCase() === 'false') && fechaFin) {
        const fechaFinDate = new Date(fechaFin);
        if (!isNaN(fechaFinDate.getTime()) && fechaFinDate <= hoy) {
          // Reactivar agente
          row[m.Disponible] = true;
          row[m.MotivoAusencia] = '';
          row[m.FechaInicioAusencia] = '';
          row[m.FechaFinAusencia] = '';
          
          sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
          
          // Notificar al agente
          notifyUser(row[m.Email], 'info', '¡Bienvenido de vuelta!',
            'Tu período de ausencia ha terminado y has sido marcado como disponible automáticamente.',
            {});
          
          reactivados++;
        }
      }
    });
    
    if (reactivados > 0) {
      clearCache(DB.USERS);
      Logger.log(`Agentes reactivados automáticamente: ${reactivados}`);
    }
    
    return { ok: true, reactivados };
    
  } catch (e) {
    Logger.log('Error en verificarFinAusenciasAgentes: ' + e.message);
    return { ok: false, error: e.message };
  }
}

// ============================================================================
// FUNCIONES DE DISPONIBILIDAD DE AGENTES
// ============================================================================

/**
 * Obtener estado de disponibilidad de un agente
 * @param {string} email - Email del agente
 * @returns {Object} Estado de disponibilidad
 */
function getDisponibilidadAgente(email) {
  if (!email) return { disponible: true };
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    
    const row = rows.find(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
    if (!row) return { disponible: true };
    
    // Si no existe la columna, asumir disponible
    if (m.Disponible == null) return { disponible: true };
    
    const disponible = row[m.Disponible] !== false && 
                       String(row[m.Disponible]).toLowerCase() !== 'false' &&
                       String(row[m.Disponible]).toLowerCase() !== 'no';
    
    return {
      disponible,
      motivo: m.MotivoAusencia != null ? row[m.MotivoAusencia] : '',
      fechaInicio: m.FechaInicioAusencia != null ? row[m.FechaInicioAusencia] : '',
      fechaFin: m.FechaFinAusencia != null ? row[m.FechaFinAusencia] : ''
    };
  } catch (e) {
    Logger.log('Error en getDisponibilidadAgente: ' + e.message);
    return { disponible: true };
  }
}

/**
 * Cambiar disponibilidad de un agente
 * @param {string} email - Email del agente
 * @param {boolean} disponible - true = disponible, false = no disponible
 * @param {string} motivo - Motivo de la ausencia (Vacaciones, Incapacidad, Permiso)
 * @param {string} fechaFin - Fecha estimada de regreso (opcional)
 * @returns {Object} Resultado de la operación
 */
function cambiarDisponibilidadAgente(email, disponible, motivo, fechaFin, fechaInicioVac) {
  if (!email) return { ok: false, error: 'Email requerido' };

  return withLock_(() => {
    try {
      const sh = getSheet(DB.USERS);
      const { headers, rows } = _readTableByHeader_(DB.USERS);
      const m = _headerMap_(headers);

      const idx = rows.findIndex(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
      if (idx < 0) return { ok: false, error: 'Usuario no encontrado' };

      const row = rows[idx];
      const nombreAgente = row[m.Nombre] || email;
      const rolAgente = row[m.Rol] || '';

      if (!rolAgente.toLowerCase().includes('agente') && rolAgente.toLowerCase() !== 'admin') {
        return { ok: false, error: 'Solo los agentes pueden cambiar su disponibilidad' };
      }

      // Asegurar columnas — igual que antes + FechaInicioVacaciones
      const colsRequeridas = ['Disponible','MotivoAusencia','FechaInicioAusencia','FechaFinAusencia','FechaInicioVacaciones'];
      colsRequeridas.forEach(col => {
        if (m[col] == null) {
          sh.insertColumnAfter(headers.length);
          sh.getRange(1, headers.length + 1).setValue(col);
          m[col] = headers.length;
          headers.push(col);
        }
      });

      while (row.length < headers.length) row.push('');

      if (disponible) {
        row[m.Disponible]             = true;
        row[m.MotivoAusencia]         = '';
        row[m.FechaInicioAusencia]    = '';
        row[m.FechaFinAusencia]       = '';
        row[m.FechaInicioVacaciones]  = '';
      } else {
        row[m.Disponible]             = false;
        row[m.MotivoAusencia]         = motivo || 'Sin especificar';
        row[m.FechaInicioAusencia]    = motivo === 'Vacaciones' && fechaInicioVac ? new Date(fechaInicioVac) : new Date();
        row[m.FechaFinAusencia]       = fechaFin ? new Date(fechaFin) : '';
        row[m.FechaInicioVacaciones]  = motivo === 'Vacaciones' && fechaInicioVac ? new Date(fechaInicioVac) : '';
      }

      sh.getRange(idx + 2, 1, 1, row.length).setValues([row]);
      clearCache(DB.USERS);

      let ticketsReasignados = 0, alertasEnviadas = 0;
      if (!disponible) {
        const res = manejarAusenciaAgente(email, nombreAgente, motivo, fechaFin);
        ticketsReasignados = res.ticketsReasignados || 0;
        alertasEnviadas    = res.alertasEnviadas    || 0;
      }

      const periodoMsg = (motivo === 'Vacaciones' && fechaInicioVac && fechaFin)
        ? ` (${fechaInicioVac} → ${fechaFin})`
        : (fechaFin ? ` hasta ${fechaFin}` : '');

      return {
        ok: true, disponible, ticketsReasignados, alertasEnviadas,
        mensaje: disponible
          ? 'Has marcado tu disponibilidad como activa'
          : `Ausencia por "${motivo}"${periodoMsg} registrada. ${ticketsReasignados} tickets reasignados.`
      };

    } catch (e) {
      Logger.log('Error en cambiarDisponibilidadAgente: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

/**
 * Revisar agentes con fecha de regreso cumplida y reactivarlos.
 * Configurar como trigger: Time-driven → Day timer → 7am-8am
 */
function autoReactivarAgentesVacaciones() {
  try {
    const sh = getSheet(DB.USERS);
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);

    if (m.Disponible == null || m.FechaFinAusencia == null) return;

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    let reactivados = 0;

    rows.forEach((row, i) => {
      const disponible = row[m.Disponible];
      if (disponible === true || String(disponible).toLowerCase() === 'true') return;

      const fechaFinVal = row[m.FechaFinAusencia];
      if (!fechaFinVal) return;

      const fechaFin = new Date(fechaFinVal);
      fechaFin.setHours(0, 0, 0, 0);
      if (fechaFin > hoy) return; // aún no es el día de regreso

      // Reactivar
      row[m.Disponible]            = true;
      row[m.MotivoAusencia]        = '';
      row[m.FechaInicioAusencia]   = '';
      row[m.FechaFinAusencia]      = '';
      if (m.FechaInicioVacaciones != null) row[m.FechaInicioVacaciones] = '';

      sh.getRange(i + 2, 1, 1, row.length).setValues([row]);

      const emailAgente = row[m.Email];
      const nombre = row[m.Nombre] || emailAgente;
      Logger.log(`✅ Auto-reactivado: ${nombre} (${emailAgente})`);

      // Notificar al agente
      try {
        crearNotificacion(emailAgente, 'info', '¡Bienvenido de vuelta!',
          'Tu ausencia venció hoy. Ahora estás marcado como disponible.', '');
        const chatId = getTelegramChatIdGrupo(emailAgente);
        if (chatId) telegramSendToGrupo(chatId, `✅ <b>¡Bienvenido de vuelta!</b>\nTu período de ausencia terminó hoy. Ya estás marcado como <b>Disponible</b>.`);
      } catch(e) {}

      reactivados++;
    });

    if (reactivados > 0) clearCache(DB.USERS);
    Logger.log(`[autoReactivar] ${reactivados} agente(s) reactivados.`);
  } catch (e) {
    Logger.log('Error en autoReactivarAgentesVacaciones: ' + e.message);
  }
}

function manejarAusenciaAgente(emailAgente, nombreAgente, motivo, fechaFin) {
  try {
    const { headers: ticketHeaders, rows: ticketRows } = _readTableByHeader_(DB.TICKETS);
    const tm = _headerMap_(ticketHeaders);

    // 1. Tickets activos del agente — comparar contra username (columna Email de Usuarios)
    const usernameLower = emailAgente.toLowerCase();
    const ticketsActivos = ticketRows.filter(r => {
      const asignado = String(r[tm.AsignadoA] || '').toLowerCase();
      const estatus  = String(r[tm.Estatus]   || '').toLowerCase();
      return asignado === usernameLower &&
             !['cerrado', 'resuelto', 'cancelado'].includes(estatus);
    });

    if (ticketsActivos.length === 0) {
      return { ticketsReasignados: 0, alertasEnviadas: 0 };
    }

    // 2. Obtener área del agente y gerente correspondiente
    const agenteInfo = getUser(emailAgente);
    const areaAgente = agenteInfo.area || '';
    const gerenteArea = getGerenteDelArea(areaAgente);

    // username  → se guarda en AsignadoA del ticket
    // emailReal → se usa para notificaciones (Email, Telegram, campanita)
    let nuevoAgenteUsername = '';
    let nuevoAgenteEmail    = '';  // email real para notificaciones
    let nuevoAgenteNombre   = '';

    if (gerenteArea && gerenteArea.username &&
        gerenteArea.username.toLowerCase() !== usernameLower) {
      nuevoAgenteUsername = gerenteArea.username;  // ej: RGNava
      nuevoAgenteEmail    = gerenteArea.email;     // ej: rgnava@bexalta.com
      nuevoAgenteNombre   = gerenteArea.nombre || gerenteArea.username;
    } else {
      // Fallback: admin disponible con menos carga
      const admins = getAdmins(); // retorna { email: username, nombre }
      if (admins.length > 0) {
        nuevoAgenteUsername = admins[0].email;
        nuevoAgenteEmail    = _resolverEmailUsuario_(admins[0].email);
        nuevoAgenteNombre   = admins[0].nombre || 'Administrador';
      }
    }

    if (!nuevoAgenteUsername) {
      Logger.log('manejarAusenciaAgente: sin destinatario para reasignar.');
      return { ticketsReasignados: 0, alertasEnviadas: 0 };
    }

    let ticketsReasignados = 0;
    let alertasEnviadas    = 0;
    const sh = getSheet(DB.TICKETS);
    const listaTicketsTelegram = [];

    // 3. Reasignar cada ticket
    ticketsActivos.forEach(ticketRow => {
      const ticketId   = ticketRow[tm.ID];
      const folio      = ticketRow[tm.Folio];
      const titulo     = ticketRow[tm['Título']] || '';
      const reportaEmail = String(ticketRow[tm.ReportaEmail] || '').trim();

      const idx = ticketRows.indexOf(ticketRow);
      if (idx < 0) return;

      // A. Guardar username en AsignadoA (no email)
      ticketRow[tm.AsignadoA]              = nuevoAgenteUsername;
      ticketRow[tm['ÚltimaActualización']] = new Date();
      sh.getRange(idx + 2, 1, 1, ticketHeaders.length).setValues([ticketRow]);

      registrarBitacora(
        ticketId,
        'Reasignación por ausencia',
        `De ${nombreAgente} a ${nuevoAgenteNombre} (Gerencia). Motivo: ${motivo}`
      );

      listaTicketsTelegram.push(`• #<b>${folio}</b>: ${titulo}`);

      // B. Notificar al usuario que reportó
      if (reportaEmail) {
        crearNotificacion(
          reportaEmail, 'asignacion', 'Cambio de Agente',
          `Tu ticket #${folio} fue asignado temporalmente a Gerencia por ausencia de tu agente.`,
          ticketId
        );

        try {
          enviarEmailNotificacion(reportaEmail,
            `Actualización Ticket #${folio} - Cambio de Agente`,
            `<h2>🔄 Actualización en tu Ticket</h2>
             <div style="background:#f8fafc;padding:15px;border-left:4px solid #3b82f6;margin-bottom:20px;">
               <p style="margin:0;"><strong>Ticket #${folio}</strong>: ${titulo}</p>
             </div>
             <p>Tu agente asignado (<strong>${nombreAgente}</strong>) se encuentra ausente.</p>
             <p>El ticket fue derivado a Gerencia, quien lo reasignará a la brevedad.</p>`
          );
        } catch(e) { Logger.log('Email usuario ausencia: ' + e.message); }

        try {
          const chatUser = getTelegramChatIdGrupo(reportaEmail);
          if (chatUser) telegramSendToGrupo(chatUser,
            `🔄 <b>Cambio de Agente</b>\nTu ticket <code>#${folio}</code> fue derivado a Gerencia ` +
            `por ausencia de tu agente. Pronto te reasignarán.`
          );
        } catch(e) {}
      }

      ticketsReasignados++;
    });

    clearCache(DB.TICKETS);

    // 4. Notificar al gerente/admin que recibe los tickets
    // Todas las notificaciones usan nuevoAgenteEmail (email real)
    if (ticketsReasignados > 0 && nuevoAgenteEmail) {

      const msgTelegram =
        `⚠️ <b>AGENTE AUSENTE</b>\n\n` +
        `👤 <b>${nombreAgente}</b> marcó su estado como <b>No Disponible</b>.\n` +
        `📝 <b>Motivo:</b> ${motivo}\n` +
        (fechaFin ? `📅 <b>Regreso est.:</b> ${fechaFin}\n\n` : '\n') +
        `🔄 Se te asignaron <b>${ticketsReasignados} tickets</b> para que los reasignes o atiendas:\n\n` +
        listaTicketsTelegram.join('\n');

      try {
        const chatIdGerente = getTelegramChatIdGrupo(nuevoAgenteEmail);
        if (chatIdGerente) telegramSendToGrupo(chatIdGerente, msgTelegram);
      } catch(e) { Logger.log('Telegram gerente ausencia: ' + e.message); }

      try {
        const listaHtml = listaTicketsTelegram
          .map(t => t.replace(/<b>/g, '<strong>').replace(/<\/b>/g, '</strong>'))
          .join('<br>');

        enviarEmailNotificacion(
          nuevoAgenteEmail,
          `⚠️ Ausencia: ${nombreAgente} — ${ticketsReasignados} tickets por reasignar`,
          `<h2>⚠️ Alerta de Ausencia de Agente</h2>
           <div style="background:#fffbeb;border-left:4px solid #f59e0b;padding:15px;margin-bottom:20px;">
             <p style="margin:0;">El agente <strong>${nombreAgente}</strong> se marcó como No Disponible.</p>
             <p style="margin:5px 0 0;"><strong>Motivo:</strong> ${motivo}</p>
             ${fechaFin ? `<p style="margin:5px 0 0;"><strong>Regreso estimado:</strong> ${fechaFin}</p>` : ''}
           </div>
           <p>Los siguientes <strong>${ticketsReasignados} tickets activos</strong> han sido derivados 
              a tu bandeja para que los reasignes o atiendas:</p>
           <div style="background:#f8fafc;padding:15px;border-radius:8px;line-height:1.8;">
             ${listaHtml}
           </div>`
        );
      } catch(e) { Logger.log('Email gerente ausencia: ' + e.message); }

      try {
        crearNotificacion(
          nuevoAgenteEmail, 'alerta',
          `Ausencia de ${nombreAgente}`,
          `Tienes ${ticketsReasignados} tickets pendientes por reasignar.`,
          ''
        );
      } catch(e) {}

      alertasEnviadas++;
    }

    return { ticketsReasignados, alertasEnviadas };

  } catch(e) {
    Logger.log('Error crítico en manejarAusenciaAgente: ' + e.message);
    return { ticketsReasignados: 0, alertasEnviadas: 0 };
  }
}

/**
 * Obtener lista de administradores
 * @returns {Array} Lista de admins con email y nombre
 */
function getAdmins() {
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);
  
  return rows
    .filter(r => String(r[m.Rol] || '').toLowerCase() === 'admin')
    .map(r => ({
      email: r[m.Email],
      nombre: r[m.Nombre] || r[m.Email]
    }));
}

/**
 * Obtener agentes de un área específica
 * @param {string} area - Nombre del área
 * @returns {Array} Lista de agentes
 */
function getAgentesPorArea(area) {
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);
  const areaLower = (area || '').toLowerCase();
  
  return rows
    .filter(r => {
      const rol = String(r[m.Rol] || '').toLowerCase();
      const userArea = String(r[m['Área']] || '').toLowerCase();
      return (rol.includes('agente') || rol === 'admin') && 
             (userArea === areaLower || rol === 'admin');
    })
    .map(r => ({
      email: r[m.Email],
      nombre: r[m.Nombre] || r[m.Email],
      rol: r[m.Rol],
      area: r[m['Área']] || '',
      ubicacion: r[m['Ubicación']] || ''
    }));
}

/**
 * Obtener agentes DISPONIBLES de un área, ordenados por carga de trabajo
 * @param {string} area - Nombre del área
 * @param {string} excluirEmail - Email a excluir (el agente ausente)
 * @returns {Array} Lista de agentes disponibles ordenados por carga
 */
function getAgentesDisponiblesPorArea(area, excluirEmail) {
  const { headers: userHeaders, rows: userRows } = _readTableByHeader_(DB.USERS);
  const um = _headerMap_(userHeaders);
  const { headers: ticketHeaders, rows: ticketRows } = _readTableByHeader_(DB.TICKETS);
  const tm = _headerMap_(ticketHeaders);
  
  const areaLower = (area || '').toLowerCase();
  const excluirLower = (excluirEmail || '').toLowerCase();
  
  // Filtrar agentes del área que estén disponibles
  const agentes = userRows.filter(r => {
    const email = String(r[um.Email] || '').toLowerCase();
    const rol = String(r[um.Rol] || '').toLowerCase();
    const userArea = String(r[um['Área']] || '').toLowerCase();
    
    // Excluir al agente especificado
    if (email === excluirLower) return false;
    
    // Verificar rol
    if (!rol.includes('agente') && rol !== 'admin') return false;
    
    // Verificar área (admin puede atender cualquier área)
    if (rol !== 'admin' && userArea !== areaLower) return false;
    
    // Verificar disponibilidad
    if (um.Disponible != null) {
      const disponible = r[um.Disponible];
      if (disponible === false || 
          String(disponible).toLowerCase() === 'false' ||
          String(disponible).toLowerCase() === 'no') {
        return false;
      }
    }
    
    return true;
  }).map(r => ({
    email: r[um.Email],
    nombre: r[um.Nombre] || r[um.Email],
    rol: r[um.Rol],
    area: r[um['Área']] || '',
    ubicacion: r[um['Ubicación']] || ''
  }));
  
  // Contar tickets activos por agente
  const cargaPorAgente = {};
  agentes.forEach(a => cargaPorAgente[a.email.toLowerCase()] = 0);
  
  ticketRows.forEach(r => {
    const asignado = String(r[tm.AsignadoA] || '').toLowerCase();
    const estatus = String(r[tm.Estatus] || '').toLowerCase();
    if (cargaPorAgente[asignado] != null && !['cerrado', 'resuelto'].includes(estatus)) {
      cargaPorAgente[asignado]++;
    }
  });
  
  // Ordenar por carga (menos tickets primero)
  agentes.sort((a, b) => {
    const cargaA = cargaPorAgente[a.email.toLowerCase()] || 0;
    const cargaB = cargaPorAgente[b.email.toLowerCase()] || 0;
    return cargaA - cargaB;
  });
  
  return agentes;
}

// ============================================================================
// ASIGNACIÓN AUTOMÁTICA POR UBICACIÓN (MEJORADA)
// Reemplazar la función asignarAgenteEquilibrado existente
// ============================================================================

/**
 * asignarAgenteEquilibrado corregido:
 * - Fallback es el gerente del área (username), no el primer admin
 * - Verifica disponibilidad del gerente antes de asignarlo
 */
function asignarAgenteEquilibrado(area, ubicacionTicket, categoria) {
  const { headers: userHeaders, rows: userRows } = _readTableByHeader_(DB.USERS);
  const um = _headerMap_(userHeaders);
  const { headers: ticketHeaders, rows: ticketRows } = _readTableByHeader_(DB.TICKETS);
  const tm = _headerMap_(ticketHeaders);

  const areaLower      = (area || '').toLowerCase().trim();
  const ubicacionLower = (ubicacionTicket || '').toLowerCase().trim();
  const categoriaLower = (categoria || '').toLowerCase().trim();

  Logger.log(`🔍 Buscando agente: Área="${area}", Ubicación="${ubicacionTicket}", Categoría="${categoria}"`);

  // ── Helper: verificar disponibilidad de un username en userRows ──────────
  function esDisponible(username) {
    if (!username) return false;
    const usernameLower = username.toLowerCase();
    const row = userRows.find(r => String(r[um.Email] || '').toLowerCase() === usernameLower);
    if (!row) return false;
    if (um.Disponible == null) return true;
    const d = row[um.Disponible];
    return d !== false && String(d).toLowerCase() !== 'false' && String(d).toLowerCase() !== 'no';
  }

  // ── PRIORIDAD 0: Agente específico de categoría ──────────────────────────
  if (categoriaLower && area) {
    const agenteCategoria = buscarAgenteEnCategoria(area, categoriaLower, ubicacionLower);
    if (agenteCategoria && esDisponible(agenteCategoria)) {
      Logger.log(`✅ Por categoría: ${agenteCategoria}`);
      return agenteCategoria;
    }
    if (agenteCategoria) Logger.log(`⚠️ Agente de categoría no disponible: ${agenteCategoria}`);
  }

  // ── Determinar rol requerido ─────────────────────────────────────────────
  let rolRequerido = '';
  if (['sistemas','ti','tecnología'].includes(areaLower))        rolRequerido = 'agente_sistemas';
  else if (['mantenimiento','mtto'].includes(areaLower))         rolRequerido = 'agente_mantenimiento';

  // ── Filtrar agentes disponibles del área ─────────────────────────────────
  const agentesDelArea = userRows.filter(r => {
    const rol = String(r[um.Rol] || '').toLowerCase().trim();
    if (rolRequerido && rol !== rolRequerido) return false;
    if (!rolRequerido && !['agente_sistemas','agente_mantenimiento'].includes(rol)) return false;

    // Disponibilidad
    if (um.Disponible != null) {
      const d = r[um.Disponible];
      if (d === false || String(d).toLowerCase() === 'false' || String(d).toLowerCase() === 'no') {
        Logger.log(`   ⏸️ No disponible: ${r[um.Email]}`);
        return false;
      }
    }
    return true;
  }).map(r => ({
    username:    String(r[um.Email] || '').trim(),   // username para AsignadoA
    nombre:      r[um.Nombre] || '',
    rol:         r[um.Rol]    || '',
    ubicaciones: String(r[um['Ubicación']] || '').toLowerCase().split(',').map(u => u.trim()).filter(Boolean)
  }));

  Logger.log(`📊 Agentes disponibles en "${area}": ${agentesDelArea.length}`);

  // ── Contar tickets activos por username ──────────────────────────────────
  const carga = {};
  agentesDelArea.forEach(a => carga[a.username] = 0);
  ticketRows.forEach(r => {
    const asig   = String(r[tm.AsignadoA] || '').trim();
    const estatus = String(r[tm.Estatus] || '').toLowerCase();
    if (carga[asig] != null && !['cerrado','resuelto'].includes(estatus)) carga[asig]++;
  });

  // ── PRIORIDAD 1: Área + Ubicación ────────────────────────────────────────
  if (ubicacionLower && agentesDelArea.length > 0) {
    const conUbicacion = agentesDelArea
      .filter(a => a.ubicaciones.includes(ubicacionLower))
      .sort((a, b) => carga[a.username] - carga[b.username]);

    if (conUbicacion.length > 0) {
      Logger.log(`✅ Por área+ubicación: ${conUbicacion[0].username} (${carga[conUbicacion[0].username]} tickets)`);
      return conUbicacion[0].username;
    }
    Logger.log(`⚠️ Sin agentes con ubicación "${ubicacionTicket}"`);
  }

  // ── PRIORIDAD 2: Área + menor carga ──────────────────────────────────────
  if (agentesDelArea.length > 0) {
    const elegido = agentesDelArea.sort((a, b) => carga[a.username] - carga[b.username])[0];
    Logger.log(`✅ Por menor carga: ${elegido.username} (${carga[elegido.username]} tickets)`);
    return elegido.username;
  }

  // ── PRIORIDAD 3: Gerente del área (no primer admin) ───────────────────────
  Logger.log(`⚠️ Sin agentes disponibles en "${area}", buscando gerente...`);
  const gerente = getGerenteDelArea(area);
  if (gerente && gerente.username) {
    if (esDisponible(gerente.username)) {
      Logger.log(`✅ Asignado a GERENTE del área: ${gerente.username}`);
      return gerente.username;
    }
    Logger.log(`⚠️ Gerente ${gerente.username} tampoco disponible, buscando admin...`);
  }

  // ── PRIORIDAD 4: Admin disponible con menos carga (último recurso) ────────
  const admins = userRows
    .filter(r => {
      const rol = String(r[um.Rol] || '').toLowerCase();
      if (rol !== 'admin') return false;
      if (um.Disponible != null) {
        const d = r[um.Disponible];
        if (d === false || String(d).toLowerCase() === 'false' || String(d).toLowerCase() === 'no') return false;
      }
      return true;
    })
    .map(r => ({ username: String(r[um.Email] || '').trim(), nombre: r[um.Nombre] || '' }));

  if (admins.length > 0) {
    admins.forEach(a => {
      if (carga[a.username] == null) {
        carga[a.username] = ticketRows.filter(r =>
          String(r[tm.AsignadoA] || '').trim() === a.username &&
          !['cerrado','resuelto'].includes(String(r[tm.Estatus] || '').toLowerCase())
        ).length;
      }
    });
    admins.sort((a, b) => (carga[a.username] || 0) - (carga[b.username] || 0));
    Logger.log(`✅ Asignado a ADMIN: ${admins[0].username}`);
    return admins[0].username;
  }

  Logger.log(`❌ Sin agente disponible para asignar`);
  return '';
}



/**
 * Buscar agente asignado en la tabla de categorías
 */
function buscarAgenteEnCategoria(area, categoria, ubicacion) {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheetName = area.toLowerCase() === 'sistemas' ? 'Categorias_TI' : 'Categorias_Mtto';
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet || sheet.getLastRow() < 2) return null;
    
    // Leer todas las categorías (6 columnas)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    
    // Buscar la categoría que coincida
    for (const row of data) {
      const [nombre, areaRow, ubicRaw, prio, sla, agenteAsignado] = row.map(x => String(x ?? '').trim());
      
      // Comparar nombre de categoría
      if (nombre.toLowerCase() !== categoria) continue;
      
      // Verificar ubicación si está especificada en la categoría
      if (ubicRaw) {
        const ubicaciones = ubicRaw.toLowerCase().split(',').map(u => u.trim());
        // Si la categoría tiene ubicaciones específicas y el ticket tiene ubicación,
        // verificar que coincida
        if (ubicacion && ubicaciones.length > 0 && !ubicaciones.includes(ubicacion)) {
          continue;
        }
      }
      
      // Retornar el agente si existe
      if (agenteAsignado) {
        Logger.log(`📋 Categoría "${nombre}" tiene agente asignado: ${agenteAsignado}`);
        return agenteAsignado;
      }
    }
    
    return null;
  } catch (e) {
    Logger.log('Error buscando agente en categoría: ' + e.message);
    return null;
  }
}


/**
 * Verificar que un agente existe y está disponible
 */
function verificarAgenteDisponible(emailAgente, userRows, um) {
  if (!emailAgente) return false;
  
  const emailLower = emailAgente.toLowerCase().trim();
  const agente = userRows.find(r => 
    String(r[um.Email] || '').toLowerCase() === emailLower
  );
  
  if (!agente) {
    Logger.log(`⚠️ Agente "${emailAgente}" no encontrado en usuarios`);
    return false;
  }
  
  // Verificar disponibilidad
  if (um.Disponible != null) {
    const disponible = agente[um.Disponible];
    if (disponible === false || 
        String(disponible).toLowerCase() === 'false' ||
        String(disponible).toLowerCase() === 'no') {
      Logger.log(`⚠️ Agente "${emailAgente}" no está disponible`);
      return false;
    }
  }
  
  return true;
}


function getAgentesParaReasignar(areaTicket, ubicacionTicket) {
  var _udata = _readTableByHeader_(DB.USERS);
  var userHeaders = _udata.headers;
  var userRows = _udata.rows;
  var um = _headerMap_(userHeaders);
  
  var _tdata = _readTableByHeader_(DB.TICKETS);
  var ticketHeaders = _tdata.headers;
  var ticketRows = _tdata.rows;
  var tm = _headerMap_(ticketHeaders);
  
  var areaTicketLower = (areaTicket || '').toLowerCase().trim();
  var ubicacionLower = (ubicacionTicket || '').toLowerCase().trim();
  
  var agentesArea = [];  // Agentes que matchean el área del ticket
  var admins = [];       // Admins como fallback al final
  
  Logger.log('[getAgentesParaReasignar] Buscando para área: "' + areaTicket + '", ubicación: "' + ubicacionTicket + '"');
  
  userRows.forEach(function(r) {
    var rolRaw = String(r[um.Rol] || '').trim();
    var rolLower = rolRaw.toLowerCase();
    
    // Campo Área del usuario (puede ser "Sistemas", "Mantenimiento", etc.)
    var areaUsuario = String(r[um['Área']] || r[um.Area] || '').toLowerCase().trim();
    
    // Verificar si está activo
    var activo = String(r[um.Activo] || 'si').toLowerCase();
    if (activo === 'no' || activo === 'false' || activo === 'inactivo') return;
    
    // Clasificar: ¿es admin? ¿es agente? ¿es gerente?
    var esAdmin = rolLower === 'admin' || rolLower.includes('admin');
    var esAgente = rolLower.includes('agente');
    var esGerente = rolLower.includes('gerente');
    
    // Solo incluir roles relevantes (agentes, gerentes, admins)
    if (!esAdmin && !esAgente && !esGerente) return;
    
    // Si es admin → va al grupo de fallback
    if (esAdmin && !esAgente) {
      // Admin puro (sin rol de agente)
      admins.push(buildAgenteInfo(r, um, rolRaw, true));
      return;
    }
    
    // Para agentes/gerentes: verificar que coincidan con el área del ticket
    var coincideArea = false;
    
    if (areaTicketLower) {
      // Método 1: El ROL contiene el nombre del área
      // Ej: 'agente_sistemas' contiene 'sistemas', 'agente_mantenimiento' contiene 'mantenimiento'
      if (rolLower.includes(areaTicketLower)) {
        coincideArea = true;
      }
      
      // Método 2: El campo ÁREA del usuario coincide
      // Ej: Área = "Sistemas" y ticket es de "Sistemas"
      if (areaUsuario === areaTicketLower) {
        coincideArea = true;
      }
      
      // Método 3: El campo Área contiene el nombre (para áreas compuestas)
      // Ej: Área = "Sistemas,Mantenimiento" contiene "sistemas"
      if (areaUsuario.indexOf(areaTicketLower) >= 0) {
        coincideArea = true;
      }
    } else {
      // Si el ticket no tiene área, incluir todos los agentes
      coincideArea = true;
    }
    
    if (coincideArea) {
      agentesArea.push(buildAgenteInfo(r, um, rolRaw, false));
    }
  });
  
  // Contar tickets activos por agente para mostrar carga de trabajo
  var todosAgentes = agentesArea.concat(admins);
  todosAgentes.forEach(function(a) {
    a.ticketsActivos = ticketRows.filter(function(r) {
      var asignado = String(r[tm.AsignadoA] || '').toLowerCase().trim();
      var estatus = String(r[tm.Estatus] || '').toLowerCase().trim();
      return asignado === a.email.toLowerCase() && 
             ['cerrado', 'resuelto', 'cancelado'].indexOf(estatus) < 0;
    }).length;
  });
  
  // Ordenar agentes del área: Disponibles > Misma Ubicación > Menos Carga
  agentesArea.sort(function(a, b) {
    if (a.disponible !== b.disponible) return (b.disponible ? 1 : 0) - (a.disponible ? 1 : 0);
    if (a.tieneUbicacion !== b.tieneUbicacion) return (b.tieneUbicacion ? 1 : 0) - (a.tieneUbicacion ? 1 : 0);
    return a.ticketsActivos - b.ticketsActivos;
  });
  
  // Admins al final, ordenados por carga
  admins.sort(function(a, b) { return a.ticketsActivos - b.ticketsActivos; });
  
  // Resultado: Agentes del área PRIMERO, luego admins como fallback
  var resultado = agentesArea.concat(admins);
  
  Logger.log('[getAgentesParaReasignar] Área: ' + areaTicket + 
             ' | Agentes del área: ' + agentesArea.length + 
             ' | Admins (fallback): ' + admins.length +
             ' | Total: ' + resultado.length);
  
  // Log detallado para debug
  resultado.forEach(function(a) {
    Logger.log('  → ' + a.nombre + ' | Rol: ' + a.rol + ' | Admin: ' + a.esAdmin + ' | Disp: ' + a.disponible);
  });
  
  return resultado;
  
  // ── Función auxiliar para construir objeto de agente ──
  function buildAgenteInfo(r, um, rolRaw, esAdminFlag) {
    var disponible = true;
    var motivoAusencia = '';
    
    if (um.Disponible != null) {
      var disp = r[um.Disponible];
      disponible = disp !== false && 
                   String(disp).toLowerCase() !== 'false' &&
                   String(disp).toLowerCase() !== 'no';
      if (!disponible && um.MotivoAusencia != null) {
        motivoAusencia = r[um.MotivoAusencia] || 'Ausente';
      }
    }
    
    var ubicaciones = String(r[um['Ubicación']] || '').split(',').map(function(u) { return u.trim(); }).filter(Boolean);
    var tieneUbicacion = ubicacionLower && ubicaciones.map(function(u) { return u.toLowerCase(); }).indexOf(ubicacionLower) >= 0;
    
    return {
      email: r[um.Email],
      nombre: r[um.Nombre] || r[um.Email],
      rol: rolRaw,
      ubicaciones: ubicaciones,
      tieneUbicacion: tieneUbicacion,
      disponible: disponible,
      motivoAusencia: motivoAusencia,
      esAdmin: esAdminFlag
    };
  }
}

// ============================================================================
// FUNCIÓN PARA OBTENER INFORMACIÓN COMPLETA DEL AGENTE (incluye disponibilidad)
// ============================================================================

/**
 * Obtener información completa del usuario incluyendo disponibilidad
 * @param {string} email - Email del usuario
 * @returns {Object} Información completa
 */
function getUserInfoCompleto(email) {
  if (!email) return null;
  
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);
  const row = rows.find(r => String(r[m.Email] || '').toLowerCase() === email.toLowerCase());
  
  if (!row) return null;
  
  // Verificar disponibilidad
  let disponible = true;
  let motivoAusencia = '';
  let fechaInicioAusencia = '';
  let fechaFinAusencia = '';
  
  if (m.Disponible != null) {
    const disp = row[m.Disponible];
    disponible = disp !== false && 
                 String(disp).toLowerCase() !== 'false' &&
                 String(disp).toLowerCase() !== 'no';
    
    if (!disponible) {
      motivoAusencia = m.MotivoAusencia != null ? row[m.MotivoAusencia] : '';
      fechaInicioAusencia = m.FechaInicioAusencia != null ? row[m.FechaInicioAusencia] : '';
      fechaFinAusencia = m.FechaFinAusencia != null ? row[m.FechaFinAusencia] : '';
    }
  }
  
  return {
    email: row[m.Email],
    nombre: row[m.Nombre] || '',
    rol: row[m.Rol] || 'usuario',
    area: (m['Área'] != null ? row[m['Área']] : '') || '',
    ubicacion: (m['Ubicación'] != null ? row[m['Ubicación']] : '') || '',
    puesto: (m.Puesto != null ? row[m.Puesto] : '') || '',
    passwordHash: (m.PasswordHash != null ? row[m.PasswordHash] : '') || '',
    // Campos de disponibilidad
    disponible,
    motivoAusencia,
    fechaInicioAusencia,
    fechaFinAusencia
  };
}


// ============================================================================
// SISTEMA DE NOTIFICACIONES POR EMAIL - HELPDESK BEXALTA
// ============================================================================
/**
 * Enviar notificación por email con template profesional
 * MODIFICADO: Usa getEmailNotificacion() para obtener el email correcto
 */
function enviarEmailNotificacion(destinatario, asunto, cuerpoHtml, opciones = {}) {
  if (!destinatario) {
    Logger.log('⚠️ No se especificó destinatario para email');
    return false;
  }
  
  // =====================================================
  // NUEVO: Obtener email real de notificación
  // =====================================================
  const emailReal = getEmailNotificacion(destinatario);
  
  if (!emailReal || !esEmailValido(emailReal)) {
    Logger.log(`⚠️ No se pudo obtener email válido para: ${destinatario}`);
    return false;
  }
  
  try {
    const config = getConfig();
    const nombreSistema = config.nombreSistema || 'Bexalta HelpDesk';
    
    // Template base del email
    const htmlCompleto = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 0; background: #f4f6f9; }
    .container { max-width: 600px; margin: 20px auto; background: white; border-radius: 12px; overflow: hidden; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #1e3a5f 0%, #2d5a87 100%); color: white; padding: 25px 30px; }
    .header h1 { margin: 0; font-size: 22px; font-weight: 600; }
    .header p { margin: 5px 0 0; opacity: 0.9; font-size: 14px; }
    .content { padding: 30px; }
    .ticket-box { background: #f8fafc; border-left: 4px solid #3b82f6; padding: 15px 20px; margin: 20px 0; border-radius: 0 8px 8px 0; }
    .ticket-box h3 { margin: 0 0 10px; color: #1e3a5f; font-size: 16px; }
    .ticket-box p { margin: 5px 0; color: #64748b; font-size: 14px; }
    .info-row { display: flex; justify-content: space-between; padding: 10px 0; border-bottom: 1px solid #e2e8f0; }
    .info-row:last-child { border-bottom: none; }
    .info-label { color: #64748b; font-weight: 500; }
    .info-value { color: #1e293b; font-weight: 600; }
    .btn { display: inline-block; padding: 12px 24px; background: #3b82f6; color: white !important; text-decoration: none; border-radius: 8px; font-weight: 600; margin: 10px 5px 10px 0; }
    .btn-success { background: #22c55e; }
    .btn-danger { background: #ef4444; }
    .btn-warning { background: #f59e0b; }
    .footer { background: #f1f5f9; padding: 20px 30px; text-align: center; color: #64748b; font-size: 12px; }
    .priority-alta, .priority-crítica, .priority-urgente { color: #ef4444; font-weight: bold; }
    .priority-media { color: #f59e0b; font-weight: bold; }
    .priority-baja { color: #22c55e; font-weight: bold; }
    .alert-box { padding: 15px; border-radius: 8px; margin: 15px 0; }
    .alert-warning { background: #fef3c7; border: 1px solid #fcd34d; }
    .alert-success { background: #d1fae5; border: 1px solid #6ee7b7; }
    .alert-danger { background: #fee2e2; border: 1px solid #fca5a5; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>${nombreSistema}</h1>
      <p>Sistema de Mesa de Ayuda</p>
    </div>
    <div class="content">
      ${cuerpoHtml}
    </div>
    <div class="footer">
      <p>Este es un correo automático del sistema ${nombreSistema}.</p>
      <p>Por favor no responda directamente a este correo.</p>
    </div>
  </div>
</body>
</html>`;
    
    const mailOptions = {
      to: emailReal,
      subject: `[${nombreSistema}] ${asunto}`,
      htmlBody: htmlCompleto
    };
    
    if (opciones.cc) mailOptions.cc = getEmailNotificacion(opciones.cc);
    if (opciones.bcc) mailOptions.bcc = getEmailNotificacion(opciones.bcc);
    if (opciones.replyTo) mailOptions.replyTo = getEmailNotificacion(opciones.replyTo);
    
    MailApp.sendEmail(mailOptions);
    Logger.log(`✅ Email enviado a ${emailReal} (usuario: ${destinatario}): ${asunto}`);
    return true;
    
  } catch (e) {
    Logger.log(`❌ Error enviando email a ${emailReal}: ${e.message}`);
    return false;
  }
}


/**
 * Notificar al agente sobre nuevo ticket asignado (INCLUYE VISITA)
 */
function notificarAgenteNuevoTicket(emailAgente, ticket) {
  const { folio, titulo, area, ubicacion, prioridad, reportaNombre, descripcion, visitaFecha, visitaHora } = ticket;
  
  const prioClass = (prioridad || '').toLowerCase().includes('alta') || 
                    (prioridad || '').toLowerCase().includes('crítica') || 
                    (prioridad || '').toLowerCase().includes('urgente') ? 'priority-alta' : 
                    (prioridad || '').toLowerCase().includes('media') ? 'priority-media' : 'priority-baja';
  
  // Bloque HTML condicional para la visita
  const visitaHtml = visitaFecha ? `
    <div style="margin: 15px 0; padding: 12px; background-color: #e0f2fe; border-left: 4px solid #0284c7; border-radius: 4px;">
      <h4 style="margin: 0 0 5px 0; color: #0284c7; font-size: 14px;">📅 Visita Programada Solicitada</h4>
      <p style="margin: 0; color: #0c4a6e;">
        <strong>Fecha:</strong> ${visitaFecha} <br>
        <strong>Hora:</strong> ${visitaHora}
      </p>
    </div>
  ` : '';

  const cuerpo = `<h2>📋 Nuevo Ticket Asignado</h2>
    <p>Se te ha asignado un nuevo ticket que requiere tu atención.</p> 
    
    <div class="ticket-box">
      <h3>Ticket #${folio}</h3>
      <p><strong>${titulo || 'Sin título'}</strong></p>
    </div>

    ${visitaHtml}

    <div class="info-row">
      <span class="info-label">Área:</span>
      <span class="info-value">${area || 'N/A'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Ubicación:</span>
      <span class="info-value">${ubicacion || 'No especificada'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Prioridad:</span>
      <span class="info-value ${prioClass}">${prioridad || 'Media'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Reportado por:</span>
      <span class="info-value">${reportaNombre || 'Usuario'}</span>
    </div>
    ${descripcion ? `<div style="margin-top:20px;padding:15px;background:#f8fafc;border-radius:8px;"><strong>Descripción:</strong><br><span style="color:#475569;">${String(descripcion).substring(0, 500)}${String(descripcion).length > 500 ? '...' : ''}</span></div>` : ''}
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn">🔗 Abrir Sistema</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailAgente, `Nuevo Ticket #${folio} Asignado`, cuerpo);
}


/**
 * Notificar al usuario sobre cambio de estado de su ticket
 */
function notificarUsuarioCambioEstado(emailUsuario, ticket, nuevoEstado, comentario) {
  const { folio, titulo } = ticket;
  
  let mensaje = '';
  let icono = '📋';
  let alertClass = '';
  
  switch ((nuevoEstado || '').toLowerCase()) {
    case 'en proceso':
      icono = '🔄';
      mensaje = 'Tu ticket está siendo atendido por nuestro equipo de soporte.';
      alertClass = 'alert-success';
      break;
    case 'resuelto':
      icono = '✅';
      mensaje = 'Tu ticket ha sido marcado como resuelto. Por favor verifica que el problema haya sido solucionado correctamente.';
      alertClass = 'alert-success';
      break;
    case 'cerrado':
      icono = '🔒';
      mensaje = 'Tu ticket ha sido cerrado. Gracias por usar el sistema de soporte.';
      alertClass = 'alert-success';
      break;
    case 'en espera':
      icono = '⏳';
      mensaje = 'Tu ticket está en espera. Puede requerir información adicional o recursos externos.';
      alertClass = 'alert-warning';
      break;
    case 'escalado':
      icono = '⚠️';
      mensaje = 'Tu ticket ha sido escalado a un nivel superior para su atención prioritaria.';
      alertClass = 'alert-warning';
      break;
    case 'en cotización':
      icono = '💰';
      mensaje = 'Tu ticket está en proceso de cotización. Te notificaremos cuando se requiera aprobación.';
      alertClass = 'alert-warning';
      break;
    case 'reabierto':
      icono = '🔄';
      mensaje = 'Tu ticket ha sido reabierto y será atendido nuevamente.';
      alertClass = 'alert-warning';
      break;
    default:
      mensaje = `El estado de tu ticket ha cambiado a: <strong>${nuevoEstado}</strong>`;
      alertClass = '';
  }
  
  const cuerpo = `
    <h2>${icono} Actualización de Ticket</h2>
    
    <div class="ticket-box">
      <h3>Ticket #${folio}</h3>
      <p><strong>${titulo || 'Sin título'}</strong></p>
      <p>Nuevo estado: <strong style="color:#3b82f6;">${nuevoEstado}</strong></p>
    </div>
    
    <div class="alert-box ${alertClass}" style="margin:20px 0;">
      <p style="margin:0;">${mensaje}</p>
    </div>
    
    ${comentario ? `<div style="margin-top:20px;padding:15px;background:#f0f9ff;border-radius:8px;border-left:4px solid #3b82f6;"><strong>💬 Comentario del agente:</strong><br><span style="color:#475569;">${comentario}</span></div>` : ''}
    
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn">🔗 Ver Ticket</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailUsuario, `Ticket #${folio} - ${nuevoEstado}`, cuerpo);
}


/**
 * Notificar al usuario sobre aprobación/rechazo de cotización
 */
function notificarUsuarioAprobacion(emailUsuario, ticket, aprobado, comentario) {
  const { folio, titulo, presupuesto } = ticket;
  
  const icono = aprobado ? '✅' : '❌';
  const estado = aprobado ? 'APROBADA' : 'RECHAZADA';
  const colorBtn = aprobado ? 'btn-success' : 'btn-danger';
  const alertClass = aprobado ? 'alert-success' : 'alert-danger';
  
  const cuerpo = `
    <h2>${icono} Cotización ${estado}</h2>
    
    <div class="ticket-box">
      <h3>Ticket #${folio}</h3>
      <p><strong>${titulo || 'Sin título'}</strong></p>
      ${presupuesto ? `<p style="font-size:18px;margin-top:10px;">Presupuesto: <strong style="color:#059669;">$${presupuesto}</strong></p>` : ''}
    </div>
    
    <div class="alert-box ${alertClass}">
      <p style="margin:0;">La cotización de tu ticket ha sido <strong>${aprobado ? 'aprobada' : 'rechazada'}</strong>.</p>
      ${aprobado ? '<p style="margin:10px 0 0;">El trabajo procederá según lo acordado.</p>' : '<p style="margin:10px 0 0;">Por favor contacta al equipo de soporte para más información.</p>'}
    </div>
    
    ${comentario ? `<div style="margin-top:20px;padding:15px;background:#f8fafc;border-radius:8px;"><strong>Comentario:</strong><br><span style="color:#475569;">${comentario}</span></div>` : ''}
    
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn ${colorBtn}">🔗 Ver Detalles</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailUsuario, `Cotización Ticket #${folio} - ${estado}`, cuerpo);
}


/**
 * Notificar reasignación de ticket a nuevo agente
 */
function notificarReasignacion(emailNuevoAgente, ticket, motivo) {
  const { folio, titulo, area, ubicacion, prioridad, reportaNombre } = ticket;
  
  const prioClass = (prioridad || '').toLowerCase().includes('alta') ? 'priority-alta' : 
                    (prioridad || '').toLowerCase().includes('media') ? 'priority-media' : 'priority-baja';
  
  const cuerpo = `
    <h2>🔄 Ticket Reasignado</h2>
    <p>Se te ha reasignado un ticket que requiere tu atención.</p>
    
    <div class="alert-box alert-warning">
      <strong>Motivo de reasignación:</strong> ${motivo || 'Reasignación manual'}
    </div>
    
    <div class="ticket-box">
      <h3>Ticket #${folio}</h3>
      <p><strong>${titulo || 'Sin título'}</strong></p>
    </div>
    
    <div class="info-row">
      <span class="info-label">Área:</span>
      <span class="info-value">${area || 'N/A'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Ubicación:</span>
      <span class="info-value">${ubicacion || 'No especificada'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Prioridad:</span>
      <span class="info-value ${prioClass}">${prioridad || 'Media'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Reportado por:</span>
      <span class="info-value">${reportaNombre || 'Usuario'}</span>
    </div>
    
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn btn-warning">⚡ Atender Ticket</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailNuevoAgente, `Ticket #${folio} Reasignado`, cuerpo);
}


/**
 * Notificar alerta de tickets sin atender por ausencia de agente
 */
function notificarAlertaTicketsSinAtender(emailAgente, tickets, agenteAusente, motivo) {
  const listaTickets = tickets.map(t => `
    <div style="padding:12px;border-bottom:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center;">
      <div>
        <strong style="color:#1e3a5f;">#${t.folio}</strong> - ${t.titulo || 'Sin título'}
        <br><small style="color:#64748b;">Ubicación: ${t.ubicacion || 'N/A'}</small>
      </div>
      <span class="priority-${(t.prioridad || 'media').toLowerCase()}" style="font-size:12px;">${t.prioridad || 'Media'}</span>
    </div>
  `).join('');
  
  const cuerpo = `
    <h2>⚠️ Alerta: Tickets Requieren Atención</h2>
    
    <div class="alert-box alert-warning">
      <p style="margin:0;"><strong>${agenteAusente}</strong> se ha marcado como ausente.</p>
      <p style="margin:5px 0 0;">Motivo: ${motivo || 'No especificado'}</p>
    </div>
    
    <p>Los siguientes tickets necesitan ser atendidos:</p>
    
    <div style="background:#fff;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;margin:20px 0;">
      ${listaTickets}
    </div>
    
    <p style="color:#64748b;">Por favor revisa estos tickets y coordina con tu equipo para su atención oportuna.</p>
    
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn btn-warning">📋 Revisar Tickets</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailAgente, `⚠️ Alerta: ${tickets.length} Ticket(s) Requieren Atención`, cuerpo);
}


/**
 * Notificar nuevo comentario en ticket
 */
function notificarNuevoComentario(emailDestino, ticket, autorComentario, comentario, esInterno) {
  if (esInterno) return false; // No notificar comentarios internos a usuarios
  
  const { folio, titulo } = ticket;
  
  const cuerpo = `
    <h2>💬 Nuevo Comentario</h2>
    <p>Se ha agregado un nuevo comentario a tu ticket.</p>
    
    <div class="ticket-box">
      <h3>Ticket #${folio}</h3>
      <p><strong>${titulo || 'Sin título'}</strong></p>
    </div>
    
    <div style="margin:20px 0;padding:20px;background:#f0f9ff;border-radius:8px;border-left:4px solid #3b82f6;">
      <p style="margin:0 0 10px;color:#64748b;font-size:13px;"><strong>${autorComentario}</strong> comentó:</p>
      <p style="margin:0;color:#1e293b;">${comentario}</p>
    </div>
    
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn">💬 Ver Conversación</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailDestino, `Nuevo comentario en Ticket #${folio}`, cuerpo);
}


/**
 * Notificar ticket próximo a vencer (SLA)
 */
function notificarTicketPorVencer(emailAgente, ticket, horasRestantes) {
  const { folio, titulo, area, ubicacion, prioridad, vencimiento } = ticket;
  
  const alertClass = horasRestantes <= 2 ? 'alert-danger' : 'alert-warning';
  const urgencia = horasRestantes <= 2 ? '🚨 URGENTE' : '⏰ Próximo a vencer';
  
  const cuerpo = `
    <h2>${urgencia}</h2>
    
    <div class="alert-box ${alertClass}">
      <p style="margin:0;"><strong>El ticket vence en ${horasRestantes} hora(s)</strong></p>
      ${vencimiento ? `<p style="margin:5px 0 0;">Fecha límite: ${new Date(vencimiento).toLocaleString('es-MX')}</p>` : ''}
    </div>
    
    <div class="ticket-box">
      <h3>Ticket #${folio}</h3>
      <p><strong>${titulo || 'Sin título'}</strong></p>
    </div>
    
    <div class="info-row">
      <span class="info-label">Área:</span>
      <span class="info-value">${area || 'N/A'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Ubicación:</span>
      <span class="info-value">${ubicacion || 'No especificada'}</span>
    </div>
    <div class="info-row">
      <span class="info-label">Prioridad:</span>
      <span class="info-value priority-alta">${prioridad || 'Alta'}</span>
    </div>
    
    <p style="margin-top:25px;">
      <a href="${getScriptUrl()}" class="btn btn-danger">⚡ Atender Ahora</a>
    </p>
  `;
  
  return enviarEmailNotificacion(emailAgente, `${urgencia}: Ticket #${folio}`, cuerpo);
}


// ============================================================================
// FUNCIÓN PARA ENVIAR TELEGRAM A AGENTES INDIVIDUALES
// ============================================================================

/**
 * Enviar mensaje de Telegram a un agente específico
 * MODIFICADO: Usa getEmailNotificacion() para buscar la clave correcta
 */
function enviarTelegramAgente(emailAgente, mensaje) {
  if (!emailAgente) return false;
  
  try {
    // Obtener el email real del agente
    const emailReal = getEmailNotificacion(emailAgente);
    
    if (!emailReal) {
      Logger.log(`⚠️ No se pudo obtener email para Telegram: ${emailAgente}`);
      return false;
    }
    
    // Buscar el chat_id del agente en la configuración
    // Formato de clave: tg_email@dominio.com
    const keyTelegram = 'tg_' + emailReal;
    const chatIdAgente = getConfig(keyTelegram);
    
    if (!chatIdAgente) {
      // También intentar con el username original (por compatibilidad)
      const keyUsername = 'tg_' + String(emailAgente).trim().toLowerCase();
      const chatIdPorUsername = getConfig(keyUsername);
      
      if (chatIdPorUsername) {
        return telegramSend(mensaje, chatIdPorUsername);
      }
      
      Logger.log(`ℹ️ El agente ${emailAgente} (${emailReal}) no tiene Telegram configurado`);
      return false;
    }
    
    return telegramSend(mensaje, chatIdAgente);
    
  } catch (e) {
    Logger.log(`❌ Error enviando Telegram a ${emailAgente}: ${e.message}`);
    return false;
  }
}



function notificarAgenteTodosCanales(emailAgente, tipo, datos) {
  const { folio, titulo, area, ubicacion, prioridad, reportaNombre, descripcion, motivo } = datos;

  // 1. Panel de notificaciones interno
  try {
    notifyUser(emailAgente, tipo, datos.tituloNotif || titulo, datos.mensajeNotif || descripcion || '', {
      ticketId: datos.ticketId,
      folio
    });
  } catch (e) {
    Logger.log('Error en panel notificaciones: ' + e.message);
  }
  
  // 2. Email
  try {
    switch (tipo) {
      case 'nuevo_ticket':
      case 'asignacion':
        notificarAgenteNuevoTicket(emailAgente, { folio, titulo, area, ubicacion, prioridad, reportaNombre, descripcion });
        break;
      case 'reasignacion':
        notificarReasignacion(emailAgente, { folio, titulo, area, ubicacion, prioridad, reportaNombre }, motivo);
        break;
      case 'vencimiento':
        notificarTicketPorVencer(emailAgente, datos, datos.horasRestantes || 4);
        break;
      case 'reapertura':
        enviarEmailReapertura(emailAgente, datos);
        break;
      case 'vencido':
        notificarTicketVencido(emailAgente, datos);
        break;
      case 'escalamiento_aprobado':
        enviarEmailConEnlaceDirecto(emailAgente, '✅ Escalamiento Aprobado - Ticket #' + folio, datos.ticketId, folio, '<h2>✅ Escalamiento Aprobado</h2><p>El ticket <strong>#' + folio + '</strong> ha sido aprobado para atención prioritaria.</p><p><strong>' + (titulo || '') + '</strong></p>');
        break;
      case 'escalamiento_rechazado':
        enviarEmailConEnlaceDirecto(emailAgente, '❌ Escalamiento Rechazado - Ticket #' + folio, datos.ticketId, folio, '<h2>❌ Escalamiento Rechazado</h2><p>La solicitud de escalamiento del ticket <strong>#' + folio + '</strong> fue rechazada.</p>' + (datos.motivo ? '<p><strong>Motivo:</strong> ' + datos.motivo + '</p>' : ''));
        break;
      case 'resolucion_confirmada':
      case 'resolucion_rechazada':
      case 'reprog_aprobada':
      case 'reprog_rechazada':
        enviarEmailNotificacion(emailAgente, datos.tituloNotif, datos.htmlCuerpo);
        break;
    }
  } catch (e) {
    Logger.log('Error en email omnicanal: ' + e.message);
  }
  
  // 3. Telegram

  Logger.log('[notificarAgenteTodosCanales] Intentando Telegram para: ' + emailAgente + ' | Tipo: ' + tipo);
  
  try {
    let msgTelegram = '';
    switch (tipo) {
      case 'nuevo_ticket':
      case 'asignacion':
        msgTelegram = `📋 <b>Nuevo Ticket Asignado</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n📍 ${ubicacion || 'Sin ubicación'}\n⚡ Prioridad: ${prioridad || 'Media'}\n👤 Reporta: ${reportaNombre || 'Usuario'}`;
        break;
      case 'reasignacion':
        msgTelegram = `🔄 <b>Ticket Reasignado</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n📝 Motivo: ${motivo || 'Reasignación'}`;
        break;
      case 'vencimiento':
        msgTelegram = `⏰ <b>Ticket por Vencer</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n⚠️ Vence en ${datos.horasRestantes || '?'} horas`;
        break;
      case 'comentario':
        msgTelegram = `💬 <b>Nuevo Comentario</b>\nTicket: <code>#${folio}</code>\n${datos.autorComentario || 'Usuario'}: ${(datos.comentario || '').substring(0, 100)}`;
        break;
      case 'reapertura':
        msgTelegram = `🔄 <b>Ticket REABIERTO</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n📍 ${ubicacion || ''}\n📝 Motivo: ${motivo || 'Sin motivo'}\n⏰ Nuevo SLA: ${datos.vencimiento ? new Date(datos.vencimiento).toLocaleString('es-MX') : 'N/A'}`;
        break;
      case 'vencido':
        msgTelegram = `🚨 <b>TICKET VENCIDO</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n📍 ${ubicacion || ''}\n⚡ Prioridad: ${prioridad || 'Media'}\n⚠️ Requiere atención URGENTE`;
        break;
      case 'escalamiento_aprobado':
        msgTelegram = `✅ <b>Escalamiento APROBADO</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n⚡ Atención prioritaria requerida`;
        break;
      case 'escalamiento_rechazado':
        msgTelegram = `❌ <b>Escalamiento RECHAZADO</b>\nFolio: <code>#${folio}</code>\n<b>${titulo || 'Sin título'}</b>\n📝 Motivo: ${motivo || 'No especificado'}`;
        break;
      case 'resolucion_confirmada':
        msgTelegram = `✅ <b>Resolución Confirmada</b>\nFolio: <code>#${folio}</code>\n¡Excelente trabajo! El usuario cerró el ticket.`;
        break;
      case 'resolucion_rechazada':
        msgTelegram = `⚠️ <b>Resolución RECHAZADA</b>\nFolio: <code>#${folio}</code>\nEl usuario indicó que no quedó resuelto.\n📝 Motivo: ${motivo}`;
        break;
      case 'reprog_aprobada':
        msgTelegram = `📅 <b>Visita Aprobada</b>\nFolio: <code>#${folio}</code>\nEl usuario aprobó la fecha propuesta.`;
        break;
      case 'reprog_rechazada':
        msgTelegram = `❌ <b>Visita Rechazada</b>\nFolio: <code>#${folio}</code>\nEl usuario rechazó la reprogramación de la visita.`;
        break;
      default:
        msgTelegram = `📌 <b>${datos.tituloNotif || 'Notificación'}</b>\n${datos.mensajeNotif || ''}`;
    }
    
    enviarTelegramAgente(emailAgente, msgTelegram);
  } catch (e) {
    Logger.log('Error en Telegram omnicanal: ' + e.message);
  }
}

function enviarEmailReapertura(emailAgente, datos) {
  var folio = datos.folio || '';
  var titulo = datos.titulo || 'Sin título';
  var area = datos.area || '';
  var ubicacion = datos.ubicacion || '';
  var prioridad = datos.prioridad || '';
  var motivo = datos.motivo || 'Sin motivo especificado';
  var vencimiento = datos.vencimiento ? new Date(datos.vencimiento).toLocaleString('es-MX') : 'N/A';
  var reportaNombre = datos.reportaNombre || 'Usuario';

  var cuerpo = '<h2>🔄 Ticket Reabierto</h2>' +
    '<div class="alert-box alert-warning">' +
      '<p style="margin:0;"><strong>Un ticket previamente cerrado ha sido reabierto</strong></p>' +
      '<p style="margin:5px 0 0;">El SLA se ha reiniciado. Nuevo vencimiento: ' + vencimiento + '</p>' +
    '</div>' +
    '<div class="ticket-box">' +
      '<h3>Ticket #' + folio + '</h3>' +
      '<p><strong>' + titulo + '</strong></p>' +
    '</div>' +
    '<div class="info-row"><span class="info-label">Motivo de reapertura:</span><span class="info-value">' + motivo + '</span></div>' +
    '<div class="info-row"><span class="info-label">Área:</span><span class="info-value">' + area + '</span></div>' +
    '<div class="info-row"><span class="info-label">Ubicación:</span><span class="info-value">' + ubicacion + '</span></div>' +
    '<div class="info-row"><span class="info-label">Prioridad:</span><span class="info-value">' + prioridad + '</span></div>' +
    '<div class="info-row"><span class="info-label">Solicitado por:</span><span class="info-value">' + reportaNombre + '</span></div>' +
    '<p style="margin-top:25px;"><a href="' + getScriptUrl() + '" class="btn btn-warning">🔄 Atender Ticket</a></p>';

  return enviarEmailNotificacion(emailAgente, '🔄 Ticket Reabierto: #' + folio + ' - ' + titulo, cuerpo);
}

function notificarTicketVencido(emailAgente, datos) {
  var folio = datos.folio || '';
  var titulo = datos.titulo || 'Sin título';
  var area = datos.area || '';
  var ubicacion = datos.ubicacion || '';
  var prioridad = datos.prioridad || '';
  var vencimiento = datos.vencimiento ? new Date(datos.vencimiento).toLocaleString('es-MX') : 'N/A';

  var cuerpo = '<h2>🚨 SLA VENCIDO</h2>' +
    '<div class="alert-box alert-danger">' +
      '<p style="margin:0;"><strong>Este ticket ha superado su fecha límite de atención</strong></p>' +
      '<p style="margin:5px 0 0;">Venció: ' + vencimiento + '</p>' +
    '</div>' +
    '<div class="ticket-box">' +
      '<h3>Ticket #' + folio + '</h3>' +
      '<p><strong>' + titulo + '</strong></p>' +
    '</div>' +
    '<div class="info-row"><span class="info-label">Área:</span><span class="info-value">' + area + '</span></div>' +
    '<div class="info-row"><span class="info-label">Ubicación:</span><span class="info-value">' + ubicacion + '</span></div>' +
    '<div class="info-row"><span class="info-label">Prioridad:</span><span class="info-value priority-alta">' + prioridad + '</span></div>' +
    '<p style="margin-top:25px;"><a href="' + getScriptUrl() + '" class="btn btn-danger">⚡ Atender URGENTE</a></p>';

  return enviarEmailNotificacion(emailAgente, '🚨 SLA VENCIDO: Ticket #' + folio + ' - ' + titulo, cuerpo);
}

// ============================================================================
// PASO 2: AGREGAR FUNCIÓN getEmailNotificacion()
// Agregar después de las funciones de usuario existentes
// ============================================================================

/**
 * Obtener el email de notificación de un usuario
 * Prioridad: EmailNotificacion > Email (si tiene @) > Email + dominio
 * 
 * @param {string} username - Username o email del usuario
 * @returns {string} - Email válido para notificaciones
 */
function getEmailNotificacion(username) {
  if (!username) return '';
  
  try {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    
    // Buscar usuario por Email (username)
    const usernameLC = String(username).trim().toLowerCase();
    const row = rows.find(r => 
      String(r[m.Email] || '').toLowerCase() === usernameLC ||
      String(r[m.EmailNotificacion] || '').toLowerCase() === usernameLC
    );
    
    if (!row) {
      // Usuario no encontrado, intentar normalizar el input
      return normalizarEmail(username);
    }
    
    // Prioridad 1: EmailNotificacion si existe y es válido
    if (m.EmailNotificacion != null) {
      const emailNotif = String(row[m.EmailNotificacion] || '').trim();
      if (emailNotif && esEmailValido(emailNotif)) {
        return emailNotif.toLowerCase();
      }
    }
    
    // Prioridad 2: Email si ya tiene @ (es un email completo)
    const emailCampo = String(row[m.Email] || '').trim();
    if (emailCampo.includes('@') && esEmailValido(emailCampo)) {
      return emailCampo.toLowerCase();
    }
    
    // Prioridad 3: Email + dominio de configuración
    return normalizarEmail(emailCampo);
    
  } catch (e) {
    Logger.log('Error en getEmailNotificacion: ' + e.message);
    return normalizarEmail(username);
  }
}


/**
 * Normaliza un email agregando el dominio si es necesario
 */
function normalizarEmail(email) {
  if (!email) return '';
  
  const emailStr = String(email).trim().toLowerCase();
  
  // Si ya tiene @, está completo
  if (emailStr.includes('@')) {
    return emailStr;
  }
  
  // Obtener dominio de la configuración
  const config = getConfig();
  const dominio = config.dominio || 'bexalta.com';
  
  return `${emailStr}@${dominio}`;
}


/**
 * Verifica si un string es un email válido
 */
function esEmailValido(email) {
  if (!email) return false;
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(String(email).trim());
}

function getTodasLasUbicaciones() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ubicaciones = new Set();

    ['Categorias_TI', 'Categorias_Mtto'].forEach(nombreHoja => {
      const sh = ss.getSheetByName(nombreHoja);
      if (!sh || sh.getLastRow() <= 1) return;

      const data = sh.getDataRange().getValues();
      // Columna C (índice 2) típicamente tiene ubicaciones
      for (let i = 1; i < data.length; i++) {
        const ubic = String(data[i][2] || '').trim();
        if (ubic) {
          // Puede ser lista separada por comas
          ubic.split(',').forEach(u => {
            const limpio = u.trim();
            if (limpio) ubicaciones.add(limpio);
          });
        }
      }
    });

    // También incluir ubicaciones de los usuarios
    const users = getCachedData(DB.USERS);
    const hdr = HEADERS.Usuarios;
    const idxUbic = hdr.indexOf('Ubicación');
    if (idxUbic >= 0) {
      users.forEach(r => {
        const ubic = String(r[idxUbic] || '').trim();
        if (ubic) {
          ubic.split(',').forEach(u => {
            const limpio = u.trim();
            if (limpio) ubicaciones.add(limpio);
          });
        }
      });
    }

    return Array.from(ubicaciones).sort((a, b) => a.localeCompare(b, 'es'));

  } catch (e) {
    Logger.log('Error en getTodasLasUbicaciones: ' + e.message);
    return [];
  }
}

function registrarAuditoriaAdmin_(accion, detalle, emailAdmin) {
  try {
    registrarBitacora('ADMIN', accion, `${detalle} | Por: ${emailAdmin || 'sistema'}`);
  } catch (e) {
    Logger.log('Error auditoría admin: ' + e.message);
  }
}

/**
 * Helper interno: resuelve username → email real desde hoja Usuarios
 * Si ya es un email (@) lo retorna directo.
 */
function _resolverEmailUsuario_(usernameOrEmail) {
  if (!usernameOrEmail) return '';
  const val = String(usernameOrEmail).trim();
  if (val.includes('@')) return val; // ya es email

  try {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);
    // Buscar por columna Email (que puede tener el username como clave)
    const row = rows.find(r => {
      const emailCol = String(r[m.Email] || '').trim().toLowerCase();
      return emailCol === val.toLowerCase();
    });
    if (!row) return val; // no encontrado, devolver tal cual

    // Prioridad: EmailNotificacion > columna Email si contiene @
    const emailNotif = m.EmailNotificacion != null ? String(row[m.EmailNotificacion] || '').trim() : '';
    const emailCol   = String(row[m.Email] || '').trim();

    if (emailNotif && emailNotif.includes('@')) return emailNotif;
    if (emailCol.includes('@')) return emailCol;
    return val; // fallback
  } catch(e) {
    Logger.log('_resolverEmailUsuario_ error: ' + e.message);
    return val;
  }
}

/**
 * Obtener gerente del área.
 * Retorna { username, email, nombre } donde:
 *   username = valor en ConfigGerentes (para AsignadoA en tickets)
 *   email    = email real para notificaciones
 */
function getGerenteDelArea(area) {
  const areaLower = String(area || '').trim().toLowerCase();

  // 1. Prioridad: constante en código
  if (typeof GERENTES_AREAS !== 'undefined' && GERENTES_AREAS[areaLower]) {
    const g = GERENTES_AREAS[areaLower];
    const emailReal = _resolverEmailUsuario_(g.email);
    return { username: g.email, email: emailReal, nombre: g.nombre };
  }

  // 2. Fallback: hoja ConfigGerentes
  try {
    const { headers, rows } = _readTableByHeader_('ConfigGerentes');
    const m = _headerMap_(headers);

    const gerenteRow = rows.find(r => {
      const areaGerente = String(r[m.Area] || r[m['Área']] || '').toLowerCase();
      const activo = String(r[m.Activo] || 'si').toLowerCase();
      return areaGerente === areaLower && activo !== 'no';
    });

    if (gerenteRow) {
      const username  = String(gerenteRow[m.GerenteEmail]  || '').trim(); // ej: RGNava
      const nombre    = String(gerenteRow[m.GerenteNombre] || '').trim();
      const emailReal = _resolverEmailUsuario_(username);                  // ej: rgnava@bexalta.com
      return { username, email: emailReal, nombre };
    }
  } catch(e) {
    Logger.log('getGerenteDelArea error: ' + e.message);
  }

  return null;
}


// ============================================================
// PARTE 7: TELEGRAM A AGENTES ESPECÍFICOS
// ============================================================

/**
 * Obtener chat ID de Telegram para un agente
 * Busca en Config: telegram_NombreAgente o tg_email
 */
function getTelegramChatId(email) {
  const cfg = getConfig();
  const emailLower = (email || '').toLowerCase().trim();
  const usuario = emailLower.split('@')[0];
  
  // Buscar por diferentes formatos
  // Formato 1: tg_usuario@dominio.com
  const key1 = 'tg_' + emailLower;
  if (cfg[key1]) return cfg[key1];
  
  // Formato 2: tg_usuario
  const key2 = 'tg_' + usuario;
  if (cfg[key2]) return cfg[key2];
  
  // Formato 3: telegram_NombreUsuario (buscar por nombre)
  for (const key in cfg) {
    if (key.startsWith('telegram_') && !['telegram_token', 'telegram_chat_id', 'telegram_chat_admin'].includes(key)) {
      // Verificar si el nombre coincide
      const nombre = key.replace('telegram_', '').toLowerCase();
      if (usuario.includes(nombre) || nombre.includes(usuario)) {
        return cfg[key];
      }
    }
  }
  
  return null;
}

/**
 * Enviar notificación Telegram a agente específico
 */
function notificarAgenteTelegram(email, mensaje, ticketInfo) {
  const chatId = getTelegramChatId(email);
  
  if (!chatId) {
    Logger.log(`⚠️ No se encontró chat ID de Telegram para: ${email}`);
    return false;
  }
  
  const cfg = getConfig();
  const token = cfg.telegram_token;
  
  if (!token) {
    Logger.log('⚠️ Falta telegram_token en Config');
    return false;
  }
  
  // Construir mensaje con formato HTML
  let texto = mensaje;
  
  if (ticketInfo) {
    texto = `${mensaje}\n\n`;
    texto += `<b>Ticket #${ticketInfo.folio || ''}</b>\n`;
    if (ticketInfo.titulo) texto += `📋 ${ticketInfo.titulo}\n`;
    if (ticketInfo.area) texto += `📁 Área: ${ticketInfo.area}\n`;
    if (ticketInfo.prioridad) texto += `⚡ Prioridad: ${ticketInfo.prioridad}\n`;
    if (ticketInfo.reporta) texto += `👤 Reportó: ${ticketInfo.reporta}\n`;
  }
  
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  
  const payload = {
    chat_id: chatId,
    text: texto,
    parse_mode: 'HTML',
    disable_web_page_preview: true
  };
  
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    Logger.log(`✅ Telegram enviado a ${email} (${chatId})`);
    return true;
  } catch (err) {
    Logger.log(`⚠️ Error enviando Telegram a ${email}: ${err.message}`);
    return false;
  }
}

/**
 * Notificar escalamiento al gerente vía Email y Telegram
 */
function notificarEscalamientoGerente(ticketId, area, motivo, solicitante) {
  // Obtener gerente del área
  const gerente = getGerenteDelArea(area);
  
  if (!gerente || !gerente.email) {
    Logger.log(`⚠️ No se encontró gerente para el área: ${area}`);
    return { ok: false, error: 'No hay gerente configurado para esta área' };
  }
  
  // Obtener datos del ticket
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);
  const row = rows.find(r => String(r[m.ID]) === String(ticketId));
  
  if (!row) {
    return { ok: false, error: 'Ticket no encontrado' };
  }
  
  const folio = row[m.Folio];
  const titulo = row[m['Título']] || row[m.Titulo] || '';
  const prioridad = row[m.Prioridad];
  const reportaEmail = row[m.ReportaEmail];
  
  // 1. Enviar Email al gerente
  try {
    const appUrl = ScriptApp.getService().getUrl();
    const approveLink = `${appUrl}?action=approve_escalar&id=${ticketId}&by=${encodeURIComponent(solicitante)}`;
    const rejectLink = `${appUrl}?action=reject_escalar&id=${ticketId}&by=${encodeURIComponent(solicitante)}`;
    
    MailApp.sendEmail({
      to: gerente.email,
      subject: `🔺 Escalamiento - Ticket #${folio} - ${area}`,
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
          <div style="background: #dc2626; color: white; padding: 20px; border-radius: 8px 8px 0 0;">
            <h2 style="margin: 0;">🔺 Solicitud de Escalamiento</h2>
          </div>
          <div style="padding: 20px; border: 1px solid #e5e7eb;">
            <table style="width: 100%;">
              <tr><td style="color: #6b7280; padding: 8px 0;">Ticket:</td><td><strong>#${folio}</strong></td></tr>
              <tr><td style="color: #6b7280; padding: 8px 0;">Título:</td><td>${titulo}</td></tr>
              <tr><td style="color: #6b7280; padding: 8px 0;">Área:</td><td>${area}</td></tr>
              <tr><td style="color: #6b7280; padding: 8px 0;">Prioridad:</td><td>${prioridad}</td></tr>
              <tr><td style="color: #6b7280; padding: 8px 0;">Solicitante:</td><td>${solicitante}</td></tr>
            </table>
            <div style="background: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; margin: 20px 0;">
              <strong>Motivo:</strong><br>${motivo}
            </div>
          </div>
          <div style="background: #1f2937; padding: 20px; border-radius: 0 0 8px 8px; text-align: center;">
            <a href="${approveLink}" style="display: inline-block; background: #16a34a; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; margin: 0 8px;">✓ Aprobar</a>
            <a href="${rejectLink}" style="display: inline-block; background: #dc2626; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; margin: 0 8px;">✗ Rechazar</a>
          </div>
        </div>
      `
    });
    Logger.log(`✅ Email de escalamiento enviado a gerente: ${gerente.email}`);
  } catch (e) {
    Logger.log(`⚠️ Error enviando email de escalamiento: ${e.message}`);
  }
  
  // 2. Enviar Telegram al gerente
  notificarAgenteTelegram(gerente.email, '🔺 <b>ESCALAMIENTO PENDIENTE</b>', {
    folio: folio,
    titulo: titulo,
    area: area,
    prioridad: prioridad,
    reporta: reportaEmail
  });
  
  return { ok: true, gerente: gerente.email };
}

// ============================================================
// FUNCIÓN: Obtener Chat ID de Telegram por usuario
// ============================================================

function getTelegramChatIdGrupo(identificador) {
  if (!identificador) return null;
  
  const id = String(identificador).trim();
  const idLower = id.toLowerCase();
  const usuario = idLower.split('@')[0];
  
  // 1. Buscar en el mapa de grupos
  if (TELEGRAM_GRUPOS[id]) return TELEGRAM_GRUPOS[id];
  if (TELEGRAM_GRUPOS[idLower]) return TELEGRAM_GRUPOS[idLower];
  if (TELEGRAM_GRUPOS[usuario]) return TELEGRAM_GRUPOS[usuario];
  
  // 2. Buscar parcialmente (por si el email tiene variaciones)
  for (const key in TELEGRAM_GRUPOS) {
    const keyLower = key.toLowerCase();
    if (keyLower.includes(usuario) || usuario.includes(keyLower.split('@')[0])) {
      return TELEGRAM_GRUPOS[key];
    }
  }
  
  // 3. Buscar en Config como fallback
  const cfg = getConfig();
  const keyTg = 'tg_' + idLower;
  if (cfg[keyTg]) return cfg[keyTg];
  
  const keyTgUsuario = 'tg_' + usuario;
  if (cfg[keyTgUsuario]) return cfg[keyTgUsuario];
  
  Logger.log(`⚠️ No se encontró chat de Telegram para: ${identificador}`);
  return null;
}


// ============================================================
// FUNCIÓN: Obtener gerente del área con su Telegram
// ============================================================

function getGerenteAreaConTelegram(area) {
  const areaLower = (area || '').toLowerCase().trim();
  
  // Buscar en la configuración de gerentes
  const gerente = GERENTES_AREAS[areaLower];
  if (gerente) {
    return gerente;
  }
  
  // Fallback: buscar en hoja ConfigGerentes
  try {
    const { headers, rows } = _readTableByHeader_('ConfigGerentes');
    const m = _headerMap_(headers);
    
    const found = rows.find(r => {
      const gerenteArea = String(r[m.Area] || '').toLowerCase();
      const activo = String(r[m.Activo] || 'si').toLowerCase();
      return gerenteArea === areaLower && activo !== 'no';
    });
    
    if (found) {
      const email = found[m.Email] || '';
      return {
        email: email,
        nombre: found[m.Nombre] || email,
        chatId: getTelegramChatIdGrupo(email)
      };
    }
  } catch (e) {
    Logger.log('Error buscando gerente en hoja: ' + e.message);
  }
  
  return null;
}

// ============================================================
// FUNCIÓN MEJORADA: Enviar Telegram
// ============================================================

function telegramSendToGrupo(chatId, texto, parseMode, fileId) {
  const cfg = getConfig();
  const token = cfg.telegram_token;
  
  if (!token || !chatId) return false;

  try {
    let response;
    const options = {
      method: 'post',
      muteHttpExceptions: true,
      connectTimeout: 8000,
      readTimeout: 8000
    };

    // CASO A: HAY ARCHIVO ADJUNTO
    if (fileId && fileId !== '') {
      const file = DriveApp.getFileById(fileId);
      const blob = file.getBlob();
      
      const url = `https://api.telegram.org/bot${token}/sendDocument`;
      
      // Para enviar archivos físicos usamos un payload multipart
      options.payload = {
        chat_id: String(chatId),
        document: blob,
        caption: texto, // El texto del comentario va como leyenda del archivo
        parse_mode: parseMode || 'HTML'
      };
      
      response = UrlFetchApp.fetch(url, options);
    } 
    // CASO B: SOLO TEXTO
    else {
      const url = `https://api.telegram.org/bot${token}/sendMessage`;
      options.contentType = 'application/json';
      options.payload = JSON.stringify({
        chat_id: String(chatId),
        text: texto,
        parse_mode: parseMode || 'HTML',
        disable_web_page_preview: true
      });
      
      response = UrlFetchApp.fetch(url, options);
    }

    const result = JSON.parse(response.getContentText());
    if (result.ok) {
      Logger.log(`✅ Telegram enviado con éxito`);
      return true;
    } else {
      Logger.log(`⚠️ Error API Telegram: ${result.description}`);
      // Reintento en texto plano si falla el HTML
      if (result.description.includes("can't parse entities")) {
        return telegramSendToGrupo(chatId, texto.replace(/<[^>]*>?/gm, ''), '', fileId);
      }
      return false;
    }
  } catch (err) {
    Logger.log('⚠️ Error crítico Telegram: ' + err.message);
    return false;
  }
}

// ============================================================
// FUNCIÓN: Notificar agente vía Telegram (grupo)
// ============================================================

function notificarAgenteTelegramGrupo(email, mensaje, ticketInfo) {
  const chatId = getTelegramChatIdGrupo(email);
  if (!chatId) {
    Logger.log(`⚠️ No hay chat de Telegram para: ${email}`);
    return false;
  }
  
  let texto = mensaje;
  if (ticketInfo) {
    texto += '\n\n';
    texto += `📋 <b>Ticket #${ticketInfo.folio || ''}</b>\n`;
    if (ticketInfo.titulo) texto += `📝 ${ticketInfo.titulo}\n`;
    if (ticketInfo.area) texto += `📁 Área: ${ticketInfo.area}\n`;
    if (ticketInfo.ubicacion) texto += `📍 Ubic: ${ticketInfo.ubicacion}\n`;
    if (ticketInfo.prioridad) texto += `⚡ Prioridad: ${ticketInfo.prioridad}\n`;
    if (ticketInfo.reporta) texto += `👤 Reportó: ${ticketInfo.reporta}\n`;
    
    if (ticketInfo.descripcion) {
      // Limpiamos HTML y caracteres que rompen Telegram
      const descLimpia = ticketInfo.descripcion.replace(/<[^>]*>?/gm, ''); 
      const desc = descLimpia.substring(0, 150);
      texto += `\n💬 ${desc}${descLimpia.length > 150 ? '...' : ''}`;
    }
  }
  
  return telegramSendToGrupo(chatId, texto);
}

// ============================================================
// FUNCIÓN: Notificar gerente de área vía Telegram
// ============================================================

function notificarGerenteTelegram(area, mensaje, ticketInfo) {
  const gerente = getGerenteAreaConTelegram(area);
  if (!gerente || !gerente.chatId) {
    Logger.log(`⚠️ No hay gerente o chat de Telegram para área: ${area}`);
    return false;
  }
  
  let texto = `👔 <b>${mensaje}</b>`;
  if (ticketInfo) {
    texto += '\n\n';
    texto += `📋 <b>Ticket #${ticketInfo.folio || ''}</b>\n`;
    if (ticketInfo.titulo) texto += `📝 ${ticketInfo.titulo}\n`;
    if (ticketInfo.area) texto += `📁 Área: ${ticketInfo.area}\n`;
    if (ticketInfo.prioridad) texto += `⚡ Prioridad: ${ticketInfo.prioridad}\n`;
    
    if (ticketInfo.motivo) {
      const motivoLimpio = ticketInfo.motivo.replace(/<[^>]*>?/gm, '');
      texto += `\n⚠️ <b>Motivo:</b> ${motivoLimpio}`;
    }
    if (ticketInfo.solicitante) texto += `\n👤 Solicitado por: ${ticketInfo.solicitante}`;
  }
  
  return telegramSendToGrupo(gerente.chatId, texto);
}

function reprogramarVisitaConValidacion(ticketId, nuevaFecha, nuevaHora, notas, userEmail) {
  return withLock_(() => {
    try {
      const { headers, rows } = _readTableByHeader_(DB.TICKETS);
      const m = _headerMap_(headers);
      const sh = getSheet(DB.TICKETS);
      
      const idx = rows.findIndex(r => String(r[m.ID]) === String(ticketId));
      if (idx < 0) return { ok: false, error: 'Ticket no encontrado' };
      
      const row = rows[idx];
      const folio = row[m.Folio];
      const vencimiento = new Date(row[m.Vencimiento]);
      const reportaEmail = row[m.ReportaEmail];
      const titulo = row[m['Título']] || '';

      // === NUEVO CANDADO DE HORARIO LABORAL ===
      const validacion = validarFechaHoraLaboral(nuevaFecha, nuevaHora);
      if (!validacion.valido) {
        return { ok: false, error: validacion.error };
      }
      // ========================================
      
      // Calcular si la nueva fecha está fuera del SLA
      const nuevaFechaObj = new Date(nuevaFecha + 'T' + nuevaHora);
      const fueraDeSLA = nuevaFechaObj > vencimiento;
      
      if (fueraDeSLA) {
        // Requiere aprobación del usuario
        Logger.log(`Reprogramación fuera de SLA - Ticket #${folio}. Solicitando aprobación.`);
        
        // Guardar solicitud pendiente
        row[m.NotasVisita] = `PENDIENTE APROBACIÓN: ${nuevaFecha} ${nuevaHora} - ${notas}`;
        sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
        
        // Enviar email al usuario para aprobar
        const appUrl = ScriptApp.getService().getUrl();
        const approveLink = `${appUrl}?action=approve_reprog&id=${ticketId}&fecha=${encodeURIComponent(nuevaFecha)}&hora=${encodeURIComponent(nuevaHora)}`;
        const rejectLink = `${appUrl}?action=reject_reprog&id=${ticketId}`;
        
        try {
          MailApp.sendEmail({
            to: reportaEmail,
            subject: `⚠️ Aprobación Requerida - Reprogramación Ticket #${folio}`,
            htmlBody: `
              <div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto;">
                <div style="background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%); color: white; padding: 20px; border-radius: 12px 12px 0 0;">
                  <h2 style="margin: 0;">⚠️ Solicitud de Reprogramación</h2>
                </div>
                
                <div style="background: #f8fafc; padding: 25px; border: 1px solid #e2e8f0;">
                  <p>El agente solicita reprogramar la visita de tu ticket:</p>
                  
                  <div style="background: white; padding: 15px; border-radius: 8px; margin: 15px 0;">
                    <h3 style="margin: 0 0 10px;">Ticket #${folio}</h3>
                    <p style="margin: 0;"><strong>${titulo}</strong></p>
                  </div>
                  
                  <div style="background: #fef3c7; border-left: 4px solid #f59e0b; padding: 15px; margin: 15px 0;">
                    <p style="margin: 0;"><strong>⚠️ IMPORTANTE:</strong> La nueva fecha propuesta está <strong>FUERA del tiempo de SLA</strong> originalmente acordado.</p>
                  </div>
                  
                  <table style="width: 100%; margin: 15px 0;">
                    <tr>
                      <td style="padding: 8px 0; color: #64748b;">Nueva fecha propuesta:</td>
                      <td style="padding: 8px 0; font-weight: bold;">${nuevaFecha} a las ${nuevaHora}</td>
                    </tr>
                    <tr>
                      <td style="padding: 8px 0; color: #64748b;">SLA original:</td>
                      <td style="padding: 8px 0;">${vencimiento.toLocaleString('es-MX')}</td>
                    </tr>
                    <tr>
                      <td style="padding: 8px 0; color: #64748b;">Motivo:</td>
                      <td style="padding: 8px 0;">${notas || 'No especificado'}</td>
                    </tr>
                  </table>
                </div>
                
                <div style="background: #1e293b; padding: 25px; border-radius: 0 0 12px 12px; text-align: center;">
                  <p style="color: #94a3b8; margin: 0 0 20px 0;">¿Aceptas la nueva fecha?</p>
                  
                  <a href="${approveLink}" style="display: inline-block; background: #16a34a; color: white; padding: 12px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; margin: 0 10px;">
                    ✓ Aprobar
                  </a>
                  
                  <a href="${rejectLink}" style="display: inline-block; background: #dc2626; color: white; padding: 12px 30px; text-decoration: none; border-radius: 8px; font-weight: bold; margin: 0 10px;">
                    ✗ Rechazar
                  </a>
                </div>
              </div>
            `
          });
        } catch (e) {
          Logger.log('Error enviando email de aprobación: ' + e.message);
        }
        
        addSystemComment(ticketId, `⚠️ Reprogramación solicitada fuera de SLA. Pendiente aprobación del usuario.\nNueva fecha: ${nuevaFecha} ${nuevaHora}\nMotivo: ${notas}`, true);
        
        return { 
          ok: true, 
          requiresApproval: true, 
          message: 'La reprogramación está fuera del SLA. Se envió solicitud de aprobación al usuario.' 
        };
      }
      
      // Si está dentro del SLA, reprogramar directamente
      row[m.FechaVisita] = nuevaFecha;
      row[m.HoraVisita] = nuevaHora;
      row[m.NotasVisita] = notas;
      row[m['ÚltimaActualización']] = new Date();
      
      sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
      clearCache(DB.TICKETS);
      
      addSystemComment(ticketId, `Visita reprogramada para ${nuevaFecha} a las ${nuevaHora}. Motivo: ${notas}`, true);
      
      // Notificar al usuario
      notifyUser(reportaEmail, 'visita_reprogramada', 'Visita reprogramada',
        `La visita para tu ticket #${folio} ha sido reprogramada para el ${nuevaFecha} a las ${nuevaHora}`,
        { ticketId, folio, fecha: nuevaFecha, hora: nuevaHora });
      
      return { ok: true, message: 'Visita reprogramada correctamente' };
      
    } catch (e) {
      Logger.log('Error en reprogramarVisitaConValidacion: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

function aprobarReprogramacion(ticketId, fecha, hora) {
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);
  const sh = getSheet(DB.TICKETS);
  
  const idx = rows.findIndex(r => String(r[m.ID]) === String(ticketId));
  if (idx < 0) throw new Error('Ticket no encontrado');
  
  const row = rows[idx];
  const folio = row[m.Folio];
  const asignadoA = row[m.AsignadoA];
  
  row[m.FechaVisita] = decodeURIComponent(fecha);
  row[m.HoraVisita] = decodeURIComponent(hora);
  row[m.NotasVisita] = (row[m.NotasVisita] || '').replace('PENDIENTE APROBACIÓN: ', 'APROBADO: ');
  row[m['ÚltimaActualización']] = new Date();
  
  sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
  clearCache(DB.TICKETS);
  
  addSystemComment(ticketId, `✓ Usuario aprobó reprogramación de visita para ${fecha} a las ${hora}`, false);
  
  // Notificar al agente
if (asignadoA) {
    notificarAgenteTodosCanales(asignadoA, 'reprog_aprobada', {
      ticketId: ticketId,
      folio: folio,
      fecha: fecha,
      hora: hora,
      tituloNotif: '📅 Reprogramación aprobada: #' + folio,
      mensajeNotif: `El usuario aprobó la reprogramación del ticket para ${fecha} ${hora}`,
      htmlCuerpo: `<h2>📅 Reprogramación Aprobada</h2><p>El usuario aceptó la nueva fecha de visita para el <strong>${fecha} a las ${hora}</strong> para el ticket #${folio}.</p>`
    });
  }
  
  return `Visita reprogramada para el ${fecha} a las ${hora}`;
}

function rechazarReprogramacion(ticketId) {
  const { headers, rows } = _readTableByHeader_(DB.TICKETS);
  const m = _headerMap_(headers);
  const sh = getSheet(DB.TICKETS);
  
  const idx = rows.findIndex(r => String(r[m.ID]) === String(ticketId));
  if (idx < 0) throw new Error('Ticket no encontrado');
  
  const row = rows[idx];
  const folio = row[m.Folio];
  const asignadoA = row[m.AsignadoA];
  
  row[m.NotasVisita] = (row[m.NotasVisita] || '').replace('PENDIENTE APROBACIÓN: ', 'RECHAZADO: ');
  row[m['ÚltimaActualización']] = new Date();
  
  sh.getRange(idx + 2, 1, 1, headers.length).setValues([row]);
  clearCache(DB.TICKETS);
  
  addSystemComment(ticketId, `✗ Usuario rechazó la reprogramación de visita`, false);
  
  // Notificar al agente
if (asignadoA) {
    notificarAgenteTodosCanales(asignadoA, 'reprog_rechazada', {
      ticketId: ticketId,
      folio: folio,
      tituloNotif: '❌ Reprogramación rechazada: #' + folio,
      mensajeNotif: `El usuario rechazó la reprogramación. Debe coordinar una nueva fecha.`,
      htmlCuerpo: `<h2>❌ Reprogramación Rechazada</h2><p>El usuario <strong>rechazó</strong> la fecha propuesta para el ticket #${folio}. Debes coordinar una nueva fecha.</p>`
    });
  }
  
  return 'Has rechazado la reprogramación. El agente deberá coordinar una nueva fecha contigo.';
}

function getMisTickets(email, filter) {
  if (!email) return { items: [], total: 0 };
  
  const data = getCachedData(DB.TICKETS);
  if (!data || !Array.isArray(data) || !data.length) {
    return { items: [], total: 0 };
  }
  
  const hdr = HEADERS.Tickets;
  const emailLower = String(email).toLowerCase().trim();
  
  // 1. Convertir a objetos
  let rows = data.map(r => {
    const obj = {};
    hdr.forEach((h, i) => obj[h] = r[i] ?? '');
    // Normalizar claves importantes
    obj['Área'] = obj['Área'] || obj['Area'] || '';
    obj['Título'] = obj['Título'] || obj['Titulo'] || '';
    obj['Estatus'] = obj['Estatus'] || '';
    obj['Folio'] = obj['Folio'] || '';
    return obj;
  });

  // 2. FILTRO ESTRICTO: Solo tickets donde el usuario es quien REPORTA
  // (Ignora si está asignado a él, eso va en el módulo de Agente)
  rows = rows.filter(obj => {
    const reporta = String(obj['ReportaEmail'] || '').toLowerCase().trim();
    return reporta === emailLower;
  });
  
  // 3. Filtros adicionales del frontend
  if (filter) {
    if (filter.estatus && filter.estatus.length) {
      const estatusValidos = filter.estatus.filter(e => e && e.toLowerCase() !== 'todos');
      if (estatusValidos.length > 0) {
        rows = rows.filter(x => estatusValidos.some(e => String(x.Estatus).toLowerCase() === e.toLowerCase()));
      }
    }
    if (filter.q) {
      const q = String(filter.q).toLowerCase();
      rows = rows.filter(x => 
        String(x.Folio).toLowerCase().includes(q) ||
        String(x['Título']).toLowerCase().includes(q)
      );
    }
  }
  
  // Ordenar descendente
  rows.sort((a, b) => new Date(b.Fecha) - new Date(a.Fecha));
  
  // Paginación
  const page = filter.page || 1;
  const pageSize = filter.pageSize || 20;
  const start = (page - 1) * pageSize;
  
  return { 
    items: rows.slice(start, start + pageSize), 
    total: rows.length 
  };
}

function getManualesUsuario(rolUsuario) {
  try {
    // ID de la carpeta de manuales en Drive
    const CARPETA_MANUALES_ID = '1l1cwstGyj9cK4FzOCZM5wZN-VpwGUFXi'; 
    const folder = DriveApp.getFolderById(CARPETA_MANUALES_ID);
    
    const manuales = [];
    

    // Función auxiliar inteligente para saber si el manual le toca a este rol
    const esParaRol = (nombreArchivo) => {
      // Si entra desde el login (sin rol), solo mostramos los genéricos o de usuario
      if (!rolUsuario) {
        const n = nombreArchivo.toLowerCase();
        return n.includes('general') || n.includes('todos') || n.includes('usuario');
      }
      
      const n = nombreArchivo.toLowerCase();
      const r = String(rolUsuario).toLowerCase();
      
      // 1. Administradores ven absolutamente todos los manuales
      if (r.includes('admin') || r.includes('superadmin')) return true;
      
      // 2. Manuales genéricos que todo el mundo (logueado) puede ver
      if (n.includes('general') || n.includes('todos') || n.includes('usuario')) return true;
      
      // 3. Filtros por área (muy útiles para los agentes/gerentes)
      if (r.includes('sistemas') && n.includes('sistemas')) return true;
      if (r.includes('mantenimiento') && (n.includes('mantenimiento') || n.includes('mtto'))) return true;
      if (r.includes('gerente') && n.includes('gerente')) return true;
      
      // 4. Coincidencia estricta (fallback)
      return n.includes(r);
    };

    // 1. Buscar Google Docs
    const files = folder.getFilesByType(MimeType.GOOGLE_DOCS);
    while (files.hasNext()) {
      const file = files.next();
      if (esParaRol(file.getName())) {
        manuales.push({
          id: file.getId(),
          nombre: file.getName(),
          url: file.getUrl(),
          embedUrl: `https://docs.google.com/document/d/${file.getId()}/preview`
        });
      }
    }
    
    // 2. Buscar PDFs
    const pdfs = folder.getFilesByType(MimeType.PDF);
    while (pdfs.hasNext()) {
      const file = pdfs.next();
      if (esParaRol(file.getName())) {
        manuales.push({
          id: file.getId(),
          nombre: file.getName(),
          url: file.getUrl(),
          embedUrl: `https://drive.google.com/file/d/${file.getId()}/preview`
        });
      }
    }
    
    return { ok: true, manuales };
  } catch (e) {
    Logger.log('Error obteniendo manuales: ' + e.message);
    return { ok: false, error: e.message, manuales: [] };
  }
}



// ============================================================================
// INTEGRACIÓN TELEGRAM BOT (RECEPCIÓN DE TICKETS)
// ============================================================================


function doPost(e) {
  // Preparamos la respuesta de éxito por adelantado
  const output = ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);

  try {
    if (!e || !e.postData || !e.postData.contents) return output;

    const update = JSON.parse(e.postData.contents);
    
    // Si es un mensaje de texto, lo procesamos
    if (update.message) {
      // Importante: Si esto falla, el catch lo atrapa, pero Telegram recibe su OK
      handleTelegramMessage(update.message);
    }
    
    return output;

  } catch (err) {
    // Si algo falla, lo registramos en tu hoja Debug pero NO le decimos a Telegram que falló
    // para que no reintente infinitamente.
    debugSheet("❌ Error en doPost: " + err.message);
    return output; 
  }
}
/**
 * Crea el ticket validando reglas de negocio (Área/Ubicación/Categoría)
 */
function crearTicketDesdeTelegram(usuario, texto, chatId, token) {
  // 1. Limpiar y preparar datos
  const contenido = texto.replace(/^\/nuevo\s*/i, '').trim();
  
  if (contenido.length < 5) {
    telegramSendSimple(chatId, "⚠️ Por favor describe el problema.\nEjemplo: <code>/nuevo Falla de Internet | No conecta</code>", token);
    return;
  }

  // Separar Título | Descripción
  let titulo, descripcion;
  if (contenido.includes('|')) {
    const partes = contenido.split('|');
    titulo = partes[0].trim();
    descripcion = partes.slice(1).join('|').trim();
  } else {
    titulo = contenido;
    descripcion = contenido; // Si no hay detalle, repetimos el título
  }

  // 2. USAR LA IA PARA CATEGORIZAR (Igual que el sistema web) [cite: 242]
  // Le pasamos la ubicación del usuario para que la IA priorice categorías de esa sede
  const sugerencia = sugerirCategoriaIA(descripcion, usuario.Ubicación);
  
  // 3. Validar y definir Área/Categoría
  let ticketArea = sugerencia.area;
  let ticketCategoria = sugerencia.categoria;
  
  // Si la IA no está segura (< 40%) o no encontró nada, usar defaults
  if (!ticketArea || sugerencia.confianza < 40) {
    // Si no detectamos área tecnológica/mtto, asignamos al área del usuario o "Sistemas" por defecto
    ticketArea = usuario.Área || 'Sistemas'; 
    ticketCategoria = 'General';
  }

  // 4. VALIDACIÓN ESTRICTA (Delimitar por Ubicación) [cite: 71]
  // Verificamos si la categoría existe para esa ubicación específica
  const esValida = validarCategoriaEnUbicacion(ticketArea, ticketCategoria, usuario.Ubicación);
  
  if (!esValida) {
    // Si la categoría sugerida no está disponible en su ubicación, forzamos "General" o "Soporte"
    // para evitar errores de asignación
    ticketCategoria = 'General'; 
  }

  // 5. Crear el Payload
  const payload = {
    reportaEmail: usuario.Email,
    reportaNombre: usuario.Nombre,
    area: ticketArea,                // Área detectada (Sistemas/Mantenimiento)
    ubicacion: usuario.Ubicación,    // Ubicación del usuario [cite: 1627]
    titulo: titulo,
    descripcion: descripcion,
    prioridad: 'Media',              // Prioridad inicial (el sistema la recalculará si la categoría tiene SLA)
    categoria: ticketCategoria,
    origen: 'Telegram Bot'
  };

  try {
    // 6. Crear Ticket (Backend procesará SLA y Asignación automática) [cite: 92]
    const resultado = createTicket(payload); 
    
    if (resultado.ok) {
      // Mensaje de éxito con detalles
      const msg = `✅ <b>Ticket Creado</b>\n` +
                  `Folio: <b>#${resultado.folio}</b>\n` +
                  `📂 Área: ${ticketArea}\n` +
                  `📋 Cat: ${ticketCategoria}\n` +
                  `📍 Ubic: ${usuario.Ubicación || 'N/A'}\n\n` +
                  `Un agente atenderá tu solicitud.`;
      telegramSendSimple(chatId, msg, token);
    } else {
      telegramSendSimple(chatId, "❌ Error creando ticket: " + resultado.error, token);
    }
  } catch (e) {
    telegramSendSimple(chatId, "❌ Error interno: " + e.message, token);
    Logger.log(e);
  }
}

/**
 * Helper para validar si una categoría es válida en una ubicación (Simula el filtro del UI)
 */
function validarCategoriaEnUbicacion(area, categoriaNombre, ubicacionUsuario) {
  if (!categoriaNombre || categoriaNombre === 'General') return true; // Siempre permitir generales
  
  // Obtener catálogo completo
  const catalogo = getCatalogosDesdeSheets().categories[area]; // [cite: 64, 65]
  if (!catalogo) return false;

  // Buscar la categoría
  const catObj = catalogo.find(c => c.nombre === categoriaNombre);
  if (!catObj) return false;

  // Si la categoría no tiene restricciones de ubicación, es válida
  if (!catObj.ubicaciones || catObj.ubicaciones.length === 0) return true;

  // Si tiene restricciones, verificar si la ubicación del usuario está permitida
  const ubicacionNorm = String(ubicacionUsuario || '').toLowerCase().trim();
  return catObj.ubicaciones.some(u => u.toLowerCase().trim() === ubicacionNorm);
}

/**
 * Busca al usuario en la hoja 'Usuarios' comparando el TelegramID
 */
function identificarUsuarioPorTelegram(chatId) {
  const { headers, rows } = _readTableByHeader_(DB.USERS);
  const m = _headerMap_(headers);
  
  // Asegúrate de tener una columna 'TelegramID' en tu hoja Usuarios
  if (m.TelegramID == null) return null; 

  const row = rows.find(r => String(r[m.TelegramID]).trim() === String(chatId).trim());
  
  if (!row) return null;
  
  return {
    Email: row[m.Email],
    Nombre: row[m.Nombre],
    Área: row[m['Área']],
    Ubicación: row[m['Ubicación']]
  };
}

/**
 * Envío simple de respuesta (auxiliar para no depender de la otra compleja)
 */
function telegramSendSimple(chatId, text, token) {
  const url = `https://api.telegram.org/bot${token}/sendMessage`;
  const payload = {
    chat_id: String(chatId),
    text: text,
    parse_mode: 'HTML'
  };
  try {
    UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload)
    });
  } catch (e) {
    Logger.log("Error enviando Telegram: " + e.message);
  }
}

// ============================================================================
// LÓGICA DEL BOT (INTELIGENTE: ACEPTA USUARIO O CORREO)
// ============================================================================

// ============================================================================
// BOT CONVERSACIONAL (PASO A PASO)
// ============================================================================

function handleTelegramMessage(msg) {
  const chatId = String(msg.chat.id);
  const text = (msg.text || '').trim();
  const token = getConfig('telegram_token'); 
  
  // 1. Identificar usuario (Revisa que tu función identificarUsuarioPorTelegram esté en el código)
  const usuario = identificarUsuarioPorTelegram(chatId);
  
  // --- FLUJO DE REGISTRO (Si no está vinculado) ---
  if (!usuario) {
    manejarRegistro(chatId, text, token);
    return;
  }

  // --- FLUJO CONVERSACIONAL ---
  
  // A. Comando cancelar (para salir de cualquier flujo)
  if (text === '/cancelar') {
    limpiarEstado(chatId);
    telegramSendSimple(chatId, "🚫 Operación cancelada. Escribe /nuevo para empezar.", token);
    return;
  }

  // B. Verificar si el usuario está en medio de un paso (tiene "memoria")
  const estadoActual = getEstado(chatId);

  if (estadoActual) {
    procesarPasoTicket(chatId, text, estadoActual, usuario, token);
    return;
  }

  // C. Comandos Iniciales (cuando no está haciendo nada)
  if (text.toLowerCase() === '/nuevo') {
    // INICIO DEL FLUJO: Guardamos que estamos esperando el TÍTULO
    guardarEstado(chatId, 'ESPERANDO_TITULO');
    telegramSendSimple(chatId, "📝 <b>Nuevo Ticket</b>\n\nPor favor, escribe el <b>título</b> o asunto del problema:\n\n<i>(Escribe /cancelar para salir)</i>", token);
  } 
  else if (text.toLowerCase() === '/ayuda' || text.toLowerCase() === '/start') {
    telegramSendSimple(chatId, `Hola <b>${usuario.Nombre}</b>.\n\nUsa <b>/nuevo</b> para crear un ticket paso a paso.`, token);
  } 
  else {
    telegramSendSimple(chatId, "No entendí ese comando. Escribe <b>/nuevo</b> para reportar un problema.", token);
  }
}

/**
 * Máquina de estados para la conversación
 */
function procesarPasoTicket(chatId, texto, estado, usuario, token) {
  
  // PASO 1: Recibimos el TÍTULO
  if (estado.paso === 'ESPERANDO_TITULO') {
    if (texto.length < 5) {
      telegramSendSimple(chatId, "⚠️ El título es muy corto. Por favor sé más descriptivo:", token);
      return;
    }
    
    // Guardamos el título y avanzamos al siguiente paso
    guardarEstado(chatId, 'ESPERANDO_DESC', { titulo: texto });
    telegramSendSimple(chatId, `✅ Título: <b>${texto}</b>\n\nAhora describe el <b>detalle</b> del problema (o envía una foto):`, token);
    return;
  }

  // PASO 2: Recibimos la DESCRIPCIÓN y creamos el ticket
  if (estado.paso === 'ESPERANDO_DESC') {
    const tituloGuardado = estado.datos.titulo;
    
    // Preparar datos finales
    const payload = {
      reportaEmail: usuario.Email,
      reportaNombre: usuario.Nombre,
      area: usuario.Área || 'Sistemas',
      ubicacion: usuario.Ubicación || 'Sin especificar',
      titulo: tituloGuardado,
      descripcion: texto,
      prioridad: 'Media',
      origen: 'Telegram Bot'
    };

    // Usamos la IA para categorizar (si tienes la función, sino usa default)
    if (typeof sugerirCategoriaIA === 'function') {
      const sugerencia = sugerirCategoriaIA(texto, usuario.Ubicación);
      payload.categoria = sugerencia.categoria || 'General';
      payload.area = sugerencia.area || payload.area; // Ajustar área si la IA detecta otra
    } else {
      payload.categoria = 'General';
    }

    telegramSendSimple(chatId, "⏳ Creando ticket, espera un momento...", token);

    try {
      [cite_start]// Crear ticket en Sheets [cite: 92]
      const resultado = createTicket(payload);
      
      if (resultado.ok) {
        telegramSendSimple(chatId, `✅ <b>¡Ticket Creado!</b>\n\nFolio: <b>#${resultado.folio}</b>\nÁrea: ${payload.area}\n\nUn agente te contactará pronto.`, token);
      } else {
        telegramSendSimple(chatId, "❌ Hubo un error al guardar el ticket. Intenta de nuevo.", token);
      }
    } catch (e) {
      telegramSendSimple(chatId, "❌ Error interno: " + e.message, token);
    }

    // ¡Importante! Limpiar la memoria al terminar
    limpiarEstado(chatId);
  }
}

// --- FUNCIONES AUXILIARES DE MEMORIA (CACHE) ---

/** Guarda el paso actual y datos temporales (dura 10 min) */
function guardarEstado(chatId, paso, datos = {}) {
  const cache = CacheService.getScriptCache();
  const data = JSON.stringify({ paso: paso, datos: datos });
  cache.put('chat_' + chatId, data, 600); // 600 segundos = 10 minutos
}

/** Recupera el estado actual */
function getEstado(chatId) {
  const cache = CacheService.getScriptCache();
  const data = cache.get('chat_' + chatId);
  return data ? JSON.parse(data) : null;
}

/** Borra la memoria */
function limpiarEstado(chatId) {
  const cache = CacheService.getScriptCache();
  cache.remove('chat_' + chatId);
}

// --- AUXILIAR PARA REGISTRO (Lo que ya tenías pero separado) ---
function manejarRegistro(chatId, text, token) {
  if (text.toLowerCase().startsWith('/soy')) {
    const inputUsuario = text.substring(4).trim();
    // Aquí llamas a tu función vincularUsuarioTelegram corregida
    const resultado = vincularUsuarioTelegram(inputUsuario, chatId);
    telegramSendSimple(chatId, resultado.mensaje, token);
  } else {
    telegramSendSimple(chatId, "👋 Bienvenido. Para empezar dime quién eres:\n<code>/soy tu_usuario</code>", token);
  }
}

function vincularUsuarioTelegram(usuarioInput, chatId) {
  debugSheet(`Iniciando vinculación inteligente para: "${usuarioInput}" en ChatID: ${chatId}`);

  return withLock_(() => {
    const { headers, rows } = _readTableByHeader_(DB.USERS);
    const m = _headerMap_(headers);

    if (m.Email === undefined) {
      debugSheet("ERROR CRÍTICO: No existe columna 'Email' o 'Usuario'.");
      return { ok: false, mensaje: "⛔ Error técnico: No encuentro la columna de usuarios." };
    }
    
    // Extraemos solo la parte del usuario (antes del @) del input recibido
    // Ejemplo: si recibe "rgnava@bexalta.com" -> busca "rgnava"
    // Ejemplo: si recibe "RGNava" -> busca "rgnava"
    const usuarioBuscado = String(usuarioInput).toLowerCase().trim().split('@')[0];
    
    debugSheet(`🔎 Buscando usuario base: "${usuarioBuscado}"`);

    // Buscar en la tabla comparando también solo la base del usuario
    const rowIndex = rows.findIndex((r, i) => {
      const valorCelda = String(r[m.Email] || '').toLowerCase().trim();
      const usuarioCelda = valorCelda.split('@')[0]; // Quitamos el dominio si la celda lo tuviera
      
      // Loguear solo las primeras 3 filas para verificar
      if (i < 3) debugSheet(`   Fila ${i+2}: "${usuarioCelda}" vs "${usuarioBuscado}"`);
      
      return usuarioCelda === usuarioBuscado;
    });
    
    if (rowIndex === -1) {
      debugSheet(`❌ No se encontró coincidencia.`);
      return { ok: false, mensaje: `⛔ No encontré el usuario "${usuarioBuscado}" en la base de datos.` };
    }
    
    debugSheet(`✅ ¡Encontrado! Fila ${rowIndex + 2}`);

    const sh = getSheet(DB.USERS);
    
    // Determinar columna TelegramID o crearla
    let colTelegram = m.TelegramID;
    if (colTelegram == null) {
      const currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0];
      colTelegram = currentHeaders.indexOf('TelegramID');
      if (colTelegram === -1) {
        colTelegram = currentHeaders.length; 
        sh.getRange(1, colTelegram + 1).setValue('TelegramID');
      }
    }
    
    // Guardar el Chat ID
    sh.getRange(rowIndex + 2, colTelegram + 1).setValue(String(chatId));
    
    clearCache(DB.USERS); 
    
    return { ok: true, mensaje: `✅ <b>¡Vinculación exitosa!</b>\n\nUsuario: <b>${usuarioBuscado}</b>\nYa puedes crear tickets con /nuevo.` };
  });
}


/**
 * Busca la Prioridad y SLA correctos.
 * VERSIÓN DEPURADA: Ajustada para estructura Col A=Nombre, C=Ubicación, D=Prioridad, E=SLA
 */
function obtenerConfiguracionSLA(area, categoriaNombre, ubicacionUsuario) {
  // Valores por defecto
  const resultadoDefault = { prioridad: 'Media', sla: 24 }; 

  if (!area || !categoriaNombre) return resultadoDefault;

  // 1. SELECCIÓN DE HOJA
  let nombreHoja = '';
  const areaNorm = String(area).trim().toLowerCase();

  if (areaNorm === 'sistemas') nombreHoja = 'Categorias_TI';
  else if (areaNorm === 'mantenimiento') nombreHoja = 'Categorias_Mtto';
  else nombreHoja = 'Categorias_' + area;

  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (!sh) {
    console.warn(`⚠️ No existe la hoja: ${nombreHoja}`);
    return resultadoDefault;
  }

  const data = sh.getDataRange().getValues();
  
  // Helper de normalización
  const normalizar = (txt) => {
    return String(txt || '')
      .toLowerCase()
      .trim()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/[\u2013\u2014]/g, "-")
      .replace(/\s+/g, " ");
  };

  const catBuscada = normalizar(categoriaNombre);
  const ubicacionBuscada = normalizar(ubicacionUsuario);

  let mejorMatch = null;
  let matchGeneral = null;

  console.log(`🔍 Buscando: Cat=[${catBuscada}] | Ubic=[${ubicacionBuscada}] en hoja [${nombreHoja}]`);

  // 2. RECORRIDO (Saltando encabezado fila 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // === MAPEO DE COLUMNAS (Ajustado a tu estructura) ===
    // Columna A (Índice 0) = Nombre Categoría
    // Columna C (Índice 2) = Ubicación (Saltamos B que parece ser Área)
    // Columna D (Índice 3) = Prioridad
    // Columna E (Índice 4) = SLA
    
    const catRow = normalizar(row[0]); // <--- CAMBIO AQUÍ: Columna A
    
    // Debug para la primera fila (para verificar si estamos leyendo bien)
    if (i === 1) {
      console.log(`📝 Leyendo Fila 1 de datos: Col A="${row[0]}", Col C="${row[2]}", Col D="${row[3]}"`);
    }

    if (catRow === catBuscada) {
      const ubicacionRaw = String(row[2] || ''); // <--- CAMBIO AQUÍ: Columna C
      const prioridadRow = row[3];              // Columna D
      const slaRow = row[4];                    // Columna E

      // Lógica de coincidencia
      if (ubicacionRaw.trim() === '') {
        matchGeneral = { prioridad: prioridadRow, sla: slaRow };
      } else {
        const listaUbicaciones = ubicacionRaw.split(',').map(u => normalizar(u));
        
        // Coincidencia flexible
        const matchEncontrado = listaUbicaciones.some(u => 
          u === ubicacionBuscada || 
          ubicacionBuscada.includes(u) || 
          (u.length > 3 && u.includes(ubicacionBuscada))
        );

        if (matchEncontrado) {
          mejorMatch = { prioridad: prioridadRow, sla: slaRow };
          console.log(`🎯 Match exacto en fila ${i+1}: Prio=${prioridadRow}, SLA=${slaRow}`);
          break; 
        }
      }
    }
  }

  if (mejorMatch) return mejorMatch;
  if (matchGeneral) {
    console.log(`ℹ️ Usando match general de la categoría (sin ubicación específica).`);
    return matchGeneral;
  }
  
  console.warn("⚠️ No se encontró coincidencia. Usando Default 24h.");
  return resultadoDefault;
}


/**
 * NUEVA — Registra un escalamiento en la tabla de auditoría EscalamientosLog.
 * Retorna el ID del registro para poder actualizarlo después (aprobación/rechazo).
 */
function registrarEscalamiento_(ticketId, folio, area, solicitante, gerente, motivo, nivelUrgencia) {
  try {
    const sh = getSheet('EscalamientosLog');
    const id = genId();
    
    sh.appendRow([
      id,
      ticketId,
      folio,
      area,
      solicitante,
      gerente || 'Sin gerente',
      new Date(),        // FechaSolicitud
      '',                // FechaRespuesta (se llena al aprobar/rechazar)
      'Pendiente',       // Estado inicial
      motivo,
      nivelUrgencia,
      ''                 // Respuesta del gerente
    ]);

    Logger.log(`✅ Escalamiento registrado: ${id} para ticket ${folio}`);
    return id;

  } catch (e) {
    Logger.log('Error registrando escalamiento: ' + e.message);
    return null;
  }
}


/**
 * NUEVA — Actualiza el estado de un escalamiento cuando el gerente aprueba/rechaza.
 */
function actualizarEscalamiento_(ticketId, nuevoEstado, respuesta) {
  try {
    const sh = getSheet('EscalamientosLog');
    const lr = sh.getLastRow();
    if (lr <= 1) return;

    const data = sh.getRange(2, 1, lr - 1, 12).getValues();
    // Buscar el escalamiento pendiente más reciente para este ticket
    let idx = -1;
    for (let i = data.length - 1; i >= 0; i--) {
      if (String(data[i][1]).trim() === String(ticketId).trim() &&
          String(data[i][8]).trim() === 'Pendiente') {
        idx = i;
        break;
      }
    }

    if (idx < 0) {
      Logger.log('No se encontró escalamiento pendiente para: ' + ticketId);
      return;
    }

    // Actualizar FechaRespuesta (col 8), Estado (col 9), Respuesta (col 12)
    sh.getRange(idx + 2, 8).setValue(new Date());          // FechaRespuesta
    sh.getRange(idx + 2, 9).setValue(nuevoEstado);         // Estado: Aprobado/Rechazado
    sh.getRange(idx + 2, 12).setValue(respuesta || '');     // Respuesta

    Logger.log(`✅ Escalamiento actualizado a ${nuevoEstado} para ticket ${ticketId}`);

  } catch (e) {
    Logger.log('Error actualizando escalamiento: ' + e.message);
  }
}


/**
 * NUEVA — Ejecuta las notificaciones de un ticket recién creado.
 * Se invoca via trigger temporal para no bloquear la creación.
 * Lee los datos del ticket desde PropertiesService y ejecuta todas las notificaciones.
 */
function postCreacionTicketNotificaciones() {
  const props = PropertiesService.getScriptProperties();

  // Buscar tickets pendientes de notificación
  const allProps = props.getProperties();
  const pendientes = Object.keys(allProps).filter(k => k.startsWith('notif_pendiente_'));

  pendientes.forEach(key => {
    try {
      const payload = JSON.parse(allProps[key]);
      const {
        ticketId, folio, area, ubicacion, prioridad, titulo, descripcion,
        reportaNombre, asignadoA, requiereVisita, visitaFecha, visitaHora
      } = payload;

      // A) Notificación interna
      if (asignadoA) {
        notifyUser(asignadoA, 'nuevo_ticket', 'Nuevo ticket asignado',
          `Se te ha asignado el ticket #${folio}: "${titulo || 'Sin título'}"`,
          { ticketId, folio });
      }

      // B) Email al agente
      if (asignadoA) {
        try {
          notificarAgenteNuevoTicket(asignadoA, {
            folio, titulo: titulo || 'Sin título', area, ubicacion, prioridad,
            reportaNombre, descripcion: descripcion || '',
            visitaFecha: visitaFecha || '', visitaHora: visitaHora || ''
          });
        } catch (e) { Logger.log('⚠️ Error email agente (trigger): ' + e.message); }
      }

      // C) Telegram al agente
      if (asignadoA) {
        try {
          notificarAgenteTelegramGrupo(asignadoA, '🎫 <b>NUEVO TICKET ASIGNADO</b>', {
            folio, titulo: titulo || 'Sin título', area, ubicacion, prioridad,
            reporta: reportaNombre, descripcion: descripcion || ''
          });
        } catch (e) { Logger.log('⚠️ Error Telegram agente (trigger): ' + e.message); }
      }

      // D) Telegram al gerente del área
      try {
        const gerente = getGerenteAreaConTelegram(area);
        if (gerente && gerente.chatId) {
          notificarGerenteTelegram(area, 'Nuevo ticket en tu área', {
            folio, titulo: titulo || 'Sin título', area, prioridad
          });
        }
      } catch (e) { Logger.log('⚠️ Error Telegram gerente (trigger): ' + e.message); }

      // E) Telegram admin general
      try {
        const cfg = getConfig();
        if (cfg.telegram_chat_admin && cfg.telegram_token) {
          let msg = `🎫 <b>Nuevo Ticket</b>\n<b>${titulo || 'Sin título'}</b>\nÁrea: ${area}\nUbicación: ${ubicacion || 'No especificada'}\nPrioridad: ${prioridad}\nAsignado a: ${asignadoA || 'Sin asignar'}`;
          if (requiereVisita) msg += `\n📅 Visita: ${visitaFecha} ${visitaHora}`;
          telegramSend(msg, cfg.telegram_chat_admin);
        }
      } catch (e) { Logger.log('⚠️ Error Telegram admin (trigger): ' + e.message); }

      // Limpiar propiedad procesada
      props.deleteProperty(key);
      Logger.log(`✅ Notificaciones enviadas para ticket ${folio}`);

    } catch (e) {
      Logger.log(`❌ Error procesando notif pendiente ${key}: ${e.message}`);
      // Limpiar para no reintentar infinitamente
      props.deleteProperty(key);
    }
  });

  // Limpiar el trigger temporal (solo se ejecuta una vez)
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'postCreacionTicketNotificaciones') {
      ScriptApp.deleteTrigger(t);
    }
  });
}


function getDashboardData(emailManual) {
  try {
    const email = (emailManual && String(emailManual).trim())
      ? String(emailManual).trim()
      : ((Session.getActiveUser && Session.getActiveUser().getEmail()) || '');

    const user = getUser(email);
    if (!user || !user.email) return { items: [], total: 0 };

    const uMail = String(user.email || '').trim().toLowerCase();

    // Roles
    const SUPERADMINS_HARDCODED = ['rgnava@bexalta.com', 'rgnava@bexalta.mx', 'admin', 'RCEsquivel'];
    const esSuperAdmin = SUPERADMINS_HARDCODED.some(sa => uMail.includes(sa.toLowerCase().split('@')[0]));
    const areaGerente = getGerenteArea(user.email);
    const esGerente = !!areaGerente;
    const role = (user.rol || 'usuario').toLowerCase();

    // Datos crudos
    const data = getCachedData(DB.TICKETS);
    if (!data || !data.length) return { items: [], total: 0 };

    const hdr = HEADERS.Tickets;
    const idx = {};
    hdr.forEach((h, i) => idx[h] = i);
    const tz = Session.getScriptTimeZone();

    // Permisos (misma lógica que listTickets paso 5)
    let filteredData;
    if (esSuperAdmin) {
      filteredData = data;
    } else if (esGerente) {
      const areaNorm = String(areaGerente || '').toLowerCase();
      if (areaNorm === 'clientes') {
        const supervisados = getUsuariosSupervisados().map(u => String(u || '').trim().toLowerCase());
        filteredData = data.filter(r =>
          supervisados.includes(String(r[idx['ReportaEmail']] || '').trim().toLowerCase())
        );
      } else {
        filteredData = data.filter(r =>
          String(r[idx['Área']] || '').trim().toLowerCase() === areaNorm
        );
      }
    } else if (role.includes('agente')) {
      const agArea = role === 'agente_sistemas' ? 'sistemas' : 'mantenimiento';
      filteredData = data.filter(r =>
        String(r[idx['Área']] || '').trim().toLowerCase() === agArea
      );
    } else {
      filteredData = data.filter(r =>
        String(r[idx['ReportaEmail']] || '').trim().toLowerCase() === uMail
      );
    }

    // Helper para formatear fechas de Sheets
    function fmtDate(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, tz, "yyyy-MM-dd'T'HH:mm:ss");
      }
      return val ? String(val) : '';
    }

    // Construir objetos con los campos necesarios para TODO el dashboard
    const items = [];
    for (var i = 0; i < filteredData.length; i++) {
      var r = filteredData[i];
      items.push({
        // Campos originales (KPIs + charts)
        Estatus:       String(r[idx['Estatus']] || ''),
        Prioridad:     String(r[idx['Prioridad']] || ''),
        Vencimiento:   fmtDate(r[idx['Vencimiento']]),
        Área:          String(r[idx['Área']] || r[idx['Area']] || ''),
        Fecha:         fmtDate(r[idx['Fecha']]),
        // Campos nuevos (secciones adicionales)
        ID:            String(r[idx['ID']] || ''),
        Folio:         String(r[idx['Folio']] || ''),
        Título:        String(r[idx['Título']] || r[idx['Titulo']] || ''),
        AsignadoA:     String(r[idx['AsignadoA']] || ''),
        ReportaNombre: String(r[idx['ReportaNombre']] || ''),
        Ubicación:     String(r[idx['Ubicación']] || r[idx['Ubicacion']] || '')
      });
    }

    return { items: items, total: items.length };

  } catch (e) {
    Logger.log('[getDashboardData] ❌ ERROR: ' + e.message);
    return { items: [], total: 0, error: e.message };
  }
}

// DESPUÉS — con SuperAdmin, caché y eliminación de duplicado
function getEscalamientosPendientes(emailGerente) {
  try {
    var user = getUser(emailGerente || '');
    var emailNorm = String(emailGerente || '').trim().toLowerCase();

    // ── Detección SuperAdmin (igual que el resto del sistema) ──
    var SUPERADMINS_HARDCODED = ['rgnava', 'rcesquivel', 'admin'];
    var esSuperAdmin = SUPERADMINS_HARDCODED.some(function(sa) {
      return emailNorm.indexOf(sa.toLowerCase()) !== -1;
    });

    var esAdmin = esSuperAdmin || (user.rol || '').toLowerCase() === 'admin';

    // ── Determinar áreas del gerente ──
    var areasGerente = [];
    if (esAdmin) {
      // SuperAdmin / Admin: ve TODAS las áreas, no necesitamos filtrar
      areasGerente = []; // señal de "todas"
    } else {
      // Buscar en ConfigGerentes las áreas que maneja
      try {
        var cfg = _readTableByHeader_('ConfigGerentes');
        var cm = _headerMap_(cfg.headers);
        cfg.rows.forEach(function(r) {
          var emailCfg = String(r[cm.Email] || '').toLowerCase().trim();
          if (emailCfg === emailNorm) {
            var areaCfg = String(r[cm.Area] || r[cm['Área']] || '').toLowerCase().trim();
            if (areaCfg) areasGerente.push(areaCfg);
          }
        });
      } catch (e) {
        Logger.log('Error leyendo ConfigGerentes: ' + e.message);
      }
      // Fallback: usar área del usuario
      if (areasGerente.length === 0) {
        var areaUser = (user.area || '').toLowerCase().trim();
        if (areaUser) areasGerente.push(areaUser);
      }
      // Si no es admin y no tiene áreas, no ve nada
      if (areasGerente.length === 0) {
        return { ok: true, pendientes: [] };
      }
    }

    // ── Usar caché en lugar de _readTableByHeader_ para mejor rendimiento ──
    var data = getCachedData(DB.TICKETS);
    if (!data || !data.length) return { ok: true, pendientes: [] };

    var hdr = HEADERS.Tickets;
    var idx = {};
    hdr.forEach(function(h, i) { idx[h] = i; });

    var f = function(s) { return String(s || '').trim().toLowerCase(); };
    var pendientes = [];

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      if (f(r[idx['Estatus']]) !== 'escalado') continue;

      var areaTicket = f(r[idx['Área']]);

      // SuperAdmin/Admin: sin filtro de área
      if (!esAdmin && areasGerente.indexOf(areaTicket) < 0) continue;

      var fechaEsc = r[idx['FechaEscalamiento']];
      var fechaStr = '';
      if (fechaEsc instanceof Date) {
        fechaStr = Utilities.formatDate(fechaEsc, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      } else {
        fechaStr = String(fechaEsc || '');
      }

      pendientes.push({
        ticketId:    r[idx['ID']],
        folio:       r[idx['Folio']],
        titulo:      r[idx['Título']] || r[idx['Titulo']] || 'Sin título',
        area:        r[idx['Área']] || '',
        ubicacion:   r[idx['Ubicación']] || '',
        prioridad:   r[idx['Prioridad']] || '',
        asignadoA:   r[idx['AsignadoA']] || 'Sin asignar',
        reportaNombre: r[idx['ReportaNombre']] || '',
        motivo:      r[idx['MotivoEscalamiento']] || 'Sin motivo',
        fecha:       fechaStr,
        solicitante: r[idx['SolicitanteEscalamiento']] || '',
        vencimiento: r[idx['Vencimiento']] ? new Date(r[idx['Vencimiento']]).toLocaleString('es-MX') : ''
      });
    }

    // Ordenar: más recientes primero
    pendientes.sort(function(a, b) {
      return new Date(b.fecha || 0) - new Date(a.fecha || 0);
    });

    return { ok: true, pendientes: pendientes };

  } catch (e) {
    Logger.log('Error en getEscalamientosPendientes: ' + e.message);
    return { ok: false, error: e.message, pendientes: [] };
  }
}

function aprobarEscalamientoDesdePanel(ticketId, decision, motivoRechazo, emailGerente) {
  return withLock_(function() {
    try {
      var user = getUser(emailGerente || '');
      var esAdmin = (user.rol || '').toLowerCase() === 'admin';
      var esGerente = false;
      
      // FIX: Búsqueda correcta en ConfigGerentes usando GerenteEmail
      try {
        var cfg = _readTableByHeader_('ConfigGerentes');
        var cm = _headerMap_(cfg.headers);
        cfg.rows.forEach(function(r) {
          // Buscamos en GerenteEmail o Email por si cambias el nombre después
          if (String(r[cm.GerenteEmail] || r[cm.Email] || '').toLowerCase().trim() === emailGerente.toLowerCase().trim()) {
            esGerente = true;
          }
        });
      } catch (e) {}

      // Respaldo extra por si el gerente está en el código duro (GERENTES_AREAS)
      if (!esGerente) {
        var areaDetectada = getGerenteArea(emailGerente);
        if (areaDetectada) esGerente = true;
      }
      
      if (!esAdmin && !esGerente) {
        return { ok: false, error: '⛔ Permiso denegado: Tu cuenta no está reconocida como Gerente de esta área.' };
      }
      
      // ... A PARTIR DE AQUÍ ES TU MISMO CÓDIGO ORIGINAL ...
      var _data = _readTableByHeader_(DB.TICKETS);
      var headers = _data.headers;
      var rows = _data.rows;
      var m = _headerMap_(headers);
      var sh = getSheet(DB.TICKETS);
      
      var rowIndex = rows.findIndex(function(r) { return String(r[m.ID]) === String(ticketId); });
      if (rowIndex < 0) return { ok: false, error: 'Ticket no encontrado' };
      
      var row = rows[rowIndex];
      var folio = row[m.Folio];
      var titulo = row[m['Título']] || '';
      var area = row[m['Área']] || '';
      var solicitanteEmail = row[m.SolicitanteEscalamiento] || row[m.AsignadoA] || '';
      var asignadoA = row[m.AsignadoA] || '';
      
      if (decision === 'aprobar') {
        row[m.Estatus] = 'En Proceso';
        row[m['ÚltimaActualización']] = new Date();
        
        var prioActual = String(row[m.Prioridad] || '').toLowerCase();
        if (prioActual !== 'crítica' && prioActual !== 'critica') {
          row[m.Prioridad] = 'Alta';
        }
        
        sh.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
        clearCache(DB.TICKETS);
        
        addSystemComment(ticketId, '✅ ESCALAMIENTO APROBADO\nAprobado por: ' + (user.nombre || emailGerente) + '\nEl ticket ha sido marcado para atención prioritaria.', true);
        registrarBitacora(ticketId, 'Escalamiento aprobado', 'Aprobado por ' + (user.nombre || emailGerente));
        
        if (asignadoA) {
          notificarAgenteTodosCanales(asignadoA, 'escalamiento_aprobado', {
            ticketId: ticketId, folio: folio, titulo: titulo, area: area,
            tituloNotif: '✅ Escalamiento Aprobado: #' + folio,
            mensajeNotif: 'El gerente ha aprobado el escalamiento. Atención prioritaria requerida.'
          });
        }
        if (solicitanteEmail && solicitanteEmail !== asignadoA) {
          crearNotificacion(solicitanteEmail, 'escalamiento_aprobado', '✅ Escalamiento aprobado', 'Tu solicitud de escalamiento para el ticket #' + folio + ' ha sido APROBADA.', ticketId);
        }
        return { ok: true, message: 'Escalamiento aprobado. Ticket en atención prioritaria.' };
        
      } else if (decision === 'rechazar') {
        row[m.Estatus] = 'En Proceso';
        row[m['ÚltimaActualización']] = new Date();
        sh.getRange(rowIndex + 2, 1, 1, headers.length).setValues([row]);
        clearCache(DB.TICKETS);
        
        var motivoTexto = motivoRechazo || 'Sin motivo especificado';
        addSystemComment(ticketId, '❌ ESCALAMIENTO RECHAZADO\nRechazado por: ' + (user.nombre || emailGerente) + '\nMotivo: ' + motivoTexto + '\nEl ticket continúa en proceso normal.', true);
        registrarBitacora(ticketId, 'Escalamiento rechazado', 'Rechazado por ' + (user.nombre || emailGerente) + '. Motivo: ' + motivoTexto);
        
        if (asignadoA) {
          notificarAgenteTodosCanales(asignadoA, 'escalamiento_rechazado', {
            ticketId: ticketId, folio: folio, titulo: titulo, area: area, motivo: motivoTexto,
            tituloNotif: '❌ Escalamiento Rechazado: #' + folio,
            mensajeNotif: 'El gerente ha rechazado el escalamiento. Motivo: ' + motivoTexto
          });
        }
        if (solicitanteEmail && solicitanteEmail !== asignadoA) {
          crearNotificacion(solicitanteEmail, 'escalamiento_rechazado', '❌ Escalamiento rechazado', 'Tu solicitud de escalamiento para el ticket #' + folio + ' fue rechazada. Motivo: ' + motivoTexto, ticketId);
        }
        return { ok: true, message: 'Escalamiento rechazado. Ticket regresa a proceso normal.' };
      } else {
        return { ok: false, error: 'Decisión no válida. Use "aprobar" o "rechazar".' };
      }
    } catch (e) {
      Logger.log('Error en aprobarEscalamientoDesdePanel: ' + e.message);
      return { ok: false, error: e.message };
    }
  });
}

// Agregar fila de categoría a Categorias_TI o Categorias_Mtto
function addCategoryRow(nombre, area, ubicaciones, prioridad, sla, agente, keywords) {
  if (!nombre || !area) throw new Error('Nombre y Área son requeridos');

  var sheetName = area.toLowerCase() === 'mantenimiento' ? 'Categorias_Mtto' : 'Categorias_TI';
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sh) throw new Error('Hoja ' + sheetName + ' no encontrada');

  // Columnas: A=Nombre, B=Area, C=Ubicaciones, D=Prioridad, E=SLA, F=Agente, G=PalabrasClave
  sh.appendRow([nombre, area, ubicaciones || '', prioridad || '', Number(sla) || 0, agente || '', keywords || '']);
  return { ok: true };
}


// Eliminar fila de categoría de Categorias_TI o Categorias_Mtto
function deleteCategoryRow(nombre, area) {
  if (!nombre) throw new Error('Nombre requerido');

  // Buscar en ambas hojas
  var sheets = ['Categorias_TI', 'Categorias_Mtto'];
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  for (var s = 0; s < sheets.length; s++) {
    var sh = ss.getSheetByName(sheets[s]);
    if (!sh || sh.getLastRow() <= 1) continue;

    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
    for (var i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === nombre && (!area || String(data[i][1]).trim() === area)) {
        sh.deleteRow(i + 2);
        return { ok: true };
      }
    }
  }
  throw new Error('Categoría no encontrada: ' + nombre);
}


// CRUD Estatus
function addStatus(nombre) {
  if (!nombre) throw new Error('Nombre requerido');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Estatus');
  if (!sh) throw new Error('Hoja Estatus no encontrada');
  sh.appendRow([nombre]);
  return { ok: true };
}

function deleteStatus(nombre) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Estatus');
  if (!sh) throw new Error('Hoja Estatus no encontrada');
  var lr = sh.getLastRow();
  if (lr <= 1) throw new Error('Estatus no encontrado');
  var data = sh.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === nombre) {
      sh.deleteRow(i + 2);
      return { ok: true };
    }
  }
  throw new Error('Estatus no encontrado: ' + nombre);
}

function getHistorialVisitas(ticketId) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('VisitasProgramadas') || 
                SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Visitas programadas');
                
    if (!sheet) return { ok: false, error: 'No existe la hoja de Visitas' };
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { ok: true, visitas: [] };
    
    const hdr = data[0];
    const idx = {};
    
    // Mapeo de cabeceras en minúsculas sin espacios
    hdr.forEach((h, i) => {
       if (h) idx[String(h).toLowerCase().replace(/\s/g, '')] = i;
    });
    
    // ========================================================================
    // FIX MAESTRO: Si no encuentra el nombre exacto, usamos el número de columna.
    // En tu hoja: Columna C (2) = TicketID | Columna G (6) = Fecha | Columna H (7) = Hora
    // ========================================================================
    const colTicketId = idx['ticketid'] ?? idx['idticket'] ?? 2;
    const colRegistro = idx['timestamp'] ?? idx['marcatemporal'] ?? idx['fecharegistro'] ?? idx['fecha'] ?? 1; // Col B
    const colFechaVisita = idx['fechavisita'] ?? 6; // Col G
    const colHoraVisita = idx['horavisita'] ?? 7;   // Col H
    const colAccion = idx['accion'] ?? idx['acción'] ?? idx['estado'] ?? 5; // Col F
    const colAgente = idx['agente'] ?? idx['usuario'] ?? idx['asignadoa'] ?? 4; // Col E
    const colNotas = idx['notas'] ?? idx['notasvisita'] ?? idx['observaciones'] ?? 8; // Col I

    const visitas = [];
    const tz = Session.getScriptTimeZone();
    
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      
      // Filtrar solo las filas que corresponden al Ticket actual
      if (String(r[colTicketId]).trim() === String(ticketId).trim()) {
        
        // 1. Fecha en que se hizo el registro (Timestamp - Columna B)
        let fechaReg = r[colRegistro] || '';
        if (fechaReg instanceof Date) {
            fechaReg = Utilities.formatDate(fechaReg, tz, "yyyy-MM-dd HH:mm");
        }
        
        // 2. Fecha Programada de la Visita (Columna G)
        let fv = r[colFechaVisita] || '';
        if (fv instanceof Date) {
            fv = Utilities.formatDate(fv, tz, "yyyy-MM-dd");
        } else if (fv) {
            // Limpieza robusta de fecha (Convierte 13/3/2026 a 2026-03-13 para que el panel lo lea bien)
            let strDate = String(fv).split('T')[0].trim();
            if (strDate.includes('/')) {
               const p = strDate.split('/');
               if (p.length === 3) {
                 const year = p[2].length === 4 ? p[2] : p[2];
                 fv = `${year}-${p[1].padStart(2, '0')}-${p[0].padStart(2, '0')}`;
               }
            } else {
               fv = strDate;
            }
        }
        
        // 3. Hora Programada de la Visita (Columna H)
        let hv = r[colHoraVisita] || '';
        if (hv instanceof Date) {
            hv = Utilities.formatDate(hv, tz, "HH:mm");
        } else if (hv) {
            hv = String(hv).replace(/'/g, '').trim();
        }
        
        visitas.push({
          fechaRegistro: fechaReg,
          agente: r[colAgente] || 'Sistema',
          accion: r[colAccion] || 'Programada',
          fechaVisita: fv,
          horaVisita: hv,
          notas: r[colNotas] || ''
        });
      }
    }
    
    // Ordenar de más reciente a más antigua (La primera será la "VIGENTE")
    visitas.sort((a, b) => new Date(b.fechaRegistro || 0) - new Date(a.fechaRegistro || 0));
    
    return { ok: true, visitas: visitas };
    
  } catch (e) {
    Logger.log('Error en getHistorialVisitas: ' + e.message);
    return { ok: false, error: e.message };
  }
}


/**
 * Consulta ultrarrápida que incluye el autor del último cambio.
 */
function checkTicketUpdates(ticketId) {
  try {
    const data = getCachedData('Tickets'); 
    const hdr = HEADERS.Tickets;
    const idxId = hdr.indexOf('ID');
    const idxUpdate = hdr.indexOf('ÚltimaActualización');
    const idxUser = hdr.indexOf('ActualizadoPor'); // Asegúrate de tener esta columna o usa 'AsignadoA'

    const row = data.find(r => String(r[idxId]) === String(ticketId));
    if (!row) return null;

    return {
      id: ticketId,
      ts: row[idxUpdate] ? new Date(row[idxUpdate]).getTime() : 0,
      user: row[idxUser] || '' // Retornamos quién hizo el cambio
    };
  } catch (e) {
    return null;
  }
}

/**
 * Obtiene la lista de usuarios activos para llenar el Select de cotizaciones.
 */
function getListaUsuariosCotizacion() {
  try {
    // Si tu hoja se llama distinto a 'Usuarios', cámbialo aquí
    const { headers, rows } = _readTableByHeader_('Usuarios'); 
    const m = _headerMap_(headers);
    
    const lista = [];
    rows.forEach(r => {
      const email = r[m.Email] || r[m.Correo] || '';
      const nombre = r[m.Nombre] || r[m.Name] || email.split('@')[0];
      const estatus = r[m.Estatus] || r[m.Estado] || 'Activo';
      
      // Traemos solo a los que tienen correo y están activos
      if (email && estatus.toLowerCase() === 'activo') {
        lista.push({ nombre: nombre, email: email });
      }
    });
    
    // Ordenar alfabéticamente por nombre
    lista.sort((a, b) => a.nombre.localeCompare(b.nombre));
    
    return lista;
  } catch(e) {
    Logger.log("Error al obtener usuarios para cotización: " + e.message);
    return []; // Devuelve vacío si falla, pero no rompe el sistema
  }
}

/**
 * Valida estrictamente que una fecha y hora caigan en horario laboral
 */
function validarFechaHoraLaboral(fechaStr, horaStr) {
  if (!fechaStr || !horaStr) return { valido: false, error: 'Fecha y hora son obligatorias.' };
  
  // 1. Validar que la fecha sea un día laboral (Lunes a Viernes y no festivo)
  const partes = fechaStr.split('-');
  const dt = new Date(partes[0], partes[1] - 1, partes[2], 12, 0, 0); // Forzar mediodía para evitar saltos de zona horaria
  
  if (!esDiaLaboral(dt)) {
    return { 
      valido: false, 
      error: '📅 La fecha seleccionada cae en fin de semana o día festivo. Por favor elige un día hábil.' 
    };
  }
  
  // 2. Validar que la hora esté dentro del rango permitido (ej. 8:00 a 18:00)
  const h = parseInt(horaStr.split(':')[0], 10);
  if (h < HORARIO_LABORAL.horaInicio || h >= HORARIO_LABORAL.horaFin) {
    return { 
      valido: false, 
      error: `⏰ La hora (${horaStr}) está fuera del horario laboral (${HORARIO_LABORAL.horaInicio}:00 hrs - ${HORARIO_LABORAL.horaFin}:00 hrs).` 
    };
  }
  
  return { valido: true };
}

// ============================================================
// DÍAS FESTIVOS — Módulo completo
// ============================================================

var DB_FESTIVOS = 'DiasFestivos'; // Nombre de la hoja en el Spreadsheet

/**
 * Inicializa la hoja DiasFestivos si no existe.
 * Llama una sola vez o al arrancar el sistema.
 */
function inicializarHojaFestivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(DB_FESTIVOS);

  if (!sh) {
    sh = ss.insertSheet(DB_FESTIVOS);
    sh.getRange('A1:D1').setValues([['Fecha','Descripcion','Tipo','Activo']]);
    sh.getRange('A1:D1').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    Logger.log('✅ Hoja DiasFestivos creada.');
  }

  // Poblar con festivos fijos de México para el año actual y el siguiente
  const añoActual = new Date().getFullYear();
  _poblarFestivosMexico_(sh, añoActual);
  _poblarFestivosMexico_(sh, añoActual + 1);

  return { ok: true, hoja: DB_FESTIVOS };
}

/**
 * Festivos oficiales México (Ley Federal del Trabajo)
 */
function _poblarFestivosMexico_(sh, año) {
  const festivos = [
    ['01/01/' + año, 'Año Nuevo',                       'Oficial', true],
    ['05/02/' + año, 'Día de la Constitución',           'Oficial', true],
    ['21/03/' + año, 'Natalicio de Benito Juárez',       'Oficial', true],
    ['01/05/' + año, 'Día del Trabajo',                  'Oficial', true],
    ['16/09/' + año, 'Día de la Independencia',          'Oficial', true],
    ['02/11/' + año, 'Día de Muertos',                   'Festivo', true],
    ['20/11/' + año, 'Revolución Mexicana',              'Oficial', true],
    ['12/12/' + año, 'Virgen de Guadalupe',              'Festivo', true],
    ['25/12/' + año, 'Navidad',                          'Oficial', true],
  ];

  // Verificar qué fechas ya existen para no duplicar
  const dataExistente = sh.getDataRange().getValues();
  const fechasExistentes = new Set(
    dataExistente.slice(1).map(r => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return isNaN(d) ? '' : d.toDateString();
    })
  );

  const nuevas = festivos.filter(f => {
    const partes = f[0].split('/');
    const d = new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0]));
    return !fechasExistentes.has(d.toDateString());
  }).map(f => {
    const partes = f[0].split('/');
    return [new Date(Number(partes[2]), Number(partes[1]) - 1, Number(partes[0])), f[1], f[2], f[3]];
  });

  if (nuevas.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, 4).setValues(nuevas);
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy');
    Logger.log(`✅ ${nuevas.length} festivos agregados para ${año}.`);
  }
}

/**
 * Retorna Set de fechas festivas activas (strings 'YYYY-MM-DD')
 */
function getFestivosActivos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(DB_FESTIVOS);
    if (!sh) return new Set();

    const rows = sh.getDataRange().getValues().slice(1);
    const festivos = new Set();

    rows.forEach(r => {
      const activo = r[3];
      if (activo === false || String(activo).toLowerCase() === 'false') return;
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      if (!isNaN(d)) festivos.add(Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    });

    return festivos;
  } catch(e) {
    Logger.log('Error getFestivosActivos: ' + e.message);
    return new Set();
  }
}

/**
 * Calcula días hábiles desde una fecha base sumando N días,
 * saltando sábados, domingos y festivos activos.
 * @param {Date} fechaBase
 * @param {number} diasHabiles
 * @returns {Date}
 */
function sumarDiasHabiles(fechaBase, diasHabiles) {
  const festivos = getFestivosActivos();
  const tz = Session.getScriptTimeZone();
  let fecha = new Date(fechaBase);
  let contador = 0;

  while (contador < diasHabiles) {
    fecha.setDate(fecha.getDate() + 1);
    const diaSemana = fecha.getDay(); // 0=Dom, 6=Sab
    const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    if (diaSemana !== 0 && diaSemana !== 6 && !festivos.has(fechaStr)) {
      contador++;
    }
  }
  return fecha;
}

/**
 * Calcula días hábiles ENTRE dos fechas (para SLA real).
 * @param {Date} inicio
 * @param {Date} fin
 * @returns {number}
 */
function calcularDiasHabiles(inicio, fin) {
  const festivos = getFestivosActivos();
  const tz = Session.getScriptTimeZone();
  let fecha = new Date(inicio);
  fecha.setHours(0, 0, 0, 0);
  let contador = 0;
  const finNorm = new Date(fin);
  finNorm.setHours(23, 59, 59, 0);

  while (fecha <= finNorm) {
    const diaSemana = fecha.getDay();
    const fechaStr = Utilities.formatDate(fecha, tz, 'yyyy-MM-dd');
    if (diaSemana !== 0 && diaSemana !== 6 && !festivos.has(fechaStr)) {
      contador++;
    }
    fecha.setDate(fecha.getDate() + 1);
  }
  return contador;
}

// ============================================================
// CRUD desde frontend (Admin)
// ============================================================

/**
 * Listar todos los festivos para el panel admin
 */
function listarFestivos() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(DB_FESTIVOS);
    if (!sh) return [];

    const rows = sh.getDataRange().getValues().slice(1);
    const tz = Session.getScriptTimeZone();

    return rows.map((r, i) => {
      const d = r[0] instanceof Date ? r[0] : new Date(r[0]);
      return {
        rowIndex: i + 2,
        fecha:       isNaN(d) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
        descripcion: String(r[1] || ''),
        tipo:        String(r[2] || 'Oficial'),
        activo:      r[3] !== false && String(r[3]).toLowerCase() !== 'false'
      };
    }).filter(r => r.fecha !== '');
  } catch(e) {
    Logger.log('Error listarFestivos: ' + e.message);
    return [];
  }
}

/**
 * Agregar o actualizar un festivo
 * @param {Object} festivo {fecha:'yyyy-MM-dd', descripcion, tipo, activo}
 * @param {number|null} rowIndex - null para nuevo, número para editar
 */
function guardarFestivo(festivo, rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(DB_FESTIVOS);
    if (!sh) { inicializarHojaFestivos(); sh = ss.getSheetByName(DB_FESTIVOS); }

    const partes = (festivo.fecha || '').split('-');
    if (partes.length !== 3) return { ok: false, error: 'Fecha inválida' };
    const fecha = new Date(Number(partes[0]), Number(partes[1]) - 1, Number(partes[2]));

    const fila = [fecha, festivo.descripcion || '', festivo.tipo || 'Oficial', festivo.activo !== false];

    if (rowIndex && rowIndex > 1) {
      sh.getRange(rowIndex, 1, 1, 4).setValues([fila]);
    } else {
      sh.appendRow(fila);
    }
    sh.getRange(2, 1, sh.getLastRow() - 1, 1).setNumberFormat('dd/MM/yyyy');
    return { ok: true };
  } catch(e) {
    Logger.log('Error guardarFestivo: ' + e.message);
    return { ok: false, error: e.message };
  }
}

/**
 * Eliminar un festivo por rowIndex
 */
function eliminarFestivo(rowIndex) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(DB_FESTIVOS);
    if (!sh || !rowIndex) return { ok: false, error: 'Parámetros inválidos' };
    sh.deleteRow(rowIndex);
    return { ok: true };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}
