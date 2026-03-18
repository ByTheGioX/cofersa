// ─── UTILIDADES ──────────────────────────────────────────────────
function cleanNum(n) {
  if (typeof n === 'number') return n;
  if (!n) return 0;
  // Tratar strings con comas como decimales y puntos como miles
  var s = String(n).replace(/[₡$ ]/g, '').trim();
  // Si tiene un patrón como 1.234,56
  if (s.includes(',') && s.includes('.')) {
    if (s.indexOf('.') < s.indexOf(',')) s = s.replace(/\./g, '').replace(',', '.');
    else s = s.replace(/,/g, '');
  } else if (s.includes(',')) {
    // Si solo tiene coma, asumimos que es decimal (formato CR/EU)
    s = s.replace(',', '.');
  }
  var val = parseFloat(s);
  return isNaN(val) ? 0 : val;
}

// ─── SERVIR LA APLICACIÓN ────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚀 Gestión Pro')
    .addItem('Inicializar Base de Datos', 'setupDatabase')
    .addToUi();
  setupDatabase();
}

function setupDatabase() {
  getDatabaseSheet();
  console.log('Base de datos inicializada correctamente.');
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Cobro Flow Pro | Gestión Premium')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ─── BASE DE DATOS AUTOCONFIGURABLE ──────────────────────────────
function getDatabaseSheet() {
  const props = PropertiesService.getScriptProperties();
  const MANUAL_SS_ID = '1BXrX2f8hMCSPLfh2qGmqL-riidzPaxvJ8c3lbNsYkmA';
  
  // Forzamos el ID en propiedades por si acaso
  props.setProperty('SS_ID', MANUAL_SS_ID);
  
  let ss = SpreadsheetApp.openById(MANUAL_SS_ID);
  if (!ss) throw new Error('No se pudo conectar con la base de datos oficial. Verifica permisos.');

  var sheetConfigs = [
    { name: 'Clientes',    headers: ['CODIGO','NOMBRE','TELEFONO','CORREO'] },
    { name: 'Promesas',    headers: ['ID','FECHA_CREACION','CLIENTE_COD','MONTO','FECHA_PROMESA','ESTADO','NOTAS','INTERVENTOR'] },
    { name: 'Incidencias', headers: ['ID','CLIENTE_COD','FACTURA','MONTO','MOTIVO','ESTADO','CANAL','LINK_GMAIL','FECHA','INTERVENTOR'] },
    { name: 'Llamadas',    headers: ['ID','FECHA','CLIENTE_COD','RESULTADO','INTERVENTOR','NOTAS'] },
    { name: 'Usuarios',    headers: ['EMAIL','PASSWORD','NOMBRE','ROL','ESTADO'] }
  ];

  sheetConfigs.forEach(function(config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
      sheet.appendRow(config.headers);
    } else {
      // Mantenimiento de Cabeceras: Verificar y reordenar/añadir si faltan
      var existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn() || 1).getValues()[0];
      var needsSync = false;
      config.headers.forEach(function(h, idx) {
        if (existingHeaders[idx] !== h) needsSync = true;
      });
      
      if (needsSync) {
        var lastRow = sheet.getLastRow();
        var data = [];
        if (lastRow > 1) {
          var oldData = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
          // Mapeo inteligente simple: si la columna existe por nombre, mantener valor
          data = oldData.map(function(row) {
            var newRow = new Array(config.headers.length).fill('');
            config.headers.forEach(function(h, newIdx) {
              var oldIdx = existingHeaders.indexOf(h);
              if (oldIdx !== -1) newRow[newIdx] = row[oldIdx];
            });
            return newRow;
          });
        }
        sheet.clear();
        sheet.appendRow(config.headers);
        if (data.length > 0) {
          sheet.getRange(2, 1, data.length, config.headers.length).setValues(data);
        }
      }
    }
    // Estilo estándar pro
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, config.headers.length)
      .setFontWeight('bold')
      .setBackground('#0d1117')
      .setFontColor('#58a6ff')
      .setHorizontalAlignment('center');
    
    // Auto dimensionar columnas para que se vea ordenado
    if (sheet.getLastColumn() > 0) sheet.autoResizeColumns(1, config.headers.length);
  });
  return ss;
}

// Acceso rápido a la Spreadsheet sin validaciones pesadas (Uso para lectura/login)
function getSS() {
  const MANUAL_SS_ID = '1BXrX2f8hMCSPLfh2qGmqL-riidzPaxvJ8c3lbNsYkmA';
  try {
    var ss = SpreadsheetApp.openById(MANUAL_SS_ID);
    if (!ss) throw new Error('Referencia de hoja nula.');
    return ss;
  } catch (e) {
    console.error('CRÍTICO: No se puede acceder a la Spreadsheet.', e);
    throw new Error('PERMISOS_DENEGADOS: No tienes acceso a la hoja de cálculo. ' + e.message);
  }
}

// ─── LECTURA ─────────────────────────────────────────────────────
function getDashboardData(userEmail) {
  var ss = getSS();
  console.log('--- LECTURA DASHBOARD ---');
  console.log('SS URL:', ss.getUrl());
  var usuarios = ss.getSheetByName('Usuarios').getDataRange().getDisplayValues().slice(1);
  var currentUser = usuarios.find(function(u) { return u[0].toLowerCase() === (userEmail||'').toLowerCase(); });
  var isAdmin = currentUser && currentUser[3] === 'ADMIN';

  // Leemos como DisplayValues para mantener ceros a la izquierda en códigos y teléfonos
  var promesas    = ss.getSheetByName('Promesas').getDataRange().getDisplayValues().slice(1);
  var incidencias = ss.getSheetByName('Incidencias').getDataRange().getDisplayValues().slice(1);
  var llamadas    = ss.getSheetByName('Llamadas').getDataRange().getDisplayValues().slice(1);
  var clientes    = ss.getSheetByName('Clientes').getDataRange().getDisplayValues().slice(1);

  // AISLAMIENTO DE DATOS: Si no es admin, solo ve lo propio + lo genérico antiguo
  if (!isAdmin && userEmail) {
    var refEmail = userEmail.toLowerCase();
    var genericTerms = ['usuario', 'sistema', ''];

    promesas = promesas.filter(function(p) { 
      var intv = (p[7]||'').toLowerCase().trim();
      return intv === refEmail || genericTerms.indexOf(intv) !== -1;
    });

    llamadas = llamadas.filter(function(l) { 
      var intv = (l[4]||'').toLowerCase().trim();
      return intv === refEmail || genericTerms.indexOf(intv) !== -1;
    });
    
    incidencias = incidencias.filter(function(inc) {
      var intv = (inc[9]||'').toLowerCase().trim();
      return intv === refEmail || genericTerms.indexOf(intv) !== -1;
    });
  }

  return {
    promesas:    promesas,
    incidencias: incidencias,
    llamadas:    llamadas,
    clientes:    clientes,
    dbUrl:       ss.getUrl(),
    isAdmin:     isAdmin,
    debug: {
      user: userEmail,
      isAdmin: isAdmin,
      counts: {
        promesas: promesas.length,
        incidencias: incidencias.length,
        llamadas: llamadas.length
      }
    }
  };
}

function loginUser(email, password) {
  try {
    var ss = getSS();
    console.log('--- INTENTO DE LOGIN ---');
    console.log('SS URL:', ss.getUrl());
    
    var sheet = ss.getSheetByName('Usuarios');
    if (!sheet) return { success: false, message: 'La hoja "Usuarios" no existe en la DB activa.' };

    var data = sheet.getDataRange().getDisplayValues();
    
    function clean(s) {
      if (!s) return '';
      // Limpieza agresiva de espacios invisibles y saltos de línea
      return String(s).replace(/[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF\r\n\t]/g, '').trim();
    }

    var emailInput = clean(email).toLowerCase();
    var passInput = clean(password);
    
    console.log('Email Input Normalized:', emailInput);
    
    for (var i = 1; i < data.length; i++) {
      var sheetEmail = clean(data[i][0]).toLowerCase();
      var sheetPass = clean(data[i][1]);
      
      if (sheetEmail === emailInput) {
        if (sheetPass === passInput) {
          if (data[i][4] === 'INACTIVO') return { success: false, message: 'Usuario inactivo' };
          return { 
            success: true, 
            user: { email: data[i][0], nombre: data[i][2], rol: data[i][3] } 
          };
        } else {
          // Si el email coincide pero el pass no, dar feedback sutil
          console.log('Password mismatch for:', emailInput);
          return { success: false, message: 'Contraseña incorrecta' };
        }
      }
    }
    return { success: false, message: 'Usuario no encontrado' };
  } catch (e) {
    console.error('Error en loginUser:', e);
    return { success: false, message: 'Error de servidor: ' + e.message };
  }
}

// Función de diagnóstico para el usuario
function getSystemStatus() {
  try {
    var ss = getSS();
    var sheet = ss.getSheetByName('Usuarios');
    var users = sheet ? sheet.getLastRow() - 1 : 0;
    return {
      success: true,
      ssUrl: ss.getUrl(),
      ssId: ss.getId(),
      sheets: ss.getSheets().map(function(s) { return s.getName(); }),
      userCount: users,
      version: 'v1.4.1 (Audit)',
      scriptUser: Session.getActiveUser().getEmail()
    };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function resetDatabaseLink() {
  PropertiesService.getScriptProperties().deleteProperty('SS_ID');
  var ss = getDatabaseSheet();
  return { success: true, url: ss.getUrl(), id: ss.getId() };
}

function getClientes() {
  var ss = getDatabaseSheet();
  var sheet = ss.getSheetByName('Clientes');
  if (!sheet) return [];
  return sheet.getDataRange().getValues().slice(1);
}

// ─── CREATE ──────────────────────────────────────────────────────
function addCliente(cliente) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Clientes');
    var row = [
      String(cliente.codigo || '').toUpperCase().trim(),
      String(cliente.nombre || '').trim(),
      String(cliente.telefono || '').trim(),
      String(cliente.correo || '').trim()
    ];
    sheet.appendRow(row);
    return { success: true, message: 'Cliente guardado.', data: row };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function savePromesa(promesa) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Promesas');
    var id = Utilities.getUuid();
    var now = new Date();
    var valMonto = cleanNum(promesa.monto);
    
    // ID, FECHA, CLIENTE_COD, MONTO, FECHA_PROMESA, ESTADO, NOTAS, INTERVENTOR
    var row = [
      id,
      now,
      String(promesa.clienteCod || ''),
      valMonto,
      promesa.fechaPromesa || '',
      'PENDIENTE',
      String(promesa.notas || '').trim(),
      String(promesa.interventor || 'SISTEMA')
    ];
    sheet.appendRow(row);
    
    var rowUi = [
      id,
      now.toISOString(),
      String(promesa.clienteCod || ''),
      valMonto,
      promesa.fechaPromesa || '',
      'PENDIENTE',
      String(promesa.notas || '').trim(),
      String(promesa.interventor || 'SISTEMA')
    ];
    return { success: true, message: 'Promesa registrada.', data: rowUi };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveIncidencia(inc) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Incidencias');
    var id = 'INC-' + new Date().getTime();
    var now = new Date();
    var valMonto = cleanNum(inc.monto);
    var factura = String(inc.factura || '').toUpperCase().trim();

    // Esquema 9 cols: ID(0), CLIENTE_COD(1), FACTURA(2), MONTO(3), MOTIVO(4), ESTADO(5), CANAL(6), LINK_GMAIL(7), FECHA(8)
    var row = [
      id,
      String(inc.clienteCod || ''),
      factura,
      valMonto,
      String(inc.motivo || '').trim(),
      'PENDIENTE',
      String(inc.canal || 'No definido').trim(),
      String(inc.linkGmail || ''),
      now,
      String(inc.interventor || 'SISTEMA')
    ];
    sheet.appendRow(row);
    
    var rowUi = [
      id,
      String(inc.clienteCod || ''),
      factura,
      valMonto,
      String(inc.motivo || '').trim(),
      'PENDIENTE',
      String(inc.canal || 'No definido').trim(),
      String(inc.linkGmail || ''),
      now.toISOString()
    ];
    return { success: true, message: 'Incidencia registrada.', data: rowUi };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Helper function for batch processing
function saveIncidenciaLocal(item, ss) {
  var ws = ss.getSheetByName('Incidencias');
  var id = 'INC-' + new Date().getTime();
  var now = new Date();
  var valMonto = cleanNum(item.monto);
  var factura = String(item.factura || '').toUpperCase().trim();

  // Esquema 9 cols: ID(0), CLIENTE_COD(1), FACTURA(2), MONTO(3), MOTIVO(4), ESTADO(5), CANAL(6), LINK_GMAIL(7), FECHA(8)
  var row = [
    id,
    String(item.clienteCod || ''),
    factura,
    valMonto,
    String(item.motivo || '').trim(),
    'PENDIENTE',
    String(item.canal || 'No definido').trim(),
    String(item.linkGmail || ''),
    now,
    String(item.interventor || 'SISTEMA')
  ];
  ws.appendRow(row);
  
  return [
    id,
    String(item.clienteCod || ''),
    factura,
    valMonto,
    String(item.motivo || '').trim(),
    'PENDIENTE',
    String(item.canal || 'No definido').trim(),
    String(item.linkGmail || ''),
    now.toISOString()
  ];
}

// ─── BATCH PROCESSING (ASISTENTE MÁGICO) ──────────────────────────
function processBatch(items) {
  try {
    var ss = getDatabaseSheet();
    var sheetPromesas = ss.getSheetByName('Promesas');
    var sheetClientes = ss.getSheetByName('Clientes');
    
    var results = { promesas: [], incidencias: [], clientes: [], count: 0 };
    var now = new Date();
    var nowIso = now.toISOString();

    items.forEach(function(item) {
      if (item.type === 'promesa') {
        var id = Utilities.getUuid();
        var monto = Number(item.monto) || 0;
        var interventor = String(item.interventor || 'SISTEMA');
        sheetPromesas.appendRow([ id, now, item.clienteCod, monto, item.fecha, 'PENDIENTE', item.notas, interventor ]);
        results.promesas.push([ id, nowIso, item.clienteCod, monto, item.fecha, 'PENDIENTE', item.notas, interventor ]);
        results.count++;
      } 
      else if (item.type === 'incidencia') {
        var incRow = saveIncidenciaLocal(item, ss); // Use the helper function
        results.incidencias.push(incRow);
        results.count++;
      }
      else if (item.type === 'cliente') {
        var cod = String(item.codigo).toUpperCase();
        sheetClientes.appendRow([ cod, item.nombre, item.telefono, item.correo ]);
        results.clientes.push([ cod, item.nombre, item.telefono, item.correo ]);
        results.count++;
      }
    });

    return { success: true, message: results.count + ' registros procesados correctamente.', data: results };
  } catch (e) {
    return { success: false, message: 'Error procesando lote: ' + e.message };
  }
}

// ─── CRM: LLAMADAS ──────────────────────────────────────────────
function saveLlamada(call) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Llamadas');
    var id = 'CALL-' + new Date().getTime();
    var now = new Date();
    
    // ID, FECHA, CLIENTE_COD, RESULTADO, INTERVENTOR, NOTAS
    var row = [
      id,
      now,
      String(call.clienteCod || ''),
      String(call.resultado || 'PENDIENTE'),
      String(call.interventor || 'USUARIO'),
      String(call.notas || '').trim()
    ];
    sheet.appendRow(row);
    
    var rowUi = [ id, now.toISOString(), call.clienteCod, call.resultado, String(call.interventor || 'SISTEMA'), call.notas ];
    return { success: true, message: 'Llamada registrada.', data: rowUi };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function enviarInformeGestion(userEmail) {
  try {
    var emailDestinatario = userEmail || Session.getActiveUser().getEmail();
    var ccEmails = "jdiaz@cofersa.cr";
    console.log('--- ENVIANDO INFORME --- Destinatario:', emailDestinatario, 'CC:', ccEmails);
    var ss = getDatabaseSheet();
    var hoy = new Date();
    var hoyStr = hoy.toDateString();
    var mesActual = hoy.getMonth();
    var anioActual = hoy.getFullYear();
    
    function parseSSDate(d) {
      if (!d || d instanceof Date) return d;
      var parts = String(d).match(/(\d{1,2})[\/-](\d{1,2})[\/-](\d{2,4})/);
      if (!parts) return null;
      var year = parts[3].length === 2 ? 2000 + parseInt(parts[3], 10) : parseInt(parts[3], 10);
      return new Date(year, parseInt(parts[2], 10) - 1, parseInt(parts[1], 10));
    }

    var llamadas    = ss.getSheetByName('Llamadas').getDataRange().getDisplayValues().slice(1);
    var clientes    = ss.getSheetByName('Clientes').getDataRange().getDisplayValues().slice(1);
    var incidencias = ss.getSheetByName('Incidencias').getDataRange().getDisplayValues().slice(1);
    var promesas    = ss.getSheetByName('Promesas').getDataRange().getDisplayValues().slice(1);
    
    function getCliName(cod) {
      if (!cod) return 'Cliente Desconocido';
      var cleanRef = String(cod).trim().replace(/^0+/, '');
      var c = clientes.find(function(x) { 
        var sheetCod = String(x[0]).trim().replace(/^0+/, '');
        return sheetCod === cleanRef; 
      });
      return c ? c[1] : cod;
    }

    var stats = {
      llamadasTotal: 0,
      contestaron: [],
      noContestaron: [],
      agenteGestiona: [],
      ecEnviados: [],
      recontactar: [],
      incidenciasDia: [],
      incidenciasMes: 0,
      incidenciasMesMonto: 0,
      promesasColocadasHoy: 0,
      promesasCumplidasHoy: 0,
      promesasVencidasHoy: 0,
      promesasVigentesMonto: 0,
      promesasVigentesCount: 0
    };

    llamadas.forEach(function(l) {
      var dl = parseSSDate(l[1]);
      if (dl && dl.toDateString() === hoyStr) {
        stats.llamadasTotal++;
        var res = (l[3]||'').toUpperCase();
        var item = { cli: getCliName(l[2]), nota: l[5] };
        if (res === 'CONTESTÓ') stats.contestaron.push(item);
        else if (res === 'NO CONTESTÓ') stats.noContestaron.push(item);
        else if (res === 'AGENTE GESTIONA') stats.agenteGestiona.push(item);
        else if (res === 'EC ENVIADO') stats.ecEnviados.push(item);
        else if (res === 'VOLVER A LLAMAR') stats.recontactar.push(item);
      }
    });

    incidencias.forEach(function(i) {
      var di = parseSSDate(i[8]);
      if (di) {
        var monto = cleanNum(i[3]);
        if (di.getMonth() === mesActual && di.getFullYear() === anioActual) {
          stats.incidenciasMes++;
          stats.incidenciasMesMonto += (monto || 0);
        }
        if (di.toDateString() === hoyStr) {
          stats.incidenciasDia.push({ cli: getCliName(i[1]), fact: i[2], monto: i[3] });
        }
      }
    });

    promesas.forEach(function(p) {
      var dCreacion = parseSSDate(p[1]);
      var dPromesa = parseSSDate(p[4]);
      var estado = (p[5]||'').toUpperCase();
      var monto = cleanNum(p[3]);

      if (dCreacion && dCreacion.toDateString() === hoyStr) {
        stats.promesasColocadasHoy++;
      }

      if (estado === 'PENDIENTE') {
        stats.promesasVigentesCount++;
        stats.promesasVigentesMonto += monto;
        if (dPromesa && dPromesa <= hoy) {
          stats.promesasVencidasHoy++;
        }
      } else if (estado === 'CUMPLIÓ') {
        if (dPromesa && dPromesa.toDateString() === hoyStr) {
          stats.promesasCumplidasHoy++;
        }
      }
    });

    var html = '<div style="font-family: sans-serif; background:#f4f7f6; padding: 20px; color: #2d3436;">' +
      '<div style="max-width: 600px; margin: 0 auto; background: white; border-radius: 15px; overflow: hidden; border: 1px solid #dfe6e9;">' +
        '<div style="background: #2d3436; padding: 30px; color: white; text-align: center;">' +
          '<h1 style="margin: 0; font-size: 24px;">Resumen de Gestion de Cobro</h1>' +
          '<p style="opacity: 0.7; margin: 5px 0 0;">' + hoy.toLocaleDateString('es-CR', {day:'2-digit', month:'long', year:'numeric'}) + '</p>' +
        '</div>' +
        '<div style="padding: 30px;">' +
          '<div style="background: #fdfdfd; padding: 15px; border-radius: 10px; border: 1px solid #f1f2f6; margin-bottom: 25px; display: flex; justify-content: space-between; align-items: center;">' +
            '<div style="text-align: center; flex: 1; border-right: 1px solid #eee;">' +
              '<small style="color: #636e72; font-weight: bold; display: block;">INCIDENCIAS (MES)</small>' +
              '<span style="font-size: 24px; font-weight: 900; color: #1e3a8a;">' + stats.incidenciasMes + '</span>' +
            '</div>' +
            '<div style="text-align: center; flex: 1;">' +
              '<small style="color: #636e72; font-weight: bold; display: block;">MONTO TOTAL</small>' +
              '<span style="font-size: 24px; font-weight: 900; color: #16a34a;">' + stats.incidenciasMesMonto.toLocaleString('es-CR', {style:'currency', currency:'CRC', minimumFractionDigits:0}) + '</span>' +
            '</div>' +
          '</div>' +
          
          '<div style="background: #fff; border: 1px solid #eee; border-radius: 10px; padding: 15px; margin-bottom: 25px;">' +
            '<div style="font-size: 11px; font-weight: 800; color: #2d3436; margin-bottom: 12px; border-bottom: 1px solid #f1f2f6; padding-bottom: 5px; text-transform: uppercase;">Estadísticas de Promesas</div>' +
            '<div style="display: grid; grid-template-columns: 1fr 1fr; gap: 10px;">' +
              '<div><small style="color:#636e72;">Colocadas hoy:</small> <b style="color:#2980b9;">'+stats.promesasColocadasHoy+'</b></div>' +
              '<div><small style="color:#636e72;">Cumplidas hoy:</small> <b style="color:#27ae60;">'+stats.promesasCumplidasHoy+'</b></div>' +
              '<div><small style="color:#636e72;">Vencidas hoy:</small> <b style="color:#c0392b;">'+stats.promesasVencidasHoy+'</b></div>' +
              '<div><small style="color:#636e72;">Vigentes totales:</small> <b style="color:#1e3a8a;">'+stats.promesasVigentesCount+' ('+stats.promesasVigentesMonto.toLocaleString('es-CR', {style:'currency', currency:'CRC', minimumFractionDigits:0})+')</b></div>' +
            '</div>' +
          '</div>' +

          renderEmailSection('CONTACTO EFECTIVO', stats.contestaron, '#27ae60') +
          renderEmailSection('AGENTE GESTIONA', stats.agenteGestiona, '#00AEEF') +
          renderEmailSection('ESTADOS DE CUENTA ENVIADOS', stats.ecEnviados, '#2980b9') +
          renderEmailSection('VOLVER A LLAMAR', stats.recontactar, '#f39c12') +
          renderEmailSection('NO CONTACTADOS', stats.noContestaron, '#c0392b') +
          
          '<h3 style="margin-top: 30px; color: #2d3436; border-bottom: 2px solid #f1f2f6; padding-bottom: 10px; font-size: 16px;">NUEVAS INCIDENCIAS (HOY)</h3>' +
          (stats.incidenciasDia.length ? '<ul style="padding-left: 20px; font-size: 13px;">' + stats.incidenciasDia.map(function(x){ return '<li style="margin-bottom:5px;"><b>'+x.cli+'</b> (Fact: '+x.fact+') &bull; <span style="color:#e67e22">'+x.monto+'</span></li>'; }).join('') + '</ul>' : '<p style="color:#b2bec3; font-size: 13px;">Sin incidencias hoy.</p>') +
          
          '<div style="margin-top: 40px; text-align: center; font-size: 11px; color: #b2bec3; padding-top: 20px; border-top: 1px solid #f1f2f6;">' +
            'Generado por Cobro Flow Pro - Usuario: ' + emailDestinatario +
          '</div>' +
        '</div>' +
      '</div>' +
    '</div>';

    function renderEmailSection(title, list, color) {
      if (!list.length) return '';
      var itemsArr = list.map(function(l) { 
        return '<div style="margin-bottom: 8px; border-left: 3px solid '+color+'; padding-left: 10px;">' +
                 '<span style="font-weight: bold; font-size: 14px; color: #2d3436;">'+l.cli+'</span>' +
                 (l.nota ? '<br><span style="font-size: 12px; color: #636e72; font-style: italic;">"'+l.nota+'"</span>' : '') +
               '</div>';
      }).join('');
      return '<div style="margin-bottom: 20px;">' +
               '<div style="font-size: 11px; font-weight: 800; color: '+color+'; margin-bottom: 8px; letter-spacing: 0.5px; text-transform: uppercase;">'+title+'</div>' +
               '<div>'+itemsArr+'</div>' +
             '</div>';
    }

    var subject = '📊 Resumen de Gestión - ' + emailDestinatario + ' - ' + hoy.toLocaleDateString('es-CR', {day:'2-digit', month:'2-digit'});
    
    GmailApp.sendEmail(emailDestinatario, subject, '', {
      cc: ccEmails,
      htmlBody: html
    });
    
    return { success: true, message: 'Informe enviado a ' + emailDestinatario + ' y ' + ccEmails };
  } catch (e) {
    console.error(e);
    return { success: false, message: 'Error enviando informe: ' + e.toString() };
  }
}


function deleteLlamada(payload) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Llamadas');
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Registro borrado.' };
      }
    }
    return { success: false, message: 'No encontrado.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
function updateLlamada(data) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Llamadas');
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(data.id)) {
        if (data.resultado) sheet.getRange(i + 1, 4).setValue(data.resultado);
        if (data.notas !== undefined) sheet.getRange(i + 1, 6).setValue(data.notas);
        var updatedRow = sheet.getRange(i + 1, 1, 1, 6).getValues()[0];
        return { success: true, message: 'Registro actualizado.', data: updatedRow };
      }
    }
    return { success: false, message: 'No encontrado.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─── UPDATE ──────────────────────────────────────────────────────
function updatePromesaEstado(data) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Promesas');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 6).setValue(data.estado); // col F = ESTADO
        return { success: true, message: 'Estado actualizado a ' + data.estado };
      }
    }
    return { success: false, message: 'Promesa no encontrada.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateIncidenciaCompleta(data) {
  try {
    var ss = getDatabaseSheet();
    var sheet = ss.getSheetByName('Incidencias');
    var rows = sheet.getDataRange().getValues();
    var foundIndex = -1;
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        foundIndex = i + 1;
        break;
      }
    }

    if (foundIndex === -1) throw new Error('No se encontró la incidencia');

    var newLink = data.link;

    // 📧 PROCESAR CORREO SI ESTÁ MARCADO 💎
    if (data.sendEmail && data.emailData) {
      newLink = enviarCorreoIncidencia(data, ss);
    }

    // ID(0), CLIENTE_COD(1), FACTURA(2), MONTO(3), MOTIVO(4), ESTADO(5), CANAL(6), LINK_GMAIL(7), FECHA(8), INTERVENTOR(9)
    sheet.getRange(foundIndex, 2).setValue(String(data.clienteCod || ''));
    sheet.getRange(foundIndex, 3).setValue(String(data.factura || ''));
    sheet.getRange(foundIndex, 4).setValue(cleanNum(data.monto || 0));
    sheet.getRange(foundIndex, 5).setValue(String(data.motivo || ''));
    sheet.getRange(foundIndex, 6).setValue(String(data.estado || 'PENDIENTE'));
    sheet.getRange(foundIndex, 7).setValue(String(data.canal || 'No definido'));
    sheet.getRange(foundIndex, 8).setValue(String(newLink || ''));
    
    return { success: true, message: 'Incidencia actualizada.', newLink: newLink };
  } catch (e) {
    console.error('Error en updateIncidenciaCompleta:', e);
    return { success: false, message: e.message };
  }
}

function enviarCorreoIncidencia(data, ss) {
  var e = data.emailData;
  var body = e.body;
  var subject = e.subject;

  // Reemplazar etiquetas mágicas
  var clienteNombre = 'Cliente';
  var clientes = ss.getSheetByName('Clientes').getDataRange().getValues();
  for(var i=1; i<clientes.length; i++){
    if(String(clientes[i][0]) === String(data.clienteCod)){
      clienteNombre = clientes[i][1];
      break;
    }
  }

  var hoy = new Date().toLocaleDateString('es-CR');
  
  body = body.replace(/{{CLIENTE}}/g, clienteNombre)
             .replace(/{{FACTURA}}/g, data.factura || 'S/N')
             .replace(/{{FECHA}}/g, hoy);
             
  subject = subject.replace(/{{CLIENTE}}/g, clienteNombre)
                   .replace(/{{FACTURA}}/g, data.factura || 'S/N')
                   .replace(/{{FECHA}}/g, hoy);

  var options = {
    htmlBody: body.replace(/\n/g, '<br>'),
    cc: "jdiaz@cofersa.cr"
  };

  // Manejar adjunto si existe
  if (e.attachment) {
    var blob = Utilities.newBlob(Utilities.base64Decode(e.attachment.content), e.attachment.mimeType, e.attachment.name);
    options.attachments = [blob];
  }

  try {
    GmailApp.sendEmail(e.to, subject, body, options);
    
    // Obtener enlace de seguimiento (directo al mensaje)
    Utilities.sleep(1500); // Esperar a que se procese en la bandeja de enviados
    var threads = GmailApp.search('to:' + e.to + ' subject:"' + subject + '"', 0, 1);
    if (threads && threads.length > 0) {
      var lastMsg = threads[0].getMessages().pop();
      return 'https://mail.google.com/mail/u/0/#search/rfc822msgid:' + lastMsg.getId();
    }
  } catch(err) {
    console.error('Error enviando correo:', err);
  }
  
  return '';
}

function updateIncidenciaEstado(data) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Incidencias');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.getRange(i + 1, 6).setValue(data.estado); // col F = ESTADO
        if (data.link) sheet.getRange(i + 1, 8).setValue(data.link); // col H = LINK_GMAIL
        return { success: true, message: 'Estado actualizado a ' + data.estado };
      }
    }
    return { success: false, message: 'Incidencia no encontrada.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ─── DELETE ──────────────────────────────────────────────────────
function deletePromesa(data) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Promesas');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.id)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Promesa eliminada.' };
      }
    }
    return { success: false, message: 'Promesa no encontrada.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteIncidencia(data) {
  try {
    var sheet = getDatabaseSheet().getSheetByName('Incidencias');
    var rows = sheet.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(data.factura)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Incidencia eliminada.' };
      }
    }
    return { success: false, message: 'Incidencia no encontrada.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}
