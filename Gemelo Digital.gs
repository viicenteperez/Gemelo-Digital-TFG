/**
 * SERVIDOR - Code.gs
 * Gemelo Digital Modula WMS
 */

function doGet(e) {
  if (e.parameter.action) {
    const action = e.parameter.action;
    const emailToProcess = e.parameter.uemail;
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    const data = sheet.getDataRange().getValues();
    
    let resultMsg = "Usuario no encontrado.";
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === emailToProcess.toLowerCase()) {
        if (action === 'aprobar') {
          sheet.getRange(i + 1, 3).setValue("ACTIVO"); // Columna C es la 3
          resultMsg = "✅ Acceso aprobado para: " + emailToProcess;
        } else {
          sheet.deleteRow(i + 1);
          resultMsg = "❌ Solicitud eliminada.";
        }
        break;
      }
    }
    return HtmlService.createHtmlOutput(`
      <div style="font-family: sans-serif; display: flex; flex-direction: column; align-items: center; justify-content: center; height: 100vh; background: #f8fafc;">
        <div style="background: white; padding: 40px; border-radius: 20px; box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1); text-align: center;">
          <div style="font-size: 4rem;">${icono}</div>
          <h2 style="color: #1e293b; margin: 20px 0;">${resultMsg}</h2>
          <p style="color: #64748b;">Puedes cerrar esta pestaña de forma segura.</p>
          <button onclick="window.close()" style="margin-top: 20px; padding: 10px 20px; cursor: pointer; border: none; border-radius: 8px; background: #2563eb; color: white; font-weight: bold;">Cerrar Ventana</button>
        </div>
      </div>`);
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gemelo Digital - Modula')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// =================================================================
// API ENDPOINT: ESCUCHAR ACTUALIZACIONES DE LA APP DE OPERACIONES
// =================================================================
function doPost(e) {
  try {
    var parametros = JSON.parse(e.postData.contents);
    
    if (parametros.accion === 'actualizarEstadoTransversal') {
      var idPieza = String(parametros.id).trim().toUpperCase();
      var nuevoEstado = parametros.estado === 'SACAR' ? 'En uso' : 'Operativo'; 
      var operario = parametros.usuario || 'Operario (App General)';
      
      // Abrimos el Excel por su ID exacto
      var ss = SpreadsheetApp.openById(""); 
      var sheetItems = ss.getSheetByName('Items');
      var data = sheetItems.getDataRange().getDisplayValues();
      
      var filaEncontrada = -1;
      
      // Buscamos la fila de la pieza
      for (var i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim().toUpperCase() === idPieza) {
          filaEncontrada = i + 1; // +1 porque el array empieza en 0 y las filas en 1
          break;
        }
      }
      
      if (filaEncontrada > -1) {
        // Modificamos la Columna J (Estado, index 10) y la Columna L (Operario, index 12)
        sheetItems.getRange(filaEncontrada, 10).setValue(nuevoEstado);
        sheetItems.getRange(filaEncontrada, 12).setValue(operario);
        
        // Escribimos en tu pestaña de Historial
        var sheetHistorial = ss.getSheetByName('Historial');
        if (sheetHistorial) {
          sheetHistorial.appendRow([new Date(), operario, "Módula WMS", idPieza, "Cambio desde App Operaciones", nuevoEstado]);
        }
        
        return ContentService.createTextOutput(JSON.stringify({exito: true}))
                             .setMimeType(ContentService.MimeType.JSON);
      } else {
        return ContentService.createTextOutput(JSON.stringify({exito: false, mensaje: "Pieza no encontrada"}))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({exito: false, mensaje: error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}

// =================================================================
// 1. VERIFICAR (Mira la columna C para el estado)
// =================================================================
function verificarIdentidad() {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
    if (!sheet) return { existe: false };

    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
        
        var nombre = data[i][1]; // Columna B
        var estado = data[i][2]; // Columna C (Estado)
        
        return { 
          existe: true, 
          pendiente: (estado === "PENDIENTE"), 
          nombre: nombre,
          email: email,
          esAdmin: true // Lo forzamos a true para que la UI cargue todo
        };
      }
    }
    return { existe: false, email: email };
  } catch (e) {
    return { existe: false };
  }
}

// 2. REGISTRAR (Escribe 4 columnas: Email, Nombre, Estado, Fecha)
function registrarUsuarioNuevo(datos) {
  try {
    var email = Session.getActiveUser().getEmail();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Usuarios');
    
    // Email | Nombre | Estado | Fecha
    sheet.appendRow([email, datos.nombre, "PENDIENTE", new Date()]);
    
    enviarEmailAprobacion(datos.nombre, email);
    return { exito: true };
  } catch (e) {
    return { exito: false, mensaje: e.message };
  }
}

// 3. APROBAR (Cambia la columna C a ACTIVO)
function doGet(e) {
  if (e.parameter.action) {
    var action = e.parameter.action;
    var email = e.parameter.uemail;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Usuarios");
    var data = sheet.getDataRange().getValues();
    
    for (var i = 1; i < data.length; i++) {
      if (data[i][0].toString().toLowerCase() === email.toLowerCase()) {
        if (action === 'aprobar') {
          sheet.getRange(i + 1, 3).setValue("ACTIVO"); // Columna C es la 3
        } else {
          sheet.deleteRow(i + 1);
        }
        break;
      }
    }
    return HtmlService.createHtmlOutput("<h2>Operación realizada con éxito. Puedes cerrar esta pestaña.</h2>");
  }
  
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Gemelo Digital - Modula')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function enviarEmailAprobacion(nombre, email) {
  var miCorreo = "vperez@celestica.com"; 
  var url = ScriptApp.getService().getUrl();
  
  var html = `
    <div style="font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; max-width: 500px; margin: auto; border: 1px solid #e2e8f0; border-radius: 16px; overflow: hidden; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);">
      <div style="background-color: #2563eb; padding: 20px; text-align: center;">
        <h2 style="color: white; margin: 0;">Solicitud de Acceso</h2>
      </div>
      <div style="padding: 30px; background-color: #ffffff;">
        <p style="color: #475569; font-size: 16px;">Se ha recibido una nueva solicitud de acceso al <b>Gemelo Digital Modula</b>:</p>
        <div style="background-color: #f8fafc; padding: 15px; border-left: 4px solid #2563eb; margin: 20px 0;">
          <p style="margin: 5px 0;"><b>👤 Usuario:</b> ${nombre}</p>
          <p style="margin: 5px 0;"><b>📧 Email:</b> ${email}</p>
        </div>
        <div style="text-align: center; margin-top: 30px;">
          <a href="${url}?action=aprobar&uemail=${email}" style="background-color: #16a34a; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block; margin: 5px;">✅ Aprobar Acceso</a>
          <a href="${url}?action=rechazar&uemail=${email}" style="background-color: #dc2626; color: white; padding: 12px 24px; text-decoration: none; border-radius: 8px; font-weight: bold; display: inline-block; margin: 5px;">❌ Denegar</a>
        </div>
      </div>
      <div style="background-color: #f1f5f9; padding: 15px; text-align: center; font-size: 12px; color: #64748b;">
        Este es un correo automático del sistema WMS.
      </div>
    </div>`;
  
  MailApp.sendEmail({ to: miCorreo, subject: "⚠️ Acceso WMS: " + nombre, htmlBody: html });
}

// CARGA DE DATOS
function obtenerInventario() {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaItems = libro.getSheetByName('Items');
    var hojaUbic = libro.getSheetByName('Ubicaciones');
    if (!hojaItems || !hojaUbic) return [];

    var datosItems = hojaItems.getDataRange().getValues();
    var datosUbic = hojaUbic.getDataRange().getValues();
    var mapaUbicaciones = {};
    
    for (var j = 1; j < datosUbic.length; j++) {
      var numEqUbic = datosUbic[j][4];
      if (numEqUbic) {
        mapaUbicaciones[numEqUbic.toString().trim().toUpperCase()] = {
          modula: datosUbic[j][1] || '1',
          balda: datosUbic[j][2] || '1',
          posicion: datosUbic[j][3] || 'A'
        };
      }
    }

    var inventario = [];
    for (var i = 1; i < datosItems.length; i++) {
      var numEq = datosItems[i][0];
      if (!numEq || numEq.toString().trim() === "") continue;
      var key = numEq.toString().trim().toUpperCase();
      var u = mapaUbicaciones[key] || { modula: '1', balda: '1', posicion: 'A' };

      var modLimpio = u.modula.toString().replace(/Módula|Modula|Torre/gi, '').trim();
      inventario.push({
        numEquipo: key,
        altura: (datosItems[i][1] || "").toString(), 
        largo: (datosItems[i][2] || "").toString(),  
        ancho: (datosItems[i][3] || "").toString(),  
        descripcion: (datosItems[i][4] || "Sin descripción").toString(),
        partCodeReal: (datosItems[i][6] || "").toString(),
        maquina: (datosItems[i][7] || "N/A").toString(),
        estado: (datosItems[i][9] || "Operativo").toString(),
        foto: (datosItems[i][10] || "").toString(),
        modula: (modLimpio || "1").toString(),
        balda: parseInt(u.balda) || 1,
        posicion: (u.posicion || "A").toString()
      });
    }
    return inventario;
  } catch (err) {
    return [];
  }
}

// REGISTRO
function registrarEquipo(f) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var idUbic = "M" + f.modula + "-B" + f.balda + "-P" + f.posicion;
    
    var urlFoto = ""; 
    
    if (f.fotoBase64 && f.fotoBase64 !== "") {
      // Usamos tu ID de carpeta: 
      var idCarpetaFotos = "idCarpetaFotos"; 
      var folder = DriveApp.getFolderById(idCarpetaFotos);
      
      var blob = Utilities.newBlob(Utilities.base64Decode(f.fotoBase64), f.fotoMimeType, f.fotoNombre);
      var archivoFoto = folder.createFile(blob);
      
      // ELIMINADA LA LÍNEA DE setSharing: La carpeta ya se encarga de esto.
      // Esto evita el error de "Acceso Denegado".
      
      urlFoto = archivoFoto.getUrl(); 
    }

    // REGISTRO EN ITEMS (Columna K para la foto)
    // Orden: NºEq(A), Alt(B), Lar(C), Anc(D), Desc(E), Cli(F), PC(G), Maq(H), Salvat(I), Est(J), Foto(K), User(L)
    libro.getSheetByName('Items').appendRow([
      f.numEquipo, f.altura, f.largo, f.ancho, f.descripcion, 
      f.cliente || "", f.partCodeReal, f.maquina, "", f.estado, urlFoto, ""
    ]);
    
    // REGISTRO EN UBICACIONES
    libro.getSheetByName('Ubicaciones').appendRow([
      idUbic, f.modula, f.balda, f.posicion, f.numEquipo, 1, ""
    ]);
    
    return { exito: true, mensaje: "Equipo registrado correctamente con foto." };
  } catch (error) {
    // Si hay un error, lo devolvemos para verlo en el msgForm del HTML [cite: 31]
    return { exito: false, mensaje: "Fallo en el servidor: " + error.toString() };
  }
}

// BAJA SEGURA
function eliminarEquipo(numEquipo) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hItems = libro.getSheetByName('Items');
    var hUbic = libro.getSheetByName('Ubicaciones');
    
    var dItems = hItems.getDataRange().getValues();
    for (var i = dItems.length - 1; i >= 1; i--) {
      if (dItems[i][0].toString().trim().toUpperCase() === numEquipo.toString().trim().toUpperCase()) {
        hItems.deleteRow(i + 1);
      }
    }
    
    var dUbic = hUbic.getDataRange().getValues();
    for (var j = dUbic.length - 1; j >= 1; j--) {
      if (dUbic[j][4].toString().trim().toUpperCase() === numEquipo.toString().trim().toUpperCase()) {
        hUbic.deleteRow(j + 1);
      }
    }
    
    return { exito: true };
  } catch (e) {
    return { exito: false, error: e.toString() };
  }
}

// =================================================================
// TRAZABILIDAD: GESTIÓN DE ESTADOS Y MOVIMIENTOS (BACK-END)
// =================================================================
function actualizarEstadoEquipo(numEquipo, nuevoEstado, operario, ubicacionStr) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaItems = libro.getSheetByName('Items');
    var hojaMovs = libro.getSheetByName('Movimientos'); 
    
    var datos = hojaItems.getDataRange().getValues();
    var fechaHoy = new Date();
    
    for (var i = 1; i < datos.length; i++) {
      // Normalizamos el ID para evitar fallos por espacios o mayúsculas
      if (datos[i][0].toString().trim().toUpperCase() === numEquipo.toString().trim().toUpperCase()) {
        
        // 1. Actualizamos el estado en la columna 10 (Estado)
        hojaItems.getRange(i + 1, 10).setValue(nuevoEstado);
        
        var accion = "";
        
        // 2. LÓGICA DE ACCIONES SEGÚN EL ESTADO RECIBIDO
        if (nuevoEstado === 'En uso') {
          // Si sale a planta
          hojaItems.getRange(i + 1, 12).setValue(operario);
          hojaItems.getRange(i + 1, 13).setValue(fechaHoy);
          accion = "🔴 RETIRADO A PLANTA";
        } 
        else if (nuevoEstado === 'En Mantenimiento') {
          // NUEVO: Si se bloquea por Mantenimiento
          hojaItems.getRange(i + 1, 12).setValue(operario); // Registramos quién lo bloqueó
          hojaItems.getRange(i + 1, 13).setValue(fechaHoy);
          accion = "🟡 BLOQUEADO POR MANTENIMIENTO";
        }
        else if (nuevoEstado === 'Operativo') {
          // Si vuelve a estar disponible (Devolución o fin de mantenimiento)
          hojaItems.getRange(i + 1, 12).setValue(""); // Limpiamos el operario
          hojaItems.getRange(i + 1, 13).setValue(fechaHoy);
          accion = "🟢 DEVUELTO AL MÓDULA";
        }
        else {
          // Otros estados (Incidencias, etc.)
          hojaItems.getRange(i + 1, 13).setValue(fechaHoy);
          accion = "🔄 ESTADO ACTUALIZADO: " + nuevoEstado;
        }
        
        // 3. REGISTRO EN LA HOJA DE MOVIMIENTOS
        if (hojaMovs) {
          hojaMovs.appendRow([
            fechaHoy, 
            operario, 
            numEquipo, 
            accion, 
            ubicacionStr
          ]);
        }
        
        return { exito: true };
      }
    }
    return { exito: false, mensaje: "Equipo no encontrado en la base de datos" };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

// INCIDENCIAS
function registrarIncidencia(numEquipo, motivo, ubicTemp, operario) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaItems = libro.getSheetByName('Items');
    var hojaMovs = libro.getSheetByName('Movimientos');
    var datos = hojaItems.getDataRange().getValues();
    var fechaHoy = new Date();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim().toUpperCase() === numEquipo.toString().trim().toUpperCase()) {
        
        hojaItems.getRange(i + 1, 10).setValue("Incidencia");
        hojaItems.getRange(i + 1, 12).setValue(operario);
        hojaItems.getRange(i + 1, 13).setValue(fechaHoy);
        
        if (hojaMovs) {
          var msjLog = "⚠️ INCIDENCIA: " + motivo;
          var ubicLog = "Temporal: " + ubicTemp;
          hojaMovs.appendRow([fechaHoy, operario, numEquipo, msjLog, ubicLog]);
        }
        
        return { exito: true };
      }
    }
    return { exito: false, mensaje: "Equipo no encontrado" };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

// ALTURAS BALDAS
function obtenerAlturaMaximaBalda(idBalda) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config_Baldas');
    if (!hoja) return 999;
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString().trim() === idBalda) {
        return parseFloat(datos[i][1]);
      }
    }
    return 999; 
  } catch (e) { return 999; }
}

function obtenerTodasLasAlturas() {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config_Baldas');
    if (!hoja) return [];
    return hoja.getDataRange().getValues().slice(1); 
  } catch (e) {
    return []; 
  }
}

function guardarAlturaServidor(idBalda, alturaCm, alturaIn) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config_Baldas');
    if (!hoja) return {exito: false, error: "No existe Config_Baldas"};
    
    var datos = hoja.getDataRange().getValues();
    for (var i = 1; i < datos.length; i++) {
      if (datos[i][0].toString() === idBalda) {
        hoja.getRange(i + 1, 2).setValue(alturaCm); // Columna B: CM
        hoja.getRange(i + 1, 3).setValue(alturaIn); // Columna C: Pulgadas
        return {exito: true};
      }
    }
    // Si no existe, creamos la fila con los 3 datos
    hoja.appendRow([idBalda, alturaCm, alturaIn]);
    return {exito: true};
  } catch (e) {
    return {exito: false, error: e.message}; 
  }
}

function borrarAlturaServidor(idBalda) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config_Baldas');
    if (!hoja) return {exito: false};
    
    var datos = hoja.getDataRange().getValues();
    for (var i = datos.length - 1; i >= 1; i--) {
      if (datos[i][0].toString().trim() === idBalda) {
        hoja.deleteRow(i + 1);
      }
    }
    return {exito: true};
  } catch(e) {
    return {exito: false, error: e.toString()};
  }
}

/**
 * Recupera el historial de movimientos de un equipo específico
 * para mostrarlo en la ficha técnica (Mejora 4)
 */
function obtenerHistorialEquipo(numEquipo) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaMovs = libro.getSheetByName('Movimientos');
    if (!hojaMovs) return [];

    var datos = hojaMovs.getDataRange().getValues();
    var historial = [];
    var buscado = numEquipo.toString().trim().toUpperCase();

    // Recorremos de abajo hacia arriba para obtener los más recientes primero
    for (var i = datos.length - 1; i >= 1; i--) {
      if (datos[i][2] && datos[i][2].toString().trim().toUpperCase() === buscado) {
        historial.push({
          fecha: Utilities.formatDate(datos[i][0], "GMT+1", "dd/MM/yy HH:mm"),
          operario: datos[i][1],
          accion: datos[i][3],
          ubicacion: datos[i][4]
        });
      }
      if (historial.length >= 4) break; // Limitamos a los 4 últimos para no saturar
    }
    return historial;
  } catch (e) {
    return [];
  }
}

// =================================================================
// AUDITORÍA: REGISTRO DE BÚSQUEDAS (ZERO TRUST)
// =================================================================
function registrarConsultaVisual(numEquipo, operario) {
  try {
    var libro = SpreadsheetApp.getActiveSpreadsheet();
    var hojaMovs = libro.getSheetByName('Movimientos');
    
    if (hojaMovs) {
      var fechaHoy = new Date();
      var msjLog = "👀 CONSULTA DE UBICACIÓN";
      var ubicLog = "Búsqueda en sistema cerrado";
      
      // Registramos silenciosamente que este operario ha mirado dónde está la pieza
      hojaMovs.appendRow([fechaHoy, operario, numEquipo, msjLog, ubicLog]);
    }
    return true;
  } catch (e) {
    return false;
  }
}

function aplicarOptimizacionMasivaServidor(piezas, alturas) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Apuntamos a tus hojas correctas
    const sheetUbic = ss.getSheetByName("Ubicaciones"); 
    const sheetAlturas = ss.getSheetByName("Config_Baldas");

    // 1. ACTUALIZAR COORDENADAS EN LA HOJA 'Ubicaciones'
    const dataUbic = sheetUbic.getDataRange().getValues();
    const mapPiezas = new Map();
    piezas.forEach(p => mapPiezas.set(p.numEquipo.toString().toUpperCase(), p));

    for (let i = 1; i < dataUbic.length; i++) {
      // El Nº Equipo en 'Ubicaciones' está en la columna E (índice 4)
      let idActual = (dataUbic[i][4] || "").toString().toUpperCase();
      
      if (mapPiezas.has(idActual)) {
        let nuevaData = mapPiezas.get(idActual);
        
        // ⚠️ FIX: Regeneramos el ID de la Columna A para que no se quede el antiguo
        let nuevoIdUbic = "M" + nuevaData.modula + "-B" + nuevaData.balda + "-P" + nuevaData.posicion;
        
        sheetUbic.getRange(i + 1, 1).setValue(nuevoIdUbic);        // Columna A (ID Compuesto)
        sheetUbic.getRange(i + 1, 2).setValue(nuevaData.modula);   // Columna B
        sheetUbic.getRange(i + 1, 3).setValue(nuevaData.balda);    // Columna C
        sheetUbic.getRange(i + 1, 4).setValue(nuevaData.posicion); // Columna D
      }
    }

    // 2. ACTUALIZAR LÍMITES DE ALTURA EN 'Config_Baldas'
    if (sheetAlturas) {
        const dataAlt = sheetAlturas.getDataRange().getValues();
        alturas.forEach(alt => {
          let encontrada = false;
          for (let i = 1; i < dataAlt.length; i++) {
            if (dataAlt[i][0].toString() == alt.idBalda) {
              sheetAlturas.getRange(i + 1, 2).setValue(alt.cm);
              sheetAlturas.getRange(i + 1, 3).setValue(alt.in);
              encontrada = true;
              break;
            }
          }
          if (!encontrada) {
            sheetAlturas.appendRow([alt.idBalda, alt.cm, alt.in]);
          }
        });
    }

    lock.releaseLock();
    return { exito: true, mensaje: "Optimización guardada" };
  } catch (e) {
    return { exito: false, mensaje: e.toString() };
  }
}

// =================================================================
// ESTADO GLOBAL DEL CANDADO DE SEGURIDAD (CELDA FÍSICA)
// =================================================================
function getEstadoCandado() {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config_Baldas');
    if (!hoja) return false;
    
    // Leemos la celda F1
    var valor = hoja.getRange("F1").getValue();
    return valor === true || valor === 'TRUE' || valor === 'true';
  } catch (e) {
    return false;
  }
}

function setEstadoCandado(estado) {
  try {
    var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config_Baldas');
    if (hoja) {
        // Escribimos en las celdas E1 y F1 para que quede guardado físicamente
        hoja.getRange("E1").setValue("CANDADO_GLOBAL:");
        hoja.getRange("E1").setFontWeight("bold");
        hoja.getRange("F1").setValue(estado ? 'TRUE' : 'FALSE');
    }
    return true;
  } catch (e) {
    throw new Error(e.message);
  }
}

// =================================================================
// MEMORIA DE BALDAS BLOQUEADAS (MANTENIMIENTO)
// =================================================================
function obtenerBaldasBloqueadas() {
  var props = PropertiesService.getScriptProperties();
  var data = props.getProperty('CONFIG_MANTENIMIENTO');
  // Si no hay nada guardado, devuelve un esquema limpio
  return data ? JSON.parse(data) : { '1':[], '2':[], '3':[], '4':[] };
}

function guardarBaldasBloqueadas(jsonConfig) {
  var props = PropertiesService.getScriptProperties();
  props.setProperty('CONFIG_MANTENIMIENTO', JSON.stringify(jsonConfig));
  return true;
}

// =================================================================
// EDICIÓN PROFUNDA DE EQUIPOS (MEDIDAS, FOTOS E IDs)
// =================================================================
function editarEquipoServidor(datos) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var hojaItems = ss.getSheetByName("Items"); // Tu hoja de datos reales
    var hojaUbic = ss.getSheetByName("Ubicaciones");
    
    var dataItems = hojaItems.getDataRange().getValues();
    var encontrado = false;
    
    // 1. BUSCAR Y ACTUALIZAR EN LA PESTAÑA 'Items'
    for (var i = 1; i < dataItems.length; i++) {
      if (dataItems[i][0].toString().trim().toUpperCase() === datos.idOriginal.toUpperCase()) {
        
        hojaItems.getRange(i + 1, 1).setValue(datos.numEquipo);      // Col A (1): Nº Equipo
        hojaItems.getRange(i + 1, 2).setValue(datos.altura);         // Col B (2): Altura
        hojaItems.getRange(i + 1, 3).setValue(datos.largo);          // Col C (3): Largo
        hojaItems.getRange(i + 1, 4).setValue(datos.ancho);          // Col D (4): Ancho
        hojaItems.getRange(i + 1, 5).setValue(datos.descripcion);    // Col E (5): Descripción
        hojaItems.getRange(i + 1, 6).setValue(datos.cliente);        // Col F (6): Cliente
        hojaItems.getRange(i + 1, 7).setValue(datos.partCode);       // Col G (7): Part Code
        hojaItems.getRange(i + 1, 8).setValue(datos.maquina);        // Col H (8): Máquina
        hojaItems.getRange(i + 1, 10).setValue(datos.estado);        // Col J (10): Estado
        
        // --- LÓGICA DE IMAGEN EN GOOGLE DRIVE ---
        if (datos.imagen && datos.imagen !== "") {
          var base64Data = datos.imagen;
          var mimeType = "image/jpeg";
          
          // Limpiamos el string Base64 que viene del navegador
          if (base64Data.indexOf(",") > -1) {
            var partes = base64Data.split(",");
            base64Data = partes[1];
            var mimeMatch = partes[0].match(/:(.*?);/);
            if (mimeMatch) mimeType = mimeMatch[1];
          }
          
          var idCarpetaFotos = "1C_SXGvMmaaVMTbPI7KVb6BEj7_7ReT0l";
          var folder = DriveApp.getFolderById(idCarpetaFotos);
          var nombreArchivo = "Edicion_" + datos.numEquipo + "_" + new Date().getTime() + ".jpg";
          
          var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, nombreArchivo);
          var archivoFoto = folder.createFile(blob);
          
          hojaItems.getRange(i + 1, 11).setValue(archivoFoto.getUrl()); // Col K (11): Foto URL
        }
        
        encontrado = true;
        break;
      }
    }
    
    if (!encontrado) {
      return { exito: false, mensaje: "No se encontró el Fixture original en la base de datos." };
    }

    // 2. ACTUALIZAR EN 'Ubicaciones' SI EL ID HA CAMBIADO
    if (datos.idOriginal.toUpperCase() !== datos.numEquipo.toUpperCase() && hojaUbic) {
      var dataUbic = hojaUbic.getDataRange().getValues();
      for (var j = 1; j < dataUbic.length; j++) {
        // En tu hoja 'Ubicaciones', el ID del equipo está en la Col E (índice 4)
        if (dataUbic[j][4].toString().trim().toUpperCase() === datos.idOriginal.toUpperCase()) {
          hojaUbic.getRange(j + 1, 5).setValue(datos.numEquipo);
          break; // Solo hay 1 ubicación por equipo
        }
      }
    }
    
    // 3. REGISTRAR LA EDICIÓN EN LA TRAZABILIDAD
    var hojaMovs = ss.getSheetByName('Movimientos');
    if (hojaMovs) {
      var msjHistorial = "Edición de datos (Medidas/Desc/Foto).";
      if (datos.idOriginal.toUpperCase() !== datos.numEquipo.toUpperCase()) {
        msjHistorial += " ID cambiado de [" + datos.idOriginal + "] a [" + datos.numEquipo + "]";
      }
      var operario = Session.getActiveUser().getEmail() || "Sistema";
      hojaMovs.appendRow([new Date(), operario, datos.numEquipo, "📝 EDICIÓN DE FICHA", msjHistorial]);
    }
    
    return { exito: true };
    
  } catch (e) {
    return { exito: false, mensaje: "Error del Servidor: " + e.toString() };
  }
}
