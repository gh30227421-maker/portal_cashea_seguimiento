/**
 * ============================================================================
 * 1. ENRUTADOR PRINCIPAL HTTP (WEB APP + API BOT PLAYWRIGHT)
 * ============================================================================
 */

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  if (params.accion) {
    try {
      if (params.accion === "obtener_filtros") {
        return responderJSON(apiObtenerFiltros());
      } else if (params.accion === "obtener_lote") {
        return responderJSON(apiObtenerLote(params.region, params.estatus, params.estado_hubspot));
      } else {
        return responderJSON({ status: "error", mensaje: "Acción GET no reconocida: " + params.accion });
      }
    } catch (err) {
      return responderJSON({ status: "error", mensaje: err.message });
    }
  }

  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Seguimiento Cashea - BNC')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responderJSON({ status: "error", mensaje: "No se recibió payload JSON en POST." });
    }

    const payload = JSON.parse(e.postData.contents);

    // 1. Petición del Bot de Python (Intacta)
    if (payload.accion === "reportar_resultados") {
      return responderJSON(apiReportarResultados(payload.resultados || []));
    }

    // 2. Peticiones de la Web App (Proxy Vercel / RPC)
    if (payload.accion === "webapp_rpc") {
      const fn = payload.funcion;
      const args = payload.argumentos || [];
      let data = null;

      if (fn === "obtenerPaqueteInicial") data = obtenerPaqueteInicial(args[0]);
      else if (fn === "obtenerClientes") data = obtenerClientes(args[0], args[1]);
      else if (fn === "obtenerClientesAdmin") data = obtenerClientesAdmin(args[0]);
      else if (fn === "guardarCambios") data = guardarCambios(args[0]);
      else if (fn === "eliminarClienteFila") data = eliminarClienteFila(args[0], args[1]);
      else if (fn === "agregarNuevoCliente") data = agregarNuevoCliente(args[0]);
      else if (fn === "verificarPinAdmin") data = verificarPinAdmin(args[0]);
      else if (fn === "obtenerDatosExportacionCompleta") data = obtenerDatosExportacionCompleta();
      else throw new Error("Función no reconocida: " + fn);

      return responderJSON({ status: "success", data: data });
    }

    return responderJSON({ status: "error", mensaje: "Acción POST no reconocida: " + payload.accion });
  } catch (err) {
    return responderJSON({ status: "error", mensaje: err.message });
  }
}

function responderJSON(objeto) {
  return ContentService.createTextOutput(JSON.stringify(objeto))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * ============================================================================
 * 2. FUNCIONES DE BASE DE DATOS Y ESTRUCTURA DE HOJAS
 * ============================================================================
 */

function obtenerHojaBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BD_ClientesCashea") || ss.getSheetByName("BD_Cashea") || ss.getActiveSheet();
  return sheet;
}

function asegurarColumnas(sheet) {
  const encabezados = sheet.getRange(1, 1, 1, 14).getValues()[0];
  if (String(encabezados[10] || "").trim() === "") sheet.getRange(1, 11).setValue("Estado en HubSpot");
  if (String(encabezados[11] || "").trim() === "") sheet.getRange(1, 12).setValue("Fecha de Actualización");
  if (String(encabezados[12] || "").trim() === "") sheet.getRange(1, 13).setValue("Fuera de Zona");
  if (String(encabezados[13] || "").trim() === "") sheet.getRange(1, 14).setValue("Fecha Gestión Agencia");
}

function obtenerHojaBajas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetBajas = ss.getSheetByName("Historico_Bajas");
  if (!sheetBajas) {
    sheetBajas = ss.insertSheet("Historico_Bajas");
    sheetBajas.appendRow([
      "Nombre de la Empresa", "RIF", "Contacto", "Teléfono", "Email",
      "Agencia BNC", "Región BNC", "Estatus de Gestión", "Fecha Envío Planillas",
      "Comentarios u Observaciones", "Estado en HubSpot", "Fecha de Actualización",
      "Fuera de Zona", "Fecha Gestión Agencia", "Fecha de Baja / Archivo"
    ]);
  }
  return sheetBajas;
}

/**
 * ============================================================================
 * 3. LÓGICA DE ENDPOINTS API PARA EL BOT
 * ============================================================================
 */

function apiObtenerFiltros() {
  const sheet = obtenerHojaBD();
  const datos = sheet.getDataRange().getValues();
  const mapaRegiones = {};
  let totalPendientes = 0;

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const empresa = String(fila[0] || "").trim();
    const region = String(fila[6] || "").trim();
    const estatus = String(fila[7] || "").trim();
    const estadoHubspot = String(fila[10] || "").trim().toUpperCase();

    if (esFilaEncabezadoOVacia(empresa, region)) continue;

    if (!mapaRegiones[region]) mapaRegiones[region] = 0;

    if (esEstatusGestionadoValido(estatus) && estadoHubspot === "POR ACTUALIZAR") {
      mapaRegiones[region]++;
      totalPendientes++;
    }
  }

  const listaRegiones = Object.keys(mapaRegiones).sort().map(nombre => ({
    nombre: nombre,
    pendientes: mapaRegiones[nombre]
  }));

  return {
    status: "success",
    regiones: listaRegiones,
    total_pendientes: totalPendientes
  };
}

function apiObtenerLote(regionFiltro, estatusFiltro, estadoHubspotFiltro) {
  const sheet = obtenerHojaBD();
  const datos = sheet.getDataRange().getValues();
  const clientes = [];
  const tz = Session.getScriptTimeZone() || "America/Caracas";

  const regBuscada = String(regionFiltro || "").trim().toUpperCase();
  const estBuscado = String(estatusFiltro || "").trim().toUpperCase();
  const hubBuscado = String(estadoHubspotFiltro || "POR ACTUALIZAR").trim().toUpperCase();

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const empresa = String(fila[0] || "").trim();
    const region = String(fila[6] || "").trim();
    const estatus = String(fila[7] || "").trim();
    const estadoHubspot = String(fila[10] || "").trim().toUpperCase();

    if (esFilaEncabezadoOVacia(empresa, region)) continue;
    if (!esEstatusGestionadoValido(estatus)) continue;
    if (hubBuscado !== "TODOS" && estadoHubspot !== hubBuscado) continue;
    if (regBuscada && regBuscada !== "TODAS" && region.toUpperCase() !== regBuscada) continue;
    if (estBuscado && estBuscado !== "TODOS" && estatus.toUpperCase() !== estBuscado) continue;

    clientes.push({
      fila: i + 1,
      empresa: empresa,
      rif: String(fila[1] || "").trim(),
      estatus: estatus,
      fecha: formatearFechaRapidaInput(fila[8], tz),
      comentarios: String(fila[9] || "").trim(),
      region: region,
      agencia: String(fila[5] || "").trim(),
      estado_hubspot: String(fila[10] || "").trim()
    });
  }

  return {
    status: "success",
    total: clientes.length,
    clientes: clientes
  };
}

function apiReportarResultados(resultados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    const maxFilas = sheet.getLastRow();
    const marcaTiempo = obtenerMarcaTemporalActual();
    const filasParaBaja = [];

    resultados.forEach(item => {
      const numFila = Number(item.fila);
      if (!numFila || numFila <= 1 || numFila > maxFilas) return;

      const estadoHubspot = String(item.estado_hubspot || "ACTUALIZADO").trim();
      const estatusFinal = String(item.estatus_final || "").trim();

      sheet.getRange(numFila, 11, 1, 2).setValues([[estadoHubspot, marcaTiempo]]);

      if (estatusFinal) {
        sheet.getRange(numFila, 8).setValue(estatusFinal);
      }

      const estatusEval = (estatusFinal || String(sheet.getRange(numFila, 8).getValue() || "")).trim().toLowerCase();
      const requiereBaja = Boolean(item.es_baja) || estatusEval === "no procede" || estatusEval === "ilocalizable";

      if (requiereBaja && !filasParaBaja.includes(numFila)) {
        filasParaBaja.push(numFila);
      }
    });

    filasParaBaja.sort((a, b) => b - a);
    filasParaBaja.forEach(numFila => {
      archivarYEliminarFilaInterno(sheet, numFila, marcaTiempo);
    });

    SpreadsheetApp.flush();

    return {
      status: "success",
      mensaje: "Filas actualizadas y bajas procesadas correctamente",
      actualizados: resultados.length,
      bajas_ejecutadas: filasParaBaja.length
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 4. FUNCIONES DE LA WEB APP
 * ============================================================================
 */

function obtenerPaqueteInicial(sesionGuardada) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetAgencias = ss.getSheetByName("Agencias");
  const catalogo = {};
  const diccionarioMaestro = [];

  if (sheetAgencias) {
    const datosAg = sheetAgencias.getDataRange().getValues();
    for (let i = 1; i < datosAg.length; i++) {
      const cod = String(datosAg[i][0] || "").trim();
      const nombreAgencia = String(datosAg[i][1] || "").trim();
      const region = String(datosAg[i][2] || "").trim();

      if (cod || nombreAgencia) {
        diccionarioMaestro.push({ cod: cod, nombre: nombreAgencia, region: region });
      }
      if (region && nombreAgencia) {
        if (!catalogo[region]) catalogo[region] = [];
        if (!catalogo[region].includes(nombreAgencia)) catalogo[region].push(nombreAgencia);
      }
    }
  }

  if (Object.keys(catalogo).length === 0) {
    const sheet = obtenerHojaBD();
    const datos = sheet.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const empresa = String(datos[i][0] || "").trim();
      const region = String(datos[i][6] || "").trim();
      const agencia = String(datos[i][5] || "").trim();
      if (esFilaEncabezadoOVacia(empresa, region)) continue;
      if (!catalogo[region]) catalogo[region] = [];
      if (agencia && !catalogo[region].includes(agencia)) catalogo[region].push(agencia);
    }
  }

  for (const reg in catalogo) {
    catalogo[reg].sort();
  }

  let clientes = null;
  if (sesionGuardada && sesionGuardada.modo === "ADMIN" && Array.isArray(sesionGuardada.regiones)) {
    clientes = obtenerClientesAdmin(sesionGuardada.regiones);
  } else if (sesionGuardada && sesionGuardada.modo === "AGENCIA" && sesionGuardada.region && sesionGuardada.agencia) {
    clientes = obtenerClientes(sesionGuardada.region, sesionGuardada.agencia);
  }

  return {
    catalogo: catalogo,
    catalogoMaestro: diccionarioMaestro,
    clientes: clientes
  };
}

function obtenerRegionesYAgencias() {
  return obtenerPaqueteInicial(null).catalogo;
}

function obtenerCatalogoMaestroAgencias() {
  return obtenerPaqueteInicial(null).catalogoMaestro;
}

function obtenerClientes(regionFiltro, agenciaFiltro) {
  const sheet = obtenerHojaBD();
  const datos = sheet.getDataRange().getValues();
  const resultado = [];
  const tz = Session.getScriptTimeZone() || "America/Caracas";

  const pedirTodasRegiones = !regionFiltro || regionFiltro === "TODAS";
  const pedirTodasAgencias = !agenciaFiltro || agenciaFiltro === "TODAS";

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const empresa = String(fila[0] || "").trim();
    const region = String(fila[6] || "").trim();
    if (!empresa || esFilaEncabezadoOVacia(empresa, region)) continue;

    if (!pedirTodasRegiones && region !== regionFiltro) continue;
    const agencia = String(fila[5] || "").trim();
    if (!pedirTodasAgencias && agencia !== agenciaFiltro) continue;

    resultado.push({
      numFila: i + 1,
      empresa: empresa,
      rif: String(fila[1] || "").trim(),
      contacto: String(fila[2] || "").trim(),
      telefono: String(fila[3] || "").trim(),
      email: String(fila[4] || "").trim(),
      agencia: agencia,
      region: region,
      estatus: String(fila[7] || "").trim(),
      fechaEnvio: formatearFechaRapidaInput(fila[8], tz),
      comentarios: String(fila[9] || "").trim(),
      estadoHubspot: String(fila[10] || "").trim(),
      fechaActualizacion: formatearFechaRapidaTexto(fila[11], tz),
      fueraDeZona: String(fila[12] || "").trim(),
      fechaGestionAgencia: formatearFechaRapidaTexto(fila[13], tz)
    });
  }

  return resultado;
}

function obtenerClientesAdmin(regionesSeleccionadas) {
  const sheet = obtenerHojaBD();
  const datos = sheet.getDataRange().getValues();
  const resultado = [];
  const tz = Session.getScriptTimeZone() || "America/Caracas";
  const todas = !regionesSeleccionadas || regionesSeleccionadas.includes("TODAS");
  const setRegiones = todas ? null : new Set(regionesSeleccionadas);

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const empresa = String(fila[0] || "").trim();
    const region = String(fila[6] || "").trim();
    if (!empresa || esFilaEncabezadoOVacia(empresa, region)) continue;

    if (todas || setRegiones.has(region)) {
      resultado.push({
        numFila: i + 1,
        empresa: empresa,
        rif: String(fila[1] || "").trim(),
        contacto: String(fila[2] || "").trim(),
        telefono: String(fila[3] || "").trim(),
        email: String(fila[4] || "").trim(),
        agencia: String(fila[5] || "").trim(),
        region: region,
        estatus: String(fila[7] || "").trim(),
        fechaEnvio: formatearFechaRapidaInput(fila[8], tz),
        comentarios: String(fila[9] || "").trim(),
        estadoHubspot: String(fila[10] || "").trim(),
        fechaActualizacion: formatearFechaRapidaTexto(fila[11], tz),
        fueraDeZona: String(fila[12] || "").trim(),
        fechaGestionAgencia: formatearFechaRapidaTexto(fila[13], tz)
      });
    }
  }

  return resultado;
}

function guardarCambios(cambios) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    asegurarColumnas(sheet);
    const marcaTiempoGestion = obtenerMarcaTemporalActual();

    cambios.forEach(item => {
      const fila = Number(item.numFila);
      if (fila && fila > 1) {
        if (item.agencia !== undefined) sheet.getRange(fila, 6).setValue(item.agencia);
        if (item.region !== undefined) sheet.getRange(fila, 7).setValue(item.region);
        if (item.estatus !== undefined) sheet.getRange(fila, 8).setValue(item.estatus);
        if (item.fechaEnvio !== undefined) sheet.getRange(fila, 9).setValue(item.fechaEnvio);
        if (item.comentarios !== undefined) sheet.getRange(fila, 10).setValue(item.comentarios);
        if (item.fueraDeZona !== undefined) sheet.getRange(fila, 13).setValue(item.fueraDeZona);

        sheet.getRange(fila, 14).setValue(marcaTiempoGestion);
        sheet.getRange(fila, 11).setValue("POR ACTUALIZAR");
      }
    });

    SpreadsheetApp.flush();

    return {
      exito: true,
      mensaje: "Se guardaron " + cambios.length + " registro(s) y se marcaron POR ACTUALIZAR en HubSpot.",
      fechaGestion: marcaTiempoGestion
    };
  } catch (error) {
    throw new Error("Error al guardar: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function eliminarClienteFila(numFila, nombreEmpresa) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    const fila = Number(numFila);

    if (!fila || fila <= 1) throw new Error("No se puede eliminar la fila de encabezado.");
    if (fila > sheet.getLastRow()) throw new Error("La fila solicitada ya no existe.");

    archivarYEliminarFilaInterno(sheet, fila, obtenerMarcaTemporalActual());

    return {
      exito: true,
      mensaje: "Cliente \"" + (nombreEmpresa || "") + "\" movido al histórico de bajas correctamente."
    };
  } catch (error) {
    throw new Error(error.message);
  } finally {
    lock.releaseLock();
  }
}

function archivarYEliminarFilaInterno(sheetActiva, numFila, marcaTiempoBaja) {
  const sheetBajas = obtenerHojaBajas();
  const numCols = Math.max(sheetActiva.getLastColumn(), 14);
  const valoresFila = sheetActiva.getRange(numFila, 1, 1, numCols).getValues()[0];

  while (valoresFila.length < 14) valoresFila.push("");
  valoresFila.push(marcaTiempoBaja);

  sheetBajas.appendRow(valoresFila);
  sheetActiva.deleteRow(numFila);
}

function agregarNuevoCliente(nuevo) {
  try {
    const sheet = obtenerHojaBD();
    const tieneEstatus = esEstatusGestionadoValido(nuevo.estatus);
    const ahora = obtenerMarcaTemporalActual();

    const nuevaFila = [
      nuevo.empresa,
      nuevo.rif,
      nuevo.contacto || "",
      nuevo.telefono || "",
      nuevo.email || "",
      nuevo.agencia,
      nuevo.region,
      nuevo.estatus || "",
      nuevo.fechaEnvio || "",
      nuevo.comentarios || "NUEVO REGISTRO",
      tieneEstatus ? "POR ACTUALIZAR" : "",
      "",
      "",
      ahora
    ];

    sheet.appendRow(nuevaFila);

    return {
      exito: true,
      mensaje: "Cliente agregado satisfactoriamente."
    };
  } catch (error) {
    throw new Error("Error al agregar cliente: " + error.message);
  }
}

function verificarPinAdmin(pinIngresado) {
  const PIN_OFICIAL = "2026";
  return String(pinIngresado).trim() === PIN_OFICIAL;
}

function obtenerDatosExportacionCompleta() {
  const sheet = obtenerHojaBD();
  const datos = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone() || "America/Caracas";
  
  const resultado = [[
    "Nombre de la Empresa", 
    "RIF", 
    "Contacto", 
    "Teléfono", 
    "Email", 
    "Agencia BNC", 
    "Región BNC", 
    "Estatus de Gestión", 
    "Fecha Envío Planillas", 
    "Comentarios u Observaciones", 
    "Estado en HubSpot", 
    "Fecha de Actualización", 
    "Fuera de Zona", 
    "Fecha Gestión Agencia"
  ]];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const empresa = String(fila[0] || "").trim();
    const region = String(fila[6] || "").trim();

    if (esFilaEncabezadoOVacia(empresa, region)) continue;

    resultado.push([
      empresa,
      String(fila[1] || "").trim(),
      String(fila[2] || "").trim(),
      String(fila[3] || "").trim(),
      String(fila[4] || "").trim(),
      String(fila[5] || "").trim(),
      region,
      String(fila[7] || "").trim(),
      formatearFechaRapidaInput(fila[8], tz),
      String(fila[9] || "").trim(),
      String(fila[10] || "").trim(),
      formatearFechaRapidaTexto(fila[11], tz),
      String(fila[12] || "").trim(),
      formatearFechaRapidaTexto(fila[13], tz)
    ]);
  }

  return resultado;
}

/**
 * ============================================================================
 * 5. UTILIDADES RÁPIDAS EN MEMORIA
 * ============================================================================
 */

function esFilaEncabezadoOVacia(empresa, region) {
  if (!empresa) return true;
  const empLower = empresa.toLowerCase();
  if (empLower === "empresa" || empLower === "nombre de la empresa" || empLower.startsWith("columna")) return true;
  const regLower = String(region || "").toLowerCase();
  if (regLower === "region" || regLower === "región" || regLower === "region bnc" || regLower === "región bnc") return true;
  return false;
}

function esEstatusGestionadoValido(estatus) {
  const limpio = String(estatus || "").trim();
  if (!limpio || limpio === "--" || limpio.toLowerCase().includes("seleccione")) return false;
  return true;
}

function obtenerMarcaTemporalActual() {
  const tz = Session.getScriptTimeZone() || "America/Caracas";
  return Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
}

function formatearFechaRapidaInput(valorFecha, tz) {
  if (!valorFecha) return "";
  if (typeof valorFecha === "string") {
    const s = valorFecha.trim();
    return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : "";
  }
  if (valorFecha instanceof Date && !isNaN(valorFecha.getTime())) {
    return Utilities.formatDate(valorFecha, tz, "yyyy-MM-dd");
  }
  return "";
}

function formatearFechaRapidaTexto(valor, tz) {
  if (!valor) return "";
  if (typeof valor === "string") return valor.trim();
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, tz, "dd/MM/yyyy HH:mm");
  }
  return String(valor).trim();
}