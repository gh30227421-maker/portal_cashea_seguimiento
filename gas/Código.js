/**
 * ============================================================================
 * 1. ENRUTADOR PRINCIPAL HTTP (WEB APP + API BOT PLAYWRIGHT)
 * ============================================================================
 */

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};

  if (params.accion) {
    try {
      validarSecretoOpcional(params.api_secret);

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
    .setTitle('V.P.E. Banca Comercial | Gestión de Cartera')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return responderJSON({ status: "error", mensaje: "No se recibió payload JSON en POST." });
    }

    const payload = JSON.parse(e.postData.contents);
    validarSecretoOpcional(payload._apiSecret || payload.api_secret);

    // 1. Petición del Bot de Python (Contrato 100% intacto)
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
      else if (fn === "obtenerClientesAdmin") data = obtenerClientesAdmin(args[0], args[1]);
      else if (fn === "guardarCambios") data = guardarCambios(args[0], args[1]);
      else if (fn === "eliminarClienteFila") data = eliminarClienteFila(args[0], args[1], args[2], args[3], args[4]);
      else if (fn === "agregarNuevoCliente") data = agregarNuevoCliente(args[0], args[1], args[2]);
      else if (fn === "verificarPinAdmin") data = verificarPinAdmin(args[0]);
      else if (fn === "obtenerDatosExportacionCompleta") data = obtenerDatosExportacionCompleta(args[0]);
      else if (fn === "depurarClientesListos") data = depurarClientesListos(args[0], args[1], args[2]);
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
 * 2. SEGURIDAD LIGERA, BASE DE DATOS Y ESTRUCTURA DE HOJAS
 * ============================================================================
 */

function obtenerPinAdminOficial() {
  const props = PropertiesService.getScriptProperties();
  let pin = props.getProperty("ADMIN_PIN");
  if (!pin) {
    pin = "2026";
    try {
      props.setProperty("ADMIN_PIN", pin);
    } catch (e) {}
  }
  return String(pin).trim();
}

function verificarPinAdmin(pinIngresado) {
  if (pinIngresado === undefined || pinIngresado === null) return false;
  return String(pinIngresado).trim() === obtenerPinAdminOficial();
}

function exigirPinAdmin(pinIngresado) {
  if (!verificarPinAdmin(pinIngresado)) {
    throw new Error("Acceso denegado: El PIN de supervisor es inválido o su sesión expiró.");
  }
}

function validarSecretoOpcional(secretoRecibido) {
  const props = PropertiesService.getScriptProperties();
  const secretoConfigurado = props.getProperty("GAS_API_SECRET");
  // Solo exige secreto si el administrador decide definir GAS_API_SECRET en Propiedades del Script
  if (secretoConfigurado && String(secretoRecibido || "").trim() !== String(secretoConfigurado).trim()) {
    throw new Error("Solicitud no autorizada (API Secret inválido).");
  }
}

function obtenerHojaBD() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("BD_ClientesCashea") || ss.getSheetByName("BD_Cashea") || ss.getActiveSheet();
  return sheet;
}

function asegurarColumnas(sheet) {
  if (sheet.getMaxColumns() < 15) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 15 - sheet.getMaxColumns());
  }
  const encabezados = sheet.getRange(1, 1, 1, 15).getValues()[0];
  if (String(encabezados[10] || "").trim() === "") sheet.getRange(1, 11).setValue("Estado en HubSpot");
  if (String(encabezados[11] || "").trim() === "") sheet.getRange(1, 12).setValue("Fecha de Actualización");
  if (String(encabezados[12] || "").trim() === "") sheet.getRange(1, 13).setValue("Fuera de Zona");
  if (String(encabezados[13] || "").trim() === "") sheet.getRange(1, 14).setValue("Fecha Gestión Agencia");
  if (String(encabezados[14] || "").trim() === "") sheet.getRange(1, 15).setValue("Gestionado Por");
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
      "Fuera de Zona", "Fecha Gestión Agencia", "Fecha de Baja / Archivo",
      "Gestionado / Archivado Por"
    ]);
  } else {
    if (sheetBajas.getMaxColumns() < 16) {
      sheetBajas.insertColumnsAfter(sheetBajas.getMaxColumns(), 16 - sheetBajas.getMaxColumns());
    }
    const encBajas = sheetBajas.getRange(1, 1, 1, 16).getValues()[0];
    if (String(encBajas[14] || "").trim() === "") sheetBajas.getRange(1, 15).setValue("Fecha de Baja / Archivo");
    if (String(encBajas[15] || "").trim() === "") sheetBajas.getRange(1, 16).setValue("Gestionado / Archivado Por");
  }
  return sheetBajas;
}

/**
 * Construye el arreglo estándar de 16 columnas para Historico_Bajas
 * respetando las 15 columnas históricas originales + columna 16 de trazabilidad de usuario.
 */
function construirFilaBaja(filaOrigen, marcaTiempoBaja, actorBaja) {
  const base14 = filaOrigen.slice(0, 14);
  while (base14.length < 14) base14.push("");

  const ultimoGestor = String(filaOrigen[14] || "").trim();
  const responsableBaja = String(actorBaja || "").trim();
  let trazaResponsable = ultimoGestor;

  if (ultimoGestor && responsableBaja && ultimoGestor.toLowerCase() !== responsableBaja.toLowerCase()) {
    trazaResponsable = `${ultimoGestor} (Archivado por: ${responsableBaja})`;
  } else if (!ultimoGestor && responsableBaja) {
    trazaResponsable = `Archivado por: ${responsableBaja}`;
  }

  return [...base14, marcaTiempoBaja, trazaResponsable];
}

/**
 * Localiza de forma segura el número de fila (1-indexed) de un cliente verificando RIF y Empresa.
 * Evita sobrescribir clientes equivocados si hubo bajas o depuraciones concurrentes.
 */
function resolverFilaSegura(datos, numFilaSugerida, rifEsperado, empresaEsperada) {
  const filaNum = Number(numFilaSugerida);
  const tieneIdentificador = Boolean(String(rifEsperado || "").trim() || String(empresaEsperada || "").trim());

  // Compatibilidad con llamadas legadas del Bot que solo envían { fila }
  if (!tieneIdentificador) {
    return (filaNum > 1 && filaNum <= datos.length) ? filaNum : -1;
  }

  const rifNorm = normalizarRifComparacion(rifEsperado);
  const empNorm = normalizarTextoComparacion(empresaEsperada);

  // 1. Verificación rápida O(1) en la fila esperada
  if (filaNum > 1 && filaNum <= datos.length) {
    const filaActual = datos[filaNum - 1];
    const rifFila = normalizarRifComparacion(filaActual[1]);
    const empFila = normalizarTextoComparacion(filaActual[0]);

    if ((rifNorm && rifFila === rifNorm) || (!rifNorm && empNorm && empFila === empNorm)) {
      return filaNum;
    }
  }

  // 2. La fila se desplazó por una eliminación previa: buscar coincidencia exacta RIF + Empresa
  if (rifNorm && empNorm) {
    for (let i = 1; i < datos.length; i++) {
      if (normalizarRifComparacion(datos[i][1]) === rifNorm && normalizarTextoComparacion(datos[i][0]) === empNorm) {
        return i + 1;
      }
    }
  }

  // 3. Respaldo por RIF único (si tiene longitud válida)
  if (rifNorm && rifNorm.length >= 5) {
    for (let i = 1; i < datos.length; i++) {
      if (normalizarRifComparacion(datos[i][1]) === rifNorm) {
        return i + 1;
      }
    }
  }

  // 4. Respaldo por nombre exacto de empresa cuando no tiene RIF cargado
  if (!rifNorm && empNorm) {
    for (let i = 1; i < datos.length; i++) {
      if (normalizarTextoComparacion(datos[i][0]) === empNorm) {
        return i + 1;
      }
    }
  }

  return -1;
}

/**
 * Elimina múltiples filas agrupando rangos contiguos en orden descendente.
 * Reduce drásticamente las llamadas a Google Sheets y evita timeouts en Vercel.
 */
function eliminarFilasEnBloquesDescendentes(sheet, listaFilas) {
  if (!listaFilas || listaFilas.length === 0) return;

  const unicasOrdenadas = Array.from(new Set(listaFilas.map(Number).filter(n => n > 1)))
    .sort((a, b) => b - a);

  let i = 0;
  while (i < unicasOrdenadas.length) {
    let finBloque = unicasOrdenadas[i];
    let inicioBloque = finBloque;
    while (i + 1 < unicasOrdenadas.length && unicasOrdenadas[i + 1] === inicioBloque - 1) {
      inicioBloque = unicasOrdenadas[i + 1];
      i++;
    }
    const cantidad = finBloque - inicioBloque + 1;
    sheet.deleteRows(inicioBloque, cantidad);
    i++;
  }
}

/**
 * ============================================================================
 * 3. LÓGICA DE ENDPOINTS API PARA EL BOT DE PLAYWRIGHT (CONTRATO INTACTO)
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

/**
 * Mapea cualquier variante de estatus proveniente de HubSpot al nombre canónico exacto
 * que utiliza el Portal Web (Lista, Gestión de Contacto, Faltan Recaudos, No Procede, Ilocalizable).
 */
function mapearEstatusHubspotAPortal(estatusRaw) {
  const raw = String(estatusRaw || "").trim();
  if (!raw) return "";

  const norm = raw
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  if (
    norm === "listo" ||
    norm === "lista" ||
    norm === "listos" ||
    norm === "listas" ||
    norm.includes("cuenta abierta") ||
    norm.includes("aperturada") ||
    norm.includes("planilla enviada") ||
    norm.includes("completad") ||
    norm.includes("cerrado ganado")
  ) {
    return "Lista";
  }

  if (norm.includes("recaudo") || norm.includes("documentacion pendiente") || norm.includes("documento")) {
    return "Faltan Recaudos";
  }

  if (
    norm.includes("gestion") ||
    norm.includes("apertura") ||
    norm.includes("contactad") ||
    norm.includes("en contacto") ||
    norm.includes("en proceso") ||
    norm.includes("seguimiento")
  ) {
    return "Gestión de Contacto";
  }

  if (
    norm.includes("no procede") ||
    norm.includes("no califica") ||
    norm.includes("rechazad") ||
    norm.includes("descartad") ||
    norm.includes("cerrado perdido")
  ) {
    return "No Procede";
  }

  if (
    norm.includes("ilocalizable") ||
    norm.includes("no localizable") ||
    norm.includes("no contesta") ||
    norm.includes("numero errado")
  ) {
    return "Ilocalizable";
  }

  return raw;
}

/**
 * Valida y convierte fechas provenientes de HubSpot (ej: "20/11/2026" o "2026-11-20")
 * al formato estándar YYYY-MM-DD de la hoja y del portal web.
 */
function normalizarFechaHubspotASheets(fechaRaw, tz) {
  if (!fechaRaw) return "";
  if (fechaRaw instanceof Date && !isNaN(fechaRaw.getTime())) {
    return Utilities.formatDate(fechaRaw, tz, "yyyy-MM-dd");
  }
  const s = String(fechaRaw).trim();
  if (!s || s === "--") return "";

  const mYMD = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (mYMD) {
    const yyyy = Number(mYMD[1]);
    const mm = Number(mYMD[2]);
    const dd = Number(mYMD[3]);
    if (yyyy >= 2000 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      return `${String(yyyy)}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
    }
  }

  const mDMY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (mDMY) {
    const dd = Number(mDMY[1]);
    const mm = Number(mDMY[2]);
    const yyyy = Number(mDMY[3]);
    if (yyyy >= 2000 && mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31) {
      return `${String(yyyy)}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
    }
  }

  return "";
}

function apiReportarResultados(resultados) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    asegurarColumnas(sheet);
    const datos = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone() || "America/Caracas";
    const marcaTiempo = obtenerMarcaTemporalActual();
    const filasParaBaja = [];
    const filasValoresBajas = [];
    let actualizadosReales = 0;
    let sincronizadosDesdeHubspot = 0;

    resultados.forEach(item => {
      const filaReal = resolverFilaSegura(datos, item.fila, item.rif, item.empresa);
      if (filaReal <= 1 || filaReal > datos.length) return;

      // 1. Manejo de desmarcar_fuera_de_zona (si el bot reubicó o validó zona en HubSpot)
      if (Boolean(item.desmarcar_fuera_de_zona)) {
        sheet.getRange(filaReal, 13).setValue("");
        datos[filaReal - 1][12] = "";

        const comConTag = String(datos[filaReal - 1][9] || "");
        if (/\[FUERA DE ZONA\]/i.test(comConTag)) {
          const comLimpioFdz = comConTag.replace(/\[FUERA DE ZONA\]\s*/gi, "").trim();
          sheet.getRange(filaReal, 10).setValue(comLimpioFdz);
          datos[filaReal - 1][9] = comLimpioFdz;
        }
      }

      // 2. Sincronización Bidireccional (HubSpot -> Google Sheets) cuando HubSpot está más avanzado
      if (Boolean(item.actualizar_en_sheets)) {
        const estatusSync = mapearEstatusHubspotAPortal(item.estatus_hubspot_real || item.estatus_final);
        if (estatusSync) {
          sheet.getRange(filaReal, 8).setValue(estatusSync);
          datos[filaReal - 1][7] = estatusSync;
        }

        // Rellenar Fecha Envío Planillas (Col 9) solo si está vacía y HubSpot envía una fecha válida
        const valorFechaActual = datos[filaReal - 1][8];
        const textoFechaActual = formatearFechaRapidaInput(valorFechaActual, tz) || String(valorFechaActual || "").trim();
        const celdaFechaVacia = (!textoFechaActual || textoFechaActual === "--");
        const fechaHubspotNormalizada = normalizarFechaHubspotASheets(item.fecha_hubspot, tz);

        if (celdaFechaVacia && fechaHubspotNormalizada) {
          sheet.getRange(filaReal, 9).setValue(fechaHubspotNormalizada);
          datos[filaReal - 1][8] = fechaHubspotNormalizada;
        }

        // Rellenar Comentarios (Col 10) solo si está vacía (nunca sobrescribe notas previas de agencia)
        const notaHubspot = String(item.nota_hubspot || "").trim();
        const comentarioActual = String(datos[filaReal - 1][9] || "").trim();
        if (notaHubspot && (comentarioActual === "" || comentarioActual === "--")) {
          sheet.getRange(filaReal, 10).setValue(notaHubspot);
          datos[filaReal - 1][9] = notaHubspot;
        }

        // Registrar estado del bot en Col 11 y marca de tiempo en Col 12
        const estadoSync = "SINCRONIZADO DESDE HUBSPOT";
        sheet.getRange(filaReal, 11, 1, 2).setValues([[estadoSync, marcaTiempo]]);
        datos[filaReal - 1][10] = estadoSync;
        datos[filaReal - 1][11] = marcaTiempo;

        // Si nadie había gestionado la fila aún en el portal, dejar constancia en Gestionado Por (Col 15)
        if (!String(datos[filaReal - 1][14] || "").trim()) {
          sheet.getRange(filaReal, 15).setValue("Sync HubSpot");
          datos[filaReal - 1][14] = "Sync HubSpot";
        }

        sincronizadosDesdeHubspot++;
        actualizadosReales++;
      } else {
        // 3. Flujo estándar existente (ACTUALIZADO, NO EXISTE CLIENTE, YA ESTABA AVANZADO, etc.)
        const estadoHubspot = String(item.estado_hubspot || "ACTUALIZADO").trim();
        const estatusFinal = item.estatus_final ? mapearEstatusHubspotAPortal(item.estatus_final) : "";

        sheet.getRange(filaReal, 11, 1, 2).setValues([[estadoHubspot, marcaTiempo]]);
        datos[filaReal - 1][10] = estadoHubspot;
        datos[filaReal - 1][11] = marcaTiempo;

        if (estatusFinal) {
          sheet.getRange(filaReal, 8).setValue(estatusFinal);
          datos[filaReal - 1][7] = estatusFinal;
        }

        actualizadosReales++;
      }

      // 4. Evaluación de baja automática (es_baja, No Procede o Ilocalizable)
      const estatusEval = String(datos[filaReal - 1][7] || "").trim().toLowerCase();
      const requiereBaja = Boolean(item.es_baja) || estatusEval === "no procede" || estatusEval === "ilocalizable";

      if (requiereBaja && !filasParaBaja.includes(filaReal)) {
        filasParaBaja.push(filaReal);
        filasValoresBajas.push(construirFilaBaja(datos[filaReal - 1], marcaTiempo, "Bot HubSpot"));
      }
    });

    if (filasValoresBajas.length > 0) {
      const sheetBajas = obtenerHojaBajas();
      const ultimaBaja = sheetBajas.getLastRow();
      sheetBajas.getRange(ultimaBaja + 1, 1, filasValoresBajas.length, 16).setValues(filasValoresBajas);
      eliminarFilasEnBloquesDescendentes(sheet, filasParaBaja);
    }

    SpreadsheetApp.flush();

    return {
      status: "success",
      mensaje: "Filas actualizadas, sincronizaciones y bajas procesadas correctamente",
      actualizados: actualizadosReales,
      sincronizados_hubspot: sincronizadosDesdeHubspot,
      bajas_ejecutadas: filasParaBaja.length
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================================
 * 4. FUNCIONES DE LA WEB APP (PORTAL VERCEL)
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
  let requierePinAdmin = false;

  if (sesionGuardada && sesionGuardada.modo === "ADMIN" && Array.isArray(sesionGuardada.regiones)) {
    if (verificarPinAdmin(sesionGuardada.pin)) {
      clientes = obtenerClientesAdmin(sesionGuardada.regiones, sesionGuardada.pin);
    } else {
      requierePinAdmin = true;
    }
  } else if (sesionGuardada && sesionGuardada.modo === "AGENCIA" && sesionGuardada.region && sesionGuardada.agencia) {
    clientes = obtenerClientes(sesionGuardada.region, sesionGuardada.agencia);
  }

  return {
    catalogo: catalogo,
    catalogoMaestro: diccionarioMaestro,
    clientes: clientes,
    requierePinAdmin: requierePinAdmin
  };
}

function mapearFilaAClienteWeb(fila, numFila, tz) {
  return {
    numFila: numFila,
    empresa: String(fila[0] || "").trim(),
    rif: String(fila[1] || "").trim(),
    contacto: String(fila[2] || "").trim(),
    telefono: String(fila[3] || "").trim(),
    email: String(fila[4] || "").trim(),
    agencia: String(fila[5] || "").trim(),
    region: String(fila[6] || "").trim(),
    estatus: String(fila[7] || "").trim(),
    fechaEnvio: formatearFechaRapidaInput(fila[8], tz),
    comentarios: String(fila[9] || "").trim(),
    estadoHubspot: String(fila[10] || "").trim(),
    fechaActualizacion: formatearFechaRapidaTexto(fila[11], tz),
    fueraDeZona: String(fila[12] || "").trim(),
    fechaGestionAgencia: formatearFechaRapidaTexto(fila[13], tz),
    gestionadoPor: String(fila[14] || "").trim()
  };
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

    resultado.push(mapearFilaAClienteWeb(fila, i + 1, tz));
  }

  return resultado;
}

function obtenerClientesAdmin(regionesSeleccionadas, pinAdmin) {
  exigirPinAdmin(pinAdmin);

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
      resultado.push(mapearFilaAClienteWeb(fila, i + 1, tz));
    }
  }

  return resultado;
}

function guardarCambios(cambios, usuarioAccion) {
  if (!Array.isArray(cambios) || cambios.length === 0) {
    return { exito: true, mensaje: "No se recibieron cambios para procesar." };
  }

  const gestorLimpio = limpiarNombrePersona(usuarioAccion);
  if (!gestorLimpio) {
    throw new Error("Debe indicar su Nombre y Apellido para registrar quién realizó la actualización.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    asegurarColumnas(sheet);
    const datos = sheet.getDataRange().getValues();
    const marcaTiempoGestion = obtenerMarcaTemporalActual();

    let guardados = 0;
    let marcadosHubspot = 0;
    const noEncontrados = [];

    cambios.forEach(item => {
      const filaReal = resolverFilaSegura(datos, item.numFila, item.rif, item.empresa);
      if (filaReal <= 1 || filaReal > datos.length) {
        noEncontrados.push(item.empresa || item.rif || ("Fila " + item.numFila));
        return;
      }

      const filaActual = datos[filaReal - 1];
      const nuevaAgencia = item.agencia !== undefined ? item.agencia : String(filaActual[5] || "");
      const nuevaRegion = item.region !== undefined ? item.region : String(filaActual[6] || "");
      const nuevoEstatus = item.estatus !== undefined ? item.estatus : String(filaActual[7] || "");
      const nuevaFechaEnvio = item.fechaEnvio !== undefined ? item.fechaEnvio : filaActual[8];
      const nuevosComentarios = item.comentarios !== undefined ? item.comentarios : String(filaActual[9] || "");
      const nuevoFueraZona = item.fueraDeZona !== undefined ? item.fueraDeZona : String(filaActual[12] || "");

      // Solo marcamos POR ACTUALIZAR en HubSpot si el cliente tiene un estatus de gestión válido
      const tieneEstatusValido = esEstatusGestionadoValido(nuevoEstatus);
      const estadoHubspotActual = String(filaActual[10] || "").trim();
      const nuevoEstadoHubspot = tieneEstatusValido ? "POR ACTUALIZAR" : estadoHubspotActual;
      if (tieneEstatusValido) marcadosHubspot++;

      const fechaActualizacionBot = filaActual[11] || "";

      // Escritura optimizada de 10 columnas contiguas (Col 6 a Col 15) en 1 sola llamada por fila
      const bloqueCols6a15 = [[
        nuevaAgencia,
        nuevaRegion,
        nuevoEstatus,
        nuevaFechaEnvio,
        nuevosComentarios,
        nuevoEstadoHubspot,
        fechaActualizacionBot,
        nuevoFueraZona,
        marcaTiempoGestion,
        gestorLimpio
      ]];

      sheet.getRange(filaReal, 6, 1, 10).setValues(bloqueCols6a15);
      guardados++;
    });

    SpreadsheetApp.flush();

    let msg = `Se guardaron ${guardados} registro(s) por ${gestorLimpio}`;
    if (marcadosHubspot > 0) {
      msg += ` (${marcadosHubspot} marcado(s) POR ACTUALIZAR en HubSpot).`;
    } else {
      msg += ".";
    }
    if (noEncontrados.length > 0) {
      msg += ` Nota: ${noEncontrados.length} registro(s) ya habían sido dados de baja previamente (${noEncontrados.join(", ")}).`;
    }

    return {
      exito: true,
      mensaje: msg,
      fechaGestion: marcaTiempoGestion,
      gestionadoPor: gestorLimpio,
      guardados: guardados
    };
  } catch (error) {
    throw new Error("Error al guardar: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function eliminarClienteFila(numFila, nombreEmpresa, rifEmpresa, usuarioAccion, pinAdmin) {
  exigirPinAdmin(pinAdmin);
  const responsable = limpiarNombrePersona(usuarioAccion) || "Supervisor";

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    asegurarColumnas(sheet);
    const datos = sheet.getDataRange().getValues();
    const filaReal = resolverFilaSegura(datos, numFila, rifEmpresa, nombreEmpresa);

    if (filaReal <= 1 || filaReal > datos.length) {
      throw new Error("El cliente \"" + (nombreEmpresa || "") + "\" ya no se encuentra en la cartera activa (posiblemente ya fue archivado).");
    }

    const marcaBaja = obtenerMarcaTemporalActual();
    const sheetBajas = obtenerHojaBajas();
    const filaArchivada = construirFilaBaja(datos[filaReal - 1], marcaBaja, responsable);

    sheetBajas.appendRow(filaArchivada);
    sheet.deleteRow(filaReal);
    SpreadsheetApp.flush();

    return {
      exito: true,
      mensaje: `Cliente "${nombreEmpresa || ""}" movido al histórico de bajas por ${responsable}.`
    };
  } catch (error) {
    throw new Error(error.message);
  } finally {
    lock.releaseLock();
  }
}

function depurarClientesListos(regionesSeleccionadas, usuarioAccion, pinAdmin) {
  exigirPinAdmin(pinAdmin);
  const responsable = limpiarNombrePersona(usuarioAccion) || "Supervisor";

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    asegurarColumnas(sheet);
    const sheetBajas = obtenerHojaBajas();
    const datos = sheet.getDataRange().getValues();
    const tz = Session.getScriptTimeZone() || "America/Caracas";
    const marcaTiempoBaja = obtenerMarcaTemporalActual();

    const todas = !regionesSeleccionadas || regionesSeleccionadas.includes("TODAS");
    const setRegiones = todas ? null : new Set(regionesSeleccionadas);

    const filasParaArchivar = [];

    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const empresa = String(fila[0] || "").trim();
      const region = String(fila[6] || "").trim();
      if (!empresa || esFilaEncabezadoOVacia(empresa, region)) continue;

      if (!todas && !setRegiones.has(region)) continue;

      const estatus = String(fila[7] || "").trim().toLowerCase();
      const fechaEnvioStr = formatearFechaRapidaInput(fila[8], tz);
      const estadoHubspot = String(fila[10] || "").trim().toUpperCase();

      const esLista = estatus.includes("lista");
      const tieneFechaEnvio = Boolean(fechaEnvioStr && fechaEnvioStr !== "--");
      const esHubspotValido = estadoHubspot !== "POR ACTUALIZAR" && (
        estadoHubspot.includes("ACTUALIZADO") ||
        estadoHubspot.includes("AVANZADO") ||
        estadoHubspot.includes("SINCRONIZADO")
      );

      if (esLista && tieneFechaEnvio && esHubspotValido) {
        filasParaArchivar.push({
          numFila: i + 1,
          valores: construirFilaBaja(fila, marcaTiempoBaja, responsable)
        });
      }
    }

    if (filasParaArchivar.length === 0) {
      return {
        exito: true,
        archivados: 0,
        mensaje: "No se encontraron clientes que cumplan los criterios (Estatus: Lista, con fecha de planillas y procesados en HubSpot como ACTUALIZADO, REASIGNADO o AVANZADO)."
      };
    }

    // 1. Guardar en bloque en Historico_Bajas (16 columnas)
    const filasValores = filasParaArchivar.map(item => item.valores);
    const ultimaFilaBajas = sheetBajas.getLastRow();
    sheetBajas.getRange(ultimaFilaBajas + 1, 1, filasValores.length, 16).setValues(filasValores);

    // 2. Eliminar de la hoja activa agrupando bloques contiguos en orden descendente
    const indicesEliminar = filasParaArchivar.map(item => item.numFila);
    eliminarFilasEnBloquesDescendentes(sheet, indicesEliminar);

    SpreadsheetApp.flush();

    return {
      exito: true,
      archivados: filasParaArchivar.length,
      mensaje: `Se depuraron y archivaron exitosamente ${filasParaArchivar.length} cliente(s) en Historico_Bajas (Ejecutado por: ${responsable}).`
    };
  } catch (error) {
    throw new Error("Error al depurar clientes: " + error.message);
  } finally {
    lock.releaseLock();
  }
}

function agregarNuevoCliente(nuevo, usuarioAccion, pinAdmin) {
  exigirPinAdmin(pinAdmin);
  const responsable = limpiarNombrePersona(usuarioAccion) || "Supervisor";

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sheet = obtenerHojaBD();
    asegurarColumnas(sheet);
    const datos = sheet.getDataRange().getValues();

    const empresaLimpia = String(nuevo.empresa || "").trim();
    const rifLimpio = String(nuevo.rif || "").trim().toUpperCase();
    const regionLimpia = String(nuevo.region || "").trim();
    const agenciaLimpia = String(nuevo.agencia || "").trim();

    if (!empresaLimpia || !rifLimpio || !regionLimpia || !agenciaLimpia) {
      throw new Error("Empresa, RIF, Región y Agencia son obligatorios.");
    }

    // Validar RIF duplicado en la cartera activa
    const rifComparar = normalizarRifComparacion(rifLimpio);
    if (rifComparar.length >= 5) {
      for (let i = 1; i < datos.length; i++) {
        if (normalizarRifComparacion(datos[i][1]) === rifComparar) {
          const empExistente = String(datos[i][0] || "").trim();
          const agExistente = String(datos[i][5] || "").trim();
          throw new Error(`Ya existe un cliente activo con el RIF ${rifLimpio} ("${empExistente}" asignado a ${agExistente}).`);
        }
      }
    }

    const tieneEstatus = esEstatusGestionadoValido(nuevo.estatus);
    const ahora = obtenerMarcaTemporalActual();

    const nuevaFila = [
      empresaLimpia,
      rifLimpio,
      String(nuevo.contacto || "").trim(),
      String(nuevo.telefono || "").trim(),
      String(nuevo.email || "").trim(),
      agenciaLimpia,
      regionLimpia,
      String(nuevo.estatus || "").trim(),
      String(nuevo.fechaEnvio || "").trim(),
      String(nuevo.comentarios || "NUEVO REGISTRO").trim(),
      tieneEstatus ? "POR ACTUALIZAR" : "",
      "",
      "",
      ahora,
      responsable
    ];

    sheet.appendRow(nuevaFila);
    SpreadsheetApp.flush();

    return {
      exito: true,
      mensaje: `Cliente "${empresaLimpia}" registrado satisfactoriamente por ${responsable}.`
    };
  } catch (error) {
    throw new Error(error.message);
  } finally {
    lock.releaseLock();
  }
}

function obtenerDatosExportacionCompleta(pinAdmin) {
  exigirPinAdmin(pinAdmin);

  const tz = Session.getScriptTimeZone() || "America/Caracas";
  const sheet = obtenerHojaBD();
  asegurarColumnas(sheet);
  const datos = sheet.getDataRange().getValues();

  const activa = [[
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
    "Fecha Gestión Agencia",
    "Gestionado Por"
  ]];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    const empresa = String(fila[0] || "").trim();
    const region = String(fila[6] || "").trim();
    if (esFilaEncabezadoOVacia(empresa, region)) continue;

    activa.push([
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
      formatearFechaRapidaTexto(fila[13], tz),
      String(fila[14] || "").trim()
    ]);
  }

  // Exportar también Historico_Bajas como segunda hoja del libro Excel
  const sheetBajas = obtenerHojaBajas();
  const datosBajas = sheetBajas.getDataRange().getValues();
  const bajas = [[
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
    "Fecha Gestión Agencia",
    "Fecha de Baja / Archivo",
    "Gestionado / Archivado Por"
  ]];

  for (let j = 1; j < datosBajas.length; j++) {
    const fb = datosBajas[j];
    const empB = String(fb[0] || "").trim();
    if (!empB) continue;

    bajas.push([
      empB,
      String(fb[1] || "").trim(),
      String(fb[2] || "").trim(),
      String(fb[3] || "").trim(),
      String(fb[4] || "").trim(),
      String(fb[5] || "").trim(),
      String(fb[6] || "").trim(),
      String(fb[7] || "").trim(),
      formatearFechaRapidaInput(fb[8], tz),
      String(fb[9] || "").trim(),
      String(fb[10] || "").trim(),
      formatearFechaRapidaTexto(fb[11], tz),
      String(fb[12] || "").trim(),
      formatearFechaRapidaTexto(fb[13], tz),
      formatearFechaRapidaTexto(fb[14], tz),
      String(fb[15] || "").trim()
    ]);
  }

  return {
    activa: activa,
    bajas: bajas
  };
}

/**
 * ============================================================================
 * 5. UTILIDADES RÁPIDAS EN MEMORIA
 * ============================================================================
 */

function limpiarNombrePersona(nombre) {
  return String(nombre || "")
    .replace(/[<>"]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizarRifComparacion(rif) {
  return String(rif || "").toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function normalizarTextoComparacion(txt) {
  return String(txt || "").toUpperCase().replace(/\s+/g, " ").trim();
}

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
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    // Soportar fechas pegadas como DD/MM/YYYY o DD-MM-YYYY en el Excel
    const mDMY = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (mDMY) {
      const dd = mDMY[1].padStart(2, "0");
      const mm = mDMY[2].padStart(2, "0");
      const yyyy = mDMY[3];
      return `${yyyy}-${mm}-${dd}`;
    }
    return "";
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
