export default async function handler(req, res) {
  // 1. Cabeceras CORS seguras (restringidas a dominios propios en Vercel y entorno local)
  const origin = req.headers.origin || '';
  const host = req.headers.host || '';
  const esOrigenPermitido =
    !origin ||
    origin.endsWith('.vercel.app') ||
    origin.includes('localhost') ||
    origin.includes('127.0.0.1') ||
    (host && origin.includes(host));

  if (origin && esOrigenPermitido) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
  } else if (!origin) {
    res.setHeader('Access-Control-Allow-Origin', '*');
  }

  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Methods', 'POST,OPTIONS');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  // Manejar preflight OPTIONS
  if (req.method === 'OPTIONS') {
    return esOrigenPermitido ? res.status(200).end() : res.status(403).end();
  }

  if (!esOrigenPermitido) {
    return res.status(403).json({
      status: 'error',
      mensaje: 'Origen no autorizado para consumir este proxy.'
    });
  }

  // Validar método POST
  if (req.method !== 'POST') {
    return res.status(405).json({
      status: 'error',
      mensaje: 'Método no permitido. Solo se aceptan solicitudes POST.'
    });
  }

  // URL del script de Google Apps Script (terminada en /exec)
  const SCRIPT_URL = process.env.GOOGLE_SCRIPT_URL || "https://script.google.com/macros/s/AKfycbzJikXwP9p37vyf-c7XXmmP39p_rPbeQbMj5E-h9VXMAfJ6kwvtWPwG3unJaY6Q88ld/exec";

  try {
    let bodyObj = {};
    if (typeof req.body === 'string') {
      try {
        bodyObj = JSON.parse(req.body);
      } catch {
        bodyObj = {};
      }
    } else if (req.body && typeof req.body === 'object') {
      bodyObj = { ...req.body };
    }

    // Inyectar token compartido opcional si está configurado en variables de entorno
    if (process.env.GAS_API_SECRET) {
      bodyObj._apiSecret = process.env.GAS_API_SECRET;
    }

    const payload = JSON.stringify(bodyObj);

    // Petición a Google Apps Script siguiendo redirecciones HTTP 302/307
    const response = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: payload,
      redirect: 'follow'
    });

    const responseText = await response.text();

    // Parseo seguro de la respuesta JSON
    try {
      const data = JSON.parse(responseText);
      return res.status(200).json(data);
    } catch (parseError) {
      console.error('La respuesta de Google Apps Script no es un JSON válido:', responseText);

      let mensajeError = 'Google Apps Script no devolvió un JSON válido.';
      if (responseText.includes('accounts.google.com') || responseText.includes('ServiceLogin')) {
        mensajeError = 'Error de autorización: La Web App de Google Apps Script requiere autenticación. Verifique que esté configurada con acceso para "Cualquiera" ("Anyone").';
      } else if (responseText.includes('ScriptError') || responseText.includes('Exceeded maximum execution time')) {
        mensajeError = 'Error de ejecución en Google Apps Script: Se excedió el tiempo límite o falló el script interno.';
      } else if (responseText.trim().startsWith('<') || responseText.includes('<html')) {
        mensajeError = 'Google Apps Script devolvió una página HTML en lugar de JSON. Verifique la implementación de la Web App.';
      }

      return res.status(502).json({
        status: 'error',
        mensaje: mensajeError,
        detalle: responseText.length > 500 ? responseText.substring(0, 500) + '...' : responseText
      });
    }
  } catch (error) {
    console.error('Error al conectar con Google Apps Script:', error);
    return res.status(500).json({
      status: 'error',
      mensaje: 'Error conectando con Google Apps Script: ' + error.message
    });
  }
}
