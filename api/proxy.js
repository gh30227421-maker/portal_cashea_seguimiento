export default async function handler(req, res) {
  // 1. Cabeceras CORS adecuadas para permitir peticiones desde cualquier origen y preflight OPTIONS
  res.setHeader('Access-Control-Allow-Credentials', 'true');
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET,OPTIONS,PATCH,DELETE,POST,PUT');
  res.setHeader(
    'Access-Control-Allow-Headers',
    'X-CSRF-Token, X-Requested-With, Accept, Accept-Version, Content-Length, Content-MD5, Content-Type, Date, X-Api-Version'
  );

  // Manejar preflight OPTIONS
  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  // Validar método POST
  if (req.method !== 'POST') {
    return res.status(405).json({
      status: 'error',
      mensaje: 'Método no permitido. Solo se aceptan solicitudes POST.'
    });
  }

  // URL del script de Google Apps Script (terminada en /exec)
  // Puede configurarse mediante la variable de entorno GOOGLE_SCRIPT_URL en Vercel o directamente aquí
  const SCRIPT_URL = process.env.GOOGLE_SCRIPT_URL || "https://script.google.com/macros/s/AKfycbzJikXwP9p37vyf-c7XXmmP39p_rPbeQbMj5E-h9VXMAfJ6kwvtWPwG3urnyF2jK2z/exec";

  try {
    // Asegurar que el cuerpo de la petición se envíe como cadena JSON limpia
    const payload = typeof req.body === 'string' ? req.body : JSON.stringify(req.body || {});

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
      // Capturar casos donde Google devuelve HTML (error de permisos, cuota o login)
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
