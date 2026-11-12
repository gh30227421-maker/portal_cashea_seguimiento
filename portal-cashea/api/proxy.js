export default async function handler(req, res) {
  if (req.method !== 'POST') {
    return res.status(405).json({ status: 'error', mensaje: 'Método no permitido' });
  }

  // Pega aquí la URL terminada en /exec de tu Apps Script
  const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbzJikXwP9p37vyf-c7XXmmP39p_rPbeQbMj5E-h9VXMAfJ6kwvtWPwG3urnyF2jK2z/exec";

  try {
    const response = await fetch(SCRIPT_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(req.body),
      redirect: 'follow'
    });

    const data = await response.json();
    return res.status(200).json(data);
  } catch (error) {
    return res.status(500).json({ status: 'error', mensaje: 'Error conectando con Google: ' + error.message });
  }
}