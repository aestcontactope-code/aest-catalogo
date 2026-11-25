exports.handler = async function (event) {
  const base = process.env.AEST_APPS_SCRIPT_URL;
  if (!base) return json(500, { ok: false, message: 'Falta configurar AEST_APPS_SCRIPT_URL en Netlify.' });
  if (event.httpMethod !== 'POST') return json(405, { ok: false, message: 'Método no permitido.' });
  try {
    const input = JSON.parse(event.body || '{}');
    if (!Array.isArray(input.items)) return json(400, { ok: false, message: 'No se recibieron productos para validar.' });
    const url = new URL(base);
    url.searchParams.set('api', 'validate');
    url.searchParams.set('items', JSON.stringify(input.items));
    const response = await fetch(url.toString(), { redirect: 'follow' });
    const body = await response.text();
    if (!response.ok) return json(502, { ok: false, message: 'Google Apps Script respondió con estado ' + response.status + '.' });
    let result;
    try { result = JSON.parse(body); } catch (_) {
      return json(502, { ok: false, message: 'Google Apps Script devolvió una respuesta no válida.' });
    }
    return json(200, result);
  } catch (error) {
    return json(502, { ok: false, message: 'No se pudo validar el stock: ' + error.message });
  }
};

function json(statusCode, body) {
  return {
    statusCode,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
    body: JSON.stringify(body)
  };
}
