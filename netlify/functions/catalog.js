exports.handler = async function () {
  const base = process.env.AEST_APPS_SCRIPT_URL;
  if (!base) return json(500, { error: 'Falta configurar AEST_APPS_SCRIPT_URL en Netlify.' });
  try {
    const url = new URL(base);
    url.searchParams.set('api', 'catalog');
    const response = await fetch(url.toString(), { redirect: 'follow' });
    const body = await response.text();
    if (!response.ok) return json(502, { error: 'Google Apps Script respondió con estado ' + response.status + '.' });
    let data;
    try { data = JSON.parse(body); } catch (_) {
      return json(502, { error: 'La respuesta de Google no es JSON. Revisa que la implementación esté disponible para cualquier usuario.' });
    }
    if (!data || !data.tienda || !Array.isArray(data.productos)) return json(502, { error: 'La respuesta de Google no contiene el catálogo esperado.' });
    return json(200, data);
  } catch (error) {
    return json(502, { error: 'No se pudo consultar Google Apps Script: ' + error.message });
  }
};

function json(statusCode, body) {
  return {
    statusCode,
    headers: { 'Content-Type': 'application/json; charset=utf-8', 'Cache-Control': 'no-store' },
    body: JSON.stringify(body)
  };
}
