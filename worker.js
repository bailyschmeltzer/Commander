const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, PUT, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, X-User-Name, X-Pod-Token',
};

function jsonResponse(body, status = 200) {
  return new Response(JSON.stringify(body), {
    status,
    headers: {
      'Content-Type': 'application/json',
      ...CORS_HEADERS,
    },
  });
}

function getRequestUser(request) {
  return (request.headers.get('X-User-Name') || '').trim();
}

function getRequestToken(request) {
  return (request.headers.get('X-Pod-Token') || '').trim();
}

function hasValidAuth(request, env) {
  const user = getRequestUser(request);
  const token = getRequestToken(request);
  const configuredToken = (env.POD_ACCESS_TOKEN || '').trim();

  if (!configuredToken) {
    return { ok: false, reason: 'Server missing POD_ACCESS_TOKEN.' };
  }

  if (!env.POD_STATE) {
    return { ok: false, reason: 'Server missing POD_STATE KV binding.' };
  }

  if (!user || !token) {
    return { ok: false, reason: 'Missing credentials.' };
  }

  if (token !== configuredToken) {
    return { ok: false, reason: 'Invalid pod access code.' };
  }

  return { ok: true, user };
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS });
    }

    if (url.pathname === '/api/state') {
      const auth = hasValidAuth(request, env);
      if (!auth.ok) {
        return jsonResponse({ error: auth.reason }, 401);
      }

      const stateKey = 'pod:default:state';

      if (request.method === 'GET') {
        const raw = await env.POD_STATE.get(stateKey, 'json');
        const state = raw && typeof raw === 'object' ? raw : { games: [], powerLevels: {}, deckLists: [], records: [] };
        state.deckLists = Array.isArray(state.deckLists) ? state.deckLists : [];
        state.records = Array.isArray(state.records) ? state.records : [];
        return jsonResponse(state, 200);
      }

      if (request.method === 'PUT') {
        let body;
        try {
          body = await request.json();
        } catch (error) {
          return jsonResponse({ error: 'Invalid JSON payload.' }, 400);
        }

        const games = Array.isArray(body.games) ? body.games : [];
        const powerLevels = body.powerLevels && typeof body.powerLevels === 'object' ? body.powerLevels : {};
        const deckLists = Array.isArray(body.deckLists) ? body.deckLists : [];
        const records = Array.isArray(body.records) ? body.records : [];

        await env.POD_STATE.put(stateKey, JSON.stringify({
          games,
          powerLevels,
          deckLists,
          records,
          updatedAt: new Date().toISOString(),
          updatedBy: auth.user,
        }));

        return jsonResponse({ ok: true }, 200);
      }

      return jsonResponse({ error: 'Method not allowed.' }, 405);
    }

    return env.ASSETS.fetch(request);
  },
};
