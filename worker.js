const CORS_HEADERS = {
  'Access-Control-Allow-Origin': '*',
  'Access-Control-Allow-Methods': 'GET, HEAD, PUT, OPTIONS',
  'Access-Control-Allow-Headers': 'Content-Type, X-User-Name, X-Pod-Token, X-State-Revision',
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
        state.revision = Number.isFinite(Number(state.revision)) ? Number(state.revision) : 0;
        state.updatedAt = String(state.updatedAt || '').trim();
        state.updatedBy = String(state.updatedBy || '').trim();

        if (url.searchParams.get('meta') === '1') {
          return jsonResponse({
            revision: state.revision,
            updatedAt: state.updatedAt,
            updatedBy: state.updatedBy,
          }, 200);
        }

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
        const expectedRevisionHeader = request.headers.get('X-State-Revision');
        const expectedRevision = Number.parseInt(String(expectedRevisionHeader || '').trim(), 10);

        if (!Number.isFinite(expectedRevision)) {
          return jsonResponse({
            error: 'Sync client is missing cloud revision metadata. Refresh the page and reconnect before syncing again.',
          }, 409);
        }

        const currentRaw = await env.POD_STATE.get(stateKey, 'json');
        const currentState = currentRaw && typeof currentRaw === 'object' ? currentRaw : null;
        const currentRevision = Number.isFinite(Number(currentState?.revision)) ? Number(currentState.revision) : 0;

        if (expectedRevision !== currentRevision) {
          return jsonResponse({
            error: 'Cloud state changed on another device before this sync completed.',
            conflict: {
              revision: currentRevision,
              updatedAt: String(currentState?.updatedAt || '').trim(),
              updatedBy: String(currentState?.updatedBy || '').trim(),
            },
          }, 409);
        }

        const updatedAt = new Date().toISOString();
        const nextRevision = currentRevision + 1;

        await env.POD_STATE.put(stateKey, JSON.stringify({
          games,
          powerLevels,
          deckLists,
          records,
          revision: nextRevision,
          updatedAt,
          updatedBy: auth.user,
        }));

        return jsonResponse({ ok: true, revision: nextRevision, updatedAt, updatedBy: auth.user }, 200);
      }

      return jsonResponse({ error: 'Method not allowed.' }, 405);
    }

    return env.ASSETS.fetch(request);
  },
};
