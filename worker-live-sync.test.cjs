const assert = require('node:assert/strict');
const { readFile } = require('node:fs/promises');
const { webcrypto } = require('node:crypto');
const path = require('node:path');
const { test } = require('node:test');

globalThis.crypto ||= webcrypto;

const stateKey = 'pod:default:state';
const sessionKey = 'pod:default:session:test-session';

async function createWorker() {
  const source = await readFile(path.join(__dirname, 'worker.js'), 'utf8');
  return (await import(`data:text/javascript;base64,${Buffer.from(source).toString('base64')}`)).default;
}

async function sendUpdate(worker, initialState, body) {
  const values = new Map([
    [stateKey, JSON.stringify(initialState)],
    [sessionKey, JSON.stringify({
      expiresAt: new Date(Date.now() + 60000).toISOString(),
      auth: { user: 'Test', userId: 'test', displayName: 'Test', role: 'member' },
    })],
  ]);
  const env = {
    POD_STATE: {
      get: async (key, type) => {
        const value = values.get(key);
        return value && type === 'json' ? JSON.parse(value) : value || null;
      },
      put: async (key, value) => values.set(key, value),
      delete: async (key) => values.delete(key),
    },
  };
  const request = new Request('https://example.test/api/state', {
    method: 'PUT',
    headers: {
      Cookie: 'commanderSession=test-session',
      'Content-Type': 'application/json',
      'X-State-Revision': String(initialState.revision),
    },
    body: JSON.stringify(body),
  });
  const response = await worker.fetch(request, env);
  return { response, state: JSON.parse(values.get(stateKey)) };
}

test('live-only sync preserves other collections', async () => {
  const worker = await createWorker();
  const initialState = {
    revision: 1,
    games: [],
    powerLevels: { Example: 7 },
    deckLists: [{ id: 'list' }],
    decks: [{ id: 'deck' }],
    records: [{ id: 'record' }],
    draftResults: [{ id: 'draft' }],
    activeGame: { id: 'game' },
    activeGameUndo: [],
  };
  const { response, state } = await sendUpdate(worker, initialState, {
    games: [],
    activeGame: { id: 'game', turnNumber: 2 },
    activeGameUndo: [{ id: 'game', turnNumber: 1 }],
  });
  assert.equal(response.status, 200, await response.text());
  assert.equal(state.activeGame.turnNumber, 2);
  assert.deepEqual(state.powerLevels, initialState.powerLevels);
  assert.deepEqual(state.deckLists, initialState.deckLists);
  assert.deepEqual(state.decks, initialState.decks);
  assert.deepEqual(state.records, initialState.records);
  assert.deepEqual(state.draftResults, initialState.draftResults);
});

test('live-only sync respects omitted fields and blocks empty overwrites', async () => {
  const worker = await createWorker();
  const initialState = { revision: 1, games: [], activeGame: { id: 'game' }, activeGameUndo: [] };
  const preserved = await sendUpdate(worker, initialState, { games: [], activeGameUndo: [] });
  assert.equal(preserved.response.status, 200, await preserved.response.text());
  assert.deepEqual(preserved.state.activeGame, initialState.activeGame);

  const emptied = await sendUpdate(worker, initialState, { games: [], activeGame: null, activeGameUndo: [] });
  assert.equal(emptied.response.status, 409, await emptied.response.text());
  assert.deepEqual(emptied.state, initialState);
});

test('live-only sync rejects stale revisions', async () => {
  const worker = await createWorker();
  const initialState = { revision: 2, games: [{ id: 'game' }], activeGame: null, activeGameUndo: [] };
  const staleState = { ...initialState, revision: 1 };
  const values = new Map([[stateKey, JSON.stringify(initialState)], [sessionKey, JSON.stringify({
    expiresAt: new Date(Date.now() + 60000).toISOString(),
    auth: { user: 'Test', userId: 'test', displayName: 'Test', role: 'member' },
  })]]);
  const env = { POD_STATE: {
    get: async (key, type) => type === 'json' ? JSON.parse(values.get(key) || 'null') : values.get(key) || null,
    put: async (key, value) => values.set(key, value),
  } };
  const request = new Request('https://example.test/api/state', {
    method: 'PUT',
    headers: { Cookie: 'commanderSession=test-session', 'X-State-Revision': String(staleState.revision) },
    body: JSON.stringify({ games: [], activeGame: { id: 'new' }, activeGameUndo: [] }),
  });
  const response = await worker.fetch(request, env);
  assert.equal(response.status, 409, await response.text());
  assert.deepEqual(JSON.parse(values.get(stateKey)), initialState);
});

test('state reads run alongside authentication auditing', async () => {
  const worker = await createWorker();
  let releaseAudit;
  let markStateRead;
  let auditWritten = false;
  let responseSettled = false;
  const auditBlocked = new Promise((resolve) => { releaseAudit = resolve; });
  const stateRead = new Promise((resolve) => { markStateRead = resolve; });
  const env = { POD_STATE: {
    get: async (key) => {
      if (key === sessionKey) {
        return {
          expiresAt: new Date(Date.now() + 60000).toISOString(),
          auth: { user: 'Test', userId: 'test', displayName: 'Test', role: 'member' },
        };
      }
      if (key === stateKey) {
        markStateRead();
        return { games: [], revision: 1 };
      }
      if (key === 'pod:default:auth-audit-log') {
        await auditBlocked;
      }
      return null;
    },
    put: async (key) => { auditWritten = key === 'pod:default:auth-audit-log'; },
  } };
  const request = new Request('https://example.test/api/state', {
    headers: { Cookie: 'commanderSession=test-session' },
  });
  const responsePromise = worker.fetch(request, env).then((response) => {
    responseSettled = true;
    return response;
  });
  let timeout;
  try {
    await Promise.race([
      stateRead,
      new Promise((resolve, reject) => {
        timeout = setTimeout(() => reject(new Error('State read waited for the audit log.')), 1000);
      }),
    ]);
    assert.equal(responseSettled, false);
  } finally {
    clearTimeout(timeout);
    releaseAudit();
  }
  const response = await responsePromise;
  assert.equal(response.status, 200);
  assert.equal(auditWritten, true);
  assert.equal((await response.json()).revision, 1);
});