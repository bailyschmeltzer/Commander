const fs = require('fs');
const vm = require('vm');
const path = 'vscode-vfs://github/bailyschmeltzer/Commander/app.js';
const source = fs.readFileSync(path, 'utf8');
const start = source.indexOf('function readLocalStorageValue(key) {');
const end = source.indexOf('function loadSyncPendingChangesState()', start);
const helperSource = source.slice(start, end);
const memory = new Map();
const context = {
  storageErrorMessage: '',
  window: {
    localStorage: {
      getItem: (key) => (memory.has(key) ? memory.get(key) : null),
      setItem: (key, value) => memory.set(key, String(value)),
      removeItem: (key) => memory.delete(key),
    },
  },
  console,
};
vm.createContext(context);
vm.runInContext(helperSource, context);
const result = context.readLocalStorageValue('sample');
const wrote = context.writeLocalStorageValue('sample', 'value');
const roundTrip = context.readLocalStorageValue('sample');
if (result !== null || wrote !== true || roundTrip !== 'value') {
  throw new Error(`Unexpected storage helper behavior: ${JSON.stringify({ result, wrote, roundTrip })}`);
}
console.log('storage helper regression check passed');
