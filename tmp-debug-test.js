const normalizeIdentityLabel = (value) => String(value || '')
  .normalize('NFKD')
  .replace(/[\u0300-\u036f]/g, '')
  .replace(/[^a-zA-Z0-9\s]/g, ' ')
  .replace(/\s+/g, ' ')
  .trim();

const getIdentityKey = (value) => {
  const normalizedValue = normalizeIdentityLabel(value);
  return normalizedValue ? normalizedValue.toLocaleLowerCase() : '';
};

const getIdentityDisplayScore = (value) => {
  const stringValue = String(value || '');
  const normalizedValue = normalizeIdentityLabel(value);
  if (!normalizedValue) {
    return 0;
  }

  let score = 0;
  if (normalizedValue !== normalizedValue.toLocaleLowerCase()) {
    score += 2;
  }
  if (/\b[A-Z]/.test(normalizedValue)) {
    score += 1;
  }
  if (/[^a-zA-Z0-9\s]/.test(stringValue)) {
    score += 2;
  }
  return score;
};

const getStringEditDistance = (a, b) => {
  const aLength = a.length;
  const bLength = b.length;
  if (!aLength) {
    return bLength;
  }
  if (!bLength) {
    return aLength;
  }

  const row = Array.from({ length: bLength + 1 }, (_, index) => index);
  let prev;
  for (let i = 1; i <= aLength; i += 1) {
    prev = row.slice();
    row[0] = i;
    for (let j = 1; j <= bLength; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      row[j] = Math.min(prev[j] + 1, row[j - 1] + 1, prev[j - 1] + cost);
    }
  }
  return row[bLength];
};

const getStringSimilarity = (a, b) => {
  const normalizedA = String(a || '').trim().toLowerCase();
  const normalizedB = String(b || '').trim().toLowerCase();
  if (!normalizedA || !normalizedB) {
    return 0;
  }
  return 1 - (getStringEditDistance(normalizedA, normalizedB) / Math.max(normalizedA.length, normalizedB.length));
};

const getUniqueValuesBySimilarity = (values, threshold = 0.95) => {
  const buckets = [];
  (Array.isArray(values) ? values : []).forEach((value) => {
    const normalized = String(value || '').trim();
    if (!normalized) {
      return;
    }
    const normalizedKey = getIdentityKey(normalized);
    if (!normalizedKey) {
      return;
    }
    let bucket = buckets.find((entry) => getStringSimilarity(normalizedKey, entry.key) >= threshold);
    if (!bucket) {
      bucket = { key: normalizedKey, variants: new Map() };
      buckets.push(bucket);
    }
    bucket.variants.set(normalized, (bucket.variants.get(normalized) || 0) + 1);
  });
  return buckets.map((bucket) => {
    let bestValue = '';
    let bestScore = -1;
    let bestCount = -1;
    bucket.variants.forEach((count, value) => {
      const score = getIdentityDisplayScore(value);
      if (
        score > bestScore ||
        (score === bestScore && count > bestCount) ||
        (score === bestScore && count === bestCount && value.localeCompare(bestValue) < 0)
      ) {
        bestValue = value;
        bestScore = score;
        bestCount = count;
      }
    });
    return bestValue;
  }).sort((a, b) => a.localeCompare(b));
};

console.log(getUniqueValuesBySimilarity(['Y shtola Night s Blessed', "Y'shtola, Night's Blessed", 'Zhulodok Void Gorger', 'Zhulodok, Void Gorger']));
console.log(getUniqueValuesBySimilarity(['Maralen Fae Ascendent', 'Maralen, Fae Ascendant']));
console.log(getIdentityDisplayScore('Maralen Fae Ascendent'), getIdentityDisplayScore('Maralen, Fae Ascendant'));