/* eslint-disable no-console */
// State management and computation utilities for the docent planner.
// The module exposes constants, helpers, pure computation functions and
// a lightweight reactive store implementation following the specification.

export const EVAL_MODE = Object.freeze({
  COMPETENCIAL: 'competencial',
  NUMERIC: 'numeric',
});

export const QUALI = Object.freeze(['NA', 'AS', 'AN', 'AE']);

const QUALI_SET = new Set(QUALI);
const QUALI_STEP = 0.1;
const RANGE_EPSILON = 1e-6;

export const DEFAULTS = Object.freeze({
  rounding: Object.freeze({ decimals: 1, mode: 'half-up' }),
  qualitativeRanges: Object.freeze({
    NA: Object.freeze([0.0, 4.9]),
    AS: Object.freeze([5.0, 6.9]),
    AN: Object.freeze([7.0, 8.9]),
    AE: Object.freeze([9.0, 10.0]),
  }),
  categoriesInit: Object.freeze([
    'Deures',
    'Actitud',
    'Treball a classe',
    'Examen',
    'Projecte',
  ]),
  attendanceThresholds: Object.freeze({
    tardancaMenor: 10,
    tardancaMajor: 30,
  }),
});

function structuredClonePolyfill(value) {
  if (typeof globalThis.structuredClone === 'function') {
    return globalThis.structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

export function assertQuali(value) {
  assert(QUALI_SET.has(value), `Valor qualitatiu desconegut: ${value}`);
}

function assertDecimals(decimals) {
  assert(
    Number.isInteger(decimals) && decimals >= 0 && decimals <= 3,
    `Decimals permesos 0..3, rebut: ${decimals}`,
  );
}

function assertNonNegative(number, context) {
  assert(
    typeof number === 'number' && Number.isFinite(number) && number >= 0,
    `${context || 'valor'} ha de ser un número ≥ 0`,
  );
}

function cloneState(state) {
  return structuredClonePolyfill(state);
}

function isEqual(a, b) {
  return JSON.stringify(a) === JSON.stringify(b);
}

export function formatCE({ textBetween = '', index }) {
  assert(Number.isInteger(index) && index >= 1, "L'índex del CE ha de ser ≥ 1");
  return `CE${textBetween}${index}`;
}

export function formatCA({ textBetween = '', ceIndex, index }) {
  assert(Number.isInteger(ceIndex) && ceIndex >= 1, "L'índex del CE ha de ser ≥ 1");
  assert(Number.isInteger(index) && index >= 1, "L'índex del CA ha de ser ≥ 1");
  return `CA${textBetween}${ceIndex}.${index}`;
}

export function assertCEIdRules(id) {
  const match = /([0-9]+)\s*$/u.exec(String(id));
  assert(match, `Identificador CE invàlid: ${id}`);
  const index = Number(match[1]);
  assert(index >= 1, `L'índex de CE ha de ser ≥ 1 (${id})`);
  return index;
}

export function assertCAIdRules(id) {
  const match = /([0-9]+)\.([0-9]+)\s*$/u.exec(String(id));
  assert(match, `Identificador CA invàlid: ${id}`);
  const ceIndex = Number(match[1]);
  const index = Number(match[2]);
  assert(ceIndex >= 1, `L'índex CE del CA ha de ser ≥ 1 (${id})`);
  assert(index >= 1, `L'índex del CA ha de ser ≥ 1 (${id})`);
  return { ceIndex, index };
}

export function roundTo(value, decimals = 0, mode = 'half-up') {
  assert(typeof value === 'number' && Number.isFinite(value), 'Cal un nombre finit a roundTo');
  assertDecimals(decimals);
  if (mode !== 'half-up') {
    throw new Error(`Mode d'arrodoniment no suportat: ${mode}`);
  }
  const factor = 10 ** decimals;
  const rounded = Math.round(value * factor + Number.EPSILON) / factor;
  return Number(rounded.toFixed(decimals));
}

export function validateQualitativeRanges(ranges) {
  assert(ranges && typeof ranges === 'object', 'Les franges qualitatives han de ser un objecte');
  const clean = {};
  QUALI.forEach((key) => {
    const range = ranges[key];
    assert(Array.isArray(range) && range.length === 2, `Franja qualitativa ${key} ha de ser [min,max]`);
    const min = Number(range[0]);
    const max = Number(range[1]);
    assert(Number.isFinite(min) && Number.isFinite(max), `Franja qualitativa ${key} ha de contenir números`);
    assert(min <= max, `Franja qualitativa ${key} ha de complir min ≤ max`);
    clean[key] = [min, max];
  });

  const firstRange = clean[QUALI[0]];
  const lastRange = clean[QUALI[QUALI.length - 1]];
  assert(firstRange[0] <= 0 + RANGE_EPSILON, 'Les franges qualitatives han de començar a 0');
  assert(lastRange[1] >= 10 - RANGE_EPSILON, 'Les franges qualitatives han d\'arribar fins a 10');

  const ticks = [];
  for (let i = 0; i <= Math.round(10 / QUALI_STEP); i += 1) {
    const value = Number((i * QUALI_STEP).toFixed(1));
    ticks.push(value);
  }

  ticks.forEach((value) => {
    let matches = 0;
    QUALI.forEach((key) => {
      const [min, max] = clean[key];
      if (value >= min - RANGE_EPSILON && value <= max + RANGE_EPSILON) {
        matches += 1;
      }
    });
    assert(matches > 0, `Valor ${value} fora de les franges qualitatives`);
    assert(matches === 1, `Valor ${value} apareix en múltiples franges qualitatives`);
  });

  return clean;
}

export function validateWeights(weights = {}) {
  assert(weights && typeof weights === 'object', 'Cal un objecte de pesos');
  const clean = {};
  Object.entries(weights).forEach(([key, value]) => {
    if (value === undefined || value === null) return;
    const num = Number(value);
    assert(Number.isFinite(num) && num >= 0, `El camp ${key} ha de ser un nombre finit ≥ 0`);
    clean[key] = num;
  });
  return clean;
}

function getAssignatura(state, assignaturaId) {
  const assignatura = state.assignatures?.byId?.[assignaturaId];
  assert(assignatura, `Assignatura ${assignaturaId} no trobada`);
  return assignatura;
}

const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/u;

function toISODate(value) {
  if (!value && value !== 0) return null;
  if (typeof value === 'string') {
    if (ISO_DATE_RE.test(value)) return value;
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return null;
    return parsed.toISOString().slice(0, 10);
  }
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'number') {
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) return null;
    return date.toISOString().slice(0, 10);
  }
  return null;
}

function toUTCDateAtStartOfDay(value) {
  const iso = typeof value === 'string' && ISO_DATE_RE.test(value) ? value : toISODate(value);
  if (!iso) return null;
  const parsed = new Date(`${iso}T00:00:00Z`);
  if (Number.isNaN(parsed.getTime())) return null;
  return parsed;
}

function getUTCWeekday(value) {
  const date = toUTCDateAtStartOfDay(value);
  if (!date) return null;
  return date.getUTCDay();
}

function compareISODate(a, b) {
  if (!a && !b) return 0;
  if (!a) return -1;
  if (!b) return 1;
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

function ensureArray(value) {
  return Array.isArray(value) ? value : [];
}

function normalizeFestiu(entry) {
  if (!entry && entry !== 0) return null;
  if (typeof entry === 'string') {
    const dataISO = toISODate(entry);
    if (!dataISO) return null;
    return { dataISO, motiu: '' };
  }
  if (typeof entry === 'object') {
    const dataISO = toISODate(entry.dataISO || entry.date || entry.data);
    if (!dataISO) return null;
    return { dataISO, motiu: entry.motiu ? String(entry.motiu) : '' };
  }
  return null;
}

function normalizeExcepcio(entry) {
  if (!entry || typeof entry !== 'object') return null;
  const dataISO = toISODate(entry.dataISO || entry.date || entry.data);
  if (!dataISO) return null;
  return {
    dataISO,
    motiu: entry.motiu ? String(entry.motiu) : '',
  };
}

function getTrimestreRange(state, assignaturaId, trimestreId) {
  if (!trimestreId) return null;
  const calendaris = Object.values(state.calendaris.byId || {}).filter(
    (cal) => cal.assignaturaId === assignaturaId,
  );
  for (const calendari of calendaris) {
    for (const trimestre of ensureArray(calendari.trimestres)) {
      if (trimestre.id === trimestreId) {
        return { inici: trimestre.tInici, fi: trimestre.tFi };
      }
    }
  }
  return null;
}

function isDateWithin(date, range) {
  if (!range) return true;
  const time = new Date(date).getTime();
  const start = new Date(range.inici).getTime();
  const end = new Date(range.fi).getTime();
  return time >= start && time <= end;
}

export function qualiToNum(assignatura, quali) {
  assertQuali(quali);
  const ranges = assignatura?.qualitativeRanges || DEFAULTS.qualitativeRanges;
  const [min, max] = ranges[quali];
  return (min + max) / 2;
}

export function numToQuali(assignatura, num) {
  const ranges = assignatura?.qualitativeRanges || DEFAULTS.qualitativeRanges;
  let found = QUALI[0];
  for (const quali of QUALI) {
    const [min, max] = ranges[quali];
    if (num >= min && num <= max) {
      found = quali;
      break;
    }
  }
  return found;
}

export function createEmptyState() {
  return {
    version: 0,
    user: { id: 'me', name: 'Docent', mfaEnabled: false },
    assignatures: { byId: {}, allIds: [] },
    alumnes: { byId: {}, allIds: [] },
    matriculacions: { byId: {}, allIds: [] },
    calendaris: { byId: {}, allIds: [] },
    ces: { byId: {}, allIds: [] },
    cas: { byId: {}, allIds: [] },
    categories: { byId: {}, allIds: [] },
    activitats: { byId: {}, allIds: [] },
    activitatCA: { byId: {}, allIds: [] },
    avaluacionsComp: { byId: {}, allIds: [] },
    avaluacionsNum: { byId: {}, allIds: [] },
    assistencia: { byId: {}, allIds: [] },
    incidencies: { byId: {}, allIds: [] },
    vincles: { byId: {}, allIds: [] },
    historial: [],
    configGlobal: {
      rounding: cloneState(DEFAULTS.rounding),
      categoriesInit: [...DEFAULTS.categoriesInit],
      attendanceThresholds: cloneState(DEFAULTS.attendanceThresholds),
    },
  };
}

function collectMatriculats(state, assignaturaId) {
  const matriculacions = state.matriculacions.allIds
    .map((id) => state.matriculacions.byId[id])
    .filter((m) => m.assignaturaId === assignaturaId);
  return matriculacions.map((m) => m.alumneId);
}

function collectCAsForCE(state, ceId) {
  return state.cas.allIds
    .map((id) => state.cas.byId[id])
    .filter((ca) => ca.ceId === ceId);
}

function collectCAActivities(state, caId) {
  return state.activitatCA.allIds
    .map((id) => state.activitatCA.byId[id])
    .filter((link) => link.caId === caId);
}
function filterActivitatsByAssignatura(state, assignaturaId) {
  return state.activitats.allIds
    .map((id) => state.activitats.byId[id])
    .filter((act) => act.assignaturaId === assignaturaId);
}

function filterActivitatsByTrimestre(activitats, range) {
  if (!range) return activitats;
  return activitats.filter((act) => isDateWithin(act.data, range));
}

function sumWeights(weights) {
  return weights.reduce((sum, w) => sum + (Number.isFinite(w) ? w : 0), 0);
}

export function computeCAForAlumne(state, assignaturaId, alumneId, caId, opts = {}) {
  const assignatura = getAssignatura(state, assignaturaId);
  const ca = state.cas.byId[caId];
  assert(ca, `CA ${caId} no trobat`);
  const links = collectCAActivities(state, caId);
  const range = opts.trimestreId ? getTrimestreRange(state, assignaturaId, opts.trimestreId) : null;
  const relevantLinks = links.filter((link) => {
    const activitat = state.activitats.byId[link.activitatId];
    return activitat && activitat.assignaturaId === assignaturaId && isDateWithin(activitat.data, range);
  });
  const weights = relevantLinks.map((link) => {
    const w = Number(link.pes_ca || 0);
    return w >= 0 ? w : 0;
  });
  const totalWeight = sumWeights(weights);
  const divisor = totalWeight === 0 ? relevantLinks.length || 1 : totalWeight;
  let accumulator = 0;
  relevantLinks.forEach((link, idx) => {
    const registres = state.avaluacionsComp.allIds
      .map((id) => state.avaluacionsComp.byId[id])
      .filter(
        (av) =>
          av.alumneId === alumneId &&
          av.activitatId === link.activitatId &&
          av.caId === caId,
      );
    registres.forEach((registre) => {
      const valueNum = qualiToNum(assignatura, registre.valorQuali);
      const weight = totalWeight === 0 ? 1 : weights[idx] || 0;
      accumulator += valueNum * (weight || 1);
    });
  });
  const raw = divisor === 0 ? 0 : accumulator / divisor;
  const decimals = assignatura.rounding?.decimals ?? DEFAULTS.rounding.decimals;
  const mode = assignatura.rounding?.mode ?? DEFAULTS.rounding.mode;
  const rounded = roundTo(raw, decimals, mode);
  return {
    valueNumRaw: raw,
    valueNumRounded: rounded,
    quali: numToQuali(assignatura, raw),
  };
}

export function computeCEForAlumne(state, assignaturaId, alumneId, ceId, opts = {}) {
  const assignatura = getAssignatura(state, assignaturaId);
  const ce = state.ces.byId[ceId];
  assert(ce, `CE ${ceId} no trobat`);
  const cas = collectCAsForCE(state, ceId);
  const weights = cas.map((ca) => ca.pesDinsCE ?? 0);
  const totalWeight = sumWeights(weights);
  const divisor = totalWeight === 0 ? (cas.length || 1) : totalWeight;
  let accumulator = 0;
  cas.forEach((ca, idx) => {
    const result = computeCAForAlumne(state, assignaturaId, alumneId, ca.id, opts);
    const weight = totalWeight === 0 ? 1 : (weights[idx] || 0);
    accumulator += result.valueNumRaw * (weight || 1);
  });
  const raw = accumulator / divisor;
  const decimals = assignatura.rounding?.decimals ?? DEFAULTS.rounding.decimals;
  const mode = assignatura.rounding?.mode ?? DEFAULTS.rounding.mode;
  const rounded = roundTo(raw, decimals, mode);
  return {
    valueNumRaw: raw,
    valueNumRounded: rounded,
    quali: numToQuali(assignatura, raw),
  };
}

export function computeTaulaCompetencial(state, assignaturaId, trimestreId) {
  const assignatura = getAssignatura(state, assignaturaId);
  assert(assignatura.mode === EVAL_MODE.COMPETENCIAL, 'Assignatura no és en mode competencial');
  const alumnesIds = collectMatriculats(state, assignaturaId);
  const cas = state.cas.allIds
    .map((id) => state.cas.byId[id])
    .filter((ca) => state.ces.byId[ca.ceId]?.assignaturaId === assignaturaId);
  const values = {};
  alumnesIds.forEach((alumneId) => {
    values[alumneId] = {};
    cas.forEach((ca) => {
      values[alumneId][ca.id] = computeCAForAlumne(state, assignaturaId, alumneId, ca.id, {
        trimestreId,
      });
    });
  });
  return {
    assignaturaId,
    trimestreId: trimestreId || null,
    alumnes: alumnesIds,
    cas: cas.map((ca) => ca.id),
    values,
    rounding: assignatura.rounding,
    qualitativeRanges: assignatura.qualitativeRanges,
  };
}

export function computeNotaPerAlumne(state, assignaturaId, alumneId, trimestreId) {
  const assignatura = getAssignatura(state, assignaturaId);
  assert(assignatura.mode === EVAL_MODE.NUMERIC, 'Assignatura no és en mode numèric');
  const range = trimestreId ? getTrimestreRange(state, assignaturaId, trimestreId) : null;
  const activitats = filterActivitatsByAssignatura(state, assignaturaId);
  const activitatsFiltrades = filterActivitatsByTrimestre(activitats, range);
  const activitatsById = new Map(activitatsFiltrades.map((act) => [act.id, act]));
  const registres = state.avaluacionsNum.allIds
    .map((id) => state.avaluacionsNum.byId[id])
    .filter((av) => av.alumneId === alumneId && activitatsById.has(av.activitatId));
  const categoriaMap = new Map();
  registres.forEach((registre) => {
    const activitat = activitatsById.get(registre.activitatId);
    if (!activitat) return;
    const categoriaId = activitat.categoriaId;
    const pesActivitat = Number(activitat.pesActivitat ?? 0);
    const bucket = categoriaMap.get(categoriaId) || { total: 0, weight: 0 };
    const weight = pesActivitat >= 0 ? pesActivitat : 0;
    bucket.total += registre.valorNum * (weight || 1);
    bucket.weight += weight || 1;
    categoriaMap.set(categoriaId, bucket);
  });
  const categoriaValues = new Map();
  categoriaMap.forEach((bucket, categoriaId) => {
    const avg = bucket.weight === 0 ? 0 : bucket.total / bucket.weight;
    categoriaValues.set(categoriaId, avg);
  });
  const pesGlobal =
    (assignatura.categoriaPesos && trimestreId && assignatura.categoriaPesos[trimestreId]) || {};
  let totalWeight = 0;
  let accumulator = 0;
  if (categoriaValues.size === 0) {
    const decimals = assignatura.rounding?.decimals ?? DEFAULTS.rounding.decimals;
    const mode = assignatura.rounding?.mode ?? DEFAULTS.rounding.mode;
    return { valueNumRaw: 0, valueNumRounded: roundTo(0, decimals, mode) };
  }
  categoriaValues.forEach((value, categoriaId) => {
    let weight = pesGlobal[categoriaId];
    if (weight === undefined) {
      weight = 0;
    }
    if (weight < 0 || !Number.isFinite(weight)) {
      weight = 0;
    }
    totalWeight += weight;
    accumulator += value * weight;
  });
  if (totalWeight === 0) {
    totalWeight = categoriaValues.size;
    accumulator = 0;
    categoriaValues.forEach((value) => {
      accumulator += value;
    });
  }
  const raw = accumulator / totalWeight;
  const decimals = assignatura.rounding?.decimals ?? DEFAULTS.rounding.decimals;
  const mode = assignatura.rounding?.mode ?? DEFAULTS.rounding.mode;
  const rounded = roundTo(raw, decimals, mode);
  return { valueNumRaw: raw, valueNumRounded: rounded };
}

export function reorderCE(state, assignaturaId, ceId, newPosition, { compact = true } = {}) {
  assert(Number.isInteger(newPosition) && newPosition >= 1, 'Nova posició CE ha de ser ≥ 1');
  const ces = state.ces.allIds
    .map((id) => state.ces.byId[id])
    .filter((ce) => ce.assignaturaId === assignaturaId)
    .sort((a, b) => (a.position || 0) - (b.position || 0));
  const ce = ces.find((item) => item.id === ceId);
  assert(ce, `CE ${ceId} no trobat per reordenar`);
  const filtered = ces.filter((item) => item.id !== ceId);
  const index = Math.min(Math.max(newPosition - 1, 0), filtered.length);
  filtered.splice(index, 0, ce);
  const used = new Set();
  filtered.forEach((item, idx) => {
    const entry = state.ces.byId[item.id];
    if (compact) {
      entry.position = idx + 1;
      used.add(entry.position);
      return;
    }
    let desired = item.id === ceId ? newPosition : entry.position ?? idx + 1;
    desired = Number.isFinite(desired) && desired >= 1 ? Math.floor(desired) : idx + 1;
    while (used.has(desired)) {
      desired += 1;
    }
    entry.position = desired;
    used.add(entry.position);
  });
}

export function reorderCA(state, ceId, caId, newPosition, { compact = true } = {}) {
  assert(Number.isInteger(newPosition) && newPosition >= 1, 'Nova posició CA ha de ser ≥ 1');
  const cas = collectCAsForCE(state, ceId).sort((a, b) => (a.position || 0) - (b.position || 0));
  const ca = cas.find((item) => item.id === caId);
  assert(ca, `CA ${caId} no trobat per reordenar`);
  const filtered = cas.filter((item) => item.id !== caId);
  const index = Math.min(Math.max(newPosition - 1, 0), filtered.length);
  filtered.splice(index, 0, ca);
  const used = new Set();
  filtered.forEach((item, idx) => {
    const entry = state.cas.byId[item.id];
    if (compact) {
      entry.position = idx + 1;
      used.add(entry.position);
      return;
    }
    let desired = item.id === caId ? newPosition : entry.position ?? idx + 1;
    desired = Number.isFinite(desired) && desired >= 1 ? Math.floor(desired) : idx + 1;
    while (used.has(desired)) {
      desired += 1;
    }
    entry.position = desired;
    used.add(entry.position);
  });
}
export function isInSameVincle(state, assignaturaIdA, assignaturaIdB) {
  return state.vincles.allIds.some((id) => {
    const vincle = state.vincles.byId[id];
    const ids = vincle.assignaturaIds || [];
    return ids.includes(assignaturaIdA) && ids.includes(assignaturaIdB);
  });
}

export function getVincle(state, assignaturaId) {
  return (
    state.vincles.allIds
      .map((id) => state.vincles.byId[id])
      .find((vincle) => (vincle.assignaturaIds || []).includes(assignaturaId)) || null
  );
}

function extractSyncEntities(state, assignaturaId) {
  const assignatura = getAssignatura(state, assignaturaId);
  const ces = state.ces.allIds
    .map((id) => state.ces.byId[id])
    .filter((ce) => ce.assignaturaId === assignaturaId);
  const cas = state.cas.allIds
    .map((id) => state.cas.byId[id])
    .filter((ca) => ces.some((ce) => ce.id === ca.ceId));
  const activitats = state.activitats.allIds
    .map((id) => state.activitats.byId[id])
    .filter((act) => act.assignaturaId === assignaturaId);
  const calendaris = state.calendaris.allIds
    .map((id) => state.calendaris.byId[id])
    .filter((cal) => cal.assignaturaId === assignaturaId);
  const festius = calendaris.flatMap((cal) => cal.festius || []);
  return {
    assignatura,
    ces,
    cas,
    activitats,
    festius,
  };
}

export function diffVincle(state, sourceAssignaturaId, targetAssignaturaId) {
  const source = extractSyncEntities(state, sourceAssignaturaId);
  const target = extractSyncEntities(state, targetAssignaturaId);
  const ceIdsTarget = new Set(target.ces.map((ce) => ce.id));
  const caIdsTarget = new Set(target.cas.map((ca) => ca.id));
  const actIdsTarget = new Set(target.activitats.map((act) => act.id));

  const missingCE = source.ces.filter((ce) => !ceIdsTarget.has(ce.id));
  const missingCA = source.cas.filter((ca) => !caIdsTarget.has(ca.id));
  const missingActivitats = source.activitats.filter((act) => !actIdsTarget.has(act.id));

  const festiusSource = new Set((source.festius || []).map((d) => new Date(d).toISOString()));
  const festiusTarget = new Set((target.festius || []).map((d) => new Date(d).toISOString()));
  const festiusMissing = [...festiusSource].filter((d) => !festiusTarget.has(d));

  const configDiff = {
    rounding: !isEqual(source.assignatura.rounding, target.assignatura.rounding),
    qualitativeRanges: !isEqual(source.assignatura.qualitativeRanges, target.assignatura.qualitativeRanges),
  };

  return {
    sourceAssignaturaId,
    targetAssignaturaId,
    missing: {
      ces: missingCE,
      cas: missingCA,
      activitats: missingActivitats,
      festius: festiusMissing,
    },
    configDiff,
  };
}

export function getFitxaAlumne(state, alumneId, opts = {}) {
  const assignaturaFilter = opts.assignaturaId
    ? new Set(Array.isArray(opts.assignaturaId) ? opts.assignaturaId : [opts.assignaturaId])
    : null;
  const range = opts.trimestreId && opts.assignaturaId
    ? getTrimestreRange(
        state,
        Array.isArray(opts.assignaturaId) ? opts.assignaturaId[0] : opts.assignaturaId,
        opts.trimestreId,
      )
    : null;
  const from = opts.from ? new Date(opts.from).getTime() : null;
  const to = opts.to ? new Date(opts.to).getTime() : null;

  const withinOptionalRange = (date) => {
    const time = new Date(date).getTime();
    if (Number.isFinite(from) && time < from) return false;
    if (Number.isFinite(to) && time > to) return false;
    if (range && !isDateWithin(date, range)) return false;
    return true;
  };

  const assistencies = state.assistencia.allIds
    .map((id) => state.assistencia.byId[id])
    .filter((entry) => entry.alumneId === alumneId)
    .filter((entry) => !assignaturaFilter || assignaturaFilter.has(entry.assignaturaId))
    .filter((entry) => withinOptionalRange(entry.dataHora));
  const incidencies = state.incidencies.allIds
    .map((id) => state.incidencies.byId[id])
    .filter((entry) => entry.alumneId === alumneId)
    .filter((entry) => !assignaturaFilter || assignaturaFilter.has(entry.assignaturaId))
    .filter((entry) => withinOptionalRange(entry.dataHora));

  const attendanceThresholds = state.configGlobal.attendanceThresholds;
  const resumAssistencia = {
    total: assistencies.length,
    presents: assistencies.filter((a) => a.present).length,
    absents: assistencies.filter((a) => a.present === false).length,
    retardsMenors: assistencies.filter(
      (a) => (a.retardMin || 0) > 0 && (a.retardMin || 0) < attendanceThresholds.tardancaMenor,
    ).length,
    retardsMajors: assistencies.filter((a) => (a.retardMin || 0) >= attendanceThresholds.tardancaMenor).length,
  };

  const resumIncidencies = incidencies.reduce((acc, incidencia) => {
    const categoria = incidencia.categoria || 'General';
    acc[categoria] = (acc[categoria] || 0) + 1;
    return acc;
  }, {});

  const perAssignatura = {};
  const matricules = state.matriculacions.allIds
    .map((id) => state.matriculacions.byId[id])
    .filter((m) => m.alumneId === alumneId)
    .filter((m) => !assignaturaFilter || assignaturaFilter.has(m.assignaturaId));

  const mitjanes = {};

  matricules.forEach((matricula) => {
    const assignatura = getAssignatura(state, matricula.assignaturaId);
    const info = {
      mode: assignatura.mode,
      rounding: assignatura.rounding,
      qualitativeRanges: assignatura.qualitativeRanges,
    };
    if (assignatura.mode === EVAL_MODE.NUMERIC) {
      const calendaris = state.calendaris.allIds
        .map((id) => state.calendaris.byId[id])
        .filter((cal) => cal.assignaturaId === assignatura.id);
      const trimestres = new Set();
      calendaris.forEach((cal) => (cal.trimestres || []).forEach((t) => trimestres.add(t.id)));
      const notesTrimestre = {};
      if (trimestres.size === 0) {
        notesTrimestre.global = computeNotaPerAlumne(state, assignatura.id, alumneId, opts.trimestreId);
      } else {
        trimestres.forEach((trimestreId) => {
          if (!opts.trimestreId || opts.trimestreId === trimestreId) {
            notesTrimestre[trimestreId] = computeNotaPerAlumne(state, assignatura.id, alumneId, trimestreId);
          }
        });
      }
      info.notesTrimestre = notesTrimestre;
      perAssignatura[assignatura.id] = info;
      const notaGlobal =
        notesTrimestre[opts.trimestreId] ||
        notesTrimestre.global ||
        computeNotaPerAlumne(state, assignatura.id, alumneId, opts.trimestreId);
      mitjanes[assignatura.id] = notaGlobal.valueNumRounded;
    } else {
      const taula = computeTaulaCompetencial(state, assignatura.id, opts.trimestreId);
      info.taulaCA = taula.values[alumneId] || {};
      perAssignatura[assignatura.id] = info;
      const values = Object.values(info.taulaCA || {});
      if (values.length > 0) {
        const avg = values.reduce((sum, value) => sum + (value?.valueNumRounded ?? 0), 0) / values.length;
        mitjanes[assignatura.id] = roundTo(
          avg,
          assignatura.rounding?.decimals ?? DEFAULTS.rounding.decimals,
          assignatura.rounding?.mode ?? DEFAULTS.rounding.mode,
        );
      }
    }
  });

  return {
    alumneId,
    resum: {
      assistencia: resumAssistencia,
      incidencies: resumIncidencies,
      mitjanes,
    },
    perAssignatura,
  };
}

export function exportStateForSave(state) {
  return cloneState(state);
}
let idCounter = 1;
function createId(prefix) {
  const id = `${prefix || 'id'}_${idCounter}`;
  idCounter += 1;
  return id;
}

function ensureCollection(state, key) {
  if (!state[key]) {
    state[key] = { byId: {}, allIds: [] };
  }
}

function addToCollection(state, key, entity) {
  ensureCollection(state, key);
  state[key].byId[entity.id] = entity;
  if (!state[key].allIds.includes(entity.id)) {
    state[key].allIds.push(entity.id);
  }
}

function updateInCollection(state, key, id, patch) {
  ensureCollection(state, key);
  assert(state[key].byId[id], `${key} ${id} no trobat`);
  state[key].byId[id] = { ...state[key].byId[id], ...patch };
}

function removeFromCollection(state, key, id) {
  ensureCollection(state, key);
  if (!state[key].byId[id]) return;
  delete state[key].byId[id];
  state[key].allIds = state[key].allIds.filter((entryId) => entryId !== id);
}

export function createStore(initialState = createEmptyState()) {
  const state = cloneState(initialState);
  state.version = state.version || 0;
  state.historial = Array.isArray(state.historial) ? state.historial : [];
  const listeners = new Set();
  let batching = false;
  let batchBefore = null;
  let batchMeta = null;
  const undoStack = [];
  const redoStack = [];
  const MAX_STACK = 20;
  let suppressHistory = false;

  function notify(change) {
    const snapshot = getState();
    listeners.forEach((listener) => {
      try {
        listener(snapshot, change);
      } catch (error) {
        console.error('Listener error', error);
      }
    });
  }

  function pushUndoSnapshot(snapshot) {
    undoStack.push(cloneState(snapshot));
    while (undoStack.length > MAX_STACK) {
      undoStack.shift();
    }
  }

  function pushRedoSnapshot(snapshot) {
    redoStack.push(cloneState(snapshot));
    while (redoStack.length > MAX_STACK) {
      redoStack.shift();
    }
  }

  function recordChange(meta, before) {
    const after = cloneState(state);
    if (before && isEqual(before, after)) {
      return;
    }
    if (!suppressHistory && before) {
      pushUndoSnapshot(before);
      redoStack.length = 0;
    }
    state.version += 1;
    const entry = {
      ts: Date.now(),
      user: 'me',
      action: meta?.action || 'patch',
    };
    if (before) entry.before = before;
    entry.after = after;
    state.historial.push(entry);
    while (state.historial.length > 200) {
      state.historial.shift();
    }
    notify(entry);
  }

  function applyMutation(meta, mutator) {
    if (batching) {
      mutator();
      return;
    }
    const before = cloneState(state);
    mutator();
    recordChange(meta, before);
  }

  function getState() {
    return cloneState(state);
  }

  function subscribe(listener) {
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  function transact(fn, meta) {
    assert(!batching, 'No es permeten transaccions imbricades');
    batching = true;
    batchBefore = cloneState(state);
    batchMeta = meta;
    try {
      fn();
    } finally {
      batching = false;
      recordChange(batchMeta, batchBefore);
      batchBefore = null;
      batchMeta = null;
    }
  }

  function replaceState(nextState) {
    const snapshot = cloneState(nextState);
    Object.keys(state).forEach((key) => {
      if (!(key in snapshot)) {
        delete state[key];
      }
    });
    Object.entries(snapshot).forEach(([key, value]) => {
      state[key] = value;
    });
  }

  function undo() {
    if (!undoStack.length) return false;
    const target = undoStack.pop();
    const current = cloneState(state);
    pushRedoSnapshot(current);
    suppressHistory = true;
    try {
      applyMutation({ action: 'undo' }, () => {
        replaceState(target);
      });
    } finally {
      suppressHistory = false;
    }
    return true;
  }

  function redo() {
    if (!redoStack.length) return false;
    const target = redoStack.pop();
    const current = cloneState(state);
    pushUndoSnapshot(current);
    suppressHistory = true;
    try {
      applyMutation({ action: 'redo' }, () => {
        replaceState(target);
      });
    } finally {
      suppressHistory = false;
    }
    return true;
  }

  function canUndo() {
    return undoStack.length > 0;
  }

  function canRedo() {
    return redoStack.length > 0;
  }

  function patch(partial, meta) {
    applyMutation(meta, () => {
      Object.assign(state, partial);
    });
  }

  function addAssignatura(data) {
    const id = data.id || createId('assignatura');
    applyMutation({ action: 'addAssignatura' }, () => {
      const rounding = {
        decimals: data.rounding?.decimals ?? DEFAULTS.rounding.decimals,
        mode: data.rounding?.mode ?? DEFAULTS.rounding.mode,
      };
      assertDecimals(rounding.decimals);
      const qualitativeRanges = validateQualitativeRanges(
        data.qualitativeRanges || DEFAULTS.qualitativeRanges,
      );
      const categoriaPesos = data.categoriaPesos ? cloneState(data.categoriaPesos) : {};
      const assignatura = {
        id,
        nom: data.nom,
        anyCurs: data.anyCurs,
        mode: data.mode || EVAL_MODE.NUMERIC,
        rounding,
        qualitativeRanges,
        categoriaPesos,
      };
      addToCollection(state, 'assignatures', assignatura);
    });
    return id;
  }

  function updateAssignatura(id, patchData) {
    applyMutation({ action: 'updateAssignatura' }, () => {
      const current = getAssignatura(state, id);
      const patchObj = { ...patchData };
      if (patchData.rounding) {
        assertDecimals(patchData.rounding.decimals ?? current.rounding.decimals);
        patchObj.rounding = {
          decimals: patchData.rounding.decimals ?? current.rounding.decimals,
          mode: patchData.rounding.mode || current.rounding.mode,
        };
      }
      if (patchData.qualitativeRanges) {
        patchObj.qualitativeRanges = validateQualitativeRanges(patchData.qualitativeRanges);
      }
      updateInCollection(state, 'assignatures', id, patchObj);
    });
  }

  function addAlumne(data) {
    const id = data.id || createId('alumne');
    applyMutation({ action: 'addAlumne' }, () => {
      addToCollection(state, 'alumnes', {
        id,
        ...data,
      });
    });
    return id;
  }

  function updateAlumne(id, patchData) {
    applyMutation({ action: 'updateAlumne' }, () => {
      updateInCollection(state, 'alumnes', id, patchData);
    });
  }

  function matricula(alumneId, assignaturaId) {
    const id = createId('matricula');
    applyMutation({ action: 'matricula' }, () => {
      getAssignatura(state, assignaturaId);
      assert(state.alumnes.byId[alumneId], `Alumne ${alumneId} no trobat`);
      addToCollection(state, 'matriculacions', {
        id,
        alumneId,
        assignaturaId,
      });
    });
    return id;
  }

  function addCE(assignaturaId, data) {
    const assignatura = getAssignatura(state, assignaturaId);
    const id = data.id || createId('ce');
    applyMutation({ action: 'addCE' }, () => {
      const index = data.index ?? assertCEIdRules(id);
      const cesAssignatura = state.ces.allIds
        .map((ceId) => state.ces.byId[ceId])
        .filter((ce) => ce.assignaturaId === assignatura.id);
      const nextPosition = cesAssignatura.length + 1;
      const ce = {
        id,
        assignaturaId: assignatura.id,
        index,
        textBetween: data.textBetween || '',
        position: data.position || nextPosition,
      };
      addToCollection(state, 'ces', ce);
    });
    return id;
  }

  function moveCE(assignaturaId, ceId, newPosition, options = {}) {
    applyMutation({ action: 'reorderCE' }, () => {
      reorderCE(state, assignaturaId, ceId, newPosition, options);
    });
  }

  function addCA(ceId, data) {
    const ce = state.ces.byId[ceId];
    assert(ce, `CE ${ceId} no trobat`);
    const id = data.id || createId('ca');
    applyMutation({ action: 'addCA' }, () => {
      const indices = data.index ? { ceIndex: ce.index, index: data.index } : assertCAIdRules(id);
      const casCE = collectCAsForCE(state, ceId);
      const nextPosition = casCE.length + 1;
      let pesDinsCEValue = data.pesDinsCE;
      if (pesDinsCEValue !== undefined) {
        const sanitized = validateWeights({ pesDinsCE: pesDinsCEValue });
        pesDinsCEValue = sanitized.pesDinsCE;
      }
      const ca = {
        id,
        ceId,
        ceIndex: indices.ceIndex || ce.index,
        index: indices.index,
        textBetween: data.textBetween || '',
        position: data.position || nextPosition,
        pesDinsCE: pesDinsCEValue,
      };
      addToCollection(state, 'cas', ca);
    });
    return id;
  }

  function moveCA(ceId, caId, newPosition, options = {}) {
    applyMutation({ action: 'reorderCA' }, () => {
      reorderCA(state, ceId, caId, newPosition, options);
    });
  }

  function addCategoria(nom) {
    const id = createId('categoria');
    applyMutation({ action: 'addCategoria' }, () => {
      addToCollection(state, 'categories', { id, nom });
    });
    return id;
  }

  function setPesCategoria(assignaturaId, trimestreId, categoriaId, pes) {
    applyMutation({ action: 'setPesCategoria' }, () => {
      const assignatura = getAssignatura(state, assignaturaId);
      const sanitized = validateWeights({ pesCategoria: pes });
      const value = sanitized.pesCategoria ?? 0;
      if (!assignatura.categoriaPesos[trimestreId]) {
        assignatura.categoriaPesos[trimestreId] = {};
      }
      assignatura.categoriaPesos[trimestreId][categoriaId] = value;
    });
  }

  function bulkSetPesosDinsCE(ceId, mapCaIdToPes) {
    applyMutation({ action: 'bulkSetPesosDinsCE' }, () => {
      const ce = state.ces.byId[ceId];
      assert(ce, `CE ${ceId} no trobat`);
      const entries =
        mapCaIdToPes instanceof Map
          ? Array.from(mapCaIdToPes.entries())
          : Object.entries(mapCaIdToPes || {});
      entries.forEach(([caId, pes]) => {
        const ca = state.cas.byId[caId];
        assert(ca && ca.ceId === ceId, `CA ${caId} no pertany al CE ${ceId}`);
        const sanitized = validateWeights({ pesDinsCE: pes });
        state.cas.byId[caId].pesDinsCE = sanitized.pesDinsCE ?? 0;
      });
    });
  }

  function addActivitat(assignaturaId, data) {
    const id = data.id || createId('activitat');
    applyMutation({ action: 'addActivitat' }, () => {
      const assignatura = getAssignatura(state, assignaturaId);
      const sanitized = validateWeights({ pesActivitat: data.pesActivitat ?? 0 });
      const pesActivitat = sanitized.pesActivitat ?? 0;
      const activitat = {
        id,
        assignaturaId: assignatura.id,
        data: data.data ? new Date(data.data) : new Date(),
        categoriaId: data.categoriaId,
        pesActivitat,
        descripcio: data.descripcio,
      };
      addToCollection(state, 'activitats', activitat);
    });
    return id;
  }

  function getCalendari(assignaturaId, { createIfMissing = false } = {}) {
    let calendari = state.calendaris.allIds
      .map((id) => state.calendaris.byId[id])
      .find((cal) => cal.assignaturaId === assignaturaId);
    if (!calendari && createIfMissing) {
      const id = createId('calendari');
      calendari = {
        id,
        assignaturaId,
        cursInici: null,
        cursFi: null,
        diesSetmanals: [],
        trimestres: [],
        festius: [],
        excepcions: [],
        horariVersions: [],
        horariActivaId: null,
      };
      addToCollection(state, 'calendaris', calendari);
    }
    if (calendari) {
      calendari.trimestres = ensureArray(calendari.trimestres);
      calendari.festius = ensureArray(calendari.festius).map((festiu) => normalizeFestiu(festiu)).filter(Boolean);
      calendari.excepcions = ensureArray(calendari.excepcions)
        .map((excepcio) => normalizeExcepcio(excepcio))
        .filter(Boolean);
      calendari.horariVersions = ensureArray(calendari.horariVersions).map((versio) => ({
        id: versio.id || createId('horariVersio'),
        effectiveFrom: toISODate(versio.effectiveFrom),
        diesSetmanals: ensureArray(versio.diesSetmanals).map((d) => Number(d)).filter((d) => Number.isInteger(d)),
      }));
    }
    return calendari || null;
  }

  function updateCalendari(assignaturaId, updater, meta = {}) {
    applyMutation(meta, () => {
      const calendari = getCalendari(assignaturaId, { createIfMissing: true });
      updater(calendari);
      state.calendaris.byId[calendari.id] = { ...calendari };
    });
  }

  function setCursRange(assignaturaId, { inici, fi }) {
    updateCalendari(
      assignaturaId,
      (calendari) => {
        const iniciISO = inici ? toISODate(inici) : null;
        const fiISO = fi ? toISODate(fi) : null;
        if (iniciISO && fiISO && compareISODate(iniciISO, fiISO) === 1) {
          throw new Error('La data de fi ha de ser posterior o igual a la data d\'inici');
        }
        calendari.cursInici = iniciISO;
        calendari.cursFi = fiISO;
      },
      { action: 'setCursRange' },
    );
  }

  function setDiesSetmanals(assignaturaId, dies) {
    updateCalendari(
      assignaturaId,
      (calendari) => {
        const normalized = ensureArray(dies)
          .map((d) => Number(d))
          .filter((d) => Number.isInteger(d) && d >= 0 && d <= 6);
        const unique = Array.from(new Set(normalized)).sort((a, b) => a - b);
        calendari.diesSetmanals = unique;
      },
      { action: 'setDiesSetmanals' },
    );
  }

  function addTrimestre(assignaturaId, data) {
    const id = data.id || createId('trimestre');
    updateCalendari(
      assignaturaId,
      (calendari) => {
        const tInici = toISODate(data.tInici);
        const tFi = toISODate(data.tFi);
        if (!tInici || !tFi) {
          throw new Error('Les dates del trimestre han de ser vàlides');
        }
        if (compareISODate(tInici, tFi) === 1) {
          throw new Error('La data de fi del trimestre ha de ser posterior o igual a la data d\'inici');
        }
        const existingIndex = calendari.trimestres.findIndex((t) => t.id === id);
        const trimestre = { id, tInici, tFi, nom: data.nom || data.name || data.label || null };
        if (existingIndex >= 0) {
          calendari.trimestres[existingIndex] = trimestre;
        } else {
          calendari.trimestres.push(trimestre);
        }
        calendari.trimestres.sort((a, b) => compareISODate(a.tInici, b.tInici));
      },
      { action: 'addTrimestre' },
    );
    return id;
  }

  function removeTrimestre(assignaturaId, trimestreId) {
    updateCalendari(
      assignaturaId,
      (calendari) => {
        calendari.trimestres = calendari.trimestres.filter((t) => t.id !== trimestreId);
      },
      { action: 'removeTrimestre' },
    );
  }

  function addFestius(assignaturaId, dates) {
    updateCalendari(
      assignaturaId,
      (calendari) => {
        const festius = ensureArray(dates)
          .map((entry) => normalizeFestiu(entry))
          .filter(Boolean);
        const byDate = new Map(calendari.festius.map((festiu) => [festiu.dataISO, festiu]));
        festius.forEach((festiu) => {
          byDate.set(festiu.dataISO, { ...byDate.get(festiu.dataISO), ...festiu });
        });
        calendari.festius = Array.from(byDate.values()).sort((a, b) => compareISODate(a.dataISO, b.dataISO));
      },
      { action: 'addFestius' },
    );
  }

  function removeFestius(assignaturaId, dates) {
    const datesSet = new Set(
      ensureArray(dates)
        .map((entry) => (typeof entry === 'string' ? toISODate(entry) : toISODate(entry?.dataISO)))
        .filter(Boolean),
    );
    updateCalendari(
      assignaturaId,
      (calendari) => {
        if (!datesSet.size) return;
        calendari.festius = calendari.festius.filter((festiu) => !datesSet.has(festiu.dataISO));
      },
      { action: 'removeFestius' },
    );
  }

  function addExcepcions(assignaturaId, items) {
    updateCalendari(
      assignaturaId,
      (calendari) => {
        const map = new Map(calendari.excepcions.map((ex) => [ex.dataISO, ex]));
        ensureArray(items)
          .map((item) => normalizeExcepcio(item))
          .filter(Boolean)
          .forEach((item) => {
            map.set(item.dataISO, { ...map.get(item.dataISO), ...item });
          });
        calendari.excepcions = Array.from(map.values()).sort((a, b) => compareISODate(a.dataISO, b.dataISO));
      },
      { action: 'addExcepcions' },
    );
  }

  function removeExcepcio(assignaturaId, dataISO) {
    const normalized = toISODate(dataISO);
    if (!normalized) return;
    updateCalendari(
      assignaturaId,
      (calendari) => {
        calendari.excepcions = calendari.excepcions.filter((item) => item.dataISO !== normalized);
      },
      { action: 'removeExcepcio' },
    );
  }

  function addHorariVersio(assignaturaId, data) {
    const id = data.id || createId('horariVersio');
    updateCalendari(
      assignaturaId,
      (calendari) => {
        const effectiveFrom = toISODate(data.effectiveFrom);
        if (!effectiveFrom) {
          throw new Error('La versió d\'horari necessita una data d\'entrada en vigor vàlida');
        }
        const dies = ensureArray(data.diesSetmanals)
          .map((d) => Number(d))
          .filter((d) => Number.isInteger(d) && d >= 0 && d <= 6);
        if (!dies.length) {
          throw new Error('La versió d\'horari ha de contenir almenys un dia lectiu');
        }
        const versio = { id, effectiveFrom, diesSetmanals: Array.from(new Set(dies)).sort((a, b) => a - b) };
        const index = calendari.horariVersions.findIndex((v) => v.id === id);
        if (index >= 0) {
          calendari.horariVersions[index] = versio;
        } else {
          calendari.horariVersions.push(versio);
        }
        calendari.horariVersions.sort((a, b) => compareISODate(a.effectiveFrom, b.effectiveFrom));
      },
      { action: 'addHorariVersio' },
    );
    return id;
  }

  function activateHorariVersio(assignaturaId, id) {
    updateCalendari(
      assignaturaId,
      (calendari) => {
        if (!calendari.horariVersions.some((versio) => versio.id === id)) {
          throw new Error(`Versió d'horari ${id} no trobada`);
        }
        calendari.horariActivaId = id;
      },
      { action: 'activateHorariVersio' },
    );
  }

  function getActiveSchedule(assignaturaId, date) {
    const calendari = getCalendari(assignaturaId);
    if (!calendari) return null;
    const isoDate = toISODate(date);
    if (!isoDate) return null;
    const candidates = calendari.horariVersions.filter((versio) => !versio.effectiveFrom || versio.effectiveFrom <= isoDate);
    let selected = null;
    const preferredId = calendari.horariActivaId;
    if (preferredId) {
      selected = candidates.find((versio) => versio.id === preferredId) || null;
    }
    if (!selected) {
      for (const versio of candidates) {
        if (!selected || compareISODate(versio.effectiveFrom, selected.effectiveFrom) === 1) {
          selected = versio;
        }
      }
    }
    if (selected) {
      return {
        diesSetmanals: Array.from(new Set(selected.diesSetmanals)).sort((a, b) => a - b),
        versioId: selected.id || null,
      };
    }
    if (calendari.diesSetmanals?.length) {
      return {
        diesSetmanals: Array.from(new Set(calendari.diesSetmanals)).sort((a, b) => a - b),
        versioId: null,
      };
    }
    return null;
  }

  function evaluateLectiu(assignaturaId, date) {
    const isoDate = toISODate(date);
    if (!isoDate) {
      return { lectiu: false, motiu: 'Data invàlida' };
    }
    const calendari = getCalendari(assignaturaId);
    if (!calendari) {
      return { lectiu: false, motiu: 'Sense calendari configurat' };
    }
    if (calendari.cursInici && compareISODate(isoDate, calendari.cursInici) === -1) {
      return { lectiu: false, motiu: 'Fora del període lectiu (abans d\'inici)' };
    }
    if (calendari.cursFi && compareISODate(isoDate, calendari.cursFi) === 1) {
      return { lectiu: false, motiu: 'Fora del període lectiu (després de fi)' };
    }
    const festiu = calendari.festius.find((item) => item.dataISO === isoDate);
    if (festiu) {
      return { lectiu: false, motiu: festiu.motiu || 'Festiu' };
    }
    const excepcio = calendari.excepcions.find((item) => item.dataISO === isoDate);
    if (excepcio) {
      return { lectiu: false, motiu: excepcio.motiu || 'Excepció' };
    }
    const schedule = getActiveSchedule(assignaturaId, isoDate);
    if (!schedule || !schedule.diesSetmanals.length) {
      return { lectiu: false, motiu: 'Sense horari actiu' };
    }
    const weekday = getUTCWeekday(isoDate);
    if (weekday === null) {
      return { lectiu: false, motiu: 'Data invàlida' };
    }
    const versioId = typeof schedule.versioId === 'string' && schedule.versioId.trim().length
      ? schedule.versioId
      : null;
    if (!schedule.diesSetmanals.includes(weekday)) {
      const result = { lectiu: false, motiu: 'No hi ha classe aquest dia' };
      if (versioId) {
        result.versioId = versioId;
      }
      return result;
    }
    return versioId ? { lectiu: true, versioId } : { lectiu: true };
  }

  function isLectiu(assignaturaId, date) {
    return evaluateLectiu(assignaturaId, date).lectiu;
  }

  function listSessions(assignaturaId, { from, to }) {
    const fromISO = toISODate(from);
    const toISO = toISODate(to);
    if (!fromISO || !toISO) return [];
    if (compareISODate(fromISO, toISO) === 1) return [];
    const sessions = [];
    const cursor = toUTCDateAtStartOfDay(fromISO);
    const end = toUTCDateAtStartOfDay(toISO);
    if (!cursor || !end) return [];
    while (cursor.getTime() <= end.getTime()) {
      const iso = cursor.toISOString().slice(0, 10);
      if (isLectiu(assignaturaId, cursor)) {
        sessions.push({ dateISO: iso, weekday: cursor.getUTCDay() });
      }
      cursor.setUTCDate(cursor.getUTCDate() + 1);
    }
    return sessions;
  }

  function listSessionsTrimestre(assignaturaId, trimestreId) {
    const calendari = getCalendari(assignaturaId);
    if (!calendari) return [];
    const trimestre = ensureArray(calendari.trimestres).find((t) => t.id === trimestreId);
    if (!trimestre) return [];
    return listSessions(assignaturaId, { from: trimestre.tInici, to: trimestre.tFi });
  }

  function simulateDay(assignaturaId, date) {
    const result = evaluateLectiu(assignaturaId, date);
    return result;
  }

  function linkCAtoActivitat(activitatId, caId, pes_ca) {
    const id = createId('activitatCA');
    applyMutation({ action: 'linkCAtoActivitat' }, () => {
      assert(state.activitats.byId[activitatId], `Activitat ${activitatId} no trobada`);
      assert(state.cas.byId[caId], `CA ${caId} no trobat`);
      const sanitized = validateWeights({ pes_ca });
      addToCollection(state, 'activitatCA', {
        id,
        activitatId,
        caId,
        pes_ca: sanitized.pes_ca ?? 0,
      });
    });
    return id;
  }

  function bulkRelinkActivitatCA(activitatId, entries, { strict = false } = {}) {
    applyMutation({ action: 'bulkRelinkActivitatCA' }, () => {
      const activitat = state.activitats.byId[activitatId];
      assert(activitat, `Activitat ${activitatId} no trobada`);
      const normalized = new Map();
      const provided = Array.isArray(entries) ? entries : [];
      provided.forEach((entry, idx) => {
        assert(entry && typeof entry === 'object', `Entrada ${idx} de relació CA/activitat invàlida`);
        const caId = entry.caId;
        assert(caId, 'Cada entrada ha d\'incloure caId');
        const ca = state.cas.byId[caId];
        assert(ca, `CA ${caId} no trobat`);
        const sanitized = validateWeights({ pes_ca: entry.pes_ca });
        normalized.set(caId, sanitized.pes_ca ?? 0);
      });
      const existing = state.activitatCA.allIds
        .map((id) => state.activitatCA.byId[id])
        .filter((link) => link.activitatId === activitatId);
      const keepIds = new Set();
      normalized.forEach((pesValue, caId) => {
        const current = existing.find((link) => link.caId === caId);
        if (current) {
          state.activitatCA.byId[current.id].pes_ca = pesValue;
          keepIds.add(current.id);
        } else {
          const id = createId('activitatCA');
          addToCollection(state, 'activitatCA', {
            id,
            activitatId,
            caId,
            pes_ca: pesValue,
          });
          keepIds.add(id);
        }
      });
      if (strict) {
        existing.forEach((link) => {
          if (!keepIds.has(link.id)) {
            removeFromCollection(state, 'activitatCA', link.id);
          }
        });
      }
    });
  }

  function registraAvaluacioComp({ alumneId, activitatId, caId, valorQuali }) {
    const id = createId('avaluacioComp');
    applyMutation({ action: 'registraAvaluacioComp' }, () => {
      assert(state.alumnes.byId[alumneId], `Alumne ${alumneId} no trobat`);
      assert(state.activitats.byId[activitatId], `Activitat ${activitatId} no trobada`);
      assert(state.cas.byId[caId], `CA ${caId} no trobat`);
      assertQuali(valorQuali);
      addToCollection(state, 'avaluacionsComp', {
        id,
        alumneId,
        activitatId,
        caId,
        valorQuali,
      });
    });
    return id;
  }

  function registraAvaluacioNum({ alumneId, activitatId, valorNum }) {
    const id = createId('avaluacioNum');
    applyMutation({ action: 'registraAvaluacioNum' }, () => {
      assert(state.alumnes.byId[alumneId], `Alumne ${alumneId} no trobat`);
      const activitat = state.activitats.byId[activitatId];
      assert(activitat, `Activitat ${activitatId} no trobada`);
      assert(
        typeof valorNum === 'number' && valorNum >= 0 && valorNum <= 10,
        'La nota numèrica ha d\'estar entre 0 i 10',
      );
      addToCollection(state, 'avaluacionsNum', {
        id,
        alumneId,
        activitatId,
        valorNum,
      });
    });
    return id;
  }

  function registraAssistencia(entry) {
    const id = createId('assistencia');
    applyMutation({ action: 'registraAssistencia' }, () => {
      assert(state.alumnes.byId[entry.alumneId], `Alumne ${entry.alumneId} no trobat`);
      getAssignatura(state, entry.assignaturaId);
      addToCollection(state, 'assistencia', {
        id,
        ...entry,
        dataHora: entry.dataHora ? new Date(entry.dataHora) : new Date(),
      });
    });
    return id;
  }

  function registraIncidencia(entry) {
    const id = createId('incidencia');
    applyMutation({ action: 'registraIncidencia' }, () => {
      assert(state.alumnes.byId[entry.alumneId], `Alumne ${entry.alumneId} no trobat`);
      getAssignatura(state, entry.assignaturaId);
      addToCollection(state, 'incidencies', {
        id,
        ...entry,
        dataHora: entry.dataHora ? new Date(entry.dataHora) : new Date(),
      });
    });
    return id;
  }

  return {
    getState,
    subscribe,
    patch,
    transact,
    undo,
    redo,
    canUndo,
    canRedo,
    addAssignatura,
    updateAssignatura,
    addAlumne,
    updateAlumne,
    matricula,
    addCE,
    addCA,
    addCategoria,
    setPesCategoria,
    bulkSetPesosDinsCE,
    moveCE,
    moveCA,
    addActivitat,
    linkCAtoActivitat,
    bulkRelinkActivitatCA,
    registraAvaluacioComp,
    registraAvaluacioNum,
    registraAssistencia,
    registraIncidencia,
    setCursRange,
    setDiesSetmanals,
    addTrimestre,
    removeTrimestre,
    addFestius,
    removeFestius,
    addExcepcions,
    removeExcepcio,
    addHorariVersio,
    activateHorariVersio,
    getActiveSchedule: (...args) => getActiveSchedule(...args),
    isLectiu: (...args) => isLectiu(...args),
    listSessions: (...args) => listSessions(...args),
    listSessionsTrimestre: (...args) => listSessionsTrimestre(...args),
    simulateDay: (...args) => simulateDay(...args),
    computeCAForAlumne: (...args) => computeCAForAlumne(state, ...args),
    computeCEForAlumne: (...args) => computeCEForAlumne(state, ...args),
    computeTaulaCompetencial: (...args) => computeTaulaCompetencial(state, ...args),
    computeNotaPerAlumne: (...args) => computeNotaPerAlumne(state, ...args),
    getFitxaAlumne: (...args) => getFitxaAlumne(state, ...args),
    exportStateForSave: () => exportStateForSave(state),
  };
}

const DEV_MODE = typeof process !== 'undefined' && process.env?.NODE_ENV === 'development';

function runSampleComputation() {
  const store = createStore(createEmptyState());
  const assignaturaId = store.addAssignatura({
    id: 'assignatura_test',
    nom: 'Prova',
    mode: EVAL_MODE.COMPETENCIAL,
    rounding: { decimals: 1, mode: 'half-up' },
  });
  const alumneId = store.addAlumne({ id: 'alumne_test', nom: 'Alumne Prova' });
  store.matricula(alumneId, assignaturaId);

  const ce1Id = store.addCE(assignaturaId, { id: 'ce_test_1', index: 1, position: 1 });
  const ce2Id = store.addCE(assignaturaId, { id: 'ce_test_2', index: 2, position: 2 });

  const ca11Id = store.addCA(ce1Id, { id: 'ca_test_11', index: 1, pesDinsCE: 2, position: 1 });
  const ca12Id = store.addCA(ce1Id, { id: 'ca_test_12', index: 2, pesDinsCE: 1, position: 2 });
  const ca21Id = store.addCA(ce2Id, { id: 'ca_test_21', index: 1, pesDinsCE: 1, position: 1 });

  const activitat1Id = store.addActivitat(assignaturaId, {
    id: 'act_test_1',
    data: new Date('2024-02-01'),
    categoriaId: null,
    pesActivitat: 1,
    descripcio: 'Activitat 1',
  });
  const activitat2Id = store.addActivitat(assignaturaId, {
    id: 'act_test_2',
    data: new Date('2024-03-15'),
    categoriaId: null,
    pesActivitat: 1,
    descripcio: 'Activitat 2',
  });

  store.linkCAtoActivitat(activitat1Id, ca11Id, 2);
  store.linkCAtoActivitat(activitat1Id, ca12Id, 1);
  store.linkCAtoActivitat(activitat1Id, ca21Id, 1);
  store.linkCAtoActivitat(activitat2Id, ca21Id, 2);

  store.registraAvaluacioComp({
    alumneId,
    activitatId: activitat1Id,
    caId: ca11Id,
    valorQuali: 'AE',
  });
  store.registraAvaluacioComp({
    alumneId,
    activitatId: activitat1Id,
    caId: ca12Id,
    valorQuali: 'AS',
  });
  store.registraAvaluacioComp({
    alumneId,
    activitatId: activitat1Id,
    caId: ca21Id,
    valorQuali: 'AS',
  });
  store.registraAvaluacioComp({
    alumneId,
    activitatId: activitat2Id,
    caId: ca21Id,
    valorQuali: 'AN',
  });

  const ca11 = store.computeCAForAlumne(assignaturaId, alumneId, ca11Id);
  const ca12 = store.computeCAForAlumne(assignaturaId, alumneId, ca12Id);
  const ca21 = store.computeCAForAlumne(assignaturaId, alumneId, ca21Id);
  const ce1 = store.computeCEForAlumne(assignaturaId, alumneId, ce1Id);
  const ce2 = store.computeCEForAlumne(assignaturaId, alumneId, ce2Id);

  assert(Math.abs(ca11.valueNumRounded - 9.5) < 1e-6, 'Resultat CA11 inesperat');
  assert(Math.abs(ca12.valueNumRounded - 6.0) < 1e-6, 'Resultat CA12 inesperat');
  assert(Math.abs(ca21.valueNumRounded - 7.3) < 1e-6, 'Resultat CA21 inesperat');
  assert(Math.abs(ce1.valueNumRounded - 8.3) < 1e-6, 'Resultat CE1 inesperat');
  assert(Math.abs(ce2.valueNumRounded - 7.3) < 1e-6, 'Resultat CE2 inesperat');

  assert(ce1.quali === 'AN', 'Qualitativa CE1 inesperada');
  assert(ce2.quali === 'AN', 'Qualitativa CE2 inesperada');

  return {
    assignaturaId,
    alumneId,
    ces: {
      [ce1Id]: ce1,
      [ce2Id]: ce2,
    },
    cas: {
      [ca11Id]: ca11,
      [ca12Id]: ca12,
      [ca21Id]: ca21,
    },
  };
}

export let __selfTest;

if (typeof window !== 'undefined' && window.__DEV__) {
  __selfTest = function __selfTest() {
    const result = runSampleComputation();
    console.info('[self-test] Resultats de mostra', result);
    return result;
  };
}

export const __test_computeSample = DEV_MODE ? () => runSampleComputation() : undefined;
