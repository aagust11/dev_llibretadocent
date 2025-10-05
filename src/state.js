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

function validateQualitativeRanges(ranges) {
  assert(ranges && typeof ranges === 'object', 'Les franges qualitatives han de ser un objecte');
  const clean = {};
  QUALI.forEach((key) => {
    const range = ranges[key];
    assert(Array.isArray(range) && range.length === 2, `Franja qualitativa ${key} ha de ser [min,max]`);
    const [min, max] = range;
    assert(
      typeof min === 'number' && typeof max === 'number' && min <= max,
      `Franja qualitativa ${key} ha de contenir números i min ≤ max`,
    );
    clean[key] = [min, max];
  });
  return clean;
}

function getAssignatura(state, assignaturaId) {
  const assignatura = state.assignatures?.byId?.[assignaturaId];
  assert(assignatura, `Assignatura ${assignaturaId} no trobada`);
  return assignatura;
}

function getTrimestreRange(state, assignaturaId, trimestreId) {
  if (!trimestreId) return null;
  const calendaris = Object.values(state.calendaris.byId).filter(
    (cal) => cal.assignaturaId === assignaturaId,
  );
  for (const calendari of calendaris) {
    for (const trimestre of calendari.trimestres || []) {
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

export function reorderCE(state, assignaturaId, ceId, newPosition) {
  assert(Number.isInteger(newPosition) && newPosition >= 1, 'Nova posició CE ha de ser ≥ 1');
  const ces = state.ces.allIds
    .map((id) => state.ces.byId[id])
    .filter((ce) => ce.assignaturaId === assignaturaId)
    .sort((a, b) => (a.position || 0) - (b.position || 0));
  const ce = ces.find((item) => item.id === ceId);
  assert(ce, `CE ${ceId} no trobat per reordenar`);
  const filtered = ces.filter((item) => item.id !== ceId);
  const index = Math.min(newPosition - 1, filtered.length);
  filtered.splice(index, 0, ce);
  filtered.forEach((item, idx) => {
    state.ces.byId[item.id].position = idx + 1;
  });
}

export function reorderCA(state, ceId, caId, newPosition) {
  assert(Number.isInteger(newPosition) && newPosition >= 1, 'Nova posició CA ha de ser ≥ 1');
  const cas = collectCAsForCE(state, ceId).sort((a, b) => (a.position || 0) - (b.position || 0));
  const ca = cas.find((item) => item.id === caId);
  assert(ca, `CA ${caId} no trobat per reordenar`);
  const filtered = cas.filter((item) => item.id !== caId);
  const index = Math.min(newPosition - 1, filtered.length);
  filtered.splice(index, 0, ca);
  filtered.forEach((item, idx) => {
    state.cas.byId[item.id].position = idx + 1;
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

export function createStore(initialState = createEmptyState()) {
  const state = cloneState(initialState);
  state.version = state.version || 0;
  const listeners = new Set();
  let batching = false;
  let batchBefore = null;
  let batchMeta = null;

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

  function recordChange(meta, before) {
    const after = cloneState(state);
    if (before && isEqual(before, after)) {
      return;
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

  function addCA(ceId, data) {
    const ce = state.ces.byId[ceId];
    assert(ce, `CE ${ceId} no trobat`);
    const id = data.id || createId('ca');
    applyMutation({ action: 'addCA' }, () => {
      const indices = data.index ? { ceIndex: ce.index, index: data.index } : assertCAIdRules(id);
      const casCE = collectCAsForCE(state, ceId);
      const nextPosition = casCE.length + 1;
      const ca = {
        id,
        ceId,
        ceIndex: indices.ceIndex || ce.index,
        index: indices.index,
        textBetween: data.textBetween || '',
        position: data.position || nextPosition,
        pesDinsCE: data.pesDinsCE,
      };
      addToCollection(state, 'cas', ca);
    });
    return id;
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
      assertNonNegative(pes, 'Pes de categoria');
      if (!assignatura.categoriaPesos[trimestreId]) {
        assignatura.categoriaPesos[trimestreId] = {};
      }
      assignatura.categoriaPesos[trimestreId][categoriaId] = pes;
    });
  }

  function addActivitat(assignaturaId, data) {
    const id = data.id || createId('activitat');
    applyMutation({ action: 'addActivitat' }, () => {
      const assignatura = getAssignatura(state, assignaturaId);
      const activitat = {
        id,
        assignaturaId: assignatura.id,
        data: data.data ? new Date(data.data) : new Date(),
        categoriaId: data.categoriaId,
        pesActivitat: data.pesActivitat ?? 0,
        descripcio: data.descripcio,
      };
      assertNonNegative(activitat.pesActivitat, 'Pes d\'activitat');
      addToCollection(state, 'activitats', activitat);
    });
    return id;
  }

  function linkCAtoActivitat(activitatId, caId, pes_ca) {
    const id = createId('activitatCA');
    applyMutation({ action: 'linkCAtoActivitat' }, () => {
      assert(state.activitats.byId[activitatId], `Activitat ${activitatId} no trobada`);
      assert(state.cas.byId[caId], `CA ${caId} no trobat`);
      assertNonNegative(pes_ca, 'Pes CA');
      addToCollection(state, 'activitatCA', {
        id,
        activitatId,
        caId,
        pes_ca,
      });
    });
    return id;
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
    addAssignatura,
    updateAssignatura,
    addAlumne,
    updateAlumne,
    matricula,
    addCE,
    addCA,
    addCategoria,
    setPesCategoria,
    addActivitat,
    linkCAtoActivitat,
    registraAvaluacioComp,
    registraAvaluacioNum,
    registraAssistencia,
    registraIncidencia,
    computeCAForAlumne: (...args) => computeCAForAlumne(state, ...args),
    computeCEForAlumne: (...args) => computeCEForAlumne(state, ...args),
    computeTaulaCompetencial: (...args) => computeTaulaCompetencial(state, ...args),
    computeNotaPerAlumne: (...args) => computeNotaPerAlumne(state, ...args),
    getFitxaAlumne: (...args) => getFitxaAlumne(state, ...args),
    exportStateForSave: () => exportStateForSave(state),
  };
}
