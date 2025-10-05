import { storageEvents } from './storage.js';

const AUTOSAVE_MS = 1200;
const CustomEventCtor = typeof CustomEvent === 'function'
  ? CustomEvent
  : class CustomEventShim extends Event {
      constructor(type, params = {}) {
        super(type, params);
        this.detail = params.detail || null;
      }
    };

const EventTargetCtor = typeof EventTarget === 'function'
  ? EventTarget
  : class SimpleEventTarget {
      constructor() {
        this.listeners = new Map();
      }

      addEventListener(type, listener) {
        if (!this.listeners.has(type)) {
          this.listeners.set(type, new Set());
        }
        this.listeners.get(type).add(listener);
      }

      removeEventListener(type, listener) {
        this.listeners.get(type)?.delete(listener);
      }

      dispatchEvent(event) {
        const set = this.listeners.get(event.type);
        if (!set) return true;
        for (const listener of Array.from(set)) {
          try {
            listener.call(this, event);
          } catch (error) {
            console.error('events listener error', error);
          }
        }
        return !event.defaultPrevented;
      }
    };

export const events = new EventTargetCtor();

function createCSV(rows, { sep = ';' } = {}) {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          if (cell === null || cell === undefined) return '';
          const value = String(cell);
          if (value.includes('"')) {
            return `"${value.replace(/"/gu, '""')}"`;
          }
          if (value.includes(sep) || /[\n\r]/u.test(value)) {
            return `"${value}"`;
          }
          return value;
        })
        .join(sep),
    )
    .join('\n');
}

function fallbackDebounce(fn, ms = AUTOSAVE_MS) {
  let timeoutId = null;
  return (...args) => {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
    timeoutId = setTimeout(() => {
      timeoutId = null;
      fn(...args);
    }, ms);
  };
}

function formatDecimal(value, { decimals = 1, numberToComma } = {}) {
  if (Number.isNaN(value) || value === null || value === undefined) return '';
  if (typeof numberToComma === 'function') {
    return numberToComma(value, decimals);
  }
  const fixed = Number(value).toFixed(decimals);
  return fixed.replace('.', ',');
}

function createEvent(type, detail) {
  return new CustomEventCtor(type, { detail });
}

function extractAlumnesPerAssignatura(state, assignaturaId) {
  const matriculacions = state.matriculacions?.allIds || [];
  const alumnes = new Map();
  matriculacions
    .map((id) => state.matriculacions.byId[id])
    .filter((mat) => mat.assignaturaId === assignaturaId)
    .forEach((mat) => {
      const alumne = state.alumnes?.byId?.[mat.alumneId];
      if (alumne) {
        alumnes.set(alumne.id, alumne);
      }
    });
  return Array.from(alumnes.values());
}

function getTrimestreLabel(state, assignaturaId, trimestreId) {
  if (!trimestreId) return '';
  const calendaris = state.calendaris?.allIds || [];
  for (const calId of calendaris) {
    const cal = state.calendaris.byId[calId];
    if (!cal || cal.assignaturaId !== assignaturaId) continue;
    const trimestre = (cal.trimestres || []).find((t) => t.id === trimestreId);
    if (trimestre) {
      return trimestre.nom || trimestre.id;
    }
  }
  return trimestreId;
}

function getDocx() {
  const lib = globalThis.docx;
  if (!lib) {
    throw new Error('Llibreria DOCX no disponible');
  }
  return lib;
}

export function createActions({ store, storage, i18n, utils }) {
  if (!store || !storage) {
    throw new Error('Cal inicialitzar les accions amb store i storage');
  }

  const debounce = utils?.debounce || fallbackDebounce;
  const toCSV = utils?.toCSV || createCSV;
  const numberToComma = utils?.numberToComma;

  let saving = false;
  let lastSave = { version: 0, when: null, source: 'idb' };
  let autosaveEnabled = true;
  let pendingPatch = null;

  function isFSConnected() {
    if (typeof storage.isFSConnected === 'function') {
      return !!storage.isFSConnected();
    }
    if (storage.fs && typeof storage.fs.connected === 'boolean') {
      return storage.fs.connected;
    }
    if (typeof storage.connected === 'boolean') {
      return storage.connected;
    }
    return false;
  }

  function getLastSaveInfo() {
    return { ...lastSave };
  }

  const debouncedSave = debounce(async () => {
    await saveState({ reason: 'autosave' });
  }, AUTOSAVE_MS);

  async function applyRemoteState(remote, metaAction = 'fs:refresh') {
    if (!remote || !remote.state) return;
    const previousAutosave = autosaveEnabled;
    autosaveEnabled = false;
    try {
      if (typeof store.patch === 'function') {
        store.patch(remote.state, { action: metaAction });
      }
      pendingPatch = null;
      lastSave = {
        version: remote.version ?? lastSave.version,
        when: new Date().toISOString(),
        source: 'fs',
      };
      events.dispatchEvent(createEvent('fs:refresh', { lastSave }));
    } finally {
      autosaveEnabled = previousAutosave;
    }
  }

  async function handleConflict(remoteVersion) {
    const localVersion = lastSave.version ?? 0;
    events.dispatchEvent(
      createEvent('conflict:detected', {
        local: localVersion,
        remote: remoteVersion,
      }),
    );
    try {
      if (typeof storage.loadFromFile === 'function') {
        const remote = await storage.loadFromFile();
        await applyRemoteState(remote, 'conflict:resolved');
      } else if (storage.fs?.loadFromFile) {
        const remote = await storage.fs.loadFromFile();
        await applyRemoteState(remote, 'conflict:resolved');
      }
      events.dispatchEvent(createEvent('conflict:resolved', {}));
    } catch (error) {
      events.dispatchEvent(
        createEvent('save:error', {
          message: 'No s\'ha pogut resoldre el conflicte de guardat',
          error,
        }),
      );
    }
  }

  async function saveState({ reason = 'manual' } = {}) {
    if (saving) return;
    saving = true;
    try {
      const stateForSave = typeof store.exportStateForSave === 'function'
        ? store.exportStateForSave()
        : store.getState?.();
      const payload = stateForSave;
      pendingPatch = null;
      const result = await storage.save(payload);
      const version = result?.version ?? (lastSave.version || 0) + 1;
      if (result?.remoteVersion && result.remoteVersion > version) {
        await handleConflict(result.remoteVersion);
        return;
      }
      lastSave = {
        version,
        when: new Date().toISOString(),
        source: isFSConnected() ? 'fs' : 'idb',
      };
      events.dispatchEvent(createEvent('save:ok', lastSave));
      if (result?.code && result.code !== 'OK') {
        events.dispatchEvent(createEvent('save:warning', { code: result.code, detail: result }));
      }
    } catch (error) {
      const code = error?.code || error?.message;
      if (code === 'LOCKED' || code === 'FS_NOT_CONNECTED') {
        events.dispatchEvent(createEvent('save:warning', { error, code }));
      } else if (code === 'REMOTE_NEWER') {
        await handleConflict(error.remoteVersion ?? 0);
      } else {
        events.dispatchEvent(
          createEvent('save:error', {
            message: 'S\'ha produït un error en desar les dades',
            error,
          }),
        );
      }
    } finally {
      saving = false;
    }
  }

  function mergePending(partial) {
    if (!partial) return;
    pendingPatch = { ...(pendingPatch || {}), ...partial };
  }

  const unsubscribe = store.subscribe?.((state, change) => {
    const partial = change?.partial ?? change?.after ?? null;
    mergePending(partial);
    if (autosaveEnabled) {
      debouncedSave();
    }
  });

  function republishStorageEvent(event) {
    events.dispatchEvent(createEvent(event.type, event.detail));
  }

  const forwardedStorageEvents = [
    'fs:connected',
    'fs:disconnected',
    'fs:error',
    'fs:recovered',
    'lock:acquired',
    'lock:released',
    'lock:blocked',
    'crypto:password-needed',
    'crypto:password-wrong',
    'crypto:changed',
  ];
  forwardedStorageEvents.forEach((type) => {
    storageEvents.addEventListener(type, republishStorageEvent);
  });

  async function init() {
    const previousAutosave = autosaveEnabled;
    autosaveEnabled = false;
    try {
      let loaded = null;
      if (typeof storage.resilientLoad === 'function') {
        loaded = await storage.resilientLoad();
      }
      if (!loaded) {
        loaded = await storage.load();
      }
      if (loaded?.state) {
        if (typeof store.patch === 'function') {
          store.patch(loaded.state, { action: 'hydrate' });
        }
        pendingPatch = null;
        lastSave = {
          version: loaded.version ?? 0,
          when: new Date().toISOString(),
          source: loaded.source || (isFSConnected() ? 'fs' : 'idb'),
        };
      }
      events.dispatchEvent(createEvent('app:ready', lastSave));
    } catch (error) {
      events.dispatchEvent(
        createEvent('save:error', {
          message: 'No s\'ha pogut carregar l\'estat inicial',
          error,
        }),
      );
    } finally {
      autosaveEnabled = previousAutosave;
    }
  }

  function navigate(viewId) {
    events.dispatchEvent(createEvent('nav:change', { viewId }));
  }

  function toggleAutosave(enabled) {
    autosaveEnabled = !!enabled;
    events.dispatchEvent(createEvent('save:toggle', { enabled: autosaveEnabled }));
    if (autosaveEnabled && pendingPatch) {
      debouncedSave();
    }
  }

  async function connectAutoCopy({ encrypted } = {}) {
    try {
      if (typeof storage.connectFile === 'function') {
        await storage.connectFile({ encrypted });
        events.dispatchEvent(createEvent('fs:connected', {}));
      }
    } catch (error) {
      events.dispatchEvent(createEvent('fs:error', { error }));
      throw error;
    }
  }

  async function disconnectAutoCopy() {
    try {
      if (typeof storage.revoke === 'function') {
        await storage.revoke();
        events.dispatchEvent(createEvent('fs:disconnected', {}));
      }
    } catch (error) {
      events.dispatchEvent(createEvent('fs:error', { error }));
      throw error;
    }
  }

  async function changePassword(oldPwd, newPwd) {
    try {
      if (typeof storage.changePassword === 'function') {
        await storage.changePassword(oldPwd, newPwd);
        events.dispatchEvent(createEvent('crypto:changed', {}));
      }
    } catch (error) {
      events.dispatchEvent(createEvent('save:error', { error }));
      throw error;
    }
  }

  async function manualBackup(note) {
    if (typeof storage.backupNow !== 'function') return null;
    try {
      const result = await storage.backupNow(note);
      events.dispatchEvent(createEvent('backup:done', result));
      return result;
    } catch (error) {
      events.dispatchEvent(createEvent('save:error', { error }));
      throw error;
    }
  }

  function listBackups() {
    if (typeof storage.listBackups === 'function') {
      return storage.listBackups();
    }
    return Promise.resolve([]);
  }

  function runMutation(actionName, callback, meta) {
    const metaData = meta && typeof meta === 'object' ? meta : { action: actionName };
    if (typeof store.transact === 'function') {
      let result;
      store.transact(() => {
        result = callback();
      }, metaData);
      return result;
    }
    return callback();
  }

  function addAssignatura(data, meta) {
    return runMutation('addAssignatura', () => store.addAssignatura?.(data), meta);
  }

  function updateAssignatura(id, patch, meta) {
    return runMutation('updateAssignatura', () => store.updateAssignatura?.(id, patch), meta);
  }

  function addAlumne(data, meta) {
    return runMutation('addAlumne', () => store.addAlumne?.(data), meta);
  }

  function updateAlumne(id, patch, meta) {
    return runMutation('updateAlumne', () => store.updateAlumne?.(id, patch), meta);
  }

  function matricula(alumneId, assignaturaId, meta) {
    return runMutation('matricula', () => store.matricula?.(alumneId, assignaturaId), meta);
  }

  function addCE(assignaturaId, data, meta) {
    return runMutation('addCE', () => store.addCE?.(assignaturaId, data), meta);
  }

  function addCA(ceId, data, meta) {
    return runMutation('addCA', () => store.addCA?.(ceId, data), meta);
  }

  function addCategoria(nom, meta) {
    return runMutation('addCategoria', () => store.addCategoria?.(nom), meta);
  }

  function setPesCategoria(assignaturaId, trimestreId, categoriaId, pes, meta) {
    return runMutation('setPesCategoria', () => store.setPesCategoria?.(assignaturaId, trimestreId, categoriaId, pes), meta);
  }

  function addActivitat(assignaturaId, data, meta) {
    return runMutation('addActivitat', () => store.addActivitat?.(assignaturaId, data), meta);
  }

  function linkCAtoActivitat(activitatId, caId, pes, meta) {
    return runMutation('linkCAtoActivitat', () => store.linkCAtoActivitat?.(activitatId, caId, pes), meta);
  }

  function registraAvaluacioComp(payload, meta) {
    return runMutation('registraAvaluacioComp', () => store.registraAvaluacioComp?.(payload), meta);
  }

  function registraAvaluacioNum(payload, meta) {
    return runMutation('registraAvaluacioNum', () => store.registraAvaluacioNum?.(payload), meta);
  }

  function registraAssistencia(entry, meta) {
    return runMutation('registraAssistencia', () => store.registraAssistencia?.(entry), meta);
  }

  function registraIncidencia(entry, meta) {
    return runMutation('registraIncidencia', () => store.registraIncidencia?.(entry), meta);
  }

  function exportEncrypted(password) {
    return storage.exportEncrypted?.(password);
  }

  function exportCSV_AvaluacionsNumeric(assignaturaId, trimestreId) {
    const state = store.getState?.();
    if (!state) {
      throw new Error('No s\'ha pogut obtenir l\'estat actual');
    }
    const assignatura = state.assignatures?.byId?.[assignaturaId];
    if (!assignatura) {
      throw new Error('Assignatura no trobada');
    }
    const alumnes = extractAlumnesPerAssignatura(state, assignaturaId);
    const rows = [
      ['Alumne', 'Assignatura', 'Trimestre', 'Nota (num)'],
    ];
    alumnes.forEach((alumne) => {
      const nota = store.computeNotaPerAlumne?.(assignaturaId, alumne.id, trimestreId);
      const decimals = assignatura.rounding?.decimals ?? 1;
      const formatted = formatDecimal(nota?.valueNumRounded ?? 0, { decimals, numberToComma });
      rows.push([
        alumne.nom || alumne.name || alumne.id,
        assignatura.nom || assignatura.name || assignatura.id,
        getTrimestreLabel(state, assignaturaId, trimestreId),
        formatted,
      ]);
    });
    return toCSV(rows, { sep: ';', decimalComma: true });
  }

  function exportCSV_AvaluacionsCompetencial(assignaturaId, trimestreId) {
    const state = store.getState?.();
    if (!state) {
      throw new Error('No s\'ha pogut obtenir l\'estat actual');
    }
    const assignatura = state.assignatures?.byId?.[assignaturaId];
    if (!assignatura) {
      throw new Error('Assignatura no trobada');
    }
    const taula = store.computeTaulaCompetencial?.(assignaturaId, trimestreId);
    if (!taula) {
      throw new Error('No s\'ha pogut construir la taula competencial');
    }
    const header = ['Alumne'];
    taula.cas.forEach((caId) => {
      const ca = state.cas?.byId?.[caId];
      header.push(ca?.nom || ca?.id || caId);
    });
    const rows = [header];
    taula.alumnes.forEach((alumneId) => {
      const alumne = state.alumnes?.byId?.[alumneId];
      const line = [alumne?.nom || alumne?.name || alumneId];
      taula.cas.forEach((caId) => {
        const value = taula.values?.[alumneId]?.[caId];
        if (value) {
          const decimals = taula.rounding?.decimals ?? 1;
          const formatted = formatDecimal(value.valueNumRounded ?? 0, {
            decimals,
            numberToComma,
          });
          const cellValue = value?.quali ? `${formatted} (${value.quali})` : formatted;
          line.push(cellValue);
        } else {
          line.push('');
        }
      });
      rows.push(line);
    });
    return toCSV(rows, { sep: ';', decimalComma: true });
  }

  async function exportDOCX_ButlletiAlumne(alumneId, opts = {}) {
    const state = store.getState?.();
    if (!state) {
      throw new Error('No s\'ha pogut obtenir l\'estat actual');
    }
    const alumne = state.alumnes?.byId?.[alumneId];
    if (!alumne) {
      throw new Error('Alumne no trobat');
    }
    const { Document, Paragraph, Table, TableRow, TableCell, HeadingLevel, WidthType, Packer } = getDocx();
    const assignatures = state.assignatures?.allIds
      .map((id) => state.assignatures.byId[id])
      .filter(Boolean);
    const children = [];
    children.push(
      new Paragraph({
        text: `Butlletí de notes - ${alumne.nom || alumne.name || alumne.id}`,
        heading: HeadingLevel?.HEADING_1 || undefined,
      }),
    );
    assignatures.forEach((assignatura) => {
      const matriculat = state.matriculacions.allIds
        .map((id) => state.matriculacions.byId[id])
        .some((mat) => mat.assignaturaId === assignatura.id && mat.alumneId === alumne.id);
      if (!matriculat) return;
      children.push(
        new Paragraph({
          text: assignatura.nom || assignatura.id,
          heading: HeadingLevel?.HEADING_2 || undefined,
        }),
      );
      if (assignatura.mode === 'numeric') {
        const nota = store.computeNotaPerAlumne?.(assignatura.id, alumne.id, opts.trimestreId);
        const decimals = assignatura.rounding?.decimals ?? 1;
        const formatted = formatDecimal(nota?.valueNumRounded ?? 0, { decimals, numberToComma });
        children.push(new Paragraph({ text: `Nota: ${formatted}` }));
      } else {
        const taula = store.computeTaulaCompetencial?.(assignatura.id, opts.trimestreId);
        if (taula) {
          const rows = [];
          rows.push(
            new TableRow({
              children: [
                new TableCell({ children: [new Paragraph({ text: 'CA' })] }),
                new TableCell({ children: [new Paragraph({ text: 'Valoració' })] }),
              ],
            }),
          );
          taula.cas.forEach((caId) => {
            const ca = state.cas?.byId?.[caId];
            const value = taula.values?.[alumne.id]?.[caId];
            const decimals = taula.rounding?.decimals ?? 1;
            const formatted = value
              ? formatDecimal(value.valueNumRounded ?? 0, { decimals, numberToComma })
              : '';
            const cellText = value?.quali ? `${formatted} (${value.quali})` : formatted;
            rows.push(
              new TableRow({
                children: [
                  new TableCell({ children: [new Paragraph({ text: ca?.nom || ca?.id || caId })] }),
                  new TableCell({ children: [new Paragraph({ text: cellText })] }),
                ],
              }),
            );
          });
          children.push(
            new Table({
              width: { size: 100, type: WidthType?.PERCENTAGE || 0 },
              rows,
            }),
          );
        }
      }
    });
    if (state.configGlobal) {
      const signature = state.configGlobal.signature || state.configGlobal.signatura || '';
      if (signature) {
        children.push(new Paragraph({ text: signature }));
      }
    }
    const doc = new Document({ sections: [{ children }] });
    return Packer.toBlob(doc);
  }

  async function exportDOCX_ActaAssignatura(assignaturaId, trimestreId) {
    const state = store.getState?.();
    if (!state) {
      throw new Error('No s\'ha pogut obtenir l\'estat actual');
    }
    const assignatura = state.assignatures?.byId?.[assignaturaId];
    if (!assignatura) {
      throw new Error('Assignatura no trobada');
    }
    const alumnes = extractAlumnesPerAssignatura(state, assignaturaId);
    const { Document, Paragraph, Table, TableRow, TableCell, HeadingLevel, WidthType, Packer } = getDocx();
    const children = [];
    children.push(
      new Paragraph({
        text: `Acta - ${assignatura.nom || assignatura.id}`,
        heading: HeadingLevel?.HEADING_1 || undefined,
      }),
    );
    if (assignatura.mode === 'numeric') {
      const rows = [];
      rows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: 'Alumne' })] }),
            new TableCell({ children: [new Paragraph({ text: 'Nota' })] }),
          ],
        }),
      );
      alumnes.forEach((alumne) => {
        const nota = store.computeNotaPerAlumne?.(assignaturaId, alumne.id, trimestreId);
        const decimals = assignatura.rounding?.decimals ?? 1;
        const formatted = formatDecimal(nota?.valueNumRounded ?? 0, { decimals, numberToComma });
        rows.push(
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ text: alumne.nom || alumne.name || alumne.id })] }),
              new TableCell({ children: [new Paragraph({ text: formatted })] }),
            ],
          }),
        );
      });
      children.push(new Table({ width: { size: 100, type: WidthType?.PERCENTAGE || 0 }, rows }));
    } else {
      const taula = store.computeTaulaCompetencial?.(assignaturaId, trimestreId);
      const headerCells = [new TableCell({ children: [new Paragraph({ text: 'Alumne' })] })];
      taula.cas.forEach((caId) => {
        const ca = state.cas?.byId?.[caId];
        headerCells.push(new TableCell({ children: [new Paragraph({ text: ca?.nom || ca?.id || caId })] }));
      });
      const rows = [new TableRow({ children: headerCells })];
      alumnes.forEach((alumne) => {
        const cells = [new TableCell({ children: [new Paragraph({ text: alumne.nom || alumne.name || alumne.id })] })];
        taula.cas.forEach((caId) => {
          const value = taula.values?.[alumne.id]?.[caId];
          const decimals = taula.rounding?.decimals ?? 1;
          const formatted = value
            ? formatDecimal(value.valueNumRounded ?? 0, { decimals, numberToComma })
            : '';
          const cellText = value?.quali ? `${formatted} (${value.quali})` : formatted;
          cells.push(new TableCell({ children: [new Paragraph({ text: cellText })] }));
        });
        rows.push(new TableRow({ children: cells }));
      });
      children.push(new Table({ width: { size: 100, type: WidthType?.PERCENTAGE || 0 }, rows }));
    }
    const doc = new Document({ sections: [{ children }] });
    return Packer.toBlob(doc);
  }

  async function refreshFromDiskIfNewer() {
    try {
      const target = storage.fs?.loadFromFile || storage.loadFromFile;
      if (!target) return null;
      const remote = await target.call(storage.fs || storage);
      if (remote && remote.version > (lastSave.version || 0)) {
        await applyRemoteState(remote, 'fs:refresh');
        return remote;
      }
      return null;
    } catch (error) {
      events.dispatchEvent(createEvent('fs:error', { error }));
      throw error;
    }
  }

  return {
    init,
    navigate,
    toggleAutosave,
    connectAutoCopy,
    disconnectAutoCopy,
    changePassword,
    manualBackup,
    listBackups,
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
    exportEncrypted,
    exportCSV_AvaluacionsNumeric,
    exportCSV_AvaluacionsCompetencial,
    exportDOCX_ButlletiAlumne,
    exportDOCX_ActaAssignatura,
    save: saveState,
    isFSConnected,
    getLastSaveInfo,
    refreshFromDiskIfNewer,
    events,
    destroy() {
      unsubscribe?.();
      forwardedStorageEvents.forEach((type) => {
        storageEvents.removeEventListener(type, republishStorageEvent);
      });
    },
  };
}

export default createActions;
