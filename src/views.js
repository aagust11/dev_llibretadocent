/* eslint-disable no-param-reassign */
const lucide = globalThis.lucide;

const VIEW_IDS = [
  'view-welcome',
  'view-assignatures',
  'view-alumnes',
  'view-rubrica',
  'view-calendari',
  'view-fitxa',
  'view-exportacions',
  'view-configuracio',
];

const TOAST_TIMEOUT = 5000;
const KEY_COMBOS = Object.freeze({
  help: '?',
  gotoAssignatures: 'g+a',
  gotoRubrica: 'g+r',
});

let appContext = null;
let containers = {};
let activeView = 'view-welcome';
let unsubscribeStore = null;
let renderQueued = false;
let lastRenderedView = null;
let lastRenderedVersion = null;

const dirtyFlags = {
  header: true,
  view: true,
  badges: true,
};

const uiState = {
  header: {
    autosave: 'Local',
    tone: 'emerald',
    lastEvent: '',
    fsConnected: false,
    locked: false,
    error: null,
  },
  assignatures: {
    selectedId: null,
    selectedCE: null,
  },
  alumnes: {
    search: '',
  },
  rubrica: {
    selectedAssignatura: null,
    selectedTrimestre: null,
    mode: 'single',
    selectedActivitat: null,
    numericInput: false,
  },
  numeric: {
    selectedAssignatura: null,
    selectedTrimestre: null,
  },
  fitxa: {
    filter: '',
    selectedAlumne: null,
    selectedTab: 'resum',
    filtreAssignatura: 'totes',
    filtrePeriode: 'tot',
    filtreTrimestre: '',
    rangInici: '',
    rangFi: '',
  },
  calendari: {
    selectedAssignatura: null,
    simulatorDate: '',
    simulatorResult: null,
    importTarget: 'festius',
    pendingSave: false,
    newTrimestre: { tInici: '', tFi: '' },
    newVersio: { effectiveFrom: '', dies: new Set([1, 2, 3, 4, 5]) },
    exportMode: 'curs',
    exportFrom: '',
    exportTo: '',
    exportSummary: '',
    exportLocation: '',
    viewMonth: null,
  },
};

const CALENDAR_MUTATIONS = new Set([
  'setCursRange',
  'setDiesSetmanals',
  'addTrimestre',
  'removeTrimestre',
  'addFestius',
  'removeFestius',
  'addExcepcions',
  'removeExcepcio',
  'addHorariVersio',
  'activateHorariVersio',
]);

let toastContainer = null;
let modalBackdrop = null;
let focusBeforeModal = null;
let lastKeySequence = '';
let keySequenceTimeout = null;

function getDocument() {
  return typeof document !== 'undefined' ? document : null;
}

function qs(selector) {
  return getDocument()?.querySelector(selector) || null;
}

function createElement(tag, options = {}) {
  const doc = getDocument();
  const element = doc.createElement(tag);
  if (options.className) {
    element.className = options.className;
  }
  if (options.text) {
    element.textContent = options.text;
  }
  if (options.html) {
    element.innerHTML = options.html;
  }
  if (options.attrs) {
    Object.entries(options.attrs).forEach(([key, value]) => {
      if (value === undefined || value === null) return;
      element.setAttribute(key, value);
    });
  }
  if (options.children) {
    options.children.forEach((child) => {
      if (!child) return;
      if (typeof child === 'string') {
        element.appendChild(doc.createTextNode(child));
      } else {
        element.appendChild(child);
      }
    });
  }
  return element;
}

function clearElement(element) {
  while (element && element.firstChild) {
    element.removeChild(element.firstChild);
  }

  const selected = uiState.assignatures.selectedId
    ? state.assignatures?.byId?.[uiState.assignatures.selectedId]
    : null;
  if (selected) {
    container.appendChild(renderAssignaturaDetail(selected, state));
  }
}

function ensureLucide() {
  if (lucide?.createIcons) {
    lucide.createIcons();
  }
}

function formatNumber(value, decimals = 1) {
  if (!appContext?.i18n) return Number(value ?? 0).toFixed(decimals);
  return appContext.i18n.formatNumberComma(Number(value ?? 0), decimals);
}

function formatDate(value) {
  if (!value) return '';
  if (!appContext?.i18n) return new Date(value).toLocaleDateString('ca-ES');
  return appContext.i18n.formatDateCAT(value);
}

const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/u;

function toISODateLocal(value) {
  if (!value && value !== 0) return '';
  if (typeof value === 'string') {
    if (ISO_DATE_RE.test(value)) return value;
    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) return '';
    return parsed.toISOString().slice(0, 10);
  }
  const date = value instanceof Date ? value : new Date(value);
  if (Number.isNaN(date.getTime())) return '';
  return date.toISOString().slice(0, 10);
}

function downloadBlob(blob, filename) {
  if (!blob) return;
  const url = URL.createObjectURL(blob);
  const link = createElement('a', { attrs: { href: url, download: filename } });
  getDocument()?.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 0);
}

function getCalendariForAssignatura(state, assignaturaId) {
  if (!assignaturaId) return null;
  const calendaris = state.calendaris?.allIds || [];
  const calendari = calendaris
    .map((id) => state.calendaris.byId[id])
    .find((cal) => cal.assignaturaId === assignaturaId);
  if (!calendari) {
    return {
      assignaturaId,
      cursInici: '',
      cursFi: '',
      diesSetmanals: [],
      trimestres: [],
      festius: [],
      excepcions: [],
      horariVersions: [],
      horariActivaId: null,
    };
  }
  return {
    ...calendari,
    trimestres: (calendari.trimestres || []).map((t) => ({
      ...t,
      tInici: toISODateLocal(t.tInici),
      tFi: toISODateLocal(t.tFi),
    })),
    festius: (calendari.festius || []).map((festiu) => ({
      dataISO: toISODateLocal(festiu.dataISO || festiu.data),
      motiu: festiu.motiu || '',
    })),
    excepcions: (calendari.excepcions || []).map((excepcio) => ({
      dataISO: toISODateLocal(excepcio.dataISO || excepcio.data),
      motiu: excepcio.motiu || '',
    })),
    horariVersions: (calendari.horariVersions || []).map((versio) => ({
      ...versio,
      effectiveFrom: toISODateLocal(versio.effectiveFrom),
      diesSetmanals: Array.from(new Set(versio.diesSetmanals || [])).sort((a, b) => a - b),
    })),
  };
}

const WEEKDAY_ORDER = [1, 2, 3, 4, 5, 6, 0];
const WEEKDAY_LABEL = {
  0: 'Dg',
  1: 'Dl',
  2: 'Dt',
  3: 'Dc',
  4: 'Dj',
  5: 'Dv',
  6: 'Ds',
};

function formatDiesSetmanals(dies) {
  if (!dies?.length) return 'Cap dia definit';
  const ordered = Array.from(new Set(dies)).sort((a, b) => WEEKDAY_ORDER.indexOf(a) - WEEKDAY_ORDER.indexOf(b));
  return ordered.map((dia) => WEEKDAY_LABEL[dia] || dia).join(', ');
}

function rangesOverlap(aInici, aFi, bInici, bFi) {
  if (!aInici || !aFi || !bInici || !bFi) return false;
  return !(aFi < bInici || aInici > bFi);
}

function validateTrimestreRange(trimestres, currentId, tInici, tFi) {
  if (!tInici || !tFi) {
    return 'Les dues dates són obligatòries.';
  }
  if (tFi < tInici) {
    return 'La data de fi ha de ser posterior o igual a la data d\'inici.';
  }
  for (const trimestre of trimestres) {
    if (trimestre.id === currentId) continue;
    if (rangesOverlap(tInici, tFi, trimestre.tInici, trimestre.tFi)) {
      return `El rang se solapa amb ${trimestre.nom || trimestre.id}.`;
    }
  }
  return '';
}

function createBadge(text, tone = 'slate') {
  const palette = {
    slate: 'bg-slate-100 text-slate-700',
    emerald: 'bg-emerald-100 text-emerald-700',
    sky: 'bg-sky-100 text-sky-700',
    amber: 'bg-amber-100 text-amber-700',
    rose: 'bg-rose-100 text-rose-700',
  };
  const classes = palette[tone] || palette.slate;
  return createElement('span', {
    className: `inline-flex items-center gap-1 rounded-full px-3 py-1 text-xs font-semibold ${classes}`,
    text,
  });
}

let toastId = 0;

function ensureToastContainer() {
  if (toastContainer) return toastContainer;
  const doc = getDocument();
  toastContainer = createElement('div', {
    className:
      'fixed inset-x-0 top-4 z-50 mx-auto flex w-full max-w-sm flex-col gap-3 px-4 sm:max-w-md',
    attrs: { role: 'status', 'aria-live': 'polite' },
  });
  doc.body.appendChild(toastContainer);
  return toastContainer;
}

function removeToast(element) {
  if (!element) return;
  element.classList.add('opacity-0', 'translate-y-2');
  setTimeout(() => element.remove(), 200);
}

export function Toast(message, type = 'info', { timeout = 5000 } = {}) {
  const container = ensureToastContainer();
  const palette = {
    success: 'bg-emerald-50 text-emerald-800 border-emerald-200',
    warn: 'bg-amber-50 text-amber-800 border-amber-200',
    error: 'bg-rose-50 text-rose-800 border-rose-200',
    info: 'bg-slate-50 text-slate-800 border-slate-200',
  };
  const icon = type === 'success' ? 'check-circle' : type === 'error' ? 'alert-triangle' : type === 'warn' ? 'alert-circle' : 'info';
  const toast = createElement('div', {
    className: `flex w-full items-start gap-3 rounded-lg border px-4 py-3 text-sm shadow-lg transition ${palette[type] || palette.info}`,
    attrs: { role: 'alert', 'data-toast-id': `toast-${toastId++}` },
  });
  toast.append(
    createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': icon } }),
    createElement('p', { className: 'flex-1 font-medium leading-snug', text: message }),
  );
  const closeBtn = createElement('button', {
    className:
      'ml-auto inline-flex rounded-full p-1 text-slate-500 transition hover:bg-white hover:text-slate-900 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-500',
    attrs: { type: 'button', 'aria-label': 'Tanca avís' },
  });
  closeBtn.appendChild(createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'x' } }));
  closeBtn.addEventListener('click', () => removeToast(toast));
  toast.appendChild(closeBtn);
  container.appendChild(toast);
  ensureLucide();
  if (timeout > 0) {
    setTimeout(() => removeToast(toast), timeout);
  }
  return toast;
}

function ensureModalBackdrop() {
  if (modalBackdrop) return modalBackdrop;
  modalBackdrop = createElement('div', {
    className: 'fixed inset-0 z-40 hidden items-center justify-center bg-slate-900/60 px-4 py-6',
    attrs: { role: 'presentation' },
  });
  getDocument().body.appendChild(modalBackdrop);
  return modalBackdrop;
}

function closeModal() {
  if (!modalBackdrop) return;
  modalBackdrop.classList.add('hidden');
  clearElement(modalBackdrop);
  if (focusBeforeModal) {
    focusBeforeModal.focus();
    focusBeforeModal = null;
  }
}

export function Modal(title, content, actions = []) {
  ensureModalBackdrop();
  focusBeforeModal = getDocument().activeElement;
  const dialog = createElement('div', {
    className: 'max-h-full w-full max-w-xl overflow-hidden rounded-xl bg-white shadow-2xl',
    attrs: { role: 'dialog', 'aria-modal': 'true', tabindex: '-1' },
  });
  const header = createElement('div', {
    className: 'flex items-start justify-between border-b border-slate-200 px-6 py-4',
  });
  header.appendChild(createElement('h2', { className: 'text-lg font-semibold text-slate-900', text: title }));
  const closeBtn = createElement('button', {
    className:
      'inline-flex rounded-full p-1 text-slate-500 transition hover:bg-slate-100 hover:text-slate-900 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-500',
    attrs: { type: 'button', 'aria-label': 'Tanca' },
  });
  closeBtn.appendChild(createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'x' } }));
  closeBtn.addEventListener('click', closeModal);
  header.appendChild(closeBtn);
  const body = createElement('div', { className: 'max-h-[60vh] overflow-y-auto px-6 py-4 text-sm text-slate-700' });
  if (typeof content === 'string') {
    body.textContent = content;
  } else if (content instanceof Node) {
    body.appendChild(content);
  }
  const footer = createElement('div', {
    className: 'flex flex-col-reverse gap-2 border-t border-slate-200 px-6 py-4 sm:flex-row sm:justify-end',
  });
  actions.forEach((action) => {
    const btn = createElement('button', {
      className: `inline-flex items-center justify-center rounded-md px-4 py-2 text-sm font-semibold shadow-sm transition focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900 ${
        action.primary
          ? 'bg-slate-900 text-white hover:bg-slate-800'
          : 'bg-white text-slate-700 ring-1 ring-inset ring-slate-200 hover:bg-slate-50'
      }`,
      attrs: { type: action.type || 'button' },
      text: action.label,
    });
    btn.addEventListener('click', (event) => {
      if (action.onClick) {
        action.onClick(event, closeModal);
      } else {
        closeModal();
      }
    });
    footer.appendChild(btn);
  });
  dialog.append(header, body, footer);
  modalBackdrop.innerHTML = '';
  modalBackdrop.appendChild(dialog);
  modalBackdrop.classList.remove('hidden');
  setTimeout(() => dialog.focus(), 10);
  ensureLucide();
  return { close: closeModal, element: dialog };
}

export function Confirm(message, { confirmLabel = 'Accepta', cancelLabel = 'Cancel·la' } = {}) {
  return new Promise((resolve) => {
    const content = createElement('p', { className: 'text-sm text-slate-600', text: message });
    Modal('Confirmació', content, [
      { label: cancelLabel, onClick: () => { closeModal(); resolve(false); } },
      { label: confirmLabel, primary: true, onClick: () => { closeModal(); resolve(true); } },
    ]);
  });
}

export function DataTable({ columns, rows, caption }) {
  const table = createElement('table', {
    className: 'min-w-full divide-y divide-slate-200 overflow-hidden rounded-lg bg-white text-left text-sm shadow-sm',
  });
  if (caption) {
    table.appendChild(
      createElement('caption', {
        className: 'bg-slate-50 px-4 py-2 text-left text-xs font-semibold uppercase tracking-wide text-slate-500',
        text: caption,
      }),
    );
  }
  const thead = createElement('thead', { className: 'bg-slate-50' });
  const headRow = createElement('tr');
  columns.forEach((column) => {
    headRow.appendChild(
      createElement('th', {
        className: 'px-4 py-3 text-xs font-semibold uppercase tracking-wide text-slate-500',
        text: column.label,
        attrs: { scope: 'col' },
      }),
    );
  });
  thead.appendChild(headRow);
  table.appendChild(thead);
  const tbody = createElement('tbody', { className: 'divide-y divide-slate-200 bg-white' });
  rows.forEach((row, rowIndex) => {
    const tr = createElement('tr', {
      className: rowIndex % 2 === 0 ? 'bg-white focus-within:bg-slate-50' : 'bg-slate-50 focus-within:bg-slate-100',
    });
    columns.forEach((column, columnIndex) => {
      const td = createElement('td', {
        className: 'px-4 py-3 text-sm text-slate-700',
        attrs: { tabindex: '0', 'data-column': column.id || column.label },
      });
      const value = typeof column.accessor === 'function' ? column.accessor(row, rowIndex, columnIndex) : row[column.id];
      if (value instanceof Node) {
        td.appendChild(value);
      } else if (value !== undefined && value !== null) {
        td.textContent = value;
      }
      td.addEventListener('keydown', (event) => {
        if (event.key === 'Enter' && typeof column.onEnter === 'function') {
          column.onEnter({ event, row, rowIndex, columnIndex });
        }
      });
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
  table.appendChild(tbody);
  return table;
}

function scheduleRender(reason = 'manual') {
  if (renderQueued) return;
  renderQueued = true;
  requestAnimationFrame(() => {
    renderQueued = false;
    doRender(reason);
  });
}

function renderWelcome(container, state) {
  clearElement(container);
  container.classList.add('space-y-4');
  const assignaturesCount = state.assignatures?.allIds?.length || 0;
  const alumnesCount = state.alumnes?.allIds?.length || 0;
  container.append(
    createElement('p', {
      className: 'text-sm text-slate-600',
      text: `Tens ${assignaturesCount} assignatures i ${alumnesCount} alumnes registrats.`,
    }),
    createElement('p', {
      className: 'text-sm text-slate-600',
      text: 'Fes servir els enllaços de la barra lateral o les dreceres (g a, g r, ?) per navegar ràpidament.',
    }),
  );
}

function buildAssignaturesRows(state) {
  const assignatures = state.assignatures?.allIds || [];
  return assignatures.map((id) => state.assignatures.byId[id]);
}

function renderAssignatures(container, state) {
  clearElement(container);
  container.classList.add('space-y-6');
  const header = createElement('div', {
    className: 'flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between',
  });
  header.append(
    createElement('div', {
      children: [
        createElement('p', {
          className: 'text-sm text-slate-600',
          text: `Assignatures actives: ${(state.assignatures?.allIds || []).length}`,
        }),
      ],
    }),
    (() => {
      const btn = createElement('button', {
        className:
          'inline-flex items-center gap-2 rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
        attrs: { type: 'button' },
        children: [
          createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'plus' } }),
          createElement('span', { text: 'Nova assignatura' }),
        ],
      });
      btn.addEventListener('click', () => openAssignaturaModal());
      return btn;
    })(),
  );
  container.appendChild(header);

  const rows = buildAssignaturesRows(state);
  if (!rows.length) {
    container.appendChild(
      createElement('p', {
        className: 'rounded-md border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-sm text-slate-600',
        text: 'Encara no hi ha assignatures. Crea la primera per començar.',
      }),
    );
  } else {
    const table = DataTable({
      caption: 'Assignatures disponibles',
      columns: [
        { id: 'nom', label: 'Nom' },
        {
          id: 'anyCurs',
          label: 'Curs',
          accessor: (row) => row.anyCurs || '—',
        },
        {
          id: 'mode',
          label: 'Mode',
          accessor: (row) => (row.mode === 'competencial' ? 'Competencial' : 'Numèric'),
        },
        {
          id: 'accions',
          label: 'Accions',
          accessor: (row) => {
            const wrapper = createElement('div', { className: 'flex flex-wrap gap-2' });
            const obrir = createElement('button', {
              className:
                'inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
              attrs: { type: 'button' },
              children: [
                createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'folder-open' } }),
                createElement('span', { text: 'Obre' }),
              ],
            });
            obrir.addEventListener('click', () => {
              uiState.assignatures.selectedId = row.id;
              dirtyFlags.view = true;
              scheduleRender('assignatura-open');
            });
            const editar = createElement('button', {
              className:
                'inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
              attrs: { type: 'button' },
              children: [
                createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'pencil' } }),
                createElement('span', { text: 'Edita' }),
              ],
            });
            editar.addEventListener('click', () => openAssignaturaModal(row));
            const config = createElement('button', {
              className:
                'inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
              attrs: { type: 'button' },
              children: [
                createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'sliders-horizontal' } }),
                createElement('span', { text: 'Configura' }),
              ],
            });
            config.addEventListener('click', () => openAssignaturaConfig(row));
            const vincle = createElement('button', {
              className:
                'inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
              attrs: { type: 'button' },
              children: [
                createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'link' } }),
                createElement('span', { text: 'Vincle' }),
              ],
            });
            vincle.addEventListener('click', () => Toast('Funció de vincle pendent.', 'info'));
            wrapper.append(obrir, editar, config, vincle);
            return wrapper;
          },
        },
      ],
      rows,
    });
    container.appendChild(table);
  }
}

function openAssignaturaModal(assignatura) {
  const form = createElement('form', { className: 'flex flex-col gap-4' });
  const fields = [
    { id: 'nom', label: 'Nom', type: 'text', required: true, value: assignatura?.nom ?? '' },
    { id: 'anyCurs', label: 'Any acadèmic', type: 'text', value: assignatura?.anyCurs ?? '' },
  ];
  fields.forEach((field) => {
    const input = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: {
        id: field.id,
        name: field.id,
        type: field.type,
        required: field.required ? 'true' : undefined,
        value: field.value,
      },
    });
    form.appendChild(
      createElement('label', {
        className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
        attrs: { for: field.id },
        children: [createElement('span', { text: field.label }), input],
      }),
    );
  });
  const modeSelect = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
    attrs: { name: 'mode', id: 'mode' },
  });
  ['numeric', 'competencial'].forEach((value) => {
    const option = createElement('option', {
      attrs: { value },
      text: value === 'numeric' ? 'Avaluació numèrica' : 'Avaluació competencial',
    });
    if ((assignatura?.mode || 'numeric') === value) option.selected = true;
    modeSelect.appendChild(option);
  });
  form.appendChild(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      attrs: { for: 'mode' },
      children: [createElement('span', { text: 'Mode d\'avaluació' }), modeSelect],
    }),
  );
  Modal(assignatura ? 'Edita assignatura' : 'Nova assignatura', form, [
    { label: 'Cancel·la', onClick: closeModal },
    {
      label: assignatura ? 'Desa canvis' : 'Crea assignatura',
      primary: true,
      type: 'submit',
      onClick: (event) => {
        event.preventDefault();
        const formData = new FormData(form);
        const nom = String(formData.get('nom') || '').trim();
        if (!nom) {
          Toast('El nom és obligatori.', 'warn');
          return;
        }
        const payload = {
          nom,
          anyCurs: String(formData.get('anyCurs') || '').trim(),
          mode: formData.get('mode') || 'numeric',
        };
        try {
          if (assignatura) {
            appContext.store.updateAssignatura(assignatura.id, payload);
            Toast('Assignatura actualitzada.', 'success');
          } else {
            const id = appContext.store.addAssignatura(payload);
            uiState.assignatures.selectedId = id;
            Toast('Assignatura creada correctament.', 'success');
          }
          dirtyFlags.view = true;
          scheduleRender('assignatura-modal');
          closeModal();
        } catch (error) {
          Toast(`Error en desar l\'assignatura: ${error.message || error}`, 'error');
        }
      },
    },
  ]);
}

function openAssignaturaConfig(assignatura) {
  const form = createElement('form', { className: 'flex flex-col gap-4 text-sm text-slate-700' });
  const decimalsInput = createElement('input', {
    className:
      'w-24 rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
    attrs: { type: 'number', min: '0', max: '3', step: '1', value: assignatura?.rounding?.decimals ?? 1, name: 'decimals', id: 'decimals' },
  });
  const ranges = assignatura?.qualitativeRanges || appContext.store.getState().configGlobal.qualitativeRanges;
  const rangeWrapper = createElement('div', { className: 'grid gap-3 md:grid-cols-2' });
  Object.entries(ranges || {}).forEach(([key, value]) => {
    const [min, max] = value;
    const minInput = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: { type: 'number', step: '0.1', min: '0', max: '10', value: min, name: `${key}-min` },
    });
    const maxInput = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: { type: 'number', step: '0.1', min: '0', max: '10', value: max, name: `${key}-max` },
    });
    rangeWrapper.appendChild(
      createElement('div', {
        className: 'flex flex-col gap-2 rounded-md border border-slate-200 p-3',
        children: [
          createElement('span', { className: 'text-xs font-semibold uppercase text-slate-500', text: key }),
          createElement('label', {
            className: 'flex flex-col gap-1 text-xs',
            children: [createElement('span', { text: 'Mínim' }), minInput],
          }),
          createElement('label', {
            className: 'flex flex-col gap-1 text-xs',
            children: [createElement('span', { text: 'Màxim' }), maxInput],
          }),
        ],
      }),
    );
  });
  form.append(
    createElement('p', { className: 'text-sm font-semibold text-slate-700', text: 'Arrodoniment i franges qualitatives' }),
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Decimals' }), decimalsInput],
    }),
    rangeWrapper,
  );
  Modal('Configuració de l\'assignatura', form, [
    { label: 'Tanca', onClick: closeModal },
    {
      label: 'Desa',
      primary: true,
      type: 'submit',
      onClick: (event) => {
        event.preventDefault();
        const decimals = Number(decimalsInput.value);
        if (Number.isNaN(decimals) || decimals < 0 || decimals > 3) {
          Toast('Els decimals han de ser entre 0 i 3.', 'warn');
          return;
        }
        const qualitativeRanges = {};
        let valid = true;
        ['NA', 'AS', 'AN', 'AE'].forEach((key) => {
          if (!valid) return;
          const min = Number(form.querySelector(`[name="${key}-min"]`).value);
          const max = Number(form.querySelector(`[name="${key}-max"]`).value);
          if (Number.isNaN(min) || Number.isNaN(max) || min > max) {
            valid = false;
            Toast(`Interval invàlid per ${key}`, 'warn');
            return;
          }
          qualitativeRanges[key] = [min, max];
        });
        if (!valid) return;
        try {
          appContext.store.updateAssignatura(assignatura.id, {
            rounding: { decimals, mode: assignatura?.rounding?.mode || 'half-up' },
            qualitativeRanges,
          });
          Toast('Configuració actualitzada.', 'success');
          dirtyFlags.view = true;
          scheduleRender('assignatura-config');
          closeModal();
        } catch (error) {
          Toast(`Error en desar la configuració: ${error.message || error}`, 'error');
        }
      },
    },
  ]);
}

function renderAssignaturaDetail(assignatura, state) {
  const wrapper = createElement('section', {
    className: 'space-y-6 rounded-lg border border-slate-200 bg-white p-6 shadow-sm',
    attrs: { 'aria-label': `Detall de ${assignatura.nom}` },
  });
  wrapper.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-2 sm:flex-row sm:items-baseline sm:justify-between',
      children: [
        createElement('div', {
          children: [
            createElement('h4', { className: 'text-lg font-semibold text-slate-900', text: assignatura.nom }),
            createElement('p', {
              className: 'text-sm text-slate-500',
              text: assignatura.mode === 'competencial' ? 'Mode competencial' : 'Mode numèric',
            }),
          ],
        }),
        createElement('div', {
          className: 'flex flex-wrap gap-3 text-xs text-slate-500',
          children: [
            createElement('span', {
              className: 'inline-flex items-center gap-1 rounded-full bg-slate-100 px-3 py-1 font-semibold text-slate-700',
              text: `CE: ${countCE(assignatura, state)}`,
            }),
            createElement('span', {
              className: 'inline-flex items-center gap-1 rounded-full bg-slate-100 px-3 py-1 font-semibold text-slate-700',
              text: `CA: ${countCA(assignatura, state)}`,
            }),
            createElement('span', {
              className: 'inline-flex items-center gap-1 rounded-full bg-slate-100 px-3 py-1 font-semibold text-slate-700',
              text: `Activitats: ${countActivitats(assignatura, state)}`,
            }),
          ],
        }),
      ],
    }),
  );

  const layout = createElement('div', {
    className: 'grid gap-6 lg:grid-cols-2',
  });
  layout.appendChild(renderCEPanel(assignatura, state));
  layout.appendChild(renderCAPanel(assignatura, state));
  wrapper.appendChild(layout);

  if (assignatura.mode === 'numeric') {
    wrapper.appendChild(renderCategoriesPanel(assignatura, state));
  }
  wrapper.appendChild(renderActivitatsPanel(assignatura, state));
  return wrapper;
}

function getCEs(assignatura, state) {
  return state.ces?.allIds
    ?.map((id) => state.ces.byId[id])
    .filter((ce) => ce.assignaturaId === assignatura.id)
    .sort((a, b) => a.position - b.position);
}

function getCAsForAssignatura(assignatura, state) {
  return state.cas?.allIds
    ?.map((id) => state.cas.byId[id])
    .filter((ca) => state.ces.byId[ca.ceId]?.assignaturaId === assignatura.id)
    .sort((a, b) => a.position - b.position);
}

function countCE(assignatura, state) {
  return getCEs(assignatura, state)?.length || 0;
}

function countCA(assignatura, state) {
  return getCAsForAssignatura(assignatura, state)?.length || 0;
}

function countActivitats(assignatura, state) {
  return (
    state.activitats?.allIds?.map((id) => state.activitats.byId[id]).filter((act) => act.assignaturaId === assignatura.id)
      ?.length || 0
  );
}

function renderCEPanel(assignatura, state) {
  const ces = getCEs(assignatura, state);
  const panel = createElement('div', {
    className: 'space-y-3 rounded-lg border border-slate-200 p-4',
  });
  panel.appendChild(
    createElement('h5', {
      className: 'text-sm font-semibold uppercase tracking-wide text-slate-500',
      text: 'Competències específiques',
    }),
  );
  if (!ces?.length) {
    panel.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: 'Encara no s\'ha registrat cap CE.',
      }),
    );
    return panel;
  }
  const list = createElement('ul', { className: 'space-y-2' });
  ces.forEach((ce) => {
    const selected = uiState.assignatures.selectedCE === ce.id;
    const item = createElement('li', {
      className: `rounded-md border px-3 py-2 text-sm transition ${
        selected ? 'border-slate-500 bg-slate-100' : 'border-slate-200 bg-white'
      }`,
    });
    const code = `CE${ce.textBetween || ''}${ce.index}`;
    const button = createElement('button', {
      className: 'flex w-full items-center justify-between text-left',
      attrs: { type: 'button' },
      children: [
        createElement('span', { className: 'font-semibold text-slate-700', text: code }),
        createElement('span', { className: 'text-xs text-slate-500', text: `Posició ${ce.position}` }),
      ],
    });
    button.addEventListener('click', () => {
      uiState.assignatures.selectedCE = selected ? null : ce.id;
      dirtyFlags.view = true;
      scheduleRender('ce-select');
    });
    item.appendChild(button);
    list.appendChild(item);
  });
  panel.appendChild(list);
  return panel;
}

function renderCAPanel(assignatura, state) {
  const panel = createElement('div', {
    className: 'space-y-3 rounded-lg border border-slate-200 p-4',
  });
  const selectedCE = uiState.assignatures.selectedCE
    ? state.ces?.byId?.[uiState.assignatures.selectedCE]
    : null;
  panel.appendChild(
    createElement('h5', {
      className: 'text-sm font-semibold uppercase tracking-wide text-slate-500',
      text: selectedCE ? `Criteris d'avaluació de CE${selectedCE.textBetween || ''}${selectedCE.index}` : 'Criteris d\'avaluació',
    }),
  );
  if (!selectedCE) {
    panel.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: 'Selecciona una CE per veure els criteris associats.',
      }),
    );
    return panel;
  }
  const cas = getCAsForAssignatura(assignatura, state).filter((ca) => ca.ceId === selectedCE.id);
  if (!cas.length) {
    panel.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: 'Sense criteris registrats per aquesta CE.',
      }),
    );
    return panel;
  }
  const list = createElement('ul', { className: 'space-y-2' });
  cas.forEach((ca) => {
    const item = createElement('li', {
      className: 'rounded-md border border-slate-200 bg-white px-3 py-2 text-sm shadow-sm',
    });
    item.append(
      createElement('div', {
        className: 'flex items-center justify-between',
        children: [
          createElement('span', {
            className: 'font-semibold text-slate-700',
            text: `CA${ca.textBetween || ''}${selectedCE.index}.${ca.index}`,
          }),
          createElement('span', {
            className: 'text-xs text-slate-500',
            text: `Pes: ${formatNumber(ca.pesDinsCE ?? 0, 2)}`,
          }),
        ],
      }),
    );
    list.appendChild(item);
  });
  panel.appendChild(list);
  return panel;
}

function renderCategoriesPanel(assignatura, state) {
  const panel = createElement('div', {
    className: 'space-y-3 rounded-lg border border-slate-200 p-4',
  });
  panel.appendChild(
    createElement('h5', {
      className: 'text-sm font-semibold uppercase tracking-wide text-slate-500',
      text: 'Categories i pesos',
    }),
  );
  const categories = state.categories?.allIds?.map((id) => state.categories.byId[id]) || [];
  if (!categories.length) {
    panel.appendChild(
      createElement('p', { className: 'text-xs text-slate-500', text: 'Sense categories definides.' }),
    );
    return panel;
  }
  const table = DataTable({
    caption: 'Pesos de categoria',
    columns: [
      { id: 'nom', label: 'Categoria' },
      {
        id: 'pesos',
        label: 'Pes per trimestre',
        accessor: (row) => {
          const container = createElement('div', { className: 'flex flex-wrap gap-2' });
          const mapping = assignatura.categoriaPesos || {};
          const trimestres = Object.keys(mapping);
          if (!trimestres.length) {
            container.appendChild(
              createElement('span', {
                className: 'text-xs text-slate-500',
                text: 'Sense trimestres configurats.',
              }),
            );
            return container;
          }
          trimestres.forEach((trimId) => {
            const value = mapping[trimId]?.[row.id] ?? 0;
            const input = createElement('input', {
              className:
                'w-24 rounded-md border border-slate-300 px-2 py-1 text-xs shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
              attrs: { type: 'number', min: '0', step: '0.1', value },
            });
            input.addEventListener('change', (event) => {
              const pes = Number(event.currentTarget.value);
              if (Number.isNaN(pes) || pes < 0) {
                Toast('El pes ha de ser ≥ 0.', 'warn');
                return;
              }
              try {
                appContext.store.setPesCategoria(assignatura.id, trimId, row.id, pes);
                Toast('Pes actualitzat.', 'success');
              } catch (error) {
                Toast(`Error en actualitzar el pes: ${error.message || error}`, 'error');
              }
            });
            const badge = createElement('span', {
              className: 'inline-flex items-center gap-2 rounded-md border border-slate-200 px-2 py-1 text-xs',
              children: [createElement('span', { className: 'font-semibold text-slate-600', text: trimId }), input],
            });
            container.appendChild(badge);
          });
          return container;
        },
      },
    ],
    rows: categories,
  });
  panel.appendChild(table);
  return panel;
}

function renderActivitatsPanel(assignatura, state) {
  const activitats = state.activitats?.allIds
    ?.map((id) => state.activitats.byId[id])
    .filter((act) => act.assignaturaId === assignatura.id)
    .sort((a, b) => new Date(a.data) - new Date(b.data));
  const panel = createElement('div', {
    className: 'space-y-3 rounded-lg border border-slate-200 p-4',
  });
  panel.appendChild(
    createElement('h5', {
      className: 'text-sm font-semibold uppercase tracking-wide text-slate-500',
      text: 'Activitats',
    }),
  );
  const addBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-1.5 text-xs font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [
      createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'calendar-plus' } }),
      createElement('span', { text: 'Nova activitat' }),
    ],
  });
  addBtn.addEventListener('click', () => {
    const form = createElement('form', { className: 'flex flex-col gap-4' });
    const dateInput = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: { type: 'date', name: 'data' },
    });
    const pesInput = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: { type: 'number', name: 'pes', step: '0.1', min: '0', value: '1' },
    });
    form.append(
      createElement('label', {
        className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
        children: [createElement('span', { text: 'Data' }), dateInput],
      }),
      createElement('label', {
        className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
        children: [createElement('span', { text: 'Pes' }), pesInput],
      }),
    );
    Modal('Nova activitat', form, [
      { label: 'Cancel·la', onClick: closeModal },
      {
        label: 'Crea',
        primary: true,
        type: 'submit',
        onClick: (event) => {
          event.preventDefault();
          const pes = Number(pesInput.value);
          if (Number.isNaN(pes) || pes < 0) {
            Toast('El pes ha de ser ≥ 0.', 'warn');
            return;
          }
          try {
            appContext.store.addActivitat(assignatura.id, {
              data: dateInput.value ? new Date(dateInput.value) : new Date(),
              pesActivitat: pes,
            });
            Toast('Activitat creada.', 'success');
            dirtyFlags.view = true;
            scheduleRender('activitat-add');
            closeModal();
          } catch (error) {
            Toast(`Error en crear l\'activitat: ${error.message || error}`, 'error');
          }
        },
      },
    ]);
  });
  panel.appendChild(addBtn);
  if (!activitats?.length) {
    panel.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: 'Sense activitats registrades.',
      }),
    );
    return panel;
  }
  const table = DataTable({
    caption: 'Activitats de l\'assignatura',
    columns: [
      {
        id: 'data',
        label: 'Data',
        accessor: (row) => formatDate(row.data),
      },
      {
        id: 'pesActivitat',
        label: 'Pes',
        accessor: (row) => formatNumber(row.pesActivitat ?? 0, 2),
      },
      {
        id: 'descripcio',
        label: 'Descripció',
        accessor: (row) => row.descripcio || '—',
      },
    ],
    rows: activitats,
  });
  panel.appendChild(table);
  return panel;
}

function renderAlumnes(container, state) {
  clearElement(container);
  container.classList.add('space-y-6');
  const header = createElement('div', { className: 'flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between' });
  const searchInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200 sm:w-72',
    attrs: { type: 'search', placeholder: 'Cerca alumne…', value: uiState.alumnes.search },
  });
  searchInput.addEventListener('input', (event) => {
    uiState.alumnes.search = event.currentTarget.value;
    dirtyFlags.view = true;
    scheduleRender('alumnes-filter');
  });
  header.appendChild(searchInput);
  container.appendChild(header);

  const alumnes = state.alumnes?.allIds?.map((id) => state.alumnes.byId[id]) || [];
  const term = uiState.alumnes.search.trim().toLowerCase();
  const filtered = term
    ? alumnes.filter((alumne) => `${alumne.nom || ''} ${alumne.cognoms || ''}`.toLowerCase().includes(term))
    : alumnes;

  if (!filtered.length) {
    container.appendChild(
      createElement('p', {
        className: 'rounded-md border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-sm text-slate-600',
        text: alumnes.length ? 'Cap alumne coincideix amb la cerca.' : 'Encara no hi ha alumnes registrats.',
      }),
    );
    return;
  }

  const rows = filtered.map((alumne) => ({
    ...alumne,
    nomComplet: `${alumne.nom || ''} ${alumne.cognoms || ''}`.trim(),
  }));
  const table = DataTable({
    caption: 'Relació d\'alumnes',
    columns: [
      { id: 'nomComplet', label: 'Nom complet' },
      { id: 'email', label: 'Correu', accessor: (row) => row.email || '—' },
      {
        id: 'accions',
        label: 'Accions',
        accessor: (row) => {
          const btn = createElement('button', {
            className:
              'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-1 text-xs font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
            attrs: { type: 'button' },
            children: [
              createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'id-card' } }),
              createElement('span', { text: 'Obre fitxa' }),
            ],
          });
          btn.addEventListener('click', () => {
            uiState.fitxa.selectedAlumne = row.id;
            showView('view-fitxa');
            dirtyFlags.view = true;
            scheduleRender('fitxa-open');
          });
          return btn;
        },
      },
    ],
    rows,
  });
  container.appendChild(table);
}

function getTrimestresForAssignatura(state, assignaturaId) {
  if (!assignaturaId) return [];
  const calendari = getCalendariForAssignatura(state, assignaturaId);
  if (!calendari) return [];
  return (calendari.trimestres || []).map((trim) => ({
    id: trim.id,
    nom: trim.nom || trim.id,
    tInici: trim.tInici,
    tFi: trim.tFi,
  }));
}

function getAlumneName(state, alumneId) {
  const alumne = state.alumnes?.byId?.[alumneId];
  if (!alumne) return '—';
  return `${alumne.nom || ''} ${alumne.cognoms || ''}`.trim() || alumne.nom || alumne.id;
}

function getCAName(state, caId) {
  const ca = state.cas?.byId?.[caId];
  if (!ca) return caId;
  const ce = state.ces?.byId?.[ca.ceId];
  const ceIndex = ce ? ce.index : ca.ceIndex;
  return `CA${ca.textBetween || ''}${ceIndex}.${ca.index}`;
}

function findQualiFor(state, alumneId, caId) {
  const matches = state.avaluacionsComp?.allIds
    ?.map((id) => state.avaluacionsComp.byId[id])
    .filter((entry) => entry.alumneId === alumneId && entry.caId === caId);
  if (!matches?.length) return null;
  return matches[matches.length - 1].valorQuali;
}

function findActivitatForCA(state, caId) {
  const link = state.activitatCA?.allIds
    ?.map((id) => state.activitatCA.byId[id])
    .find((entry) => entry.caId === caId);
  return link?.activitatId || null;
}

function renderRubrica(container, state) {
  clearElement(container);
  container.classList.add('space-y-6');
  const assignatures = state.assignatures?.allIds
    ?.map((id) => state.assignatures.byId[id])
    .filter((assignatura) => assignatura.mode === 'competencial');
  if (!assignatures?.length) {
    container.appendChild(
      createElement('p', {
        className: 'rounded-md border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-sm text-slate-600',
        text: 'Cap assignatura està configurada en mode competencial.',
      }),
    );
    return;
  }

  if (!uiState.rubrica.selectedAssignatura || !assignatures.some((a) => a.id === uiState.rubrica.selectedAssignatura)) {
    uiState.rubrica.selectedAssignatura = assignatures[0].id;
  }

  const controls = createElement('div', {
    className: 'flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between',
  });
  const assignaturaSelect = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200 lg:w-80',
  });
  assignatures.forEach((assignatura) => {
    const option = createElement('option', { attrs: { value: assignatura.id }, text: assignatura.nom });
    if (assignatura.id === uiState.rubrica.selectedAssignatura) option.selected = true;
    assignaturaSelect.appendChild(option);
  });
  assignaturaSelect.addEventListener('change', (event) => {
    uiState.rubrica.selectedAssignatura = event.currentTarget.value;
    dirtyFlags.view = true;
    scheduleRender('rubrica-assignatura');
  });
  controls.appendChild(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Assignatura' }), assignaturaSelect],
    }),
  );

  const trimestres = getTrimestresForAssignatura(state, uiState.rubrica.selectedAssignatura);
  if (
    uiState.rubrica.selectedTrimestre &&
    !trimestres.some((trim) => trim.id === uiState.rubrica.selectedTrimestre)
  ) {
    uiState.rubrica.selectedTrimestre = null;
  }
  const trimestreSelect = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200 lg:w-48',
  });
  trimestreSelect.appendChild(createElement('option', { attrs: { value: '' }, text: 'Tots els trimestres' }));
  trimestres.forEach((trim) => {
    const labelRange = trim.tInici && trim.tFi ? `${formatDate(trim.tInici)} – ${formatDate(trim.tFi)}` : 'Dates pendents';
    const option = createElement('option', {
      attrs: { value: trim.id },
      text: `${trim.nom} · ${labelRange}`,
    });
    if (uiState.rubrica.selectedTrimestre === trim.id) option.selected = true;
    trimestreSelect.appendChild(option);
  });
  trimestreSelect.addEventListener('change', (event) => {
    const value = event.currentTarget.value;
    uiState.rubrica.selectedTrimestre = value || null;
    dirtyFlags.view = true;
    scheduleRender('rubrica-trimestre');
  });
  controls.appendChild(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Trimestre' }), trimestreSelect],
    }),
  );

  const exportButtons = createElement('div', { className: 'flex flex-wrap gap-3' });
  const exportCSV = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'file-down' } }), createElement('span', { text: 'Exporta CSV' })],
  });
  exportCSV.addEventListener('click', () => {
    appContext.actions.exportCSV_AvaluacionsCompetencial?.(
      uiState.rubrica.selectedAssignatura,
      uiState.rubrica.selectedTrimestre,
    );
  });
  const exportDOCX = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'file-text' } }), createElement('span', { text: 'Exporta DOCX' })],
  });
  exportDOCX.addEventListener('click', () => {
    appContext.actions.exportDOCX_ActaAssignatura?.(
      uiState.rubrica.selectedAssignatura,
      uiState.rubrica.selectedTrimestre,
    );
  });
  exportButtons.append(exportCSV, exportDOCX);
  controls.appendChild(exportButtons);

  container.appendChild(controls);

  let tableData = null;
  try {
    tableData = appContext.store.computeTaulaCompetencial(
      uiState.rubrica.selectedAssignatura,
      uiState.rubrica.selectedTrimestre,
    );
  } catch (error) {
    container.appendChild(
      createElement('p', {
        className: 'rounded-md border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700',
        text: `No s'ha pogut calcular la rúbrica: ${error.message || error}`,
      }),
    );
    return;
  }

  const columns = [
    {
      id: 'alumne',
      label: 'Alumne',
      accessor: (row) => getAlumneName(state, row.alumneId),
    },
  ];
  tableData.cas.forEach((caId) => {
    columns.push({ id: caId, label: getCAName(state, caId) });
  });
  const rows = tableData.alumnes.map((alumneId) => ({ alumneId }));

  const table = DataTable({
    caption: 'Rúbrica competencial',
    columns,
    rows,
  });
  const tbody = table.querySelector('tbody');
  Array.from(tbody.rows).forEach((tr, rowIndex) => {
    const alumneId = rows[rowIndex].alumneId;
    tableData.cas.forEach((caId, colIndex) => {
      const cell = tr.cells[colIndex + 1];
      const currentValue = findQualiFor(state, alumneId, caId) || tableData.values?.[alumneId]?.[caId]?.quali || 'NA';
      const select = createElement('select', {
        className:
          'w-full rounded-md border border-slate-300 px-2 py-1 text-xs shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      });
      ['NA', 'AS', 'AN', 'AE'].forEach((optionValue) => {
        const option = createElement('option', { attrs: { value: optionValue }, text: optionValue });
        if (optionValue === currentValue) option.selected = true;
        select.appendChild(option);
      });
      select.addEventListener('change', (event) => {
        const value = event.currentTarget.value;
        const activitatId = findActivitatForCA(state, caId);
        if (!activitatId) {
          Toast('Cap activitat està vinculada a aquest CA.', 'warn');
          event.currentTarget.value = currentValue;
          return;
        }
        try {
          appContext.store.registraAvaluacioComp({
            alumneId,
            activitatId,
            caId,
            valorQuali: value,
          });
          Toast('Valor registrat.', 'success');
          dirtyFlags.view = true;
          scheduleRender('rubrica-update');
        } catch (error) {
          Toast(`Error en registrar la rúbrica: ${error.message || error}`, 'error');
        }
      });
      clearElement(cell);
      cell.appendChild(select);
    });
  });
  container.appendChild(table);

  const summaryRow = createElement('div', {
    className: 'flex flex-wrap gap-3 rounded-md border border-slate-200 bg-slate-50 px-4 py-3 text-xs text-slate-600',
  });
  tableData.cas.forEach((caId) => {
    const values = tableData.alumnes.map((alumneId) => tableData.values?.[alumneId]?.[caId]?.valueNumRounded ?? 0);
    const average = values.reduce((acc, num) => acc + num, 0) / (values.length || 1);
    summaryRow.appendChild(
      createElement('span', {
        className: 'inline-flex items-center gap-2 rounded-md bg-white px-3 py-1 font-medium shadow-sm',
        text: `${getCAName(state, caId)} · Mitjana ${formatNumber(average, 2)}`,
      }),
    );
  });
  container.appendChild(summaryRow);
}

function renderCalendari(container, state) {
  clearElement(container);
  container.classList.add('space-y-6');
  const assignatures = state.assignatures?.allIds?.map((id) => state.assignatures.byId[id]) || [];
  if (!assignatures.length) {
    container.appendChild(
      createElement('p', {
        className: 'rounded-md border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-sm text-slate-600',
        text: 'Crea una assignatura per començar a configurar el calendari.',
      }),
    );
    return;
  }

  if (!uiState.calendari.selectedAssignatura || !assignatures.some((a) => a.id === uiState.calendari.selectedAssignatura)) {
    uiState.calendari.selectedAssignatura = assignatures[0].id;
  }

  const assignatura = assignatures.find((a) => a.id === uiState.calendari.selectedAssignatura);
  const calendari = getCalendariForAssignatura(state, uiState.calendari.selectedAssignatura);
  const trimestres = getTrimestresForAssignatura(state, uiState.calendari.selectedAssignatura);

  if (!uiState.calendari.exportSummary && assignatura) {
    uiState.calendari.exportSummary = `${assignatura.nom || assignatura.id} (classe)`;
  }
  if (!uiState.calendari.viewMonth) {
    const now = new Date();
    const firstDay = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1));
    uiState.calendari.viewMonth = toISODateLocal(firstDay);
  }

  const header = createElement('div', {
    className: 'flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between',
  });
  const titleWrapper = createElement('div', { className: 'flex items-center gap-3' });
  const title = createElement('h4', { className: 'text-lg font-semibold text-slate-900', text: 'Calendari i horaris' });
  titleWrapper.appendChild(title);
  if (uiState.calendari.pendingSave) {
    titleWrapper.appendChild(
      createElement('span', {
        className: 'text-xl leading-none text-emerald-500',
        text: '•',
        attrs: { role: 'status', 'aria-label': 'Canvis pendents de desar' },
      }),
    );
  }
  header.appendChild(titleWrapper);

  const assignaturaSelect = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200 lg:w-80',
  });
  assignatures.forEach((item) => {
    const option = createElement('option', { attrs: { value: item.id }, text: item.nom || item.id });
    if (item.id === uiState.calendari.selectedAssignatura) option.selected = true;
    assignaturaSelect.appendChild(option);
  });
  assignaturaSelect.addEventListener('change', (event) => {
    uiState.calendari.selectedAssignatura = event.currentTarget.value;
    uiState.calendari.simulatorResult = null;
    dirtyFlags.view = true;
    scheduleRender('calendar-assignatura');
  });
  header.appendChild(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Assignatura' }), assignaturaSelect],
    }),
  );
  container.appendChild(header);

  if (!assignatura) {
    container.appendChild(
      createElement('p', {
        className: 'rounded-md border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-700',
        text: 'Selecciona una assignatura per gestionar-ne el calendari.',
      }),
    );
    return;
  }

  container.appendChild(renderCalendariRangeSection(assignatura, calendari));
  container.appendChild(renderDiesSetmanalsSection(assignatura, calendari));
  container.appendChild(renderHorariVersionsSection(assignatura, calendari));
  container.appendChild(renderTrimestresSection(assignatura, calendari));
  container.appendChild(renderFestiusSection(assignatura, calendari));
  container.appendChild(renderExcepcionsSection(assignatura, calendari));
  container.appendChild(renderSimuladorSection(assignatura, calendari));
  container.appendChild(renderICSExportSection(assignatura, calendari, trimestres));
  container.appendChild(renderCalendariPreview(assignatura, calendari));
}

function renderCalendariRangeSection(assignatura, calendari) {
  const section = createElement('section', {
    className: 'space-y-3 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  section.appendChild(
    createElement('header', {
      className: 'flex items-center justify-between',
      children: [
        createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: 'Període del curs' }),
        createBadge(
          calendari.cursInici && calendari.cursFi
            ? `${formatDate(calendari.cursInici)} – ${formatDate(calendari.cursFi)}`
            : 'Sense dates',
          calendari.cursInici && calendari.cursFi ? 'emerald' : 'slate',
        ),
      ],
    }),
  );
  const grid = createElement('div', { className: 'grid gap-4 sm:grid-cols-2' });
  const errorMsg = createElement('p', {
    className: 'hidden text-xs font-medium text-rose-600',
    attrs: { 'aria-live': 'polite' },
  });
  const iniciInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: calendari.cursInici || '' },
  });
  const fiInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: calendari.cursFi || '' },
  });

  const showError = (message) => {
    if (!message) {
      errorMsg.classList.add('hidden');
      errorMsg.textContent = '';
      return;
    }
    errorMsg.textContent = message;
    errorMsg.classList.remove('hidden');
  };

  const handleChange = () => {
    const iniciISO = toISODateLocal(iniciInput.value) || null;
    const fiISO = toISODateLocal(fiInput.value) || null;
    if (iniciISO && fiISO && fiISO < iniciISO) {
      showError('La data de fi ha de ser posterior o igual a la data d\'inici.');
      return;
    }
    showError('');
    try {
      appContext.store.setCursRange(assignatura.id, { inici: iniciISO, fi: fiISO });
    } catch (error) {
      showError(error.message || 'No s\'ha pogut actualitzar el període.');
      iniciInput.value = calendari.cursInici || '';
      fiInput.value = calendari.cursFi || '';
    }
  };

  iniciInput.addEventListener('change', handleChange);
  fiInput.addEventListener('change', handleChange);

  grid.append(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Data d\'inici' }), iniciInput],
    }),
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Data de fi' }), fiInput],
    }),
  );
  section.append(grid, errorMsg);
  return section;
}

function renderDiesSetmanalsSection(assignatura, calendari) {
  const section = createElement('section', {
    className: 'space-y-3 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  section.appendChild(
    createElement('header', {
      className: 'flex items-center justify-between',
      children: [
        createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: 'Dies setmanals base' }),
        createBadge(formatDiesSetmanals(calendari.diesSetmanals), 'sky'),
      ],
    }),
  );
  const diesSeleccionats = new Set(calendari.diesSetmanals || []);
  const checklist = createElement('div', { className: 'flex flex-wrap gap-3' });
  WEEKDAY_ORDER.forEach((dia) => {
    const wrapper = createElement('label', {
      className: `inline-flex items-center gap-2 rounded-md border px-3 py-1 text-sm transition ${
        diesSeleccionats.has(dia) ? 'border-emerald-300 bg-emerald-50 text-emerald-700' : 'border-slate-200 bg-white text-slate-700'
      }`,
    });
    const checkbox = createElement('input', {
      className: 'h-4 w-4 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500',
      attrs: { type: 'checkbox', value: String(dia) },
    });
    checkbox.checked = diesSeleccionats.has(dia);
    checkbox.addEventListener('change', (event) => {
      const checked = event.currentTarget.checked;
      if (checked) {
        diesSeleccionats.add(dia);
      } else {
        diesSeleccionats.delete(dia);
      }
      try {
        const updated = Array.from(diesSeleccionats).sort((a, b) => WEEKDAY_ORDER.indexOf(a) - WEEKDAY_ORDER.indexOf(b));
        appContext.store.setDiesSetmanals(assignatura.id, updated);
      } catch (error) {
        event.currentTarget.checked = !checked;
        Toast(`Error en actualitzar els dies lectius: ${error.message || error}`, 'error');
      }
    });
    wrapper.append(checkbox, createElement('span', { text: WEEKDAY_LABEL[dia] }));
    checklist.appendChild(wrapper);
  });
  section.appendChild(checklist);
  section.appendChild(
    createElement('p', {
      className: 'text-xs text-slate-500',
      text: 'Aquest horari base s\'utilitza quan no hi ha cap versió d\'horari específica activa.',
    }),
  );
  return section;
}

function renderHorariVersionsSection(assignatura, calendari) {
  const section = createElement('section', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  const activeSchedule = typeof appContext.store.getActiveSchedule === 'function'
    ? appContext.store.getActiveSchedule(assignatura.id, new Date())
    : null;
  const activeVersioId =
    activeSchedule && typeof activeSchedule.versioId === 'string' && activeSchedule.versioId.trim().length
      ? activeSchedule.versioId
      : null;
  section.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between',
      children: [
        createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: "Versions d'horari" }),
        createBadge(
          activeVersioId
            ? `Activa avui: ${activeVersioId}`
            : 'En ús l\'horari base',
          activeVersioId ? 'emerald' : 'slate',
        ),
      ],
    }),
  );

  const versions = calendari.horariVersions || [];
  if (!versions.length) {
    section.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: 'Encara no hi ha versions d\'horari. Pots crear-ne una per canviar els dies lectius a partir d\'una data.',
      }),
    );
  } else {
    const table = createElement('table', { className: 'min-w-full divide-y divide-slate-200 overflow-hidden rounded-md border border-slate-200 text-sm' });
    const thead = createElement('thead', { className: 'bg-slate-50 text-xs uppercase tracking-wide text-slate-500' });
    thead.appendChild(
      createElement('tr', {
        children: [
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Versió' }),
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Activa des de' }),
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Dies' }),
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Estat' }),
          createElement('th', { className: 'px-4 py-2 text-right', text: 'Accions' }),
        ],
      }),
    );
    table.appendChild(thead);
    const tbody = createElement('tbody', { className: 'divide-y divide-slate-200 bg-white' });
    versions.forEach((versio, index) => {
      const row = createElement('tr', {
        className: activeVersioId === versio.id ? 'bg-emerald-50' : index % 2 === 0 ? 'bg-white' : 'bg-slate-50',
      });
      row.append(
        createElement('td', {
          className: 'px-4 py-3 font-medium text-slate-700',
          text: versio.id || `Versió ${index + 1}`,
        }),
        createElement('td', {
          className: 'px-4 py-3 text-slate-600',
          text: versio.effectiveFrom ? formatDate(versio.effectiveFrom) : '—',
        }),
        createElement('td', {
          className: 'px-4 py-3 text-slate-600',
          text: formatDiesSetmanals(versio.diesSetmanals),
        }),
        createElement('td', {
          className: 'px-4 py-3 text-slate-600',
          children: [
            activeVersioId === versio.id
              ? createBadge('Activa', 'emerald')
              : createElement('span', { className: 'text-xs text-slate-500', text: 'Inactiva' }),
          ],
        }),
        createElement('td', {
          className: 'px-4 py-3 text-right',
          children: [
            (() => {
              const button = createElement('button', {
                className:
                  'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-1 text-xs font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-500',
                attrs: { type: 'button' },
                text: 'Activa',
              });
              if (activeVersioId === versio.id) {
                button.disabled = true;
                button.classList.add('cursor-not-allowed', 'opacity-50');
              } else {
                button.addEventListener('click', () => {
                  try {
                    appContext.store.activateHorariVersio(assignatura.id, versio.id);
                    Toast('Versió activada correctament.', 'success');
                  } catch (error) {
                    Toast(`No s\'ha pogut activar la versió: ${error.message || error}`, 'error');
                  }
                });
              }
              return button;
            })(),
          ],
        }),
      );
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    section.appendChild(table);
  }

  const form = createElement('div', {
    className: 'space-y-3 rounded-md border border-dashed border-slate-300 bg-slate-50 p-4 text-sm text-slate-600',
  });
  form.appendChild(createElement('p', { className: 'font-medium text-slate-700', text: 'Nova versió d\'horari' }));
  const dateInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: uiState.calendari.newVersio.effectiveFrom || '' },
  });
  const diesSet = uiState.calendari.newVersio.dies instanceof Set
    ? uiState.calendari.newVersio.dies
    : new Set(uiState.calendari.newVersio.dies || []);
  uiState.calendari.newVersio.dies = diesSet;
  const diesChecklist = createElement('div', { className: 'flex flex-wrap gap-2' });
  WEEKDAY_ORDER.forEach((dia) => {
    const label = createElement('label', {
      className: `inline-flex items-center gap-2 rounded-md border px-2 py-1 text-xs transition ${
        diesSet.has(dia) ? 'border-emerald-300 bg-emerald-50 text-emerald-700' : 'border-slate-200 bg-white text-slate-600'
      }`,
    });
    const checkbox = createElement('input', {
      className: 'h-3.5 w-3.5 rounded border-slate-300 text-emerald-600 focus:ring-emerald-500',
      attrs: { type: 'checkbox', value: String(dia) },
    });
    checkbox.checked = diesSet.has(dia);
    checkbox.addEventListener('change', (event) => {
      if (event.currentTarget.checked) {
        diesSet.add(dia);
      } else {
        diesSet.delete(dia);
      }
    });
    label.append(checkbox, createElement('span', { text: WEEKDAY_LABEL[dia] }));
    diesChecklist.appendChild(label);
  });
  const errorMsg = createElement('p', {
    className: 'hidden text-xs font-medium text-rose-600',
    attrs: { 'aria-live': 'polite' },
  });
  const submitBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-semibold text-emerald-700 shadow-sm transition hover:bg-emerald-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-500',
    attrs: { type: 'button' },
    text: 'Afegeix versió',
  });
  submitBtn.addEventListener('click', () => {
    const effectiveISO = toISODateLocal(dateInput.value);
    if (!effectiveISO) {
      errorMsg.textContent = 'Cal indicar la data d\'entrada en vigor.';
      errorMsg.classList.remove('hidden');
      return;
    }
    const dies = Array.from(diesSet).sort((a, b) => WEEKDAY_ORDER.indexOf(a) - WEEKDAY_ORDER.indexOf(b));
    if (!dies.length) {
      errorMsg.textContent = 'La versió ha de contenir com a mínim un dia lectiu.';
      errorMsg.classList.remove('hidden');
      return;
    }
    errorMsg.classList.add('hidden');
    try {
      appContext.store.addHorariVersio(assignatura.id, { effectiveFrom: effectiveISO, diesSetmanals: dies });
      Toast('Versió creada.', 'success');
      uiState.calendari.newVersio = { effectiveFrom: '', dies: new Set([1, 2, 3, 4, 5]) };
      dirtyFlags.view = true;
      scheduleRender('calendar-versio');
    } catch (error) {
      errorMsg.textContent = error.message || 'No s\'ha pogut crear la versió.';
      errorMsg.classList.remove('hidden');
    }
  });
  form.append(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Entrada en vigor' }), dateInput],
    }),
    createElement('div', {
      className: 'space-y-2',
      children: [createElement('span', { className: 'text-xs font-semibold uppercase text-slate-500', text: 'Dies lectius' }), diesChecklist],
    }),
    errorMsg,
    submitBtn,
  );
  section.appendChild(form);
  return section;
}

function renderTrimestresSection(assignatura, calendari) {
  const section = createElement('section', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  section.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between',
      children: [
        createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: 'Trimestres' }),
        createBadge(`${(calendari.trimestres || []).length} trimestres`, 'sky'),
      ],
    }),
  );

  const trimestres = calendari.trimestres || [];
  if (!trimestres.length) {
    section.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: 'Encara no s\'ha definit cap trimestre. Afegeix-ne per facilitar els filtres d\'avaluació i exportació.',
      }),
    );
  } else {
    const table = createElement('table', { className: 'min-w-full divide-y divide-slate-200 overflow-hidden rounded-md border border-slate-200 text-sm' });
    const thead = createElement('thead', { className: 'bg-slate-50 text-xs uppercase tracking-wide text-slate-500' });
    thead.appendChild(
      createElement('tr', {
        children: [
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Nom' }),
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Inici' }),
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Fi' }),
          createElement('th', { className: 'px-4 py-2 text-right', text: 'Accions' }),
        ],
      }),
    );
    table.appendChild(thead);
    const tbody = createElement('tbody', { className: 'divide-y divide-slate-200 bg-white' });
    trimestres.forEach((trimestre, index) => {
      const row = createElement('tr', { className: index % 2 === 0 ? 'bg-white' : 'bg-slate-50' });
      const startInput = createElement('input', {
        className:
          'w-full rounded-md border border-slate-300 px-2 py-1 text-xs shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
        attrs: { type: 'date', value: trimestre.tInici || '' },
      });
      const endInput = createElement('input', {
        className:
          'w-full rounded-md border border-slate-300 px-2 py-1 text-xs shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
        attrs: { type: 'date', value: trimestre.tFi || '' },
      });
      const errorMsg = createElement('p', {
        className: 'hidden text-xs font-medium text-rose-600',
        attrs: { 'aria-live': 'polite' },
      });

      const applyChange = (newInici, newFi) => {
        const message = validateTrimestreRange(trimestres, trimestre.id, newInici, newFi);
        if (message) {
          errorMsg.textContent = message;
          errorMsg.classList.remove('hidden');
          startInput.value = trimestre.tInici || '';
          endInput.value = trimestre.tFi || '';
          return;
        }
        errorMsg.classList.add('hidden');
        try {
          appContext.store.addTrimestre(assignatura.id, {
            id: trimestre.id,
            tInici: newInici,
            tFi: newFi,
            nom: trimestre.nom,
          });
          Toast('Trimestre actualitzat.', 'success');
        } catch (error) {
          errorMsg.textContent = error.message || 'No s\'ha pogut actualitzar el trimestre.';
          errorMsg.classList.remove('hidden');
          startInput.value = trimestre.tInici || '';
          endInput.value = trimestre.tFi || '';
        }
      };

      startInput.addEventListener('change', (event) => {
        const newInici = toISODateLocal(event.currentTarget.value) || '';
        const newFi = toISODateLocal(endInput.value) || '';
        if (!newInici) {
          event.currentTarget.value = trimestre.tInici || '';
          return;
        }
        applyChange(newInici, newFi || trimestre.tFi || '');
      });
      endInput.addEventListener('change', (event) => {
        const newInici = toISODateLocal(startInput.value) || '';
        const newFi = toISODateLocal(event.currentTarget.value) || '';
        if (!newFi) {
          event.currentTarget.value = trimestre.tFi || '';
          return;
        }
        applyChange(newInici || trimestre.tInici || '', newFi);
      });

      row.append(
        createElement('td', { className: 'px-4 py-3 font-medium text-slate-700', text: trimestre.nom || trimestre.id }),
        createElement('td', { className: 'px-4 py-3 text-slate-600', children: [startInput, errorMsg] }),
        createElement('td', { className: 'px-4 py-3 text-slate-600', children: [endInput] }),
        createElement('td', {
          className: 'px-4 py-3 text-right',
          children: [
            (() => {
              const button = createElement('button', {
                className:
                  'inline-flex items-center gap-2 rounded-md border border-rose-200 bg-rose-50 px-3 py-1 text-xs font-semibold text-rose-700 shadow-sm transition hover:bg-rose-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-rose-500',
                attrs: { type: 'button' },
                children: [createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'trash' } }), createElement('span', { text: 'Elimina' })],
              });
              button.addEventListener('click', () => {
                if (!confirm('Vols eliminar aquest trimestre?')) return;
                try {
                  appContext.store.removeTrimestre(assignatura.id, trimestre.id);
                  Toast('Trimestre eliminat.', 'success');
                } catch (error) {
                  Toast(`No s\'ha pogut eliminar: ${error.message || error}`, 'error');
                }
              });
              return button;
            })(),
          ],
        }),
      );
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    section.appendChild(table);
  }

  const form = createElement('div', {
    className: 'space-y-3 rounded-md border border-dashed border-slate-300 bg-slate-50 p-4 text-sm text-slate-600',
  });
  form.appendChild(createElement('p', { className: 'font-medium text-slate-700', text: 'Afegeix trimestre' }));
  const nomInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'text', placeholder: 'Nom del trimestre (opcional)' },
  });
  const iniciInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: uiState.calendari.newTrimestre.tInici || '' },
  });
  const fiInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: uiState.calendari.newTrimestre.tFi || '' },
  });
  const errorMsg = createElement('p', {
    className: 'hidden text-xs font-medium text-rose-600',
    attrs: { 'aria-live': 'polite' },
  });

  iniciInput.addEventListener('change', (event) => {
    uiState.calendari.newTrimestre.tInici = toISODateLocal(event.currentTarget.value) || '';
  });
  fiInput.addEventListener('change', (event) => {
    uiState.calendari.newTrimestre.tFi = toISODateLocal(event.currentTarget.value) || '';
  });

  const submitBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-semibold text-emerald-700 shadow-sm transition hover:bg-emerald-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-500',
    attrs: { type: 'button' },
    text: 'Afegeix',
  });
  submitBtn.addEventListener('click', () => {
    const { tInici, tFi } = uiState.calendari.newTrimestre;
    const iniciISO = toISODateLocal(tInici) || tInici;
    const fiISO = toISODateLocal(tFi) || tFi;
    const message = validateTrimestreRange(trimestres, null, iniciISO, fiISO);
    if (message) {
      errorMsg.textContent = message;
      errorMsg.classList.remove('hidden');
      return;
    }
    if (!iniciISO || !fiISO) {
      errorMsg.textContent = 'Cal indicar totes dues dates.';
      errorMsg.classList.remove('hidden');
      return;
    }
    errorMsg.classList.add('hidden');
    try {
      appContext.store.addTrimestre(assignatura.id, {
        tInici: iniciISO,
        tFi: fiISO,
        nom: nomInput.value.trim() || undefined,
      });
      Toast('Trimestre afegit.', 'success');
      uiState.calendari.newTrimestre = { tInici: '', tFi: '' };
      nomInput.value = '';
      iniciInput.value = '';
      fiInput.value = '';
      dirtyFlags.view = true;
      scheduleRender('calendar-trimestre');
    } catch (error) {
      errorMsg.textContent = error.message || 'No s\'ha pogut afegir el trimestre.';
      errorMsg.classList.remove('hidden');
    }
  });

  form.append(
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Nom' }), nomInput] }),
    createElement('div', {
      className: 'grid gap-3 sm:grid-cols-2',
      children: [
        createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Data d\'inici' }), iniciInput] }),
        createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Data de fi' }), fiInput] }),
      ],
    }),
    errorMsg,
    submitBtn,
  );
  section.appendChild(form);
  return section;
}

function renderCalendariEntriesSection(assignatura, calendari, type) {
  const isFestius = type === 'festius';
  const entries = isFestius ? calendari.festius || [] : calendari.excepcions || [];
  const section = createElement('section', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  section.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between',
      children: [
        createElement('h5', {
          className: 'text-sm font-semibold uppercase tracking-wide text-slate-500',
          text: isFestius ? 'Festius' : 'Excepcions (dies no lectius)',
        }),
        createBadge(`${entries.length} registres`, isFestius ? 'amber' : 'rose'),
      ],
    }),
  );

  section.appendChild(createImportExportControls(assignatura, type));

  if (!entries.length) {
    section.appendChild(
      createElement('p', {
        className: 'text-xs text-slate-500',
        text: isFestius
          ? 'No hi ha festius registrats. Pots importar-los des d\'un CSV o afegir-los manualment.'
          : 'No hi ha excepcions puntuals registrades.',
      }),
    );
  } else {
    const table = createElement('table', { className: 'min-w-full divide-y divide-slate-200 overflow-hidden rounded-md border border-slate-200 text-sm' });
    const thead = createElement('thead', { className: 'bg-slate-50 text-xs uppercase tracking-wide text-slate-500' });
    thead.appendChild(
      createElement('tr', {
        children: [
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Data' }),
          createElement('th', { className: 'px-4 py-2 text-left', text: 'Motiu' }),
          createElement('th', { className: 'px-4 py-2 text-right', text: 'Accions' }),
        ],
      }),
    );
    table.appendChild(thead);
    const tbody = createElement('tbody', { className: 'divide-y divide-slate-200 bg-white' });
    entries.forEach((item, index) => {
      const row = createElement('tr', { className: index % 2 === 0 ? 'bg-white' : 'bg-slate-50' });
      const dateInput = createElement('input', {
        className:
          'w-full rounded-md border border-slate-300 px-2 py-1 text-xs shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
        attrs: { type: 'date', value: item.dataISO || '' },
      });
      const motiuInput = createElement('input', {
        className:
          'w-full rounded-md border border-slate-300 px-2 py-1 text-xs shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
        attrs: { type: 'text', value: item.motiu || '', placeholder: 'Motiu (opcional)' },
      });
      dateInput.addEventListener('change', (event) => {
        const newISO = toISODateLocal(event.currentTarget.value);
        if (!newISO) {
          event.currentTarget.value = item.dataISO || '';
          Toast('Introdueix una data vàlida.', 'warn');
          return;
        }
        try {
          if (isFestius) {
            appContext.store.removeFestius(assignatura.id, [item.dataISO]);
            appContext.store.addFestius(assignatura.id, [{ dataISO: newISO, motiu: motiuInput.value }]);
          } else {
            appContext.store.removeExcepcio(assignatura.id, item.dataISO);
            appContext.store.addExcepcions(assignatura.id, [{ dataISO: newISO, motiu: motiuInput.value }]);
          }
          Toast('Data actualitzada.', 'success');
        } catch (error) {
          Toast(`No s\'ha pogut actualitzar la data: ${error.message || error}`, 'error');
          event.currentTarget.value = item.dataISO || '';
        }
      });
      motiuInput.addEventListener('change', (event) => {
        try {
          const motiu = event.currentTarget.value;
          if (isFestius) {
            appContext.store.addFestius(assignatura.id, [{ dataISO: item.dataISO, motiu }]);
          } else {
            appContext.store.addExcepcions(assignatura.id, [{ dataISO: item.dataISO, motiu }]);
          }
          Toast('Motiu actualitzat.', 'success');
        } catch (error) {
          Toast(`No s\'ha pogut actualitzar el motiu: ${error.message || error}`, 'error');
          event.currentTarget.value = item.motiu || '';
        }
      });
      const removeBtn = createElement('button', {
        className:
          'inline-flex items-center gap-2 rounded-md border border-rose-200 bg-rose-50 px-3 py-1 text-xs font-semibold text-rose-700 shadow-sm transition hover:bg-rose-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-rose-500',
        attrs: { type: 'button' },
        children: [createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'trash-2' } }), createElement('span', { text: 'Elimina' })],
      });
      removeBtn.addEventListener('click', () => {
        if (!confirm('Vols eliminar aquesta entrada?')) return;
        try {
          if (isFestius) {
            appContext.store.removeFestius(assignatura.id, [item.dataISO]);
          } else {
            appContext.store.removeExcepcio(assignatura.id, item.dataISO);
          }
          Toast('Entrada eliminada.', 'success');
        } catch (error) {
          Toast(`No s\'ha pogut eliminar: ${error.message || error}`, 'error');
        }
      });
      row.append(
        createElement('td', { className: 'px-4 py-3 text-slate-600', children: [dateInput] }),
        createElement('td', { className: 'px-4 py-3 text-slate-600', children: [motiuInput] }),
        createElement('td', { className: 'px-4 py-3 text-right', children: [removeBtn] }),
      );
      tbody.appendChild(row);
    });
    table.appendChild(tbody);
    section.appendChild(table);
  }

  const form = createElement('div', {
    className: 'space-y-3 rounded-md border border-dashed border-slate-300 bg-slate-50 p-4 text-sm text-slate-600',
  });
  form.appendChild(
    createElement('p', {
      className: 'font-medium text-slate-700',
      text: isFestius ? 'Afegeix festiu manualment' : 'Afegeix excepció puntual',
    }),
  );
  const dateInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date' },
  });
  const motiuInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'text', placeholder: 'Motiu' },
  });
  const errorMsg = createElement('p', {
    className: 'hidden text-xs font-medium text-rose-600',
    attrs: { 'aria-live': 'polite' },
  });
  const submitBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-semibold text-emerald-700 shadow-sm transition hover:bg-emerald-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-500',
    attrs: { type: 'button' },
    text: 'Afegeix',
  });
  submitBtn.addEventListener('click', () => {
    const dateISO = toISODateLocal(dateInput.value);
    if (!dateISO) {
      errorMsg.textContent = 'Cal indicar una data vàlida.';
      errorMsg.classList.remove('hidden');
      return;
    }
    errorMsg.classList.add('hidden');
    try {
      if (isFestius) {
        appContext.store.addFestius(assignatura.id, [{ dataISO: dateISO, motiu: motiuInput.value }]);
      } else {
        appContext.store.addExcepcions(assignatura.id, [{ dataISO: dateISO, motiu: motiuInput.value }]);
      }
      Toast('Entrada afegida.', 'success');
      dateInput.value = '';
      motiuInput.value = '';
    } catch (error) {
      errorMsg.textContent = error.message || 'No s\'ha pogut afegir l\'entrada.';
      errorMsg.classList.remove('hidden');
    }
  });

  form.append(
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Data' }), dateInput] }),
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Motiu' }), motiuInput] }),
    errorMsg,
    submitBtn,
  );
  section.appendChild(form);
  return section;
}

function createImportExportControls(assignatura, type) {
  const wrapper = createElement('div', {
    className: 'flex flex-col gap-3 rounded-md border border-slate-200 bg-slate-50 p-3 text-sm text-slate-600 sm:flex-row sm:items-end sm:justify-between',
  });
  const radioGroup = createElement('div', { className: 'flex items-center gap-3 text-xs text-slate-600' });
  ['festius', 'excepcions'].forEach((target) => {
    const id = `import-${target}-${type}`;
    const label = createElement('label', { className: 'inline-flex items-center gap-2' });
    const radio = createElement('input', {
      className: 'h-3.5 w-3.5 border-slate-300 text-emerald-600 focus:ring-emerald-500',
      attrs: { type: 'radio', name: `import-target-${type}`, id, value: target },
    });
    radio.checked = uiState.calendari.importTarget === target;
    radio.addEventListener('change', () => {
      uiState.calendari.importTarget = target;
      if (typeof appContext.actions.setCalendariImportTarget === 'function') {
        appContext.actions.setCalendariImportTarget(target);
      }
    });
    label.append(radio, createElement('span', { text: target === 'festius' ? 'Festius' : 'Excepcions' }));
    radioGroup.appendChild(label);
  });

  const fileInput = createElement('input', {
    className:
      'w-full max-w-xs rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'file', accept: '.csv,text/csv' },
  });
  const importBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-3 py-2 text-xs font-semibold text-emerald-700 shadow-sm transition hover:bg-emerald-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-500',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'file-input' } }), createElement('span', { text: 'Importa CSV' })],
  });
  importBtn.addEventListener('click', async () => {
    const file = fileInput.files?.[0];
    if (!file) {
      Toast('Selecciona un fitxer CSV.', 'warn');
      return;
    }
    if (typeof appContext.actions.setCalendariImportTarget === 'function') {
      appContext.actions.setCalendariImportTarget(uiState.calendari.importTarget);
    }
    try {
      const result = await appContext.actions.importFestiusCSV?.(assignatura.id, file);
      const tipus = result?.tipus || uiState.calendari.importTarget;
      Toast(`Importats ${result?.importats || 0} registres a ${tipus}.`, 'success');
      fileInput.value = '';
    } catch (error) {
      Toast(`Error en importar: ${error.message || error}`, 'error');
    }
  });

  const exportBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 text-xs font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'file-down' } }), createElement('span', { text: 'Exporta CSV' })],
  });
  exportBtn.addEventListener('click', () => {
    try {
      const blob = type === 'festius'
        ? appContext.actions.exportFestiusCSV?.(assignatura.id)
        : appContext.actions.exportExcepcionsCSV?.(assignatura.id);
      if (!blob) {
        Toast('No hi ha dades per exportar.', 'info');
        return;
      }
      const filename = `${assignatura.nom || assignatura.id}-${type}.csv`;
      downloadBlob(blob, filename);
      Toast('Exportació preparada.', 'success');
    } catch (error) {
      Toast(`Error en exportar: ${error.message || error}`, 'error');
    }
  });

  wrapper.append(radioGroup, fileInput, importBtn, exportBtn);
  return wrapper;
}

function renderFestiusSection(assignatura, calendari) {
  return renderCalendariEntriesSection(assignatura, calendari, 'festius');
}

function renderExcepcionsSection(assignatura, calendari) {
  return renderCalendariEntriesSection(assignatura, calendari, 'excepcions');
}

function renderSimuladorSection(assignatura, calendari) {
  const section = createElement('section', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  section.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between',
      children: [
        createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: 'Simulador “Què passa si…?”' }),
        createBadge('Consulta lectiva', 'emerald'),
      ],
    }),
  );

  const form = createElement('div', { className: 'flex flex-col gap-3 sm:flex-row sm:items-end' });
  const dateInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200 sm:w-60',
    attrs: { type: 'date', value: uiState.calendari.simulatorDate || '' },
  });
  const errorMsg = createElement('p', { className: 'hidden text-xs font-medium text-rose-600', attrs: { 'aria-live': 'polite' } });
  const button = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-emerald-200 bg-emerald-50 px-4 py-2 text-sm font-semibold text-emerald-700 shadow-sm transition hover:bg-emerald-100 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-emerald-500',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'play' } }), createElement('span', { text: 'Simula' })],
  });
  button.addEventListener('click', () => {
    const iso = toISODateLocal(dateInput.value);
    if (!iso) {
      errorMsg.textContent = 'Cal indicar una data vàlida.';
      errorMsg.classList.remove('hidden');
      return;
    }
    errorMsg.classList.add('hidden');
    try {
      const result = appContext.actions.simulate?.(assignatura.id, new Date(`${iso}T00:00:00`)) || {};
      uiState.calendari.simulatorDate = iso;
      uiState.calendari.simulatorResult = { ...result, dateISO: iso };
      dirtyFlags.view = true;
      scheduleRender('calendar-simulator');
    } catch (error) {
      errorMsg.textContent = error.message || 'No s\'ha pogut calcular la simulació.';
      errorMsg.classList.remove('hidden');
    }
  });
  form.append(
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Data a simular' }), dateInput] }),
    button,
  );
  section.append(form, errorMsg);

  const resultBox = createElement('div', {
    className: 'rounded-md border border-dashed border-slate-300 bg-slate-50 p-4 text-sm text-slate-600',
    attrs: { 'aria-live': 'polite' },
  });
  if (!uiState.calendari.simulatorResult) {
    resultBox.textContent = 'Introdueix una data i prem “Simula” per veure si hi ha classe, la versió d\'horari aplicable i el motiu en cas contrari.';
  } else {
    const { lectiu, versioId, motiu, dateISO } = uiState.calendari.simulatorResult;
    const versio = versioId ? (calendari.horariVersions || []).find((v) => v.id === versioId) : null;
    if (lectiu) {
      resultBox.classList.remove('border-dashed', 'border-slate-300', 'bg-slate-50', 'text-slate-600');
      resultBox.classList.add('border-emerald-200', 'bg-emerald-50', 'text-emerald-700');
      resultBox.append(
        createElement('p', { className: 'font-semibold', text: `${formatDate(dateISO)} és lectiu.` }),
        createElement('p', { className: 'text-sm', text: versio ? `Versió d'horari aplicable: ${versio.id} (${formatDiesSetmanals(versio.diesSetmanals)})` : 'S\'utilitza l\'horari base.' }),
      );
    } else {
      resultBox.classList.remove('border-dashed', 'border-slate-300', 'bg-slate-50', 'text-slate-600');
      resultBox.classList.add('border-rose-200', 'bg-rose-50', 'text-rose-700');
      resultBox.append(
        createElement('p', { className: 'font-semibold', text: `${formatDate(dateISO)} no és lectiu.` }),
        createElement('p', { className: 'text-sm', text: motiu || 'Motiu desconegut.' }),
      );
      if (versioId) {
        resultBox.append(createElement('p', { className: 'text-xs', text: `Versió prevista: ${versioId}.` }));
      }
    }
  }
  section.appendChild(resultBox);
  return section;
}

function renderICSExportSection(assignatura, calendari, trimestres) {
  const section = createElement('section', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  section.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between',
      children: [
        createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: 'Exporta sessions (ICS)' }),
        createBadge('Calendari digital', 'sky'),
      ],
    }),
  );

  const form = createElement('div', { className: 'grid gap-3 md:grid-cols-2 lg:grid-cols-4' });
  const select = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
  });
  const hasCursRange = calendari.cursInici && calendari.cursFi;
  const options = [];
  if (hasCursRange) {
    options.push({ value: 'curs', label: 'Període del curs' });
  }
  trimestres.forEach((trim) => {
    options.push({ value: `trimestre:${trim.id}`, label: `${trim.nom} · ${trim.tInici && trim.tFi ? `${formatDate(trim.tInici)} – ${formatDate(trim.tFi)}` : 'Dates pendents'}` });
  });
  options.push({ value: 'personalitzat', label: 'Dates personalitzades' });
  if (!options.some((opt) => opt.value === uiState.calendari.exportMode)) {
    uiState.calendari.exportMode = options[0]?.value || 'personalitzat';
  }
  options.forEach((opt) => {
    const option = createElement('option', { attrs: { value: opt.value }, text: opt.label });
    if (opt.value === uiState.calendari.exportMode) option.selected = true;
    select.appendChild(option);
  });
  select.addEventListener('change', (event) => {
    uiState.calendari.exportMode = event.currentTarget.value;
    dirtyFlags.view = true;
    scheduleRender('calendar-export');
  });

  const fromInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: uiState.calendari.exportFrom || '' },
  });
  const toInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'date', value: uiState.calendari.exportTo || '' },
  });
  fromInput.addEventListener('change', (event) => {
    uiState.calendari.exportFrom = toISODateLocal(event.currentTarget.value) || '';
  });
  toInput.addEventListener('change', (event) => {
    uiState.calendari.exportTo = toISODateLocal(event.currentTarget.value) || '';
  });
  const isCustom = uiState.calendari.exportMode === 'personalitzat';
  fromInput.disabled = !isCustom;
  toInput.disabled = !isCustom;

  const summaryInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'text', value: uiState.calendari.exportSummary || '', placeholder: 'Títol de l\'esdeveniment' },
  });
  summaryInput.addEventListener('input', (event) => {
    uiState.calendari.exportSummary = event.currentTarget.value;
  });
  const locationInput = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-emerald-500 focus:outline-none focus:ring-2 focus:ring-emerald-200',
    attrs: { type: 'text', value: uiState.calendari.exportLocation || '', placeholder: 'Ubicació (opcional)' },
  });
  locationInput.addEventListener('input', (event) => {
    uiState.calendari.exportLocation = event.currentTarget.value;
  });

  form.append(
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Rang a exportar' }), select] }),
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Data inicial' }), fromInput] }),
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Data final' }), toInput] }),
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Títol' }), summaryInput] }),
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Ubicació' }), locationInput] }),
  );
  section.appendChild(form);

  const errorMsg = createElement('p', { className: 'hidden text-xs font-medium text-rose-600', attrs: { 'aria-live': 'polite' } });
  const exportBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-4 py-2 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'calendar' } }), createElement('span', { text: 'Exporta ICS' })],
  });
  exportBtn.addEventListener('click', () => {
    let fromISO = null;
    let toISO = null;
    if (uiState.calendari.exportMode === 'curs' && hasCursRange) {
      fromISO = calendari.cursInici;
      toISO = calendari.cursFi;
    } else if (uiState.calendari.exportMode.startsWith('trimestre:')) {
      const id = uiState.calendari.exportMode.split(':')[1];
      const trimestre = trimestres.find((t) => t.id === id);
      fromISO = trimestre?.tInici || null;
      toISO = trimestre?.tFi || null;
    } else if (uiState.calendari.exportMode === 'personalitzat') {
      fromISO = uiState.calendari.exportFrom || null;
      toISO = uiState.calendari.exportTo || null;
    }
    if (!fromISO || !toISO) {
      errorMsg.textContent = 'Cal definir un interval complet per exportar.';
      errorMsg.classList.remove('hidden');
      return;
    }
    if (toISO < fromISO) {
      errorMsg.textContent = 'La data final ha de ser posterior o igual a la inicial.';
      errorMsg.classList.remove('hidden');
      return;
    }
    errorMsg.classList.add('hidden');
    try {
      const blob = appContext.actions.exportSessionsICS?.(assignatura.id, {
        from: fromISO,
        to: toISO,
        summary: summaryInput.value || undefined,
        location: locationInput.value || undefined,
      });
      if (!blob) {
        Toast('No s\'han trobat sessions en el període indicat.', 'info');
        return;
      }
      const filename = `${assignatura.nom || assignatura.id}-${fromISO}-${toISO}.ics`;
      downloadBlob(blob, filename);
      Toast('Arxiu ICS generat.', 'success');
    } catch (error) {
      errorMsg.textContent = error.message || 'No s\'ha pogut generar l\'ICS.';
      errorMsg.classList.remove('hidden');
    }
  });

  section.append(errorMsg, exportBtn);
  return section;
}

function renderCalendariPreview(assignatura, calendari) {
  const section = createElement('section', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  const baseISO = uiState.calendari.viewMonth || toISODateLocal(new Date());
  const [yearStr, monthStr] = baseISO.split('-');
  const baseDate = new Date(Date.UTC(Number(yearStr), Number(monthStr) - 1, 1));
  const mondayOffset = (baseDate.getUTCDay() + 6) % 7;
  const start = new Date(baseDate);
  start.setUTCDate(start.getUTCDate() - mondayOffset);
  const weeksToRender = 6;
  const end = new Date(start);
  end.setUTCDate(end.getUTCDate() + weeksToRender * 7 - 1);
  const sessions = typeof appContext.store.listSessions === 'function'
    ? appContext.store.listSessions(assignatura.id, { from: new Date(start), to: new Date(end) })
    : [];
  const sessionSet = new Set((sessions || []).map((s) => s.dateISO));
  const festiuMap = new Map((calendari.festius || []).map((festiu) => [festiu.dataISO, festiu.motiu || 'Festiu']));
  const excepcioMap = new Map((calendari.excepcions || []).map((ex) => [ex.dataISO, ex.motiu || 'Excepció']));

  const header = createElement('div', { className: 'flex items-center justify-between gap-3' });
  const formatter = new Intl.DateTimeFormat('ca-ES', { month: 'long', year: 'numeric' });
  header.appendChild(createElement('h5', { className: 'text-sm font-semibold uppercase tracking-wide text-slate-500', text: `Calendari ${formatter.format(baseDate)}` }));
  const nav = createElement('div', { className: 'flex items-center gap-2' });
  const prevBtn = createElement('button', {
    className:
      'inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'chevron-left' } }), createElement('span', { text: 'Anterior' })],
  });
  prevBtn.addEventListener('click', () => {
    const prev = new Date(baseDate);
    prev.setUTCMonth(prev.getUTCMonth() - 1);
    uiState.calendari.viewMonth = toISODateLocal(prev);
    dirtyFlags.view = true;
    scheduleRender('calendar-prev');
  });
  const nextBtn = createElement('button', {
    className:
      'inline-flex items-center gap-1 rounded-md border border-slate-200 bg-white px-2 py-1 text-xs font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { text: 'Següent' }), createElement('span', { className: 'lucide h-3.5 w-3.5', attrs: { 'data-lucide': 'chevron-right' } })],
  });
  nextBtn.addEventListener('click', () => {
    const next = new Date(baseDate);
    next.setUTCMonth(next.getUTCMonth() + 1);
    uiState.calendari.viewMonth = toISODateLocal(next);
    dirtyFlags.view = true;
    scheduleRender('calendar-next');
  });
  nav.append(prevBtn, nextBtn);
  header.appendChild(nav);
  section.appendChild(header);

  const table = createElement('table', { className: 'w-full table-fixed border-collapse overflow-hidden rounded-md border border-slate-200 text-sm' });
  const thead = createElement('thead', { className: 'bg-slate-50 text-xs uppercase tracking-wide text-slate-500' });
  const headRow = createElement('tr', {});
  WEEKDAY_ORDER.slice(0, 5).forEach((dia) => {
    headRow.appendChild(createElement('th', { className: 'px-3 py-2 text-left', text: WEEKDAY_LABEL[dia] }));
  });
  thead.appendChild(headRow);
  table.appendChild(thead);

  const tbody = createElement('tbody', {});
  const monthIndex = baseDate.getUTCMonth();
  for (let week = 0; week < weeksToRender; week += 1) {
    const row = createElement('tr', { className: week % 2 === 0 ? 'bg-white' : 'bg-slate-50' });
    for (let dayIndex = 0; dayIndex < 5; dayIndex += 1) {
      const cellDate = new Date(start);
      cellDate.setUTCDate(start.getUTCDate() + week * 7 + dayIndex);
      const iso = cellDate.toISOString().slice(0, 10);
      const isCurrentMonth = cellDate.getUTCMonth() === monthIndex;
      const festiuMotiu = festiuMap.get(iso) || excepcioMap.get(iso);
      const isLectiu = sessionSet.has(iso);
      const cell = createElement('td', {
        className: `h-24 border border-slate-200 align-top p-2 text-xs ${
          isLectiu
            ? 'bg-emerald-50 text-emerald-700'
            : festiuMotiu
              ? 'bg-rose-50 text-rose-700'
              : 'text-slate-600'
        } ${isCurrentMonth ? '' : 'opacity-60'}`,
      });
      cell.appendChild(createElement('div', { className: 'flex items-center justify-between text-xs font-semibold', children: [createElement('span', { text: cellDate.getUTCDate() })] }));
      if (isLectiu) {
        cell.appendChild(createElement('p', { className: 'mt-2 text-[11px]', text: 'Lectiu' }));
      } else if (festiuMotiu) {
        cell.appendChild(createElement('p', { className: 'mt-2 text-[11px]', text: festiuMotiu }));
      } else {
        cell.appendChild(createElement('p', { className: 'mt-2 text-[11px] text-slate-500', text: '—' }));
      }
      row.appendChild(cell);
    }
    tbody.appendChild(row);
  }
  table.appendChild(tbody);
  section.appendChild(table);

  section.appendChild(
    createElement('p', {
      className: 'text-xs text-slate-500',
      text: 'El calendari mostra únicament de dilluns a divendres. Les sessions lectives apareixen destacades en verd i els dies no lectius en rosa.',
    }),
  );
  return section;
}

function renderFitxa(container, state) {
  clearElement(container);
  container.classList.add('grid', 'gap-6', 'lg:grid-cols-[20rem,1fr]');
  const llista = createElement('div', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-4 shadow-sm',
  });
  const search = createElement('input', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
    attrs: { type: 'search', placeholder: 'Cerca alumne…', value: uiState.fitxa.filter },
  });
  search.addEventListener('input', (event) => {
    uiState.fitxa.filter = event.currentTarget.value;
    dirtyFlags.view = true;
    scheduleRender('fitxa-filter');
  });
  llista.appendChild(search);
  const alumnes = state.alumnes?.allIds?.map((id) => state.alumnes.byId[id]) || [];
  const term = uiState.fitxa.filter.trim().toLowerCase();
  const filtered = term
    ? alumnes.filter((alumne) => `${alumne.nom || ''} ${alumne.cognoms || ''}`.toLowerCase().includes(term))
    : alumnes;
  const list = createElement('ul', { className: 'space-y-2 max-h-[70vh] overflow-y-auto' });
  filtered.forEach((alumne) => {
    const selected = uiState.fitxa.selectedAlumne === alumne.id;
    const item = createElement('li', {
      className: `rounded-md border px-3 py-2 text-sm transition ${
        selected ? 'border-slate-500 bg-slate-100' : 'border-slate-200 bg-white'
      }`,
    });
    const button = createElement('button', {
      className: 'flex w-full items-center justify-between text-left',
      attrs: { type: 'button' },
      children: [
        createElement('span', { className: 'font-medium text-slate-800', text: getAlumneName(state, alumne.id) }),
        createElement('span', { className: 'text-xs text-slate-500', text: alumne.grup || '—' }),
      ],
    });
    button.addEventListener('click', () => {
      uiState.fitxa.selectedAlumne = alumne.id;
      dirtyFlags.view = true;
      scheduleRender('fitxa-select');
    });
    item.appendChild(button);
    list.appendChild(item);
  });
  llista.appendChild(list);
  container.appendChild(llista);

  const detail = createElement('div', {
    className: 'space-y-4 rounded-lg border border-slate-200 bg-white p-6 shadow-sm',
  });
  if (!uiState.fitxa.selectedAlumne) {
    detail.appendChild(
      createElement('p', {
        className: 'text-sm text-slate-600',
        text: 'Selecciona un alumne per veure la seva fitxa agregada.',
      }),
    );
  } else {
    detail.appendChild(renderFitxaDetail(state, uiState.fitxa.selectedAlumne));
  }
  container.appendChild(detail);
}

function renderFitxaDetail(state, alumneId) {
  const wrapper = createElement('div', { className: 'space-y-6' });
  const assignaturesList = state.assignatures?.allIds?.map((id) => state.assignatures.byId[id]) || [];
  if (uiState.fitxa.filtreAssignatura !== 'totes' && !assignaturesList.some((a) => a.id === uiState.fitxa.filtreAssignatura)) {
    uiState.fitxa.filtreAssignatura = 'totes';
  }
  const assignaturaFiltreId = uiState.fitxa.filtreAssignatura !== 'totes' ? uiState.fitxa.filtreAssignatura : null;
  const trimestresDisponibles = assignaturaFiltreId ? getTrimestresForAssignatura(state, assignaturaFiltreId) : [];
  if (uiState.fitxa.filtrePeriode === 'trimestre' && (!assignaturaFiltreId || !trimestresDisponibles.length)) {
    uiState.fitxa.filtrePeriode = 'tot';
    uiState.fitxa.filtreTrimestre = '';
  }
  if (
    uiState.fitxa.filtrePeriode === 'trimestre' &&
    trimestresDisponibles.length &&
    !trimestresDisponibles.some((trim) => trim.id === uiState.fitxa.filtreTrimestre)
  ) {
    uiState.fitxa.filtreTrimestre = trimestresDisponibles[0].id;
  }

  const opts = {};
  if (assignaturaFiltreId) {
    opts.assignaturaId = assignaturaFiltreId;
  }
  if (uiState.fitxa.filtrePeriode === 'trimestre' && uiState.fitxa.filtreTrimestre) {
    opts.trimestreId = uiState.fitxa.filtreTrimestre;
  } else if (uiState.fitxa.filtrePeriode === 'rang') {
    if (uiState.fitxa.rangInici) opts.from = uiState.fitxa.rangInici;
    if (uiState.fitxa.rangFi) opts.to = uiState.fitxa.rangFi;
  }

  let data = null;
  try {
    data = appContext.store.getFitxaAlumne(alumneId, opts);
  } catch (error) {
    return createElement('p', {
      className: 'rounded-md border border-rose-200 bg-rose-50 px-4 py-3 text-sm text-rose-700',
      text: `No s'ha pogut carregar la fitxa: ${error.message || error}`,
    });
  }
  const alumne = state.alumnes.byId[alumneId];
  wrapper.appendChild(
    createElement('header', {
      className: 'flex flex-col gap-1',
      children: [
        createElement('h4', { className: 'text-lg font-semibold text-slate-900', text: getAlumneName(state, alumneId) }),
        createElement('p', {
          className: 'text-sm text-slate-500',
          text: `Matriculat a ${Object.keys(data.perAssignatura).length} assignatures`,
        }),
      ],
    }),
  );

  const filtres = createElement('div', { className: 'grid gap-3 md:grid-cols-3' });
  const assignaturaSelect = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
  });
  assignaturaSelect.appendChild(createElement('option', { attrs: { value: 'totes' }, text: 'Totes les assignatures' }));
  assignaturesList.forEach((assignatura) => {
    const option = createElement('option', { attrs: { value: assignatura.id }, text: assignatura.nom || assignatura.id });
    if (assignatura.id === assignaturaFiltreId) option.selected = true;
    assignaturaSelect.appendChild(option);
  });
  assignaturaSelect.addEventListener('change', (event) => {
    uiState.fitxa.filtreAssignatura = event.currentTarget.value || 'totes';
    if (uiState.fitxa.filtreAssignatura === 'totes') {
      uiState.fitxa.filtreTrimestre = '';
      uiState.fitxa.filtrePeriode = 'tot';
    }
    dirtyFlags.view = true;
    scheduleRender('fitxa-filter-assignatura');
  });

  const periodeSelect = createElement('select', {
    className:
      'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
  });
  const periodeOptions = [
    { value: 'tot', text: 'Totes les dates', disabled: false },
    { value: 'trimestre', text: 'Per trimestre', disabled: !assignaturaFiltreId },
    { value: 'rang', text: 'Per rang de dates', disabled: false },
  ];
  periodeOptions.forEach((optionDef) => {
    const option = createElement('option', {
      attrs: { value: optionDef.value, disabled: optionDef.disabled ? 'disabled' : null },
      text: optionDef.text,
    });
    if (optionDef.value === uiState.fitxa.filtrePeriode) option.selected = true;
    periodeSelect.appendChild(option);
  });
  periodeSelect.addEventListener('change', (event) => {
    const value = event.currentTarget.value;
    if (value === 'trimestre' && !assignaturaFiltreId && assignaturesList.length) {
      uiState.fitxa.filtreAssignatura = assignaturesList[0].id;
    }
    uiState.fitxa.filtrePeriode = value;
    dirtyFlags.view = true;
    scheduleRender('fitxa-filter-periode');
  });

  filtres.append(
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Assignatura' }), assignaturaSelect] }),
    createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Període' }), periodeSelect] }),
  );

  if (uiState.fitxa.filtrePeriode === 'trimestre' && assignaturaFiltreId) {
    const trimestreSelect = createElement('select', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
    });
    trimestresDisponibles.forEach((trim) => {
      const option = createElement('option', { attrs: { value: trim.id }, text: `${trim.nom} (${formatDate(trim.tInici)} – ${formatDate(trim.tFi)})` });
      if (trim.id === uiState.fitxa.filtreTrimestre) option.selected = true;
      trimestreSelect.appendChild(option);
    });
    trimestreSelect.addEventListener('change', (event) => {
      uiState.fitxa.filtreTrimestre = event.currentTarget.value;
      dirtyFlags.view = true;
      scheduleRender('fitxa-filter-trimestre');
    });
    filtres.appendChild(
      createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Trimestre' }), trimestreSelect] }),
    );
  } else if (uiState.fitxa.filtrePeriode === 'rang') {
    const fromInput = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: { type: 'date', value: uiState.fitxa.rangInici || '' },
    });
    fromInput.addEventListener('change', (event) => {
      uiState.fitxa.rangInici = toISODateLocal(event.currentTarget.value) || '';
      dirtyFlags.view = true;
      scheduleRender('fitxa-range');
    });
    const toInput = createElement('input', {
      className:
        'w-full rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
      attrs: { type: 'date', value: uiState.fitxa.rangFi || '' },
    });
    toInput.addEventListener('change', (event) => {
      uiState.fitxa.rangFi = toISODateLocal(event.currentTarget.value) || '';
      dirtyFlags.view = true;
      scheduleRender('fitxa-range');
    });
    filtres.append(
      createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Des de' }), fromInput] }),
      createElement('label', { className: 'flex flex-col gap-2 text-sm font-medium text-slate-700', children: [createElement('span', { text: 'Fins a' }), toInput] }),
    );
  }

  wrapper.appendChild(filtres);
  const summary = createElement('div', {
    className: 'grid gap-4 sm:grid-cols-2',
  });
  summary.append(
    createElement('div', {
      className: 'rounded-md border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700',
      children: [
        createElement('h5', { className: 'text-xs font-semibold uppercase text-slate-500', text: 'Assistència' }),
        createElement('p', {
          className: 'mt-2 text-sm',
          text: `Sessions totals: ${data.resumAssistencia.total}. Absències: ${data.resumAssistencia.absents}. Retards: ${data.resumAssistencia.retardsMenors + data.resumAssistencia.retardsMajors}.`,
        }),
      ],
    }),
    createElement('div', {
      className: 'rounded-md border border-slate-200 bg-slate-50 p-4 text-sm text-slate-700',
      children: [
        createElement('h5', { className: 'text-xs font-semibold uppercase text-slate-500', text: 'Incidències' }),
        createElement('p', {
          className: 'mt-2 text-sm',
          text: `Categories registrades: ${Object.keys(data.resumIncidencies).join(', ') || 'Cap incidència'}`,
        }),
      ],
    }),
  );
  wrapper.appendChild(summary);

  const perAssignaturaList = createElement('div', { className: 'space-y-3' });
  Object.entries(data.perAssignatura).forEach(([assignaturaId, info]) => {
    const assignatura = state.assignatures.byId[assignaturaId];
    perAssignaturaList.appendChild(
      createElement('article', {
        className: 'rounded-md border border-slate-200 bg-white p-4 text-sm text-slate-700 shadow-sm',
        children: [
          createElement('header', {
            className: 'flex items-center justify-between',
            children: [
              createElement('h6', { className: 'font-semibold text-slate-900', text: assignatura?.nom || assignaturaId }),
              createElement('span', {
                className: 'rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-600',
                text: info.mode === 'competencial' ? 'Competencial' : 'Numèric',
              }),
            ],
          }),
          createElement('p', {
            className: 'mt-2 text-xs text-slate-500',
            text: `Última nota calculada: ${formatNumber(info.notaFinal?.valueNumRounded ?? 0, 2)} (${info.notaFinal?.quali || 'NA'})`,
          }),
        ],
      }),
    );
  });
  wrapper.appendChild(perAssignaturaList);

  const actionsRow = createElement('div', { className: 'flex flex-wrap gap-3' });
  const docxBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'file-text' } }), createElement('span', { text: 'Butlletí DOCX' })],
  });
  docxBtn.addEventListener('click', () => {
    appContext.actions.exportDOCX_ButlletiAlumne?.(alumneId);
  });
  const csvBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md border border-slate-200 bg-white px-3 py-2 text-sm font-semibold text-slate-700 shadow-sm transition hover:bg-slate-50 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'button' },
    children: [createElement('span', { className: 'lucide h-4 w-4', attrs: { 'data-lucide': 'file-down' } }), createElement('span', { text: 'Exporta CSV' })],
  });
  csvBtn.addEventListener('click', () => {
    appContext.actions.exportCSV_AvaluacionsCompetencial?.();
    Toast('Exportació CSV personalitzada pendent d\'adaptació.', 'info');
  });
  actionsRow.append(docxBtn, csvBtn);
  wrapper.appendChild(actionsRow);

  return wrapper;
}

function renderExportacions(container, state) {
  clearElement(container);
  container.classList.add('grid', 'gap-4', 'md:grid-cols-2', 'xl:grid-cols-3');
  const cards = [
    {
      title: 'Rúbrica CSV',
      description: 'Exporta les qualificacions competencials amb separador ; i coma decimal.',
      action: () => appContext.actions.exportCSV_AvaluacionsCompetencial?.(),
      icon: 'file-down',
    },
    {
      title: 'Acta DOCX',
      description: 'Genera l\'acta de l\'assignatura en format DOCX.',
      action: () => appContext.actions.exportDOCX_ActaAssignatura?.(),
      icon: 'file-text',
    },
    {
      title: 'Backup xifrat',
      description: 'Descarrega una còpia xifrada de totes les dades.',
      action: async () => {
        const password = prompt('Introdueix la contrasenya per xifrar el backup');
        if (!password) return;
        await appContext.actions.exportEncrypted?.(password);
      },
      icon: 'shield',
    },
  ];
  cards.forEach((card) => {
    const article = createElement('article', {
      className: 'flex h-full flex-col justify-between rounded-lg border border-slate-200 bg-white p-5 shadow-sm',
    });
    article.appendChild(
      createElement('div', {
        className: 'space-y-2',
        children: [
          createElement('span', { className: 'lucide h-6 w-6 text-slate-500', attrs: { 'data-lucide': card.icon } }),
          createElement('h4', { className: 'text-lg font-semibold text-slate-900', text: card.title }),
          createElement('p', { className: 'text-sm text-slate-600', text: card.description }),
        ],
      }),
    );
    const button = createElement('button', {
      className:
        'mt-4 inline-flex items-center justify-center rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
      attrs: { type: 'button' },
      text: 'Exporta',
    });
    button.addEventListener('click', card.action);
    article.appendChild(button);
    container.appendChild(article);
  });
}

function renderConfiguracio(container, state) {
  clearElement(container);
  container.classList.add('space-y-6');
  const config = state.configGlobal || {};
  const form = createElement('form', { className: 'space-y-6 rounded-lg border border-slate-200 bg-white p-6 shadow-sm' });
  const decimalsInput = createElement('input', {
    className:
      'w-24 rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
    attrs: { type: 'number', min: '0', max: '3', step: '1', value: config.rounding?.decimals ?? 1 },
  });
  form.appendChild(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Decimals per defecte' }), decimalsInput],
    }),
  );
  const modeSelect = createElement('select', {
    className:
      'w-40 rounded-md border border-slate-300 px-3 py-2 text-sm shadow-sm focus:border-slate-500 focus:outline-none focus:ring-2 focus:ring-slate-200',
  });
  ['half-up'].forEach((mode) => {
    const option = createElement('option', { attrs: { value: mode }, text: mode });
    if ((config.rounding?.mode || 'half-up') === mode) option.selected = true;
    modeSelect.appendChild(option);
  });
  form.appendChild(
    createElement('label', {
      className: 'flex flex-col gap-2 text-sm font-medium text-slate-700',
      children: [createElement('span', { text: 'Mode d\'arrodoniment per defecte' }), modeSelect],
    }),
  );
  const categoriesList = createElement('div', { className: 'space-y-2' });
  (config.categoriesInit || []).forEach((categoria) => {
    categoriesList.appendChild(
      createElement('span', {
        className: 'inline-flex items-center gap-2 rounded-full bg-slate-100 px-3 py-1 text-xs font-semibold text-slate-700',
        text: categoria,
      }),
    );
  });
  form.appendChild(
    createElement('div', {
      className: 'space-y-2',
      children: [
        createElement('span', { className: 'text-sm font-medium text-slate-700', text: 'Categories inicials' }),
        categoriesList,
      ],
    }),
  );
  const submitBtn = createElement('button', {
    className:
      'inline-flex items-center gap-2 rounded-md bg-slate-900 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-slate-800 focus-visible:outline focus-visible:outline-2 focus-visible:outline-offset-2 focus-visible:outline-slate-900',
    attrs: { type: 'submit' },
    text: 'Desa configuració global',
  });
  form.appendChild(submitBtn);
  form.addEventListener('submit', (event) => {
    event.preventDefault();
    const decimals = Number(decimalsInput.value);
    if (Number.isNaN(decimals) || decimals < 0 || decimals > 3) {
      Toast('Els decimals han de ser entre 0 i 3.', 'warn');
      return;
    }
    try {
      appContext.store.patch(
        {
          configGlobal: {
            ...config,
            rounding: { decimals, mode: modeSelect.value },
          },
        },
        { action: 'config:update' },
      );
      Toast('Configuració global actualitzada.', 'success');
    } catch (error) {
      Toast(`Error en actualitzar la configuració: ${error.message || error}`, 'error');
    }
  });
  container.appendChild(form);
}

function renderCurrentView(container, state) {
  switch (container.id) {
    case 'view-welcome':
      renderWelcome(container, state);
      break;
    case 'view-assignatures':
      renderAssignatures(container, state);
      break;
    case 'view-alumnes':
      renderAlumnes(container, state);
      break;
    case 'view-rubrica':
      renderRubrica(container, state);
      break;
    case 'view-calendari':
      renderCalendari(container, state);
      break;
    case 'view-fitxa':
      renderFitxa(container, state);
      break;
    case 'view-exportacions':
      renderExportacions(container, state);
      break;
    case 'view-configuracio':
      renderConfiguracio(container, state);
      break;
    default:
      clearElement(container);
      container.appendChild(
        createElement('p', {
          className: 'rounded-md border border-dashed border-slate-300 bg-slate-50 px-4 py-6 text-sm text-slate-600',
          text: 'Contingut pendent d\'implementació.',
        }),
      );
  }
}

function doRender() {
  if (!appContext?.store) return;
  const state = appContext.store.getState();
  if (dirtyFlags.header || lastRenderedVersion !== state.version) {
    renderHeader(state);
    dirtyFlags.header = false;
  }
  if (dirtyFlags.view || lastRenderedVersion !== state.version || lastRenderedView !== activeView) {
    const container = containers[activeView];
    if (container) {
      renderCurrentView(container, state);
    }
    lastRenderedView = activeView;
    dirtyFlags.view = false;
  }
  lastRenderedVersion = state.version;
  ensureLucide();
}

function renderHeader(state) {
  const statusValue = qs('header span[role="status"] span.uppercase');
  const statusPill = statusValue?.parentElement;
  if (!statusValue || !statusPill) return;
  statusValue.textContent = uiState.header.autosave;
  const palette = {
    emerald: 'bg-emerald-100 text-emerald-700',
    slate: 'bg-slate-200 text-slate-700',
    amber: 'bg-amber-100 text-amber-700',
    rose: 'bg-rose-100 text-rose-700',
  };
  statusPill.className = `inline-flex items-center gap-2 rounded-full px-3 py-1 text-xs font-semibold ${
    palette[uiState.header.tone] || palette.emerald
  }`;
  const connectBtn = qs('header button[aria-label="Connecta fitxer d\'autocòpia"]');
  if (connectBtn) {
    connectBtn.disabled = uiState.header.locked;
    connectBtn.classList.toggle('opacity-60', uiState.header.locked);
  }
}

function bindHeaderActions() {
  const connectBtn = qs('header button[aria-label="Connecta fitxer d\'autocòpia"]');
  if (connectBtn) {
    connectBtn.addEventListener('click', async () => {
      try {
        await appContext.actions.connectAutoCopy?.({ encrypted: true });
        Toast('S\'ha iniciat la connexió amb el fitxer d\'autocòpia.', 'info');
      } catch (error) {
        Toast(`Error en connectar el fitxer: ${error.message || error}`, 'error');
      }
    });
  }
}

function bindShortcuts() {
  const doc = getDocument();
  doc.addEventListener('keydown', (event) => {
    if (event.defaultPrevented || event.target.closest('input, textarea, select, [contenteditable="true"]')) return;
    if (event.key === '?') {
      event.preventDefault();
      Toast('Dreceres: ? ajuda, g a assignatures, g r rúbrica.', 'info');
      return;
    }
    if (event.key === 'g') {
      lastKeySequence = 'g';
      clearTimeout(keySequenceTimeout);
      keySequenceTimeout = setTimeout(() => {
        lastKeySequence = '';
      }, 1500);
      return;
    }
    if (lastKeySequence === 'g') {
      if (event.key === 'a') {
        event.preventDefault();
        showView('view-assignatures');
        lastKeySequence = '';
      } else if (event.key === 'r') {
        event.preventDefault();
        showView('view-rubrica');
        lastKeySequence = '';
      }
    }
  });
}

function attachSidebarHandlers() {
  const doc = getDocument();
  const navButtons = doc.querySelectorAll('aside .sidebar-link');
  navButtons.forEach((button) => {
    button.addEventListener('click', () => {
      const viewId = button.dataset.view;
      if (viewId) {
        showView(viewId);
      }
    });
  });
}

function subscribeEvents() {
  if (!appContext?.events) return;
  const { events } = appContext;
  const updateStatus = (tone, message) => {
    uiState.header.tone = tone;
    uiState.header.autosave = message;
    dirtyFlags.header = true;
    scheduleRender('status-update');
  };
  events.addEventListener('save:ok', () => {
    updateStatus('emerald', 'Local');
    Toast('Canvis desats correctament.', 'success');
    if (uiState.calendari.pendingSave) {
      uiState.calendari.pendingSave = false;
      dirtyFlags.view = true;
      scheduleRender('save-ok');
    }
  });
  events.addEventListener('save:warning', () => {
    updateStatus('amber', 'Avís');
    Toast('S\'ha desat amb advertiments.', 'warn');
    if (uiState.calendari.pendingSave) {
      uiState.calendari.pendingSave = false;
      dirtyFlags.view = true;
      scheduleRender('save-warning');
    }
  });
  events.addEventListener('save:error', (event) => {
    updateStatus('rose', 'Error');
    Toast(`Error en desar: ${event.detail?.error?.message || 'desconegut'}`, 'error');
  });
  events.addEventListener('fs:connected', () => {
    uiState.header.fsConnected = true;
    updateStatus('emerald', 'FS connectat');
    Toast('Connexió establerta amb el sistema d\'autocòpia.', 'success');
  });
  events.addEventListener('fs:disconnected', () => {
    uiState.header.fsConnected = false;
    updateStatus('slate', 'Local');
    Toast('S\'ha desconnectat l\'autocòpia.', 'info');
  });
  events.addEventListener('fs:error', (event) => {
    updateStatus('rose', 'Error');
    Toast(`Error de fitxer: ${event.detail?.error?.message || 'desconegut'}`, 'error');
  });
  events.addEventListener('lock:acquired', () => {
    uiState.header.locked = true;
    updateStatus('amber', 'LOCK');
    Toast('Fitxer bloquejat temporalment.', 'warn');
  });
  events.addEventListener('lock:released', () => {
    uiState.header.locked = false;
    updateStatus('emerald', uiState.header.fsConnected ? 'FS connectat' : 'Local');
    Toast('S\'ha alliberat el bloqueig.', 'info');
  });
  events.addEventListener('crypto:changed', () => {
    Toast('Contrasenya actualitzada correctament.', 'success');
  });
  events.addEventListener('crypto:error', (event) => {
    Toast(`Error criptogràfic: ${event.detail?.error?.message || 'desconegut'}`, 'error');
  });
  events.addEventListener('backup:done', () => {
    Toast('Còpia de seguretat completada.', 'success');
  });
  events.addEventListener('nav:change', (event) => {
    const { viewId } = event.detail || {};
    if (viewId) {
      showView(viewId);
    }
  });
}

function subscribeStore() {
  if (!appContext?.store) return;
  if (unsubscribeStore) unsubscribeStore();
  unsubscribeStore = appContext.store.subscribe((snapshot, change) => {
    if (change?.action && CALENDAR_MUTATIONS.has(change.action)) {
      uiState.calendari.pendingSave = true;
    }
    dirtyFlags.view = true;
    scheduleRender('store');
  });
}

function setInitialView() {
  const doc = getDocument();
  const activeButton = doc.querySelector('aside .sidebar-link.is-active');
  const defaultView = activeButton?.dataset.view || 'view-welcome';
  activeView = defaultView;
}

function ensureContainers() {
  const doc = getDocument();
  containers = {};
  VIEW_IDS.forEach((viewId) => {
    const el = doc.getElementById(viewId);
    if (el) {
      containers[viewId] = el;
    }
  });
}

function renderAllBadges() {
  dirtyFlags.badges = false;
}

function renderAll(state) {
  if (dirtyFlags.badges) {
    renderAllBadges(state);
  }
  dirtyFlags.header = true;
  dirtyFlags.view = true;
  doRender('full');
}

function handleNavigation(viewId) {
  const doc = getDocument();
  const navButtons = doc.querySelectorAll('aside .sidebar-link');
  navButtons.forEach((button) => {
    const isActive = button.dataset.view === viewId;
    button.classList.toggle('is-active', isActive);
    if (isActive) {
      button.setAttribute('aria-current', 'page');
    } else {
      button.removeAttribute('aria-current');
    }
  });
  Object.entries(containers).forEach(([id, section]) => {
    section.classList.toggle('hidden', id !== viewId);
  });
}

export function showView(id) {
  if (!VIEW_IDS.includes(id)) return;
  activeView = id;
  handleNavigation(id);
  dirtyFlags.view = true;
  scheduleRender('showView');
}

function renderHeaderInitial(state) {
  dirtyFlags.header = true;
  renderHeader(state);
}

function initKeyboard() {
  bindShortcuts();
}

function initHeader(state) {
  bindHeaderActions();
  renderHeaderInitial(state);
}

function initNavigation() {
  attachSidebarHandlers();
  setInitialView();
  handleNavigation(activeView);
}

export function render() {
  renderAll(appContext?.store?.getState?.());
}

export function initViews(context) {
  appContext = context;
  ensureContainers();
  const state = appContext.store.getState();
  initHeader(state);
  initNavigation();
  initKeyboard();
  subscribeEvents();
  subscribeStore();
  renderAll(state);
}
