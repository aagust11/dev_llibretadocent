import { createStore } from './state.js';
import { getAdapter } from './storage.js';
import { createActions } from './actions.js';
import { initViews, showView } from './views.js';
import { i18n } from './i18n.js';
import { debounce } from './utils.js';

let store;
let storage;
let actions;
let ready = false;

const DEFAULT_VIEW_ID = 'view-assignatures';
const LAST_VIEW_KEY = 'llibretadocent:lastViewId';
const REFRESH_INTERVAL_MS = 5 * 60 * 1000;

let refreshIntervalId = null;
let fsConnected = false;
let currentViewId = null;

const knownViewIds = new Set();

const updateHashDebounced = debounce((viewId) => {
  if (!viewId) return;
  const hash = `#${viewId}`;
  if (window.location.hash !== hash) {
    window.location.hash = hash;
  }
}, 120);

function emitUIEvent(type, detail) {
  document.dispatchEvent(new CustomEvent(type, { detail }));
}

function getSessionStorage() {
  try {
    return window.sessionStorage;
  } catch (error) {
    console.warn('SessionStorage not available', error);
    return null;
  }
}

function collectKnownViewIds() {
  knownViewIds.clear();
  const navButtons = document.querySelectorAll('[data-view]');
  navButtons.forEach((button) => {
    if (button.dataset.view) {
      knownViewIds.add(button.dataset.view);
    }
  });
  const sections = document.querySelectorAll('section[id^="view-"]');
  sections.forEach((section) => {
    knownViewIds.add(section.id);
  });
}

function isValidViewId(viewId) {
  if (!viewId) return false;
  if (!knownViewIds.size) {
    collectKnownViewIds();
  }
  return knownViewIds.has(viewId);
}

function storeLastView(viewId) {
  if (!isValidViewId(viewId)) return;
  const session = getSessionStorage();
  session?.setItem(LAST_VIEW_KEY, viewId);
}

function readLastView() {
  const session = getSessionStorage();
  if (!session) return null;
  const stored = session.getItem(LAST_VIEW_KEY);
  return isValidViewId(stored) ? stored : null;
}

function getViewFromHash() {
  const { hash } = window.location;
  if (!hash) return null;
  const viewId = decodeURIComponent(hash.replace(/^#/, ''));
  return isValidViewId(viewId) ? viewId : null;
}

function currentOrDefaultView() {
  return currentViewId || DEFAULT_VIEW_ID;
}

function activateView(viewId, { updateHash = true, notifyActions = true, source = 'manual' } = {}) {
  if (!isValidViewId(viewId)) return false;
  currentViewId = viewId;
  showView(viewId);
  if (document.body) {
    document.body.dataset.activeView = viewId;
  }
  if (notifyActions && typeof actions?.navigate === 'function') {
    actions.navigate(viewId);
  }
  if (updateHash) {
    updateHashDebounced(viewId);
  }
  storeLastView(viewId);
  emitUIEvent('view:change', { id: viewId, source });
  return true;
}

function determineInitialView() {
  return getViewFromHash() || readLastView() || DEFAULT_VIEW_ID;
}

function focusActiveViewContainer() {
  const activeId = currentOrDefaultView();
  let target = document.getElementById(activeId);
  if (!target || target.classList.contains('hidden')) {
    target = document.querySelector('#main-content');
  }
  if (!target) return;
  target.setAttribute('tabindex', '-1');
  target.focus({ preventScroll: false });
  target.addEventListener(
    'blur',
    () => {
      target.removeAttribute('tabindex');
    },
    { once: true },
  );
}

function setupSkipLink() {
  const skipLink = document.querySelector('.skip-link');
  if (!skipLink) return;
  const focusHandler = (event) => {
    event.preventDefault();
    window.requestAnimationFrame(() => focusActiveViewContainer());
  };
  skipLink.addEventListener('click', focusHandler);
  skipLink.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') {
      focusHandler(event);
    }
  });
}

function isEditableTarget(target) {
  if (!target) return false;
  const editable = target.closest?.('input, textarea, select, [contenteditable="true"]');
  return Boolean(editable) || target.isContentEditable;
}

let shortcutPrefix = '';
let shortcutTimer = null;

const SHORTCUT_TIMEOUT = 1500;
const SHORTCUT_VIEWS = new Map([
  ['ga', 'view-assignatures'],
  ['gr', 'view-rubrica'],
  ['gf', 'view-fitxa'],
  ['gc', 'view-configuracio'],
]);

function handleShortcutKeydown(event) {
  if (!ready) return;
  if (event.defaultPrevented) return;
  if (isEditableTarget(event.target)) return;
  const key = event.key;
  if (key === '?') {
    event.preventDefault();
    emitUIEvent('help:toggle');
    return;
  }
  if (key.toLowerCase() === 'g') {
    shortcutPrefix = 'g';
    window.clearTimeout(shortcutTimer);
    shortcutTimer = window.setTimeout(() => {
      shortcutPrefix = '';
    }, SHORTCUT_TIMEOUT);
    return;
  }
  if (shortcutPrefix === 'g') {
    const sequence = `g${key.toLowerCase()}`;
    const viewId = SHORTCUT_VIEWS.get(sequence);
    if (viewId) {
      event.preventDefault();
      activateView(viewId, { source: 'shortcut' });
      shortcutPrefix = '';
      window.clearTimeout(shortcutTimer);
    }
  }
}

function setupShortcutListeners() {
  window.addEventListener('keydown', handleShortcutKeydown);
}

function handleHashChange() {
  const viewId = getViewFromHash();
  if (!viewId) return;
  activateView(viewId, { updateHash: false, source: 'hash' });
}

function setupHashNavigation() {
  window.addEventListener('hashchange', handleHashChange);
}

function handleVisibilityChange() {
  if (!ready) return;
  if (document.visibilityState === 'visible') {
    actions?.refreshFromDiskIfNewer?.().catch((error) => {
      console.error('Refresh on visibility failed', error);
    });
  }
}

function setupVisibilityListener() {
  document.addEventListener('visibilitychange', handleVisibilityChange);
}

function startRefreshInterval() {
  if (!ready || !fsConnected) return;
  if (refreshIntervalId) return;
  refreshIntervalId = window.setInterval(() => {
    if (!ready || !fsConnected) return;
    actions?.refreshFromDiskIfNewer?.().catch((error) => {
      console.error('Periodic refresh failed', error);
    });
  }, REFRESH_INTERVAL_MS);
}

function stopRefreshInterval() {
  if (!refreshIntervalId) return;
  window.clearInterval(refreshIntervalId);
  refreshIntervalId = null;
}

function hideLoadingState() {
  document.documentElement.classList.remove('is-loading');
  if (document.body) {
    document.body.classList.remove('is-loading');
    document.body.dataset.appReady = 'true';
  }
  const loading = document.querySelector('[data-app-loading]');
  if (loading) {
    loading.classList.add('hidden');
  }
}

function updateSaveState(status, detail) {
  const body = document.body;
  if (!body) return;
  body.dataset.save = status;
  if (detail?.when) {
    body.dataset.saveTs = detail.when;
  } else {
    body.dataset.saveTs = new Date().toISOString();
  }
}

function handleAppReady(event) {
  ready = true;
  hideLoadingState();
  const detail = event?.detail || {};
  updateSaveState('ok', detail);
  const initialView = determineInitialView();
  activateView(initialView, { source: 'startup' });
  if (fsConnected) {
    startRefreshInterval();
  }
}

function handleSaveOk(event) {
  updateSaveState('ok', event.detail);
}

function handleSaveWarning(event) {
  updateSaveState('warn', event.detail);
  emitUIEvent('toast:show', { tone: 'warn', detail: event.detail });
}

function handleSaveError(event) {
  updateSaveState('err', event.detail);
  emitUIEvent('toast:show', { tone: 'error', detail: event.detail });
}

function handleFsConnected() {
  fsConnected = true;
  if (document.body) {
    document.body.dataset.fs = 'connected';
  }
  startRefreshInterval();
}

function handleFsDisconnected() {
  fsConnected = false;
  if (document.body) {
    document.body.dataset.fs = 'disconnected';
  }
  stopRefreshInterval();
}

function handleFsError() {
  if (document.body) {
    document.body.dataset.fs = 'error';
  }
}

function handleLockEvent(type, event) {
  if (document.body) {
    document.body.dataset.lock = type.split(':')[1] || type;
  }
}

function handleConflictDetected(event) {
  if (event?.detail) {
    event.detail.resolve = () => actions?.refreshFromDiskIfNewer?.();
  }
}

function handleConflictResolved(event) {
  emitUIEvent('toast:show', { tone: 'success', code: 'conflict:resolved', detail: event.detail });
}

function handleCryptoPasswordNeeded() {
  // UI will react through the emitted event from the bridge and the askPassword handler.
}

function handleCryptoPasswordWrong(event) {
  emitUIEvent('toast:show', { tone: 'error', code: 'crypto:password-wrong', detail: event.detail });
}

function handleCryptoChanged(event) {
  emitUIEvent('toast:show', { tone: 'success', code: 'crypto:changed', detail: event.detail });
}

function handleBackupDone(event) {
  emitUIEvent('toast:show', { tone: 'success', code: 'backup:done', detail: event.detail });
}

function handleNavChange(event) {
  const viewId = event.detail?.viewId;
  if (!isValidViewId(viewId)) return;
  if (viewId === currentViewId) {
    storeLastView(viewId);
    updateHashDebounced(viewId);
    return;
  }
  activateView(viewId, { notifyActions: false, source: 'actions' });
}

function setupActionEventBridge() {
  if (!actions?.events?.addEventListener) return;
  const map = new Map([
    ['app:ready', handleAppReady],
    ['save:ok', handleSaveOk],
    ['save:warning', handleSaveWarning],
    ['save:error', handleSaveError],
    ['fs:connected', handleFsConnected],
    ['fs:disconnected', handleFsDisconnected],
    ['fs:error', handleFsError],
    ['lock:acquired', (event) => handleLockEvent('lock:acquired', event)],
    ['lock:released', (event) => handleLockEvent('lock:released', event)],
    ['lock:blocked', (event) => handleLockEvent('lock:blocked', event)],
    ['conflict:detected', handleConflictDetected],
    ['conflict:resolved', handleConflictResolved],
    ['crypto:password-needed', handleCryptoPasswordNeeded],
    ['crypto:password-wrong', handleCryptoPasswordWrong],
    ['crypto:changed', handleCryptoChanged],
    ['backup:done', handleBackupDone],
    ['nav:change', handleNavChange],
  ]);
  map.forEach((handler, type) => {
    const listener = (event) => {
      try {
        handler?.(event);
      } catch (error) {
        console.error(`Error handling action event ${type}`, error);
      }
      emitUIEvent(type, event.detail);
    };
    actions.events.addEventListener(type, listener);
  });
}

function requestPassword(context = {}) {
  return new Promise((resolve) => {
    const detail = {
      ...context,
      responded: false,
      respond(password) {
        if (detail.responded) return;
        detail.responded = true;
        resolve(password || null);
      },
      cancel() {
        if (detail.responded) return;
        detail.responded = true;
        resolve(null);
      },
    };
    emitUIEvent('crypto:password-needed', detail);
  });
}

function setupPasswordHandler() {
  if (storage) {
    storage.askPassword = requestPassword;
  }
}

function setupGlobalErrorHandlers() {
  window.addEventListener('error', (event) => {
    console.error('Global error', event.error || event.message || event);
    emitUIEvent('app:error', { error: event.error || event.message || event });
  });
  window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled rejection', event.reason || event);
    emitUIEvent('app:error', { error: event.reason || event });
  });
}

function setupGlobalHandlers() {
  setupSkipLink();
  setupShortcutListeners();
  setupHashNavigation();
  setupVisibilityListener();
  setupGlobalErrorHandlers();
}

async function init() {
  if (store || actions) return;
  collectKnownViewIds();
  setupGlobalHandlers();

  store = createStore();
  storage = getAdapter();
  setupPasswordHandler();

  actions = createActions({
    store,
    storage,
    i18n,
    utils: { debounce },
  });

  setupActionEventBridge();

  initViews({ store, actions, i18n, events: actions.events });

  try {
    await actions.init();
  } catch (error) {
    console.error('Error during actions.init()', error);
    emitUIEvent('app:error', { error });
  }
}

function boot() {
  init().catch((error) => {
    console.error('Failed to initialise application', error);
    emitUIEvent('app:error', { error });
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', boot, { once: true });
} else {
  boot();
}
