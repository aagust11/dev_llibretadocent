import { createStore, DEFAULTS } from './state.js';
import { getAdapter, DB_NAME, DB_VERSION, STORE_NAME } from './storage.js';
import { actions } from './actions.js';
import { initViews, showView } from './views.js';
import { i18n } from './i18n.js';
import { debounce } from './utils.js';

const store = createStore(DEFAULTS);
const storageAdapter = getAdapter({ dbName: DB_NAME, storeName: STORE_NAME, version: DB_VERSION });

initViews({ store, actions, i18n, storage: storageAdapter });

actions.init?.({ store, storage: storageAdapter });

const navButtons = document.querySelectorAll('[data-view]');
const viewSections = document.querySelectorAll('[id^="view-"]');

function handleNavigation(event) {
  const target = event.currentTarget;
  const viewId = target.dataset.view;
  if (!viewId) return;

  navButtons.forEach((button) => {
    const isActive = button === target;
    button.classList.toggle('is-active', isActive);
    if (isActive) {
      button.setAttribute('aria-current', 'page');
    } else {
      button.removeAttribute('aria-current');
    }
  });

  viewSections.forEach((section) => {
    section.classList.toggle('hidden', section.id !== viewId);
  });

  showView(viewId);
  actions.navigate?.(viewId);
}

navButtons.forEach((button) => {
  button.addEventListener('click', handleNavigation);
});

const handleResize = debounce(() => {
  actions.onResize?.(window.innerWidth);
}, 200);

window.addEventListener('resize', handleResize);

console.log('init');
