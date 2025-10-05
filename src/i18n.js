export const i18n = {
  locale: 'ca',
  strings: {
    welcome: 'Benvingut/da a la llibreta docent',
    connect_backup: "Connecta fitxer d'autocòpia",
    search_placeholder: 'Cerca...',
    autosave_local: 'Local',
  },
};

export function t(key) {
  return i18n.strings[key] ?? key;
}

export function formatDate() {
  console.warn('formatDate pendent d\'implementació');
}

export function formatNumber() {
  console.warn('formatNumber pendent d\'implementació');
}
