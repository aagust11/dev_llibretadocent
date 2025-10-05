export const actions = {
  init(context) {
    console.info('Accions inicialitzades', context);
  },
  navigate(viewId) {
    console.info('Navegant cap a', viewId);
  },
  onResize(width) {
    console.debug('Amplada actual', width);
  },
  setLocale(locale) {
    console.info('Canvi de llengua pendent', locale);
  },
};
