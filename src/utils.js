export function debounce(fn, ms = 200) {
  let timeoutId;
  return (...args) => {
    window.clearTimeout(timeoutId);
    timeoutId = window.setTimeout(() => {
      fn(...args);
    }, ms);
  };
}

export const formatters = {
  currency() {
    console.warn('Formatador de moneda pendent');
    return null;
  },
  percent() {
    console.warn('Formatador de percentatge pendent');
    return null;
  },
};
