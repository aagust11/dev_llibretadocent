import { i18n, fail, formatNumberComma, parseNumberComma } from './i18n.js';

export function assert(condition, msg = 'S\'ha produït un error inesperat') {
  if (!condition) {
    throw new Error(msg);
  }
}

export function debounce(fn, ms = 300) {
  assert(typeof fn === 'function', 'Cal una funció a debounce');
  let timeoutId = null;
  const wrapper = (...args) => {
    if (timeoutId !== null) {
      clearTimeout(timeoutId);
    }
    timeoutId = setTimeout(() => {
      timeoutId = null;
      fn(...args);
    }, ms);
  };
  wrapper.cancel = () => {
    if (timeoutId !== null) {
      clearTimeout(timeoutId);
      timeoutId = null;
    }
  };
  return wrapper;
}

function serializeCell(value, { sep, decimalComma }) {
  if (value === null || value === undefined) {
    return '';
  }
  if (typeof value === 'number' && Number.isFinite(value)) {
    if (decimalComma) {
      const [integerPart, fractionalPart] = String(value).split('.');
      return fractionalPart ? `${integerPart},${fractionalPart}` : integerPart;
    }
    return String(value);
  }
  if (value instanceof Date) {
    return i18n.formatDateISO(value);
  }
  const text = String(value);
  const escaped = text.replace(/"/gu, '""');
  if (escaped.includes(sep) || /[\r\n]/u.test(escaped) || escaped !== text) {
    return `"${escaped}"`;
  }
  return escaped;
}

export function toCSV(rows, { sep = ';', decimalComma = true, headers = null } = {}) {
  assert(Array.isArray(rows), 'Cal una llista de files per exportar a CSV');
  const lines = [];
  let resolvedHeaders = headers;

  if (!resolvedHeaders && rows.length > 0) {
    const firstRow = rows[0];
    if (firstRow && typeof firstRow === 'object' && !Array.isArray(firstRow)) {
      resolvedHeaders = Object.keys(firstRow);
    }
  }

  if (resolvedHeaders) {
    lines.push(
      resolvedHeaders
        .map((header) => serializeCell(header, { sep, decimalComma: false }))
        .join(sep),
    );
  }

  for (const row of rows) {
    let cells;
    if (Array.isArray(row)) {
      cells = row;
    } else if (row && typeof row === 'object') {
      const keys = resolvedHeaders || Object.keys(row);
      cells = keys.map((key) => row[key]);
    } else {
      cells = [row];
    }
    const serialized = cells.map((cell) => {
      if (typeof cell === 'number' && Number.isFinite(cell) && decimalComma) {
        const [integerPart, fractionalPart = ''] = String(cell).split('.');
        const formatted = fractionalPart ? `${integerPart},${fractionalPart}` : integerPart;
        return serializeCell(formatted, { sep, decimalComma });
      }
      return serializeCell(cell, { sep, decimalComma });
    });
    lines.push(serialized.join(sep));
  }

  return lines.join('\r\n');
}

export function numberToComma(value, decimals = 1) {
  return formatNumberComma(value, decimals);
}

export function commaToNumber(value) {
  return parseNumberComma(value);
}

export function safeRound(value, decimals = 1, mode = 'half-up') {
  assert(typeof value === 'number' && Number.isFinite(value), 'Cal un nombre finit per arrodonir');
  const dec = Number(decimals);
  if (!Number.isInteger(dec) || dec < 0 || dec > 3) {
    fail('decimalsNoPermesos', { decimals });
  }
  if (mode !== 'half-up') {
    fail('valorInvalid', { mode });
  }
  const factor = 10 ** dec;
  const scaled = Math.sign(value) * Math.round(Math.abs(value) * factor + Number.EPSILON);
  const result = scaled / factor;
  return Number(result.toFixed(dec));
}

export function isISODate(value) {
  if (typeof value !== 'string') {
    return false;
  }
  if (!/^\d{4}-\d{2}-\d{2}$/u.test(value)) {
    return false;
  }
  const [year, month, day] = value.split('-').map((part) => Number(part));
  const date = new Date(Date.UTC(year, month - 1, day));
  return date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day;
}

function toDate(value) {
  if (value instanceof Date) {
    const time = value.getTime();
    if (Number.isNaN(time)) {
      throw fail('valorInvalid', { value });
    }
    return new Date(time);
  }
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    throw fail('valorInvalid', { value });
  }
  return date;
}

export function toISODate(value) {
  const date = toDate(value);
  const year = date.getUTCFullYear();
  const month = String(date.getUTCMonth() + 1).padStart(2, '0');
  const day = String(date.getUTCDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

export function fromISODate(value) {
  assert(typeof value === 'string', 'Cal una cadena ISO per convertir a data');
  if (!isISODate(value)) {
    throw fail('valorInvalid', { value });
  }
  const [year, month, day] = value.split('-').map((part) => Number(part));
  return new Date(Date.UTC(year, month - 1, day));
}

const focusableSelectors = [
  'a[href]',
  'button:not([disabled])',
  'input:not([disabled]):not([type="hidden"])',
  'select:not([disabled])',
  'textarea:not([disabled])',
  '[tabindex]:not([tabindex="-1"])',
];

function getFocusableElements(container) {
  if (!container || typeof container.querySelectorAll !== 'function') {
    return [];
  }
  const nodes = container.querySelectorAll(focusableSelectors.join(','));
  return Array.from(nodes).filter((el) => typeof el.focus === 'function');
}

export function focusTrap(container, { initialFocus = null, escapeDeactivates = true, returnFocus = true } = {}) {
  let previousActive = null;
  let active = false;
  let focusables = [];

  const handleKeyDown = (event) => {
    if (!active || event.key !== 'Tab') {
      if (escapeDeactivates && event.key === 'Escape') {
        event.stopPropagation();
        deactivate();
      }
      return;
    }

    focusables = getFocusableElements(container);
    if (focusables.length === 0) {
      event.preventDefault();
      return;
    }
    const first = focusables[0];
    const last = focusables[focusables.length - 1];
    const current = event.target;
    const contains = container && typeof container.contains === 'function';
    const isElement = typeof Element !== 'undefined' && current instanceof Element;
    const isInside = contains && isElement ? container.contains(current) : true;
    if (event.shiftKey) {
      if (current === first || !isInside) {
        event.preventDefault();
        last.focus();
      }
    } else if (current === last || !isInside) {
      event.preventDefault();
      first.focus();
    }
  };

  const activate = () => {
    if (active) return;
    if (typeof document === 'undefined') {
      active = true;
      return;
    }
    active = true;
    previousActive = document.activeElement;
    focusables = getFocusableElements(container);
    const focusTarget =
      (typeof initialFocus === 'function' ? initialFocus() : initialFocus) || focusables[0];
    if (focusTarget && typeof focusTarget.focus === 'function') {
      focusTarget.focus();
    }
    if (container && typeof container.addEventListener === 'function') {
      container.addEventListener('keydown', handleKeyDown, true);
    }
  };

  const deactivate = () => {
    if (!active) return;
    active = false;
    if (container && typeof container.removeEventListener === 'function') {
      container.removeEventListener('keydown', handleKeyDown, true);
    }
    if (returnFocus && previousActive && typeof previousActive.focus === 'function') {
      previousActive.focus();
    }
  };

  return { activate, deactivate };
}

export function restoreFocus(element) {
  if (element && typeof element.focus === 'function') {
    element.focus();
  }
  return element;
}

export function prefersReducedMotion() {
  if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
    return false;
  }
  return window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

export function uid(prefix = 'id') {
  const time = Date.now().toString(36);
  const random = Math.floor(Math.random() * 1e6).toString(36);
  return `${prefix}-${time}-${random}`;
}

export function clamp(value, min, max) {
  assert(typeof value === 'number' && typeof min === 'number' && typeof max === 'number', 'clamp requereix números');
  return Math.min(Math.max(value, min), max);
}

export function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function resolveTime(value) {
  if (value instanceof Date) {
    const time = value.getTime();
    if (Number.isNaN(time)) {
      throw fail('valorInvalid', { value });
    }
    return time;
  }
  if (typeof value === 'number') {
    return value;
  }
  if (typeof value === 'string') {
    const parsed = Date.parse(value);
    if (!Number.isNaN(parsed)) {
      return parsed;
    }
  }
  throw fail('valorInvalid', { value });
}

export function relativeTime(from, to = Date.now()) {
  const fromMs = resolveTime(from);
  const toMs = resolveTime(to);
  const diffMs = toMs - fromMs;
  const past = diffMs >= 0;
  const delta = Math.abs(diffMs) / 1000;
  if (delta < 45) {
    return 'ara mateix';
  }
  const minutes = delta / 60;
  if (minutes < 1.5) {
    return past ? 'fa 1 min' : 'd\'aquí 1 min';
  }
  if (minutes < 45) {
    const value = Math.round(minutes);
    return past ? `fa ${value} min` : `d'aquí ${value} min`;
  }
  const hours = minutes / 60;
  if (hours < 1.5) {
    return past ? 'fa 1 h' : "d'aquí 1 h";
  }
  if (hours < 22) {
    const value = Math.round(hours);
    return past ? `fa ${value} h` : `d'aquí ${value} h`;
  }
  const days = hours / 24;
  if (days < 1.5) {
    return past ? 'fa 1 dia' : "d'aquí 1 dia";
  }
  if (days < 26) {
    const value = Math.round(days);
    return past ? `fa ${value} dies` : `d'aquí ${value} dies`;
  }
  const months = days / 30;
  if (months < 1.5) {
    return past ? 'fa 1 mes' : "d'aquí 1 mes";
  }
  if (months < 18) {
    const value = Math.round(months);
    return past ? `fa ${value} mesos` : `d'aquí ${value} mesos`;
  }
  const years = days / 365;
  if (years < 1.5) {
    return past ? 'fa 1 any' : "d'aquí 1 any";
  }
  const value = Math.round(years);
  return past ? `fa ${value} anys` : `d'aquí ${value} anys`;
}
