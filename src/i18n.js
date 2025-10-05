const strings = {
  general: {
    desa: 'Desa',
    "cancel·la": 'Cancel·la',
    tanca: 'Tanca',
    accepta: 'Accepta',
    elimina: 'Elimina',
    editar: 'Edita',
    duplicar: 'Duplica',
    confirma: 'Confirma',
    copia: 'Copia',
    connectaAutocopia: "Connecta l'autocòpia",
    desconnecta: 'Desconnecta',
    canviaContrasenya: 'Canvia la contrasenya',
    backupAra: 'Fes una còpia ara',
    llistaBackups: 'Mostra les còpies de seguretat',
    fsConnectat: 'Sistema de fitxers connectat',
    fsDesconnectat: 'Sistema de fitxers desconnectat',
    autosaveOk: 'Autodesat correcte',
    autosaveWarn: 'Autodesat amb avisos',
    autosaveErr: "Error d'autodesat",
    lockActiu: 'Edició bloquejada',
    conflicteDetectat: "S'ha detectat un conflicte",
    conflicteResol: 'Resol el conflicte',
    contrasenyaIncorrecta: 'Contrasenya incorrecta',
    contrasenyaCanviada: 'Contrasenya actualitzada',
  },
  rubrica: {
    na: 'No assolit',
    as: 'Assoliment satisfactori',
    an: 'Assoliment notable',
    ae: "Assoliment excel·lent",
    mitjanaCA: "Mitjana del criteri d'avaluació",
    notaCE: 'Nota del criteri específic',
    veureUnaActivitat: 'Veure una activitat',
    veureMultiplesActivitats: 'Veure múltiples activitats',
  },
  numeric: {
    categories: 'Categories',
    pesGlobalCategoria: 'Pes global de la categoria',
    pesActivitat: "Pes de l'activitat",
    notaFinal: 'Nota final',
  },
  fitxaAlumne: {
    resum: 'Resum',
    avaluacions: 'Avaluacions',
    assistencia: 'Assistència',
    incidencies: 'Incidències',
    nese: 'Necessitats específiques',
    exportacions: 'Exportacions',
  },
  exportacions: {
    exportaCSV: 'Exporta CSV',
    exportaDOCX: 'Exporta DOCX',
    backupXifrat: 'Còpia de seguretat xifrada',
  },
  calendari: {
    trimestres: 'Trimestres',
    diesLectius: 'Dies lectius',
    festius: 'Festius',
    excepcions: 'Excepcions',
    versionsHorari: "Versions d'horari",
    activaDesDe: 'Activa des de',
    simula: 'Simula',
  },
  errors: {
    valorInvalid: 'Valor invàlid',
    intervalInvalid: 'Interval invàlid',
    decimalsNoPermesos: 'Nombre de decimals no permès',
    carregant: 'Carregant…',
    llest: 'Llest',
  },
};

const vowelMonths = new Set(['a', 'à', 'á', 'e', 'é', 'è', 'i', 'í', 'ï', 'o', 'ó', 'ò', 'u', 'ú', 'ü']);

const defaultQualitativeRanges = Object.freeze({
  NA: Object.freeze([0.0, 4.9]),
  AS: Object.freeze([5.0, 6.9]),
  AN: Object.freeze([7.0, 8.9]),
  AE: Object.freeze([9.0, 10.0]),
});

const defaultRounding = Object.freeze({ decimals: 1, mode: 'half-up' });

function ensureFiniteNumber(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    throw fail('valorInvalid', { value });
  }
  return value;
}

function normalizeDecimals(decimals) {
  const dec = Number(decimals ?? 0);
  if (!Number.isInteger(dec) || dec < 0 || dec > 3) {
    throw fail('decimalsNoPermesos', { decimals });
  }
  return dec;
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

function pad(value) {
  return String(value).padStart(2, '0');
}

export function fail(msgKey, extra) {
  const message = strings.errors[msgKey] || msgKey;
  const error = new Error(message);
  error.code = msgKey;
  if (extra !== undefined) {
    error.extra = extra;
  }
  throw error;
}

export function formatNumberComma(value, decimals = 1) {
  ensureFiniteNumber(value);
  const dec = normalizeDecimals(decimals);
  const sign = value < 0 ? '-' : '';
  const absolute = Math.abs(value);
  const fixed = absolute.toFixed(dec);
  const [integerPartRaw, fractionPart] = fixed.split('.');
  const integerPart = integerPartRaw.replace(/\B(?=(\d{3})+(?!\d))/gu, '\u202f');
  if (dec === 0) {
    return `${sign}${integerPart}`;
  }
  return `${sign}${integerPart},${fractionPart}`;
}

export function parseNumberComma(input) {
  if (input === null || input === undefined || input === '') {
    throw fail('valorInvalid', { value: input });
  }
  if (typeof input === 'number' && Number.isFinite(input)) {
    return input;
  }
  const raw = String(input).trim();
  if (!raw) {
    throw fail('valorInvalid', { value: input });
  }
  const cleaned = raw.replace(/[\u202f\u00a0\s]/gu, '');
  const sign = cleaned[0] === '-' || cleaned[0] === '+' ? cleaned[0] : '';
  const unsigned = sign ? cleaned.slice(1) : cleaned;
  if (!/^(?:\d+(?:[.,]\d+)*|[.,]\d+)$/u.test(unsigned)) {
    throw fail('valorInvalid', { value: input });
  }
  const commaCount = (unsigned.match(/,/gu) || []).length;
  const dotCount = (unsigned.match(/\./gu) || []).length;
  const lastComma = cleaned.lastIndexOf(',');
  const lastDot = cleaned.lastIndexOf('.');
  let decimalSeparator = null;
  if (commaCount > 0 && dotCount > 0) {
    decimalSeparator = lastComma > lastDot ? ',' : '.';
  } else if (commaCount === 1) {
    decimalSeparator = ',';
  } else if (dotCount === 1) {
    decimalSeparator = '.';
  }

  let normalized;
  if (decimalSeparator) {
    const index = cleaned.lastIndexOf(decimalSeparator);
    const integerDigits = cleaned
      .slice(sign ? 1 : 0, index)
      .replace(/[.,]/gu, '') || '0';
    const fractionDigits = cleaned.slice(index + 1).replace(/[.,]/gu, '');
    normalized = `${sign}${integerDigits}.${fractionDigits}`;
  } else {
    const integerDigits = unsigned.replace(/[.,]/gu, '') || '0';
    normalized = `${sign}${integerDigits}`;
  }

  if (!/^[-+]?\d*(?:\.\d*)?$/u.test(normalized) || normalized === '' || normalized === '+' || normalized === '-') {
    throw fail('valorInvalid', { value: input });
  }
  const number = Number(normalized);
  if (!Number.isFinite(number)) {
    throw fail('valorInvalid', { value: input });
  }
  return number;
}

export function formatDateCAT(value) {
  const date = toDate(value);
  const day = new Intl.DateTimeFormat('ca-ES', { day: 'numeric' }).format(date);
  let month = new Intl.DateTimeFormat('ca-ES', { month: 'short' }).format(date).toLocaleLowerCase('ca-ES');
  const year = new Intl.DateTimeFormat('ca-ES', { year: 'numeric' }).format(date);
  const needsApostrophe = month && vowelMonths.has(month[0]);
  const preposition = needsApostrophe ? 'd’' : 'de ';
  return `${day} ${preposition}${month} de ${year}`;
}

export function formatDateISO(value) {
  const date = toDate(value);
  const year = date.getUTCFullYear();
  const month = pad(date.getUTCMonth() + 1);
  const day = pad(date.getUTCDate());
  return `${year}-${month}-${day}`;
}

export function formatDateTimeCAT(value) {
  const date = toDate(value);
  const datePart = formatDateCAT(date);
  const timePart = new Intl.DateTimeFormat('ca-ES', {
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).format(date);
  return `${datePart}, ${timePart}`;
}

const qualiLabels = {
  NA: `${strings.rubrica.na} (NA)`,
  AS: `${strings.rubrica.as} (AS)`,
  AN: `${strings.rubrica.an} (AN)`,
  AE: `${strings.rubrica.ae} (AE)`,
};

export function qualiLabel(code) {
  const key = String(code || '').toUpperCase();
  return qualiLabels[key] || code;
}

export function qualiShort(code) {
  const key = String(code || '').toUpperCase();
  if (qualiLabels[key]) {
    return key;
  }
  return code;
}

export const i18n = {
  locale: 'ca',
  strings,
  formatNumberComma,
  parseNumberComma,
  formatDateCAT,
  formatDateISO,
  formatDateTimeCAT,
  qualiLabel,
  qualiShort,
  defaultQualitativeRanges,
  defaultRounding,
  fail,
};

export default i18n;

export { strings, defaultQualitativeRanges, defaultRounding };
