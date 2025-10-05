/*
 * Storage subsystem for llibretadocent
 *
 * Implements a composite storage layer backed by IndexedDB and (optionally)
 * the File System Access API. The IndexedDB layer is the source of truth and
 * all encryption/locking utilities follow the specification described in the
 * project instructions.
 */

// ---------------------------------------------------------------------------
// Constants & configuration
// ---------------------------------------------------------------------------

export const DB_NAME = 'docent_app_v1';
export const STORE_NAME = 'state';
export const DB_VERSION = 1;

export const FILE_FORMAT_VERSION = 2;

export const ENCRYPTION = {
  ALGO: 'AES-GCM',
  KEK_DERIVE: 'PBKDF2',
  HASH: 'SHA-256',
  ITERATIONS: 150000,
  KEY_LENGTH: 256, // bits
};

const DEFAULT_EMPTY_STATE = {};
const SINGLETON_ID = 'singleton';

const LOCK_FILENAME_SUFFIX = '.lock';
const PREVIOUS_SUFFIX = '.prev';
const BACKUP_PREFIX = '.backup-';
const BACKUP_INTERVAL_VERSION = 50;
const BACKUP_INTERVAL_MS = 24 * 60 * 60 * 1000; // 24h
const LOCK_TTL_MS = 120000;
const LOCK_GRACE_MS = 10000;
const LOCK_MAX_RETRY_MS = 15000;

const KEY_USAGE_DEK = ['encrypt', 'decrypt'];
const KEY_USAGE_KEK = ['wrapKey', 'unwrapKey'];

// ---------------------------------------------------------------------------
// Event handling
// ---------------------------------------------------------------------------

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
        if (!this.listeners.has(type)) this.listeners.set(type, new Set());
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
            console.error('storageEvents listener error', error);
          }
        }
        return !event.defaultPrevented;
      }
    };

export const storageEvents = new EventTargetCtor();

const EVENTS = {
  FS_CONNECTED: 'fs:connected',
  FS_DISCONNECTED: 'fs:disconnected',
  FS_ERROR: 'fs:error',
  FS_RECOVERED: 'fs:recovered',
  LOCK_ACQUIRED: 'lock:acquired',
  LOCK_RELEASED: 'lock:released',
  LOCK_BLOCKED: 'lock:blocked',
  CRYPTO_PASSWORD_NEEDED: 'crypto:password-needed',
  CRYPTO_PASSWORD_WRONG: 'crypto:password-wrong',
  CRYPTO_CHANGED: 'crypto:changed',
};

function emit(eventName, detail) {
  storageEvents.dispatchEvent(new CustomEventCtor(eventName, { detail }));
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

const encoder = new TextEncoder();
const decoder = new TextDecoder();

function getCrypto() {
  const crypto = globalThis.crypto || (globalThis.window && window.crypto);
  if (!crypto || !crypto.subtle) {
    throw new Error('WebCrypto API not available');
  }
  return crypto;
}

function nowISO() {
  return new Date().toISOString();
}

function generateRandomBytes(length) {
  const crypto = getCrypto();
  const array = new Uint8Array(length);
  crypto.getRandomValues(array);
  return array;
}

function toBase64(buffer) {
  if (!buffer) return '';
  const bytes = buffer instanceof ArrayBuffer ? new Uint8Array(buffer) : new Uint8Array(buffer.buffer || buffer);
  let binary = '';
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

function fromBase64(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

async function sha256(buffer) {
  const crypto = getCrypto();
  const buf = buffer instanceof ArrayBuffer ? buffer : encoder.encode(buffer);
  return crypto.subtle.digest('SHA-256', buf);
}

function bufferToHex(buffer) {
  const bytes = new Uint8Array(buffer);
  return Array.from(bytes)
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('');
}

function generateDeviceId() {
  const random = generateRandomBytes(16);
  return bufferToHex(random.buffer);
}

function isPlainObject(value) {
  return Object.prototype.toString.call(value) === '[object Object]';
}

function cloneStructured(value) {
  if (typeof structuredClone === 'function') {
    return structuredClone(value);
  }
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (error) {
    return value;
  }
}

function shallowMerge(base, patch) {
  if (!isPlainObject(base)) return patch;
  if (!isPlainObject(patch)) return patch;
  return { ...base, ...patch };
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

// ---------------------------------------------------------------------------
// IndexedDB helper
// ---------------------------------------------------------------------------

function getIndexedDB() {
  const db = globalThis.indexedDB || (globalThis.window && window.indexedDB);
  if (!db) {
    throw new Error('IndexedDB not available');
  }
  return db;
}

function openDatabase() {
  const indexedDBRef = getIndexedDB();
  return new Promise((resolve, reject) => {
    const request = indexedDBRef.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        db.createObjectStore(STORE_NAME, { keyPath: 'id' });
      }
    };
    request.onerror = () => reject(request.error);
    request.onsuccess = () => resolve(request.result);
  });
}

async function withTransaction(db, mode, callback) {
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, mode);
    const store = tx.objectStore(STORE_NAME);
    let finished = false;
    function done(value) {
      if (!finished) {
        finished = true;
        resolve(value);
      }
    }
    function fail(error) {
      if (!finished) {
        finished = true;
        reject(error);
      }
    }
    tx.oncomplete = () => done(undefined);
    tx.onerror = () => fail(tx.error);
    tx.onabort = () => fail(tx.error);
    Promise.resolve()
      .then(() => callback(store))
      .then((value) => {
        if (mode === 'readonly') {
          done(value);
        }
      })
      .catch((error) => {
        try {
          tx.abort();
        } catch (err) {
          // ignore
        }
        fail(error);
      });
  });
}

async function readSingleton(store) {
  return new Promise((resolve, reject) => {
    const request = store.get(SINGLETON_ID);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function writeSingleton(store, data) {
  return new Promise((resolve, reject) => {
    const request = store.put({ ...data, id: SINGLETON_ID });
    request.onsuccess = () => resolve();
    request.onerror = () => reject(request.error);
  });
}

// ---------------------------------------------------------------------------
// Crypto helpers (KEK/DEK scheme)
// ---------------------------------------------------------------------------

async function deriveKeyFromPassword(password, salt, iterations = ENCRYPTION.ITERATIONS) {
  const crypto = getCrypto();
  const keyMaterial = await crypto.subtle.importKey(
    'raw',
    encoder.encode(password),
    { name: ENCRYPTION.KEK_DERIVE },
    false,
    ['deriveBits', 'deriveKey'],
  );
  return crypto.subtle.deriveKey(
    {
      name: ENCRYPTION.KEK_DERIVE,
      hash: ENCRYPTION.HASH,
      iterations,
      salt,
    },
    keyMaterial,
    { name: 'AES-KW', length: ENCRYPTION.KEY_LENGTH },
    false,
    KEY_USAGE_KEK,
  );
}

async function exportRawKey(key) {
  return getCrypto().subtle.exportKey('raw', key);
}

async function generateDek() {
  const crypto = getCrypto();
  return crypto.subtle.generateKey({ name: ENCRYPTION.ALGO, length: ENCRYPTION.KEY_LENGTH }, true, KEY_USAGE_DEK);
}

async function wrapDek(dek, kek) {
  const crypto = getCrypto();
  return crypto.subtle.wrapKey('raw', dek, kek, 'AES-KW');
}

async function unwrapDek(wrapped, kek) {
  const crypto = getCrypto();
  return crypto.subtle.unwrapKey('raw', wrapped, kek, 'AES-KW', { name: ENCRYPTION.ALGO, length: ENCRYPTION.KEY_LENGTH }, true, KEY_USAGE_DEK);
}

async function encryptPayload(dek, iv, plaintextBuffer) {
  const crypto = getCrypto();
  return crypto.subtle.encrypt({ name: ENCRYPTION.ALGO, iv }, dek, plaintextBuffer);
}

async function decryptPayload(dek, iv, ciphertext) {
  const crypto = getCrypto();
  return crypto.subtle.decrypt({ name: ENCRYPTION.ALGO, iv }, dek, ciphertext);
}

async function encryptStateToContainer({ state, password, meta, previousHeader }) {
  const salt = previousHeader?.kdf?.salt_kek ? fromBase64(previousHeader.kdf.salt_kek) : generateRandomBytes(16).buffer;
  const iterations = previousHeader?.kdf?.iterations || ENCRYPTION.ITERATIONS;
  const kek = await deriveKeyFromPassword(password, salt, iterations);
  const rawKek = await exportRawKey(kek);
  const kekFingerprint = await sha256(rawKek);

  const dek = previousHeader?.dek_wrapped
    ? await unwrapDek(fromBase64(previousHeader.dek_wrapped), kek)
    : await generateDek();
  const dekWrapped = await wrapDek(dek, kek);
  const iv = generateRandomBytes(12).buffer;

  const plaintext = {
    schema_version: 1,
    state,
  };
  const plaintextString = JSON.stringify(plaintext);
  const plaintextBuffer = encoder.encode(plaintextString);
  const ciphertext = await encryptPayload(dek, iv, plaintextBuffer);
  const plaintextHash = await sha256(plaintextBuffer);

  const container = {
    file_format_version: FILE_FORMAT_VERSION,
    header: {
      kdf: {
        algo: ENCRYPTION.KEK_DERIVE,
        hash: ENCRYPTION.HASH,
        iterations,
        salt_kek: toBase64(salt),
      },
      kek_fingerprint: toBase64(kekFingerprint),
      dek_wrapped: toBase64(dekWrapped),
      payload: {
        algo: ENCRYPTION.ALGO,
        iv: toBase64(iv),
      },
      meta: {
        device_id: meta.device_id,
        version_counter: meta.version_counter,
        last_modified: meta.last_modified,
        prev_version_counter: meta.prev_version_counter ?? null,
      },
      integrity: {
        plaintext_sha256: toBase64(plaintextHash),
      },
    },
    ciphertext: toBase64(ciphertext),
  };

  return { container, dek, rawKek };
}

async function decryptContainer(container, password) {
  if (!container || typeof container !== 'object') {
    throw Object.assign(new Error('CORRUPTED_FILE'), { code: 'CORRUPTED_FILE' });
  }
  if (container.file_format_version !== FILE_FORMAT_VERSION) {
    throw Object.assign(new Error('FILE_FORMAT_VERSION_MISMATCH'), { code: 'FILE_FORMAT_VERSION_MISMATCH' });
  }
  const { header } = container;
  if (!header) {
    throw Object.assign(new Error('CORRUPTED_FILE'), { code: 'CORRUPTED_FILE' });
  }
  const salt = fromBase64(header.kdf.salt_kek);
  const iterations = header.kdf.iterations || ENCRYPTION.ITERATIONS;
  const kek = await deriveKeyFromPassword(password, salt, iterations);
  const rawKek = await exportRawKey(kek);
  const fingerprint = await sha256(rawKek);
  const expectedFingerprint = header.kek_fingerprint;
  if (expectedFingerprint && expectedFingerprint !== toBase64(fingerprint)) {
    emit(EVENTS.CRYPTO_PASSWORD_WRONG, { code: 'CONTRASENYA_INCORRECTA' });
    const error = new Error('CONTRASENYA_INCORRECTA');
    error.code = 'CONTRASENYA_INCORRECTA';
    throw error;
  }

  const dekWrapped = fromBase64(header.dek_wrapped);
  let dek;
  try {
    dek = await unwrapDek(dekWrapped, kek);
  } catch (error) {
    const err = new Error('CONTRASENYA_INCORRECTA');
    err.code = 'CONTRASENYA_INCORRECTA';
    throw err;
  }

  const ciphertext = fromBase64(container.ciphertext);
  const iv = fromBase64(header.payload.iv);
  let decryptedBuffer;
  try {
    decryptedBuffer = await decryptPayload(dek, iv, ciphertext);
  } catch (error) {
    const err = new Error('INTEGRITY_FAIL');
    err.code = 'INTEGRITY_FAIL';
    throw err;
  }
  const plaintext = JSON.parse(decoder.decode(decryptedBuffer));
  const hash = await sha256(decryptedBuffer);
  const expectedHash = header.integrity?.plaintext_sha256;
  if (expectedHash && expectedHash !== toBase64(hash)) {
    const err = new Error('INTEGRITY_FAIL');
    err.code = 'INTEGRITY_FAIL';
    throw err;
  }
  return { plaintext, header, dek, rawKek };
}

async function rewrapContainer(container, oldPassword, newPassword, options = {}) {
  const { header } = container;
  const saltOld = fromBase64(header.kdf.salt_kek);
  const iterationsOld = header.kdf.iterations || ENCRYPTION.ITERATIONS;
  const kekOld = await deriveKeyFromPassword(oldPassword, saltOld, iterationsOld);
  const rawOld = await exportRawKey(kekOld);
  const fingerprintOld = await sha256(rawOld);
  if (header.kek_fingerprint && header.kek_fingerprint !== toBase64(fingerprintOld)) {
    const err = new Error('CONTRASENYA_INCORRECTA');
    err.code = 'CONTRASENYA_INCORRECTA';
    throw err;
  }
  const dekWrappedOld = fromBase64(header.dek_wrapped);
  const dek = await unwrapDek(dekWrappedOld, kekOld);

  const saltNew = options.salt ? options.salt : generateRandomBytes(16).buffer;
  const iterationsNew = options.iterations || ENCRYPTION.ITERATIONS;
  const kekNew = await deriveKeyFromPassword(newPassword, saltNew, iterationsNew);
  const rawNew = await exportRawKey(kekNew);
  const fingerprintNew = await sha256(rawNew);
  const dekWrappedNew = await wrapDek(dek, kekNew);

  container.header = {
    ...container.header,
    kdf: {
      algo: ENCRYPTION.KEK_DERIVE,
      hash: ENCRYPTION.HASH,
      iterations: iterationsNew,
      salt_kek: toBase64(saltNew),
    },
    kek_fingerprint: toBase64(fingerprintNew),
    dek_wrapped: toBase64(dekWrappedNew),
  };
  return container;
}

// ---------------------------------------------------------------------------
// IndexedDB Adapter (source of truth)
// ---------------------------------------------------------------------------

class IndexedDBAdapter {
  constructor() {
    this._dbPromise = null;
    this._cache = null;
  }

  async _getDB() {
    if (!this._dbPromise) {
      this._dbPromise = openDatabase();
    }
    return this._dbPromise;
  }

  async _ensureCache() {
    if (this._cache) return this._cache;
    const db = await this._getDB();
    let record;
    await withTransaction(db, 'readonly', async (store) => {
      record = await readSingleton(store);
    });
    if (!record) {
      record = {
        id: SINGLETON_ID,
        version: 0,
        last_modified: nowISO(),
        state: DEFAULT_EMPTY_STATE,
        settings: {
          device_id: generateDeviceId(),
          fs_handle: null,
          fs_encrypted: true,
          last_backup: null,
          options: {},
          last_backup_version: 0,
        },
      };
      await withTransaction(db, 'readwrite', async (store) => {
        await writeSingleton(store, record);
      });
    }
    this._cache = record;
    return record;
  }

  async load() {
    const record = await this._ensureCache();
    return { state: cloneStructured(record.state), version: record.version };
  }

  async getSettings() {
    const record = await this._ensureCache();
    return cloneStructured(record.settings || {});
  }

  async updateSettings(patch) {
    const db = await this._getDB();
    const record = await this._ensureCache();
    const nextSettings = shallowMerge(record.settings || {}, patch || {});
    const updated = { ...record, settings: nextSettings };
    await withTransaction(db, 'readwrite', async (store) => {
      await writeSingleton(store, updated);
    });
    this._cache = updated;
    return cloneStructured(nextSettings);
  }

  async save(patchOrState) {
    const db = await this._getDB();
    const record = await this._ensureCache();
    const isPatch = isPlainObject(patchOrState);
    const nextState = isPatch ? shallowMerge(record.state || {}, patchOrState || {}) : patchOrState;
    const version = (record.version || 0) + 1;
    const last_modified = nowISO();
    const updated = {
      ...record,
      state: cloneStructured(nextState),
      version,
      last_modified,
    };
    await withTransaction(db, 'readwrite', async (store) => {
      await writeSingleton(store, updated);
    });
    this._cache = updated;
    lastSaveMeta = { version, last_modified, source: 'idb' };
    return { version };
  }

  async setState(state, version, lastModified = nowISO()) {
    const db = await this._getDB();
    const record = await this._ensureCache();
    const updated = {
      ...record,
      state: cloneStructured(state),
      version,
      last_modified: lastModified,
    };
    await withTransaction(db, 'readwrite', async (store) => {
      await writeSingleton(store, updated);
    });
    this._cache = updated;
    lastSaveMeta = { version, last_modified: lastModified, source: 'idb' };
    return { state: cloneStructured(state), version };
  }

  async exportEncrypted(password) {
    if (!password) {
      throw new Error('Password required');
    }
    const record = await this._ensureCache();
    const meta = {
      device_id: record.settings?.device_id || generateDeviceId(),
      version_counter: record.version,
      last_modified: record.last_modified,
      prev_version_counter: record.version - 1,
    };
    const { container } = await encryptStateToContainer({
      state: record.state,
      password,
      meta,
    });
    const blob = new Blob([JSON.stringify(container)], { type: 'application/octet-stream' });
    return blob;
  }

  async importEncrypted(file, password) {
    if (!file) throw new Error('File required');
    if (!password) throw new Error('Password required');
    const text = await file.text();
    const container = JSON.parse(text);
    const { plaintext, header } = await decryptContainer(container, password);
    const record = await this._ensureCache();
    if ((header.meta?.version_counter || 0) <= (record.version || 0)) {
      const err = new Error('OLDER_VERSION');
      err.code = 'OLDER_VERSION';
      throw err;
    }
    const db = await this._getDB();
    const updated = {
      ...record,
      state: cloneStructured(plaintext.state),
      version: header.meta?.version_counter || record.version,
      last_modified: header.meta?.last_modified || nowISO(),
    };
    await withTransaction(db, 'readwrite', async (store) => {
      await writeSingleton(store, updated);
    });
    this._cache = updated;
    lastSaveMeta = { version: updated.version, last_modified: updated.last_modified, source: 'idb' };
    return { state: cloneStructured(updated.state), version: updated.version };
  }
}

// ---------------------------------------------------------------------------
// File lock helper (best-effort)
// ---------------------------------------------------------------------------

class FileLock {
  constructor(fileHandle, deviceId) {
    this.handle = fileHandle;
    this.deviceId = deviceId;
    this.lockHandle = null;
    this.lockPath = null;
    this.renewTimer = null;
  }

  async _resolveLockHandle() {
    if (!this.handle || !this.handle.name) return null;
    if (this.lockHandle) return this.lockHandle;
    const parent = this.handle.getParent ? await this.handle.getParent() : null;
    if (!parent || !parent.getFileHandle) return null;
    const lockName = `${this.handle.name}${LOCK_FILENAME_SUFFIX}`;
    try {
      this.lockHandle = await parent.getFileHandle(lockName, { create: true });
      this.lockPath = lockName;
    } catch (error) {
      this.lockHandle = null;
    }
    return this.lockHandle;
  }

  async acquire(ttl = LOCK_TTL_MS) {
    const lockFile = await this._resolveLockHandle();
    if (!lockFile) return true; // fallback: no lock support
    const now = Date.now();
    const expiresAt = new Date(now + ttl).toISOString();
    const payload = {
      device_id: this.deviceId,
      owner_heartbeat: new Date(now).toISOString(),
      expires_at: expiresAt,
    };
    try {
      const existing = await this._read(lockFile);
      if (existing) {
        const exp = new Date(existing.expires_at).getTime();
        const heartbeat = existing.owner_heartbeat ? new Date(existing.owner_heartbeat).getTime() : 0;
        if (Date.now() < exp + LOCK_GRACE_MS && Date.now() - heartbeat < LOCK_TTL_MS) {
          return false;
        }
      }
      await this._write(lockFile, payload);
      emit(EVENTS.LOCK_ACQUIRED);
      return true;
    } catch (error) {
      emit(EVENTS.LOCK_BLOCKED, { error });
      return false;
    }
  }

  async renew(ttl = LOCK_TTL_MS) {
    const lockFile = await this._resolveLockHandle();
    if (!lockFile) return;
    const payload = {
      device_id: this.deviceId,
      owner_heartbeat: nowISO(),
      expires_at: new Date(Date.now() + ttl).toISOString(),
    };
    await this._write(lockFile, payload);
  }

  async release() {
    const lockFile = await this._resolveLockHandle();
    if (!lockFile) return;
    try {
      const writable = await lockFile.createWritable();
      await writable.truncate(0);
      await writable.close();
      emit(EVENTS.LOCK_RELEASED);
    } catch (error) {
      emit(EVENTS.LOCK_BLOCKED, { error });
    }
  }

  async _read(lockFile) {
    try {
      const file = await lockFile.getFile();
      if (!file.size) return null;
      const text = await file.text();
      return JSON.parse(text);
    } catch (error) {
      return null;
    }
  }

  async _write(lockFile, data) {
    const writable = await lockFile.createWritable();
    await writable.write(JSON.stringify(data));
    await writable.close();
  }
}

// ---------------------------------------------------------------------------
// File System Adapter (FSAA)
// ---------------------------------------------------------------------------

class FileSystemAdapter {
  constructor(indexedDBAdapter, options = {}) {
    this.idb = indexedDBAdapter;
    this.askPassword = options.askPassword;
    this.handle = null;
    this.settings = null;
    this.connected = false;
    this.lock = null;
    this.lastWrite = { version: 0, timestamp: 0 };
    this._initialised = false;
  }

  async init() {
    if (this._initialised) return;
    this.settings = await this.idb.getSettings();
    this.handle = this.settings.fs_handle || null;
    if (this.handle) {
      this.connected = await this.isConnected();
      if (this.connected) {
        emit(EVENTS.FS_CONNECTED, {});
        const deviceId = this.settings.device_id || generateDeviceId();
        this.lock = new FileLock(this.handle, deviceId);
      }
    }
    this._initialised = true;
  }

  async connectFile({ encrypted = true } = {}) {
    if (!globalThis.showSaveFilePicker) {
      throw new Error('File System Access API not available');
    }
    const handle = await globalThis.showSaveFilePicker({
      suggestedName: encrypted ? 'dades.docentapp.json.enc' : 'dades.docentapp.json',
      types: [
        {
          description: 'Docent data',
          accept: { 'application/json': ['.json', '.enc'] },
        },
      ],
    });
    const settings = await this.idb.updateSettings({
      fs_handle: handle,
      fs_encrypted: !!encrypted,
    });
    this.settings = settings;
    this.handle = handle;
    const deviceId = settings.device_id || generateDeviceId();
    this.lock = new FileLock(handle, deviceId);
    this.connected = true;
    emit(EVENTS.FS_CONNECTED, {});
  }

  async revoke() {
    await this.idb.updateSettings({ fs_handle: null });
    this.handle = null;
    this.connected = false;
    emit(EVENTS.FS_DISCONNECTED, {});
  }

  async isConnected() {
    if (!this.handle) return false;
    if (!this.handle.queryPermission) return true;
    const permission = await this.handle.queryPermission({ mode: 'readwrite' });
    if (permission === 'granted') return true;
    if (permission === 'prompt') {
      const result = await this.handle.requestPermission({ mode: 'readwrite' });
      return result === 'granted';
    }
    return false;
  }

  async saveToFile(stateObj, metadata) {
    if (!this.connected || !this.handle) {
      return { code: 'FS_NOT_CONNECTED' };
    }

    this.settings = await this.idb.getSettings();

    const lockAcquired = await this._acquireLockWithRetry();
    if (!lockAcquired) {
      emit(EVENTS.LOCK_BLOCKED, { code: 'LOCKED' });
      return { code: 'LOCKED' };
    }

    try {
      const password = await this._ensurePassword();
      if (!password) {
        emit(EVENTS.CRYPTO_PASSWORD_NEEDED, {});
        return { code: 'PASSWORD_REQUIRED' };
      }
      const meta = {
        device_id: this.settings.device_id || generateDeviceId(),
        version_counter: metadata.version_counter,
        prev_version_counter: metadata.prev_version_counter,
        last_modified: metadata.last_modified,
      };
      const existing = await this._readContainer();
      const previousHeader = existing?.header;
      const { container } = await encryptStateToContainer({
        state: stateObj,
        password,
        meta,
        previousHeader,
      });
      await this._writeContainer(container);
      await this._maybeBackup(container);
      this.lastWrite = { version: meta.version_counter, timestamp: Date.now() };
      return { code: 'OK' };
    } catch (error) {
      emit(EVENTS.FS_ERROR, { error, code: 'FS_WRITE_FAIL' });
      return { code: 'FS_WRITE_FAIL', error };
    } finally {
      if (lockAcquired) {
        await this.lock?.release();
      }
    }
  }

  async loadFromFile() {
    if (!this.connected || !this.handle) {
      return null;
    }
    this.settings = await this.idb.getSettings();
    try {
      const password = await this._ensurePassword({ optional: true });
      const container = await this._resilientLoad();
      if (!container) return null;
      const pass = container.header?.dek_wrapped ? password : null;
      const secret = pass || (this.settings.fs_encrypted ? password : null);
      if (this.settings.fs_encrypted && !secret) {
        emit(EVENTS.CRYPTO_PASSWORD_NEEDED, {});
        return null;
      }
      const { plaintext, header } = await decryptContainer(container, secret);
      return {
        state: plaintext.state,
        version: header.meta?.version_counter || 0,
        last_modified: header.meta?.last_modified || nowISO(),
      };
    } catch (error) {
      emit(EVENTS.FS_ERROR, { error, code: 'FS_READ_FAIL' });
      return null;
    }
  }

  async listBackups() {
    const backups = [];
    if (!this.handle || !this.handle.getParent) return backups;
    const parent = await this.handle.getParent();
    if (!parent || !parent.entries) return backups;
    for await (const [name, entry] of parent.entries()) {
      if (!name.includes(BACKUP_PREFIX) || !entry.getFile) continue;
      const file = await entry.getFile();
      let version = 0;
      try {
        const text = await file.text();
        const container = JSON.parse(text);
        version = container?.header?.meta?.version_counter || 0;
      } catch (error) {
        version = 0;
      }
      backups.push({
        name,
        date: file.lastModified ? new Date(file.lastModified).toISOString() : nowISO(),
        version,
      });
    }
    return backups;
  }

  async changePassword(oldPwd, newPwd) {
    if (!this.connected || !this.handle) return;
    this.settings = await this.idb.getSettings();
    try {
      const container = await this._resilientLoad();
      if (!container) {
        throw Object.assign(new Error('FS_READ_FAIL'), { code: 'FS_READ_FAIL' });
      }
      const originalCipher = container.ciphertext;
      await rewrapContainer(container, oldPwd, newPwd);
      container.ciphertext = originalCipher; // ensure ciphertext unchanged
      await this._writeContainer(container, { keepPrev: true });
      emit(EVENTS.CRYPTO_CHANGED, {});
    } catch (error) {
      emit(EVENTS.FS_ERROR, { error, code: error.code || 'FS_WRITE_FAIL' });
      throw error;
    }
  }

  async _ensurePassword({ optional = false } = {}) {
    if (!this.settings.fs_encrypted) return null;
    if (typeof this.askPassword !== 'function') {
      if (optional) return null;
      throw new Error('Password callback missing');
    }
    const password = await this.askPassword();
    if (!password) {
      if (!optional) {
        emit(EVENTS.CRYPTO_PASSWORD_NEEDED, {});
        const err = new Error('CONTRASENYA_INCORRECTA');
        err.code = 'CONTRASENYA_INCORRECTA';
        throw err;
      }
      return null;
    }
    return password;
  }

  async _acquireLockWithRetry() {
    if (!this.lock) return true;
    const start = Date.now();
    while (Date.now() - start < LOCK_MAX_RETRY_MS) {
      const ok = await this.lock.acquire();
      if (ok) return true;
      await delay(250 + Math.random() * 250);
    }
    return false;
  }

  async _writeContainer(container, { keepPrev = true } = {}) {
    if (!this.handle) throw new Error('FS handle missing');
    const data = JSON.stringify(container);
    const writable = await this.handle.createWritable();
    await writable.write(data);
    await writable.close();
    if (keepPrev) {
      await this._persistPrevious(container);
    }
  }

  async _persistPrevious(container) {
    if (!this.handle || !this.handle.getParent) return;
    try {
      const parent = await this.handle.getParent();
      if (!parent?.getFileHandle) return;
      const prevName = `${this.handle.name}${PREVIOUS_SUFFIX}`;
      const prevHandle = await parent.getFileHandle(prevName, { create: true });
      const writable = await prevHandle.createWritable();
      await writable.write(JSON.stringify(container));
      await writable.close();
    } catch (error) {
      // best effort backup
    }
  }

  async _maybeBackup(container) {
    if (!this.handle || !this.handle.getParent) return;
    const version = container.header?.meta?.version_counter || 0;
    const now = Date.now();
    const should =
      version - (this.settings.last_backup_version || 0) >= BACKUP_INTERVAL_VERSION ||
      now - (this.settings.last_backup_ts || 0) >= BACKUP_INTERVAL_MS;
    if (!should) return;
    try {
      const parent = await this.handle.getParent();
      if (!parent?.getFileHandle) return;
      const timestamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15);
      const backupName = `${this.handle.name}${BACKUP_PREFIX}${timestamp}.enc`;
      const backupHandle = await parent.getFileHandle(backupName, { create: true });
      const writable = await backupHandle.createWritable();
      await writable.write(JSON.stringify(container));
      await writable.close();
      await this.idb.updateSettings({ last_backup: nowISO(), last_backup_version: version, last_backup_ts: now });
    } catch (error) {
      // ignore backup errors
    }
  }

  async _readContainer() {
    if (!this.handle) return null;
    const file = await this.handle.getFile();
    if (!file || file.size === 0) return null;
    const text = await file.text();
    try {
      return JSON.parse(text);
    } catch (error) {
      const err = new Error('CORRUPTED_FILE');
      err.code = 'CORRUPTED_FILE';
      throw err;
    }
  }

  async _resilientLoad() {
    const container = await this._readContainer().catch(() => null);
    if (container) return container;
    // try .prev
    if (this.handle?.getParent) {
      try {
        const parent = await this.handle.getParent();
        const prevHandle = await parent.getFileHandle(`${this.handle.name}${PREVIOUS_SUFFIX}`);
        const file = await prevHandle.getFile();
        const text = await file.text();
        const parsed = JSON.parse(text);
        emit(EVENTS.FS_RECOVERED, { source: 'prev' });
        return parsed;
      } catch (error) {
        // fall-through
      }
      try {
        const parent = await this.handle.getParent();
        for await (const [name, entry] of parent.entries()) {
          if (!name.includes(BACKUP_PREFIX)) continue;
          const backupFile = await entry.getFile();
          const text = await backupFile.text();
          const parsed = JSON.parse(text);
          emit(EVENTS.FS_RECOVERED, { source: name });
          return parsed;
        }
      } catch (error) {
        // ignore
      }
    }
    return null;
  }
}

// ---------------------------------------------------------------------------
// Composite adapter
// ---------------------------------------------------------------------------

class CompositeAdapter {
  constructor(indexedDBAdapter, fileSystemAdapter) {
    this.idb = indexedDBAdapter;
    this.fs = fileSystemAdapter;
    this.connected = !!fileSystemAdapter?.connected;
    this.lastLoaded = null;
    this._initialised = false;
  }

  async init() {
    if (this._initialised) return;
    if (this.fs) {
      await this.fs.init();
      this.connected = this.fs.connected;
    }
    this._initialised = true;
  }

  async load() {
    await this.init();
    const idbState = await this.idb.load();
    let fsState = null;
    if (this.fs && this.fs.connected) {
      fsState = await this.fs.loadFromFile();
      if (fsState && fsState.version > idbState.version) {
        await this.idb.setState(fsState.state, fsState.version, fsState.last_modified || nowISO());
        this.lastLoaded = { ...fsState, source: 'fs' };
        return fsState;
      }
    }
    this.lastLoaded = { ...idbState, source: 'idb' };
    return idbState;
  }

  async save(patchOrState) {
    await this.init();
    const result = await this.idb.save(patchOrState);
    const version = result.version;
    const record = await this.idb._ensureCache();
    const metadata = {
      version_counter: record.version,
      prev_version_counter: record.version - 1,
      last_modified: record.last_modified,
    };
    if (this.fs && this.fs.connected) {
      const fsResult = await this.fs.saveToFile(record.state, metadata);
      if (fsResult?.code !== 'OK') {
        emit(EVENTS.FS_ERROR, fsResult);
      } else {
        lastSaveMeta = { version, last_modified: record.last_modified, source: 'fs' };
      }
    }
    return { version };
  }

  async exportEncrypted(password) {
    await this.init();
    return this.idb.exportEncrypted(password);
  }

  async importEncrypted(file, password) {
    await this.init();
    const result = await this.idb.importEncrypted(file, password);
    if (this.fs && this.fs.connected) {
      const record = await this.idb._ensureCache();
      const metadata = {
        version_counter: record.version,
        prev_version_counter: record.version - 1,
        last_modified: record.last_modified,
      };
      await this.fs.saveToFile(record.state, metadata);
    }
    return result;
  }

  async changePassword(oldPwd, newPwd) {
    await this.init();
    if (!this.fs || !this.fs.connected) {
      throw new Error('FS_NOT_CONNECTED');
    }
    await this.fs.changePassword(oldPwd, newPwd);
  }

  async listBackups() {
    await this.init();
    if (!this.fs) return [];
    return this.fs.listBackups();
  }
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

let singletonAdapter = null;
let lastSaveMeta = { version: 0, last_modified: null, source: 'idb' };

export function getAdapter(options = {}) {
  if (singletonAdapter) return singletonAdapter;
  const indexed = new IndexedDBAdapter();
  const fsAdapter = options.fileSystem === false ? null : new FileSystemAdapter(indexed, {
    askPassword: options.askPassword,
  });
  singletonAdapter = new CompositeAdapter(indexed, fsAdapter);
  return singletonAdapter;
}

export async function changePassword(oldPwd, newPwd) {
  if (!singletonAdapter) throw new Error('Adapter not initialised');
  await singletonAdapter.changePassword(oldPwd, newPwd);
}

export function lastSaveInfo() {
  return { ...lastSaveMeta };
}

export async function listBackups() {
  if (!singletonAdapter) return [];
  return singletonAdapter.listBackups();
}

// ---------------------------------------------------------------------------
// Test helpers (documentation purposes only)
// ---------------------------------------------------------------------------

/**
 * Round-trip encryption/decryption test helper (documentation).
 * Demonstrates export/import with password validation.
 */
export async function __testRoundTrip(password, state = { foo: 'bar' }) {
  const adapter = getAdapter();
  await adapter.idb.save(state);
  const blob = await adapter.exportEncrypted(password);
  const file = new File([await blob.text()], 'test.enc', { type: 'application/octet-stream' });
  return adapter.importEncrypted(file, password);
}

