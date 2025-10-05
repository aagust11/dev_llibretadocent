export const DB_NAME = 'llibretadocent';
export const STORE_NAME = 'settings';
export const DB_VERSION = 1;
export const AUTOSAVE_KEY = 'autosave';
export const PROFILE_KEY = 'profile';

export function getAdapter(options = {}) {
  console.info('Storage adapter no implementat', options);
  return {
    async load() {
      return null;
    },
    async save() {
      return null;
    },
  };
}
