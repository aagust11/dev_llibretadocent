export const DEFAULTS = {};

export function createStore(initialState = DEFAULTS) {
  const state = { ...initialState };
  const listeners = new Set();

  return {
    getState() {
      return { ...state };
    },
    setState(partial) {
      Object.assign(state, partial);
      listeners.forEach((listener) => listener({ ...state }));
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
