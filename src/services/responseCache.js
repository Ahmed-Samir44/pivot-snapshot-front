// A small in-page memory of recent API answers (2026-10-27, "anything to make it faster?"): Undo,
// going back to an earlier layout, or opening the same field's filter list again is answered at
// once, with no trip to the server. Entries expire after `ttlMs`; past `maxEntries` the least
// recently used one goes. A failed request is never kept. The backend keeps its own copies too
// (QueryResultCache, SharedTaskCache) — this one saves the network trip and the processing as well.
export function createResponseCache({ ttlMs, maxEntries, now = () => Date.now() }) {
  const entries = new Map(); // key -> { value, expires }; Map order = least recently used first

  const get = (key) => {
    const entry = entries.get(key);
    if (!entry) return undefined;
    if (entry.expires <= now()) {
      entries.delete(key);
      return undefined;
    }
    entries.delete(key);
    entries.set(key, entry);
    return entry.value;
  };

  const set = (key, value) => {
    entries.delete(key);
    entries.set(key, { value, expires: now() + ttlMs });
    while (entries.size > maxEntries) entries.delete(entries.keys().next().value);
  };

  return {
    get,
    set,
    delete: (key) => entries.delete(key),
    clear: () => entries.clear(),
    // One shared load per key: everyone asking for the same thing at once waits on the same
    // request, which isn't tied to any one caller's `signal` — a caller that gives up just stops
    // waiting (its promise rejects with an AbortError) while the others still get the answer.
    shared(key, load, signal) {
      let promise = get(key);
      if (!promise) {
        promise = load();
        set(key, promise);
        promise.catch(() => {
          if (get(key) === promise) entries.delete(key);
        });
      }
      return untilAborted(promise, signal);
    },
  };
}

function abortError(signal) {
  return signal.reason ?? new DOMException("The operation was aborted.", "AbortError");
}

// Waits for `promise`, or rejects as soon as `signal` aborts.
export function untilAborted(promise, signal) {
  if (!signal) return promise;
  if (signal.aborted) return Promise.reject(abortError(signal));
  return new Promise((resolve, reject) => {
    const onAbort = () => reject(abortError(signal));
    signal.addEventListener("abort", onAbort, { once: true });
    promise.then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      },
    );
  });
}
