import { describe, expect, it } from "vitest";
import { createResponseCache, untilAborted } from "./responseCache";

// Recent answers remembered in the page (2026-10-27).
describe("response cache", () => {
  it("remembers an answer until it expires", () => {
    let time = 0;
    const cache = createResponseCache({ ttlMs: 100, maxEntries: 5, now: () => time });
    cache.set("a", 1);
    time = 99;
    expect(cache.get("a")).toBe(1);
    time = 100;
    expect(cache.get("a")).toBeUndefined();
  });

  it("drops the least recently used answer when full", () => {
    const cache = createResponseCache({ ttlMs: 1000, maxEntries: 2 });
    cache.set("a", 1);
    cache.set("b", 2);
    cache.get("a");
    cache.set("c", 3);
    expect(cache.get("a")).toBe(1);
    expect(cache.get("b")).toBeUndefined();
    expect(cache.get("c")).toBe(3);
  });

  it("shares one load between callers asking at once", async () => {
    const cache = createResponseCache({ ttlMs: 1000, maxEntries: 5 });
    let loads = 0;
    const load = () => {
      loads++;
      return Promise.resolve(["x"]);
    };
    const [first, second] = await Promise.all([cache.shared("f", load), cache.shared("f", load)]);
    expect(first).toBe(second);
    expect(loads).toBe(1);
  });

  it("forgets a failed load so the next one tries again", async () => {
    const cache = createResponseCache({ ttlMs: 1000, maxEntries: 5 });
    await expect(cache.shared("f", () => Promise.reject(new Error("busy")))).rejects.toThrow("busy");
    await expect(cache.shared("f", () => Promise.resolve("ok"))).resolves.toBe("ok");
  });

  it("lets one caller give up without cancelling the load for the others", async () => {
    const cache = createResponseCache({ ttlMs: 1000, maxEntries: 5 });
    let finish;
    const load = () => new Promise((resolve) => (finish = resolve));
    const leaving = new AbortController();
    const abandoned = cache.shared("f", load, leaving.signal);
    const staying = cache.shared("f", load);
    leaving.abort();
    await expect(abandoned).rejects.toMatchObject({ name: "AbortError" });
    finish("items");
    await expect(staying).resolves.toBe("items");
  });

  it("rejects at once for a signal that's already aborted", async () => {
    const controller = new AbortController();
    controller.abort();
    await expect(untilAborted(new Promise(() => {}), controller.signal)).rejects.toMatchObject({ name: "AbortError" });
  });
});
