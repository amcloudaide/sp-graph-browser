import { describe, it, expect, beforeEach } from "vitest";
import "fake-indexeddb/auto";
import { CacheStore } from "../cacheStore";

describe("CacheStore", () => {
  let store: CacheStore;

  beforeEach(async () => {
    store = new CacheStore();
    await store.clear();
  });

  it("stores and retrieves a cache entry", async () => {
    await store.set("sites:root", { id: "1", name: "Test" }, "site");
    const entry = await store.get("sites:root");
    expect(entry).not.toBeNull();
    expect(entry!.data).toEqual({ id: "1", name: "Test" });
    expect(entry!.nodeType).toBe("site");
  });

  it("returns null for missing keys", async () => {
    const entry = await store.get("nonexistent");
    expect(entry).toBeNull();
  });

  it("reports fresh entries within TTL", async () => {
    await store.set("sites:root", { id: "1" }, "site");
    const fresh = await store.isFresh("sites:root", 30);
    expect(fresh).toBe(true);
  });

  it("reports stale entries beyond TTL", async () => {
    await store.set("sites:root", { id: "1" }, "site");
    const fresh = await store.isFresh("sites:root", 0);
    expect(fresh).toBe(false);
  });

  it("invalidates a single key", async () => {
    await store.set("sites:root", { id: "1" }, "site");
    await store.invalidate("sites:root");
    const entry = await store.get("sites:root");
    expect(entry).toBeNull();
  });

  it("invalidates by prefix", async () => {
    await store.set("sites:1:lists", [{ id: "a" }], "lists");
    await store.set("sites:1:columns", [{ id: "b" }], "columns");
    await store.set("sites:2:lists", [{ id: "c" }], "lists");
    await store.invalidateByPrefix("sites:1");
    expect(await store.get("sites:1:lists")).toBeNull();
    expect(await store.get("sites:1:columns")).toBeNull();
    expect(await store.get("sites:2:lists")).not.toBeNull();
  });
});
