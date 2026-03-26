import { openDB, IDBPDatabase, DBSchema } from "idb";
import type { CacheEntry, NodeType } from "../types";

interface SpBrowserDB extends DBSchema {
  cache: {
    key: string;
    value: CacheEntry;
  };
}

export class CacheStore {
  private dbPromise: Promise<IDBPDatabase<SpBrowserDB>>;

  constructor(dbName = "sp-graph-browser") {
    this.dbPromise = openDB<SpBrowserDB>(dbName, 1, {
      upgrade(db) {
        db.createObjectStore("cache", { keyPath: "key" });
      },
    });
  }

  async get(key: string): Promise<CacheEntry | null> {
    const db = await this.dbPromise;
    const entry = await db.get("cache", key);
    return entry ?? null;
  }

  async set(key: string, data: unknown, nodeType: NodeType): Promise<void> {
    const db = await this.dbPromise;
    await db.put("cache", {
      key,
      data,
      fetchedAt: Date.now(),
      nodeType,
    });
  }

  async isFresh(key: string, ttlMinutes: number): Promise<boolean> {
    const entry = await this.get(key);
    if (!entry) return false;
    const age = Date.now() - entry.fetchedAt;
    return age < ttlMinutes * 60 * 1000;
  }

  async invalidate(key: string): Promise<void> {
    const db = await this.dbPromise;
    await db.delete("cache", key);
  }

  async invalidateByPrefix(prefix: string): Promise<void> {
    const db = await this.dbPromise;
    const tx = db.transaction("cache", "readwrite");
    let cursor = await tx.store.openCursor();
    while (cursor) {
      if (cursor.key.startsWith(prefix)) {
        await cursor.delete();
      }
      cursor = await cursor.continue();
    }
    await tx.done;
  }

  async clear(): Promise<void> {
    const db = await this.dbPromise;
    await db.clear("cache");
  }
}
