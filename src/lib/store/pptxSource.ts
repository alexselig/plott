import { openDB, type IDBPDatabase } from "idb";

/**
 * Stores the raw bytes of `.pptx` files imported for the round-trip, keyed by a
 * `sourceToken`. Kept out of the main chart store (which holds JSON documents)
 * so large binaries don't bloat every `getAll` of the gallery.
 */

const DB_NAME = "plott-pptx";
const STORE = "sources";

interface SourceRecord {
  token: string;
  fileName: string;
  bytes: Uint8Array;
  savedAt: string;
}

let dbPromise: Promise<IDBPDatabase> | null = null;

function db(): Promise<IDBPDatabase> {
  if (!dbPromise) {
    dbPromise = openDB(DB_NAME, 1, {
      upgrade(d) {
        if (!d.objectStoreNames.contains(STORE)) {
          d.createObjectStore(STORE, { keyPath: "token" });
        }
      },
    });
  }
  return dbPromise;
}

/** Persist source bytes and return a token that references them. */
export async function saveSource(fileName: string, bytes: Uint8Array): Promise<string> {
  const token = `src-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
  const rec: SourceRecord = { token, fileName, bytes, savedAt: new Date().toISOString() };
  await (await db()).put(STORE, rec);
  return token;
}

/** Retrieve stored source bytes for a token (or undefined if absent). */
export async function getSource(token: string): Promise<Uint8Array | undefined> {
  const rec = (await (await db()).get(STORE, token)) as SourceRecord | undefined;
  return rec?.bytes;
}

/** Delete a stored source (e.g. once the user is done with it). */
export async function deleteSource(token: string): Promise<void> {
  await (await db()).delete(STORE, token);
}
