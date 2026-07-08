import { openDB, type IDBPDatabase } from "idb";

import { newChartId, nowIso } from "@/lib/id";
import type { ChartDocument } from "@/lib/types";

const DB_NAME = "chartforge";
const STORE = "charts";

let dbPromise: Promise<IDBPDatabase> | null = null;

function db(): Promise<IDBPDatabase> {
  if (!dbPromise) {
    dbPromise = openDB(DB_NAME, 1, {
      upgrade(d) {
        if (!d.objectStoreNames.contains(STORE)) {
          d.createObjectStore(STORE, { keyPath: "id" });
        }
      },
    });
  }
  return dbPromise;
}

export async function saveDocument(doc: ChartDocument): Promise<void> {
  await (await db()).put(STORE, doc);
}

export async function getDocument(id: string): Promise<ChartDocument | undefined> {
  return (await (await db()).get(STORE, id)) as ChartDocument | undefined;
}

export async function listDocuments(): Promise<ChartDocument[]> {
  const all = (await (await db()).getAll(STORE)) as ChartDocument[];
  return all.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
}

export async function deleteDocument(id: string): Promise<void> {
  await (await db()).delete(STORE, id);
}

/** Deep-copy a stored document under a fresh id and persist it. */
export async function duplicateDocument(id: string): Promise<ChartDocument | undefined> {
  const doc = await getDocument(id);
  if (!doc) return undefined;
  const ts = nowIso();
  const copy: ChartDocument = {
    ...structuredClone(doc),
    id: newChartId(),
    title: `${doc.title} (copy)`,
    createdAt: ts,
    updatedAt: ts,
  };
  await saveDocument(copy);
  return copy;
}
