import { openDB, type IDBPDatabase } from "idb";

import type { SlideSize } from "@/lib/types";

/**
 * A Deck groups every chart imported from one PowerPoint presentation, so they
 * can be edited together and exported back into a single `.pptx`. The charts are
 * ordinary ChartDocuments (referenced by id) that share the deck's `sourceToken`
 * (the original `.pptx` bytes, stored in `pptxSource`).
 */
export interface Deck {
  id: string;
  /** Presentation name (the imported file name, sans extension). */
  name: string;
  /** File name as imported (e.g. `Q3 review.pptx`) — used for the export name. */
  fileName: string;
  /** IndexedDB key (in `pptxSource`) for the original `.pptx` bytes. */
  sourceToken: string;
  slideSize: SlideSize;
  /** Ordered ChartDocument ids belonging to this deck. */
  chartIds: string[];
  createdAt: string;
  updatedAt: string;
}

const DB_NAME = "plott-decks";
const STORE = "decks";

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

export function newDeckId(): string {
  return `deck-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`;
}

export async function saveDeck(deck: Deck): Promise<void> {
  await (await db()).put(STORE, deck);
}

export async function getDeck(id: string): Promise<Deck | undefined> {
  return (await (await db()).get(STORE, id)) as Deck | undefined;
}

export async function listDecks(): Promise<Deck[]> {
  const all = (await (await db()).getAll(STORE)) as Deck[];
  return all.sort((a, b) => (a.updatedAt < b.updatedAt ? 1 : -1));
}

export async function deleteDeck(id: string): Promise<void> {
  await (await db()).delete(STORE, id);
}
