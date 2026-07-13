import { Suspense } from "react";

import DeckView from "@/components/DeckView";

export const metadata = { title: "Deck · Plott" };

export default function DeckPage() {
  return (
    <Suspense fallback={<p className="p-6 text-sm text-muted">Loading deck…</p>}>
      <DeckView />
    </Suspense>
  );
}
