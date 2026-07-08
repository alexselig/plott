import { Suspense } from "react";

import EditorLoader from "@/components/EditorLoader";

export default function EditorPage() {
  return (
    <Suspense fallback={<p className="text-sm text-zinc-400">Loading editor…</p>}>
      <EditorLoader />
    </Suspense>
  );
}
