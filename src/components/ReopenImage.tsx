"use client";

import { useState, type ChangeEvent } from "react";
import { useRouter } from "next/navigation";

import { dhashFromFile } from "@/lib/phash";
import {
  matchByHash,
  parseCodeInput,
  parseFilenameStamp,
  readImageStamp,
  stampToRef,
  type StampRef,
} from "@/lib/reopen";
import { getDocument } from "@/lib/store/db";

export default function ReopenImage() {
  const router = useRouter();
  const [code, setCode] = useState("");
  const [msg, setMsg] = useState<string | null>(null);

  async function resolveAndOpen(ref: StampRef | null) {
    if (!ref) {
      setMsg("Couldn't find a Plott code (like PLT-7Q2F).");
      return;
    }
    const doc = await getDocument(ref.chartId);
    if (!doc) {
      setMsg(`${ref.chartId} isn't in this browser's library.`);
      return;
    }
    const params = new URLSearchParams({ id: ref.chartId });
    if (ref.version && doc.versions.some((x) => x.version === ref.version)) {
      params.set("v", String(ref.version));
    }
    router.push(`/editor?${params.toString()}`);
  }

  async function onImage(e: ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;
    setMsg(null);
    const stamp = await readImageStamp(file);
    let ref = stamp ? stampToRef(stamp) : parseFilenameStamp(file.name);
    // Fall back to perceptual-hash matching for re-encoded / pasted images.
    if (!ref || !(await getDocument(ref.chartId))) {
      try {
        const match = await matchByHash(await dhashFromFile(file));
        if (match) ref = match;
      } catch {
        /* ignore hashing errors */
      }
    }
    await resolveAndOpen(ref);
    e.target.value = "";
  }

  return (
    <div className="rounded-xl border border-rule bg-panel p-4">
      <h2 className="plott-mono text-[11px] uppercase tracking-[0.14em] text-faint">Reopen from an exported image</h2>
      <p className="mt-1 text-xs text-muted">
        Drop a PNG exported by Plott (reads its embedded ID), or paste a code like{" "}
        <code className="plott-mono">PLT-7Q2F</code>.
      </p>
      <div className="mt-3 flex flex-wrap items-center gap-2">
        <label className="cursor-pointer rounded-md border border-border px-3 py-1.5 text-xs font-medium hover:border-accent">
          Choose PNG…
          <input type="file" accept="image/png" onChange={onImage} className="sr-only" />
        </label>
        <input
          value={code}
          onChange={(e) => setCode(e.target.value)}
          placeholder="PLT-7Q2F"
          className="plott-mono rounded-md border border-border bg-panel px-2 py-1.5 text-xs"
          aria-label="Chart code"
        />
        <button
          type="button"
          onClick={() => resolveAndOpen(parseCodeInput(code))}
          className="rounded-md bg-accent px-3 py-1.5 text-xs font-medium text-white hover:bg-accent-hover"
        >
          Open
        </button>
      </div>
      {msg && <p className="mt-2 text-xs text-accent">{msg}</p>}
      <p className="mt-2 text-xs text-faint">
        Re-encoded or pasted images are matched by visual fingerprint too.
      </p>
    </div>
  );
}
