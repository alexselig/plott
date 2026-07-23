import Link from "next/link";

import Masthead from "@/components/Masthead";

export const metadata = {
  title: "Install the add-in · Plott",
  description: "Add Plott to PowerPoint — design presentation-ready charts and drop them onto your slides.",
};

// Static assets aren't automatically base-path-prefixed, so build URLs explicitly.
const BASE = process.env.NEXT_PUBLIC_BASE_PATH || "";
const MANIFEST_HREF = `${BASE}/manifest.xml`;

function Step({ n, children }: { n: number; children: React.ReactNode }) {
  return (
    <li className="flex gap-3.5">
      <span className="plott-mono mt-0.5 flex h-6 w-6 shrink-0 items-center justify-center rounded-full border border-accent text-[12px] font-semibold text-accent">
        {n}
      </span>
      <div className="text-[14px] leading-relaxed text-ink">{children}</div>
    </li>
  );
}

function Platform({ title, subtitle, children }: { title: string; subtitle: string; children: React.ReactNode }) {
  return (
    <section className="rounded-xl border border-rule bg-panel p-6 sm:p-7">
      <h3 className="plott-serif m-0 text-[24px]">{title}</h3>
      <p className="plott-mono mt-1 mb-5 text-[11px] uppercase tracking-[0.14em] text-muted">{subtitle}</p>
      <ol className="flex list-none flex-col gap-3.5 p-0">{children}</ol>
    </section>
  );
}

export default function InstallPage() {
  return (
    <div className="flex min-h-screen flex-col">
      <Masthead
        right={
          <Link
            href="/"
            className="plott-mono rounded-md border border-border px-4 py-2 text-xs text-muted hover:border-accent hover:text-accent"
          >
            Open Plott →
          </Link>
        }
      />

      <div className="mx-auto w-full max-w-[880px] flex-1 px-6 py-12 sm:px-10">
        {/* Hero */}
        <div className="plott-mono mb-2.5 text-[11px] uppercase tracking-[0.2em] text-accent">PowerPoint add-in</div>
        <h1 className="plott-serif m-0 mb-3 text-[44px] font-normal leading-[1.05] tracking-[-.01em] sm:text-[52px]">
          Install Plott for PowerPoint
        </h1>
        <p className="m-0 mb-8 max-w-[620px] text-[15.5px] leading-relaxed text-muted">
          Plott adds a task pane to PowerPoint so you can design a chart, drop it onto the current slide as a picture
          or editable shapes, and restyle a chart that&apos;s already there — right where you&apos;re presenting.
        </p>

        {/* Download */}
        <div className="mb-10 flex flex-col gap-3 rounded-xl border border-accent/40 bg-accent/5 p-6 sm:flex-row sm:items-center sm:justify-between">
          <div>
            <div className="plott-serif text-[20px]">Step 1 — get the manifest</div>
            <p className="m-0 mt-1 text-[13.5px] text-muted">
              A small <code className="rounded bg-black/5 px-1 py-0.5 text-[12px]">.xml</code> file that tells PowerPoint
              where to load Plott from. Download it, then follow your platform below.
            </p>
          </div>
          <a
            href={MANIFEST_HREF}
            download="plott-manifest.xml"
            className="shrink-0 rounded-md bg-accent px-5 py-3 text-center font-semibold text-white hover:bg-accent-hover"
          >
            Download manifest
          </a>
        </div>

        {/* Platform instructions */}
        <h2 className="plott-serif m-0 mb-1 text-[26px]">Step 2 — side-load it</h2>
        <p className="m-0 mb-6 text-[14px] text-muted">Pick where you use PowerPoint. It takes about a minute.</p>

        <div className="flex flex-col gap-5">
          <Platform title="PowerPoint on the web" subtitle="office.com · Microsoft 365">
            <Step n={1}>
              Open any presentation at{" "}
              <a className="text-accent hover:underline" href="https://office.com" target="_blank" rel="noreferrer">
                office.com
              </a>{" "}
              (or from OneDrive / SharePoint).
            </Step>
            <Step n={2}>
              On the ribbon go to <strong>Home → Add-ins</strong> (or <strong>Insert → Add-ins</strong>), then{" "}
              <strong>More Add-ins</strong>.
            </Step>
            <Step n={3}>
              Open the <strong>My Add-ins</strong> tab and choose <strong>Upload My Add-in</strong>.
            </Step>
            <Step n={4}>
              Browse to the <strong>plott-manifest.xml</strong> you downloaded and click <strong>Upload</strong>.
            </Step>
            <Step n={5}>
              Click the <strong>Chart builder</strong> button in the <strong>Plott</strong> group on the Home tab to open
              the pane.
            </Step>
          </Platform>

          <Platform title="PowerPoint on Windows" subtitle="Microsoft 365 · 2208 or newer">
            <Step n={1}>
              Put the manifest in a folder, e.g. <code className="rounded bg-black/5 px-1 py-0.5 text-[12px]">C:\Plott</code>.
            </Step>
            <Step n={2}>
              In File Explorer, right-click that folder → <strong>Properties → Sharing → Share…</strong> and share it with
              yourself, then copy the <strong>network path</strong> (looks like{" "}
              <code className="rounded bg-black/5 px-1 py-0.5 text-[12px]">\\YOURPC\Plott</code>).
            </Step>
            <Step n={3}>
              In PowerPoint go to <strong>File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs</strong>.
            </Step>
            <Step n={4}>
              Paste the network path into <strong>Catalog Url</strong>, click <strong>Add catalog</strong>, tick{" "}
              <strong>Show in Menu</strong>, then OK and restart PowerPoint.
            </Step>
            <Step n={5}>
              <strong>Home → Add-ins → My Add-ins → Shared Folder</strong>, pick <strong>Plott</strong>, then use the{" "}
              <strong>Chart builder</strong> button.
            </Step>
          </Platform>

          <Platform title="PowerPoint on Mac" subtitle="Microsoft 365 · 16.64 or newer">
            <Step n={1}>
              Open <strong>Finder</strong>, press <strong>⇧⌘G</strong>, and go to:
              <pre className="mt-2 overflow-x-auto rounded-md border border-rule bg-white p-3 text-[12px] leading-relaxed">
                ~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef
              </pre>
              If the <strong>wef</strong> folder doesn&apos;t exist, create it.
            </Step>
            <Step n={2}>
              Copy <strong>plott-manifest.xml</strong> into that <strong>wef</strong> folder.
            </Step>
            <Step n={3}>Quit and reopen PowerPoint, then open any presentation.</Step>
            <Step n={4}>
              <strong>Home → Add-ins → My Add-ins → Plott</strong>, or click the <strong>Chart builder</strong> button in
              the Plott group on the Home tab.
            </Step>
          </Platform>
        </div>

        {/* What you can do */}
        <h2 className="plott-serif mb-4 mt-12 text-[26px]">What you can do with the pane</h2>
        <ul className="grid list-none grid-cols-1 gap-3 p-0 sm:grid-cols-2">
          {[
            ["Design a chart", "Pick a type, type or paste data, choose a palette and style — live preview as you go."],
            ["Expand & drag to edit", "Open the full-pane editor and drag bars or points to change values, just like the web app."],
            ["Insert on the slide", "Drop it as a crisp image, or as native, editable PowerPoint shapes you can recolor."],
            ["Restyle what's there", "Select a Plott chart already on a slide and re-open it to tweak and update in place."],
            ["Style an Excel chart", "Select a native chart, pull its data into Plott, restyle it, and overlay it exactly."],
          ].map(([t, d]) => (
            <li key={t} className="rounded-lg border border-rule bg-panel p-4">
              <div className="plott-serif text-[17px]">{t}</div>
              <div className="mt-1 text-[13px] leading-relaxed text-muted">{d}</div>
            </li>
          ))}
        </ul>

        {/* Notes */}
        <h2 className="plott-serif mb-3 mt-12 text-[26px]">Good to know</h2>
        <ul className="flex list-disc flex-col gap-2 pl-5 text-[13.5px] leading-relaxed text-muted">
          <li>
            Plott runs entirely in your PowerPoint — your slides and data stay on your device and in your Microsoft 365
            session. There&apos;s no separate account to create.
          </li>
          <li>
            <strong>Style Excel Chart</strong> needs to read the deck&apos;s data. If a presentation carries an{" "}
            <strong>encrypting sensitivity label or password</strong>, PowerPoint hands the add-in an encrypted copy it
            can&apos;t read — remove/lower the label (or Save As an unlabeled copy) and reopen, then try again. You can
            still design a chart and insert it either way.
          </li>
          <li>
            Requires PowerPoint on the web, Windows (Microsoft 365 2208+), or Mac (16.64+). The manifest points at{" "}
            <code className="rounded bg-black/5 px-1 py-0.5 text-[12px]">this site</code>, so keep it reachable over
            HTTPS.
          </li>
        </ul>

        <div className="mt-12 border-t border-rule pt-6 text-[13px] text-muted">
          Trouble loading it? Make sure you uploaded the manifest you downloaded from this page, that you&apos;re signed
          in to Microsoft 365, and that you restarted PowerPoint after adding it. Source &amp; issues on{" "}
          <a className="text-accent hover:underline" href="https://github.com/alexselig/plott" target="_blank" rel="noreferrer">
            GitHub
          </a>
          .
        </div>
      </div>
    </div>
  );
}
