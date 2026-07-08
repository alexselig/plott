import { customAlphabet } from "nanoid";

import { APP_VERSION, ID_PREFIX } from "@/lib/constants";
import type {
  ChartDocument,
  ChartSpec,
  ChartVersion,
  DataTable,
  ExportStamp,
} from "@/lib/types";

/**
 * Unambiguous alphabet: no 0/O/1/I/L so an ID can be read off a slide and
 * retyped by hand. Four characters => ~1M combinations, plenty per user.
 */
const nano = customAlphabet("23456789ABCDEFGHJKLMNPQRSTUVWXYZ", 4);

export function newChartId(): string {
  return `${ID_PREFIX}-${nano()}`;
}

export function nowIso(): string {
  return new Date().toISOString();
}

/** Create a brand-new document at version 1 from a spec + data snapshot. */
export function createDocument(
  spec: ChartSpec,
  data: DataTable,
  title?: string,
): ChartDocument {
  const timestamp = nowIso();
  const v1: ChartVersion = { version: 1, timestamp, spec, data };
  return {
    id: newChartId(),
    title: title ?? spec.title ?? "Untitled chart",
    currentVersion: 1,
    versions: [v1],
    createdAt: timestamp,
    updatedAt: timestamp,
    appVersion: APP_VERSION,
  };
}

/** Append a new immutable version snapshot and make it the current one. */
export function commitVersion(
  doc: ChartDocument,
  spec: ChartSpec,
  data: DataTable,
  note?: string,
): ChartDocument {
  const nextVersion = latestVersionNumber(doc) + 1;
  const timestamp = nowIso();
  const version: ChartVersion = { version: nextVersion, timestamp, spec, data, note };
  return {
    ...doc,
    currentVersion: nextVersion,
    versions: [...doc.versions, version],
    updatedAt: timestamp,
    appVersion: APP_VERSION,
  };
}

export function latestVersionNumber(doc: ChartDocument): number {
  return doc.versions.reduce((max, v) => Math.max(max, v.version), 0);
}

/** Resolve a specific version (defaults to the current one). */
export function getVersion(doc: ChartDocument, version?: number): ChartVersion {
  const target = version ?? doc.currentVersion;
  return (
    doc.versions.find((v) => v.version === target) ??
    doc.versions[doc.versions.length - 1]
  );
}

/** Build the stamp embedded into an exported image for this version. */
export function stampFor(doc: ChartDocument, version?: number): ExportStamp {
  const v = getVersion(doc, version);
  return {
    chartId: doc.id,
    version: v.version,
    timestamp: v.timestamp,
    appVersion: doc.appVersion,
  };
}

/**
 * Deterministic, information-carrying export filename,
 * e.g. `PLT-7Q2F_v3_2026-07-07.png`.
 */
export function exportFilename(
  doc: ChartDocument,
  version?: number,
  ext = "png",
): string {
  const v = getVersion(doc, version);
  const datePart = v.timestamp.slice(0, 10);
  return `${doc.id}_v${v.version}_${datePart}.${ext}`;
}

/** Record a perceptual-hash fingerprint for a version's exported image. */
export function addPreview(
  doc: ChartDocument,
  version: number,
  hash: string,
): ChartDocument {
  const previews = (doc.previews ?? []).filter((p) => p.version !== version);
  previews.push({ version, hash });
  return { ...doc, previews };
}
