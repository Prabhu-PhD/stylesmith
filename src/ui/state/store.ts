/**
 * The Zustand store — the single state surface the UI talks to. Components read
 * state and call actions; the store calls the office bridge and core model ops.
 * No Office JS in components (CLAUDE.md).
 *
 * Mutations follow one path: run a pure core op → persist via bridge.saveDocument
 * → set the new doc. Reads (scan, selection) refresh cached host state.
 */
import { create } from "zustand";
import { officeBridge, type ScannedShape, type SweepScope, type SweepResult } from "../../office/bridge";
import type { BridgeError } from "../../office/bridge";
import type { LayerName, LiteralPrimitive, StyleSmithDocument, TokenType, ValueKind } from "../../core/schema/types";
import { createEmptyDocument, newId, tokenTypeForProperty } from "../../core/schema/defaults";
import { TEXT_PROPERTIES, SHAPE_PROPERTIES } from "../../core/schema/schemas";
import {
  createStyle,
  renameStyle,
  deleteStyle,
  duplicateStyle,
  bindProperty,
  layersFromCapture,
  createToken,
  editTokenValue,
  renameToken,
  deleteToken,
} from "../../core/styles/model";
import { computeTokenUsage } from "../../core/tokens/usage";
import { resolveValue, indexTokens } from "../../core/tokens/resolve";
import { clusterBySignature } from "../../core/fingerprint/cluster";
import { matchFormatting, type Deviation } from "../../core/fingerprint/match";
import { resolveStyleToTargets, type CurrentFormatting } from "../../core/styles/diff";
import { makeResolveContext } from "./resolve";

export type View = "styles" | "tokens" | "debug";

export interface Operation {
  readonly styleName: string;
  readonly done: number;
  readonly total: number;
  readonly phase: "running" | "done";
  readonly result?: SweepResult;
  readonly cancel: () => void;
}

interface StoreState {
  hostAvailable: boolean;
  status: "loading" | "ready" | "error";
  error: string | null;
  notice: string | null;
  doc: StyleSmithDocument | null;

  view: View;
  selectedStyleId: string | null;
  selectedTokenId: string | null;
  search: string;

  shapes: ScannedShape[];
  scanning: boolean;
  selectionIds: string[];

  applyStyleId: string | null;
  operation: Operation | null;

  // Adoption
  adoptionOpen: boolean;
  adoptionScanning: boolean;
  adoptionScannedCount: number;
  adoptionClusters: AdoptionCluster[];
  adoptionNearMatches: NearMatch[];

  init: () => Promise<void>;
  setView: (v: View) => void;
  setSearch: (s: string) => void;
  selectStyle: (id: string | null) => void;
  selectToken: (id: string | null) => void;
  setNotice: (n: string | null) => void;

  refreshScan: () => Promise<void>;
  refreshSelection: () => Promise<void>;

  createFromSelection: (name: string) => Promise<string | null>;
  rename: (id: string, name: string) => Promise<void>;
  remove: (id: string) => Promise<void>;
  duplicate: (id: string) => Promise<string | null>;
  setProperty: (styleId: string, layer: "text" | "shape", property: string, value: ValueKind) => Promise<void>;
  updateFromSelection: (styleId: string) => Promise<void>;

  openApply: (styleId: string) => void;
  closeApply: () => void;
  apply: (
    styleId: string,
    scope: SweepScope,
    opts: { preserveOverrides: boolean; layers?: LayerName[] },
  ) => Promise<void>;
  applyToSelection: (styleId: string) => Promise<void>;
  dismissOperation: () => void;

  // Tokens
  addToken: (name: string, type: TokenType, value: ValueKind) => Promise<string | null>;
  setTokenValue: (tokenId: string, value: ValueKind) => Promise<boolean>;
  renameTokenAction: (tokenId: string, name: string) => Promise<void>;
  deleteTokenAction: (tokenId: string) => Promise<void>;
  bindToToken: (styleId: string, layer: "text" | "shape", property: string, tokenId: string) => Promise<void>;
  createTokenFromValue: (styleId: string, layer: "text" | "shape", property: string, name: string) => Promise<string | null>;
  unbindToLiteralAction: (styleId: string, layer: "text" | "shape", property: string) => Promise<void>;
  applyTokenCascade: (tokenId: string) => Promise<void>;

  startAdoption: () => Promise<void>;
  cancelAdoption: () => void;
  adoptClusters: (selections: { id: string; name: string }[]) => Promise<void>;
  linkNearMatch: (shapeId: string, styleId: string, normalise: boolean) => Promise<void>;
  skipNearMatch: (shapeId: string) => void;
}

export interface AdoptionCluster {
  readonly id: string;
  readonly size: number;
  readonly memberIds: string[];
  readonly representative: CurrentFormatting;
  readonly summary: string;
}

export interface NearMatch {
  readonly shapeId: string;
  readonly name: string;
  readonly slideIndex: number;
  readonly styleId: string;
  readonly styleName: string;
  readonly deviations: Deviation[];
}

/** A short human description of a cluster's formatting (UX §7 step 2). */
function summarizeFormatting(f: CurrentFormatting): string {
  const parts: string[] = [];
  const fill = f["shape.fill"];
  if (fill !== undefined) parts.push(fill === "none" ? "no fill" : `fill ${fill}`);
  const bw = f["shape.borderWeight"];
  if (bw !== undefined) parts.push(`${bw}pt border`);
  const r0 = f["geometry.adjustments[0]"];
  if (r0 !== undefined) parts.push(`⌒ ${r0}`);
  const size = f["text.fontSize"];
  if (size !== undefined) parts.push(`${size}pt`);
  return parts.join(" · ") || "plain";
}

export function bridgeErrorMessage(e: BridgeError): string {
  switch (e.code) {
    case "no-host":
      return "Not connected to PowerPoint.";
    case "cancelled":
      return "Cancelled.";
    case "office-error":
      return e.message;
    case "invalid-document":
      return e.issues.join("; ");
  }
}

/* ── Pure selectors over cached scan data ────────────────────────────────── */

export function countShapesForStyle(shapes: ScannedShape[], styleId: string): number {
  return shapes.reduce((n, s) => (s.styleId === styleId ? n + 1 : n), 0);
}

export function slideIndicesForSelection(shapes: ScannedShape[], selectionIds: string[]): number[] {
  const sel = new Set(selectionIds);
  const slides = new Set<number>();
  for (const s of shapes) if (sel.has(s.id)) slides.add(s.slideIndex);
  return [...slides].sort((a, b) => a - b);
}

export function unlinkedShapeCount(shapes: ScannedShape[]): number {
  return shapes.reduce((n, s) => (s.styleId === null && !s.isGroup ? n + 1 : n), 0);
}

export interface TokenCascade {
  readonly styleCount: number;
  readonly shapeCount: number;
  readonly slideCount: number;
  readonly styles: { styleId: string; name: string; shapeCount: number }[];
}

/** The cascade preview for a token: which styles (transitively) and how many shapes/slides. */
export function tokenCascade(
  doc: StyleSmithDocument,
  shapes: ScannedShape[],
  tokenId: string,
): TokenCascade {
  const usage = computeTokenUsage(doc, tokenId);
  const affected = new Set(usage.affectedStyleIds);
  const styles = usage.affectedStyleIds.map((id) => ({
    styleId: id,
    name: doc.styles.find((s) => s.id === id)?.name ?? id,
    shapeCount: countShapesForStyle(shapes, id),
  }));
  const shapeCount = styles.reduce((n, s) => n + s.shapeCount, 0);
  const slides = new Set<number>();
  for (const s of shapes) if (s.styleId && affected.has(s.styleId)) slides.add(s.slideIndex);
  return { styleCount: styles.length, shapeCount, slideCount: slides.size, styles };
}

export const useStore = create<StoreState>((set, get) => ({
  hostAvailable: false,
  status: "loading",
  error: null,
  notice: null,
  doc: null,
  view: "styles",
  selectedStyleId: null,
  selectedTokenId: null,
  search: "",
  shapes: [],
  scanning: false,
  selectionIds: [],
  applyStyleId: null,
  operation: null,
  adoptionOpen: false,
  adoptionScanning: false,
  adoptionScannedCount: 0,
  adoptionClusters: [],
  adoptionNearMatches: [],

  init: async () => {
    const hostAvailable = officeBridge.isHostAvailable;
    set({ status: "loading", hostAvailable });

    const loaded = await officeBridge.loadDocument();
    if (loaded.ok) {
      set({ doc: loaded.value ?? createEmptyDocument(newId), status: "ready", error: null });
    } else if (loaded.error.code === "no-host") {
      // Running in a plain browser (dev) — show an empty doc so the UI renders.
      set({ doc: createEmptyDocument(newId), status: "ready", error: null });
    } else {
      set({ status: "error", error: bridgeErrorMessage(loaded.error) });
      return;
    }

    // Keep selection (and thus scope counts / create-enabled) live.
    officeBridge.onSelectionChanged(() => void get().refreshSelection());
    await Promise.all([get().refreshScan(), get().refreshSelection()]);
  },

  setView: (view) => set({ view }),
  setSearch: (search) => set({ search }),
  selectStyle: (selectedStyleId) => set({ selectedStyleId }),
  selectToken: (selectedTokenId) => set({ selectedTokenId }),
  setNotice: (notice) => set({ notice }),

  refreshScan: async () => {
    set({ scanning: true });
    const r = await officeBridge.scanShapes();
    set({ shapes: r.ok ? r.value : [], scanning: false });
  },

  refreshSelection: async () => {
    const r = await officeBridge.getSelectedShapeIds();
    set({ selectionIds: r.ok ? r.value : [] });
  },

  createFromSelection: async (name) => {
    const doc = get().doc;
    if (!doc) return null;

    const captured = await officeBridge.captureSelection();
    if (!captured.ok) {
      set({ notice: bridgeErrorMessage(captured.error) });
      return null;
    }
    if (captured.value === null) {
      set({ notice: "Select a shape first." });
      return null;
    }

    const created = createStyle(doc, { name, layers: layersFromCapture(captured.value) }, newId);
    if (!created.ok) {
      set({ notice: created.error.code === "name-conflict" ? `A style named "${name}" already exists.` : "Could not create style." });
      return null;
    }

    if (!(await persist(set, created.value.doc))) return null;
    const linked = await officeBridge.linkSelectionToStyle(created.value.styleId);
    if (!linked.ok) set({ notice: bridgeErrorMessage(linked.error) });
    await get().refreshScan();
    return created.value.styleId;
  },

  rename: async (id, name) => {
    const doc = get().doc;
    if (!doc) return;
    const r = renameStyle(doc, id, name);
    if (!r.ok) {
      set({ notice: r.error.code === "name-conflict" ? `A style named "${name}" already exists.` : "Rename failed." });
      return;
    }
    await persist(set, r.value);
  },

  remove: async (id) => {
    const doc = get().doc;
    if (!doc) return;
    const r = deleteStyle(doc, id);
    if (!r.ok) return;
    if (await persist(set, r.value)) {
      if (get().selectedStyleId === id) set({ selectedStyleId: null });
      await get().refreshScan();
    }
  },

  duplicate: async (id) => {
    const doc = get().doc;
    if (!doc) return null;
    const r = duplicateStyle(doc, id, newId);
    if (!r.ok) return null;
    if (!(await persist(set, r.value.doc))) return null;
    return r.value.styleId;
  },

  setProperty: async (styleId, layer, property, value) => {
    const doc = get().doc;
    if (!doc) return;
    const r = bindProperty(doc, styleId, layer, property, value);
    if (!r.ok) {
      set({ notice: "That value is not valid for this property." });
      return;
    }
    await persist(set, r.value);
  },

  updateFromSelection: async (styleId) => {
    const doc = get().doc;
    const style = doc?.styles.find((s) => s.id === styleId);
    if (!doc || !style) return;

    const captured = await officeBridge.captureSelection();
    if (!captured.ok) {
      set({ notice: bridgeErrorMessage(captured.error) });
      return;
    }
    if (captured.value === null) {
      set({ notice: "Select a shape first." });
      return;
    }

    // Update literal properties from the shape; do NOT clobber token/theme
    // bindings (their value belongs to the token, not the shape) — a light
    // version of AC4.2. Geometry update is deferred (Phase 4.x).
    let next = doc;
    const groups: [("text" | "shape"), readonly string[]][] = [
      ["text", TEXT_PROPERTIES],
      ["shape", SHAPE_PROPERTIES],
    ];
    for (const [layer, props] of groups) {
      for (const property of props) {
        const v = captured.value[`${layer}.${property}`];
        if (v === undefined) continue;
        const layerObj = next.styles.find((s) => s.id === styleId)?.layers[layer];
        const existing = layerObj ? (layerObj as Record<string, ValueKind | undefined>)[property] : undefined;
        if (existing && existing.kind !== "literal") continue; // keep the binding
        const r = bindProperty(next, styleId, layer, property, { kind: "literal", value: v });
        if (r.ok) next = r.value;
      }
    }
    if (await persist(set, next)) set({ notice: "Style updated from selection." });
  },

  openApply: (styleId) => set({ applyStyleId: styleId }),
  closeApply: () => set({ applyStyleId: null }),

  apply: async (styleId, scope, opts) => {
    const doc = get().doc;
    const style = doc?.styles.find((s) => s.id === styleId);
    if (!style) return;

    const controller = new AbortController();
    set({
      applyStyleId: null,
      operation: { styleName: style.name, done: 0, total: 0, phase: "running", cancel: () => controller.abort() },
    });

    const r = await officeBridge.applyStyle(styleId, scope, {
      preserveOverrides: opts.preserveOverrides,
      ...(opts.layers ? { layers: opts.layers } : {}),
      chunkSize: 50,
      signal: controller.signal,
      onProgress: (done, total) => {
        const op = get().operation;
        if (op) set({ operation: { ...op, done, total } });
      },
    });

    if (!r.ok) {
      set({ operation: null, notice: bridgeErrorMessage(r.error) });
      return;
    }
    const op = get().operation;
    if (op) set({ operation: { ...op, phase: "done", result: r.value, done: r.value.applied, total: r.value.matched } });
    await get().refreshScan();
  },

  applyToSelection: async (styleId) => {
    const ids = get().selectionIds;
    if (ids.length === 0) {
      set({ notice: "Select one or more shapes first." });
      return;
    }
    // A direct apply to the selection is a fresh link — nothing to preserve.
    await get().apply(styleId, { kind: "shapes", shapeIds: ids }, { preserveOverrides: false });
  },

  dismissOperation: () => set({ operation: null }),

  /* ── Tokens ────────────────────────────────────────────────────────────── */

  addToken: async (name, type, value) => {
    const doc = get().doc;
    const setId = doc?.tokenSets[0]?.id;
    if (!doc || !setId) return null;
    const r = createToken(doc, setId, { name, type, value }, newId);
    if (!r.ok) {
      set({ notice: r.error.code === "name-conflict" ? `A token named "${name}" already exists.` : "Could not create token." });
      return null;
    }
    return (await persist(set, r.value.doc)) ? r.value.tokenId : null;
  },

  setTokenValue: async (tokenId, value) => {
    const doc = get().doc;
    if (!doc) return false;
    const r = editTokenValue(doc, tokenId, value);
    if (!r.ok) {
      const msg =
        r.error.code === "cycle"
          ? "That would create a cycle between tokens."
          : r.error.code === "type-mismatch"
            ? "That token is a different type."
            : "Could not update the token.";
      set({ notice: msg });
      return false;
    }
    return persist(set, r.value);
  },

  renameTokenAction: async (tokenId, name) => {
    const doc = get().doc;
    if (!doc) return;
    const r = renameToken(doc, tokenId, name);
    if (!r.ok) {
      set({ notice: r.error.code === "name-conflict" ? `A token named "${name}" already exists.` : "Rename failed." });
      return;
    }
    await persist(set, r.value);
  },

  deleteTokenAction: async (tokenId) => {
    const doc = get().doc;
    if (!doc) return;
    const r = deleteToken(doc, tokenId);
    if (!r.ok) return;
    if (await persist(set, r.value) && get().selectedTokenId === tokenId) set({ selectedTokenId: null });
  },

  bindToToken: async (styleId, layer, property, tokenId) => {
    const doc = get().doc;
    if (!doc) return;
    // Cache the token's currently-resolved literal so a later delete degrades safely.
    const resolved = resolveValue({ kind: "tokenRef", tokenId }, { tokens: indexTokens(doc) });
    const value: ValueKind =
      resolved.ok ? { kind: "tokenRef", tokenId, cached: resolved.value } : { kind: "tokenRef", tokenId };
    const r = bindProperty(doc, styleId, layer, property, value);
    if (!r.ok) {
      set({ notice: "That token can't bind to this property (type mismatch)." });
      return;
    }
    await persist(set, r.value);
  },

  createTokenFromValue: async (styleId, layer, property, name) => {
    const doc = get().doc;
    const setId = doc?.tokenSets[0]?.id;
    if (!doc || !setId) return null;
    const type = tokenTypeForProperty(layer, property);
    if (!type) {
      set({ notice: "This property can't be turned into a token." });
      return null;
    }
    const style = doc.styles.find((s) => s.id === styleId);
    const current = style?.layers[layer]
      ? (style.layers[layer] as Record<string, ValueKind | undefined>)[property]
      : undefined;
    if (!current || current.kind !== "literal") {
      set({ notice: "Only a literal value can be promoted to a token." });
      return null;
    }
    const literal: LiteralPrimitive = current.value;

    const created = createToken(doc, setId, { name, type, value: { kind: "literal", value: literal } }, newId);
    if (!created.ok) {
      set({ notice: created.error.code === "name-conflict" ? `A token named "${name}" already exists.` : "Could not create token." });
      return null;
    }
    const bound = bindProperty(created.value.doc, styleId, layer, property, {
      kind: "tokenRef",
      tokenId: created.value.tokenId,
      cached: literal,
    });
    if (!bound.ok) return null;
    return (await persist(set, bound.value)) ? created.value.tokenId : null;
  },

  unbindToLiteralAction: async (styleId, layer, property) => {
    const doc = get().doc;
    const style = doc?.styles.find((s) => s.id === styleId);
    if (!doc || !style) return;
    const current = style.layers[layer]
      ? (style.layers[layer] as Record<string, ValueKind | undefined>)[property]
      : undefined;
    if (!current) return;
    const resolved = resolveValue(current, { tokens: indexTokens(doc) });
    if (!resolved.ok) {
      set({ notice: "Can't unbind — the current value is unresolvable." });
      return;
    }
    const r = bindProperty(doc, styleId, layer, property, { kind: "literal", value: resolved.value });
    if (r.ok) await persist(set, r.value);
  },

  applyTokenCascade: async (tokenId) => {
    const doc = get().doc;
    const token = doc?.tokenSets.flatMap((s) => s.tokens).find((t) => t.id === tokenId);
    if (!doc || !token) return;

    const cascade = tokenCascade(doc, get().shapes, tokenId);
    if (cascade.styleCount === 0) {
      set({ notice: "No styles use this token yet." });
      return;
    }

    const controller = new AbortController();
    set({ operation: { styleName: token.name, done: 0, total: cascade.shapeCount, phase: "running", cancel: () => controller.abort() } });

    let applied = 0;
    let matched = 0;
    let skippedGeometryShapes = 0;
    let cancelled = false;

    for (const s of cascade.styles) {
      if (controller.signal.aborted) { cancelled = true; break; }
      const r = await officeBridge.applyStyle(s.styleId, { kind: "deck" }, {
        preserveOverrides: true,
        chunkSize: 50,
        signal: controller.signal,
        onProgress: (done) => {
          const op = get().operation;
          if (op) set({ operation: { ...op, done: applied + done } });
        },
      });
      if (!r.ok) { set({ operation: null, notice: bridgeErrorMessage(r.error) }); return; }
      applied += r.value.applied;
      matched += r.value.matched;
      skippedGeometryShapes += r.value.skippedGeometryShapes;
      if (r.value.cancelled) cancelled = true;
    }

    const result: SweepResult = {
      matched, applied, preservedProps: 0, skippedGeometryShapes,
      unresolvedPaths: [], deferredPaths: [], cancelled, chunkSize: 50, elapsedMs: 0, throughput: 0,
    };
    const op = get().operation;
    if (op) set({ operation: { ...op, phase: "done", result, done: applied, total: matched } });
    await get().refreshScan();
  },

  /* ── Adoption ──────────────────────────────────────────────────────────── */

  startAdoption: async () => {
    const doc = get().doc;
    if (!doc) return;
    set({ adoptionOpen: true, adoptionScanning: true, adoptionClusters: [], adoptionNearMatches: [] });

    const r = await officeBridge.readShapesFormatting();
    if (!r.ok) {
      set({ adoptionScanning: false, notice: bridgeErrorMessage(r.error) });
      return;
    }

    // Cluster identical formatting.
    const clusters = clusterBySignature(r.value.map((s) => ({ item: s.id, formatting: s.formatting }))).map(
      (c) => ({ id: c.hash, size: c.size, memberIds: c.members, representative: c.representative, summary: summarizeFormatting(c.representative) }),
    );

    // Near matches against EXISTING styles (shapes not obviously identical).
    const ctx = makeResolveContext(doc);
    const styleTargets = doc.styles.map((style) => {
      const resolved = resolveStyleToTargets(style, ctx);
      const targetFmt: CurrentFormatting = {};
      for (const t of resolved.targets) targetFmt[t.path] = t.value;
      return { style, targetFmt };
    });
    const nearMatches: NearMatch[] = [];
    for (const shp of r.value) {
      for (const st of styleTargets) {
        if (Object.keys(st.targetFmt).length === 0) continue;
        const m = matchFormatting(shp.formatting, st.targetFmt);
        if (m.near) {
          nearMatches.push({ shapeId: shp.id, name: shp.name, slideIndex: shp.slideIndex, styleId: st.style.id, styleName: st.style.name, deviations: m.deviations });
          break;
        }
      }
    }

    set({ adoptionScanning: false, adoptionClusters: clusters, adoptionNearMatches: nearMatches, adoptionScannedCount: r.value.length });
  },

  cancelAdoption: () => set({ adoptionOpen: false, adoptionClusters: [], adoptionNearMatches: [] }),

  adoptClusters: async (selections) => {
    const doc = get().doc;
    if (!doc) return;
    const clusters = get().adoptionClusters;

    let next = doc;
    let created = 0;
    let linked = 0;
    for (const sel of selections) {
      const cluster = clusters.find((c) => c.id === sel.id);
      const name = sel.name.trim();
      if (!cluster || name === "") continue;

      const res = createStyle(next, { name, layers: layersFromCapture(cluster.representative) }, newId);
      if (!res.ok) {
        set({ notice: res.error.code === "name-conflict" ? `A style named "${name}" already exists.` : "Could not create style." });
        continue;
      }
      next = res.value.doc;
      const link = await officeBridge.linkShapes(res.value.styleId, cluster.memberIds);
      if (link.ok) linked += link.value;
      created += 1;
    }

    if (await persist(set, next)) await get().refreshScan();
    set({ adoptionOpen: false, adoptionClusters: [], adoptionNearMatches: [], notice: `Created ${created} ${created === 1 ? "style" : "styles"}, linked ${linked} shapes.` });
  },

  linkNearMatch: async (shapeId, styleId, normalise) => {
    const r = await officeBridge.linkShapes(styleId, [shapeId]);
    if (!r.ok) {
      set({ notice: bridgeErrorMessage(r.error) });
      return;
    }
    if (normalise) {
      await get().apply(styleId, { kind: "shapes", shapeIds: [shapeId] }, { preserveOverrides: false });
    }
    set({ adoptionNearMatches: get().adoptionNearMatches.filter((m) => m.shapeId !== shapeId) });
    await get().refreshScan();
  },

  skipNearMatch: (shapeId) => set({ adoptionNearMatches: get().adoptionNearMatches.filter((m) => m.shapeId !== shapeId) }),
}));

/** Persist a new document via the bridge; returns true on success. */
async function persist(
  set: (partial: Partial<StoreState>) => void,
  doc: StyleSmithDocument,
): Promise<boolean> {
  const r = await officeBridge.saveDocument(doc);
  if (!r.ok) {
    // In a no-host dev tab, saving is unavailable — still reflect the change locally.
    if (r.error.code === "no-host") {
      set({ doc });
      return true;
    }
    set({ notice: bridgeErrorMessage(r.error) });
    return false;
  }
  set({ doc });
  return true;
}
