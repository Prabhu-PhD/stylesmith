/**
 * Apply planning + drift — the pure logic behind the apply engine.
 *
 * Two jobs:
 *  1. resolveStyleToTargets — turn a style's bound values into concrete literal
 *     targets (per property), plus the list that could not be resolved.
 *  2. planApply — given a per-shape SNAPSHOT of what the style last wrote, the
 *     shape's CURRENT formatting, and whether to preserve overrides, decide per
 *     property whether to write the new value or preserve a manual override
 *     (AC13.1). A property whose current value diverges from the snapshot was
 *     hand-edited since the last apply → preserved.
 *
 * PURE — no Office JS. The office/ layer reads current formatting and applies
 * the plan; this module never touches a shape.
 */
import type {
  LayerName,
  LiteralPrimitive,
  ResolveFailureReason,
  Style,
} from "../schema/types";
import { iterateStyleValues } from "./model";
import { resolveValue, type ResolveContext } from "../tokens/resolve";

/** Numbers within this tolerance are "equal" (adjustment floats, e.g. 0.16667). */
export const VALUE_EPSILON = 1e-4;

/** Resolved values the style last wrote to a shape, keyed by property path. */
export type Snapshot = Record<string, LiteralPrimitive>;
/** A shape's current formatting, keyed by property path. */
export type CurrentFormatting = Record<string, LiteralPrimitive>;

export interface TargetValue {
  readonly path: string;
  readonly layer: LayerName;
  readonly property: string;
  readonly index?: number;
  readonly value: LiteralPrimitive;
}

export interface UnresolvedTarget {
  readonly path: string;
  readonly layer: LayerName;
  readonly property: string;
  readonly index?: number;
  readonly reason: ResolveFailureReason;
}

export interface ResolvedStyle {
  readonly targets: TargetValue[];
  readonly unresolved: UnresolvedTarget[];
}

/** "layer.property" or "layer.property[i]". */
export function pathOf(layer: string, property: string, index?: number): string {
  return index === undefined ? `${layer}.${property}` : `${layer}.${property}[${index}]`;
}

/** Resolve every bound value in a style to a concrete target literal. */
export function resolveStyleToTargets(style: Style, ctx: ResolveContext): ResolvedStyle {
  const targets: TargetValue[] = [];
  const unresolved: UnresolvedTarget[] = [];

  for (const entry of iterateStyleValues(style)) {
    const path = pathOf(entry.layer, entry.property, entry.index);
    const base = { path, layer: entry.layer, property: entry.property } as const;
    const withIndex = entry.index !== undefined ? { ...base, index: entry.index } : base;

    const resolved = resolveValue(entry.value, ctx);
    if (resolved.ok) {
      targets.push({ ...withIndex, value: resolved.value });
    } else {
      unresolved.push({ ...withIndex, reason: resolved.reason });
    }
  }
  return { targets, unresolved };
}

function normalizeString(s: string): string {
  return s.trim().replace(/^#/, "").toLowerCase();
}

/** Tolerant equality: numbers within epsilon; strings normalised (case, #, trim). */
export function valuesEqual(a: LiteralPrimitive, b: LiteralPrimitive): boolean {
  if (typeof a === "number" && typeof b === "number") return Math.abs(a - b) <= VALUE_EPSILON;
  return normalizeString(String(a)) === normalizeString(String(b));
}

export interface ApplyPlan {
  /** Targets to write to the shape. */
  readonly toApply: TargetValue[];
  /** Property paths preserved because the shape carries a manual override. */
  readonly preserved: string[];
  /** The snapshot to store on the shape after this apply. */
  readonly nextSnapshot: Snapshot;
}

/**
 * Decide, per resolved target, whether to write it or preserve an override.
 *
 * A property is a preserved override when: preserve is on, the shape has a prior
 * snapshot for it, and its current value diverges from that snapshot (i.e. it was
 * hand-edited since we last wrote it). Everything else is (re)written and its
 * snapshot advanced to the new value. Preserved properties keep their old
 * snapshot baseline so they stay preserved on the next apply too.
 */
export function planApply(
  resolved: ResolvedStyle,
  snapshot: Snapshot | null,
  current: CurrentFormatting,
  preserve: boolean,
): ApplyPlan {
  const toApply: TargetValue[] = [];
  const preserved: string[] = [];
  const nextSnapshot: Snapshot = {};

  for (const target of resolved.targets) {
    const snap = snapshot?.[target.path];
    const cur = current[target.path];
    const isOverride = preserve && snap !== undefined && cur !== undefined && !valuesEqual(cur, snap);

    if (isOverride) {
      preserved.push(target.path);
      nextSnapshot[target.path] = snap;
    } else {
      toApply.push(target);
      nextSnapshot[target.path] = target.value;
    }
  }
  return { toApply, preserved, nextSnapshot };
}

/**
 * Properties whose current value diverges from the style's resolved target —
 * the drift indicator (S18). Independent of snapshots/overrides.
 */
export function computeDrift(resolved: ResolvedStyle, current: CurrentFormatting): string[] {
  return resolved.targets
    .filter((t) => {
      const cur = current[t.path];
      return cur !== undefined && !valuesEqual(cur, t.value);
    })
    .map((t) => t.path);
}
