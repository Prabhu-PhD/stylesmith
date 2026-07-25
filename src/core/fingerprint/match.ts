/**
 * Exact / near matching with per-kind tolerance.
 *
 * Compares a shape's formatting against a style's resolved target values and
 * reports the SPECIFIC deviating properties — this is what makes adoption
 * trustworthy rather than magic (AC15.4: "⌒ 0.18 (style: 0.15)"). It is also the
 * comparison BrandGuard performs: one engine, two products.
 *
 * PURE — no Office JS.
 */
import type { LiteralPrimitive } from "../schema/types";
import type { Formatting } from "./signature";
import { valuesEqual } from "../styles/diff";

export interface Deviation {
  readonly path: string;
  readonly shapeValue: LiteralPrimitive;
  readonly styleValue: LiteralPrimitive;
}

export interface MatchResult {
  /** Every compared property matches exactly. */
  readonly exact: boolean;
  /** Not exact, but every deviating property is within tolerance. */
  readonly near: boolean;
  /** The properties that differ (present on both, unequal). */
  readonly deviations: Deviation[];
  /** How many of the style's properties were compared. */
  readonly compared: number;
}

export interface Tolerance {
  /** Max RGB euclidean distance for two colours to be "near" (default 16). */
  readonly colourDistance?: number;
  /** Max absolute numeric delta (default 0.03, e.g. a corner radius nudge). */
  readonly numberAbsolute?: number;
  /** Max relative numeric delta (default 0.1 — 10%, e.g. 14pt vs 15pt). */
  readonly numberRelative?: number;
}

const DEFAULT_TOLERANCE: Required<Tolerance> = {
  colourDistance: 16,
  numberAbsolute: 0.03,
  numberRelative: 0.1,
};

function parseHex(s: string): [number, number, number] | null {
  const m = /^#?([0-9a-f]{6})$/i.exec(s.trim());
  if (!m || !m[1]) return null;
  const n = parseInt(m[1], 16);
  return [(n >> 16) & 0xff, (n >> 8) & 0xff, n & 0xff];
}

function withinTolerance(a: LiteralPrimitive, b: LiteralPrimitive, tol: Required<Tolerance>): boolean {
  if (typeof a === "number" && typeof b === "number") {
    const diff = Math.abs(a - b);
    return diff <= tol.numberAbsolute || diff <= tol.numberRelative * Math.max(1, Math.abs(b));
  }
  const ca = parseHex(String(a));
  const cb = parseHex(String(b));
  if (ca && cb) {
    const d = Math.sqrt((ca[0] - cb[0]) ** 2 + (ca[1] - cb[1]) ** 2 + (ca[2] - cb[2]) ** 2);
    return d <= tol.colourDistance;
  }
  return false; // non-numeric, non-colour strings must match exactly (handled by caller)
}

/**
 * Match a shape's formatting against a style's resolved target values. The
 * target's properties define the comparison set. A property the shape lacks
 * counts as a hard mismatch (breaks near).
 */
export function matchFormatting(
  shape: Formatting,
  target: Formatting,
  tolerance?: Tolerance,
): MatchResult {
  const tol = { ...DEFAULT_TOLERANCE, ...tolerance };
  const paths = Object.keys(target);
  const deviations: Deviation[] = [];
  let allExact = true;
  let allWithin = true;

  for (const path of paths) {
    const styleValue = target[path];
    const shapeValue = shape[path];
    if (styleValue === undefined) continue;

    if (shapeValue === undefined) {
      allExact = false;
      allWithin = false;
      continue;
    }
    if (valuesEqual(shapeValue, styleValue)) continue;

    allExact = false;
    deviations.push({ path, shapeValue, styleValue });
    if (!withinTolerance(shapeValue, styleValue, tol)) allWithin = false;
  }

  return {
    exact: allExact && deviations.length === 0,
    near: !allExact && allWithin && deviations.length > 0,
    deviations,
    compared: paths.length,
  };
}
