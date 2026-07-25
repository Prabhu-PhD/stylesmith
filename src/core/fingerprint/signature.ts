/**
 * Formatting → comparable signature.
 *
 * The adoption engine works precisely because it depends on NO identifier
 * (CLAUDE.md: tags don't survive paste, nothing does). Instead we compare a
 * shape's actual formatting. A signature normalises each property to a canonical
 * string and joins them into a stable hash for exact clustering; the normalised
 * per-property map is kept for computing near-match deltas.
 *
 * PURE — no Office JS. Signatures are ephemeral (computed on demand, never
 * persisted), so this introduces no schema surface.
 */
import type { LiteralPrimitive } from "../schema/types";

/** A shape's formatting keyed by property path (e.g. "shape.fill"). */
export type Formatting = Record<string, LiteralPrimitive>;

export interface Signature {
  /** Canonical per-property strings, for delta computation on near matches. */
  readonly props: Record<string, string>;
  /** Stable exact-match key: sorted, joined canonical props. */
  readonly hash: string;
}

export interface SignatureOptions {
  /** Numbers are rounded to this quantum before hashing (default 0.01). */
  readonly numberQuantum?: number;
}

/** Canonicalise a single value: quantise numbers; normalise strings/colours. */
export function canonicalize(value: LiteralPrimitive, numberQuantum = 0.01): string {
  if (typeof value === "number") {
    const q = numberQuantum > 0 ? numberQuantum : 0.01;
    return (Math.round(value / q) * q).toFixed(4);
  }
  return value.trim().toLowerCase().replace(/^#/, "");
}

/** Compute a signature from a shape's formatting. */
export function computeSignature(formatting: Formatting, opts?: SignatureOptions): Signature {
  const quantum = opts?.numberQuantum ?? 0.01;
  const props: Record<string, string> = {};
  for (const key of Object.keys(formatting)) {
    const v = formatting[key];
    if (v !== undefined) props[key] = canonicalize(v, quantum);
  }
  const hash = Object.keys(props)
    .sort()
    .map((k) => `${k}=${props[k]}`)
    .join("|");
  return { props, hash };
}
