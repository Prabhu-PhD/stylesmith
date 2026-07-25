/**
 * Group shapes by formatting signature. Shapes with an identical signature hash
 * form a cluster — the candidate for "N shapes share one format. Create a style
 * from these?" (adoption step 2, AC15.2).
 *
 * PURE — no Office JS.
 */
import type { Formatting, Signature } from "./signature";
import { computeSignature, type SignatureOptions } from "./signature";

/** An item paired with the formatting to fingerprint it by. */
export interface Fingerprinted<T> {
  readonly item: T;
  readonly formatting: Formatting;
}

export interface Cluster<T> {
  readonly hash: string;
  /** The representative signature (shared by all members). */
  readonly signature: Signature;
  /** Representative formatting (the first member's) — used to build the style. */
  readonly representative: Formatting;
  readonly members: T[];
  readonly size: number;
}

export interface ClusterOptions extends SignatureOptions {
  /** Minimum members to surface a cluster (default 2 — singletons aren't useful). */
  readonly minSize?: number;
}

/** Cluster items by identical signature, largest first. */
export function clusterBySignature<T>(items: Fingerprinted<T>[], opts?: ClusterOptions): Cluster<T>[] {
  const minSize = opts?.minSize ?? 2;
  const groups = new Map<string, Fingerprinted<T>[]>();

  for (const entry of items) {
    const sig = computeSignature(entry.formatting, opts);
    const bucket = groups.get(sig.hash);
    if (bucket) bucket.push(entry);
    else groups.set(sig.hash, [entry]);
  }

  const clusters: Cluster<T>[] = [];
  for (const bucket of groups.values()) {
    if (bucket.length < minSize) continue;
    const first = bucket[0];
    if (!first) continue;
    clusters.push({
      hash: computeSignature(first.formatting, opts).hash,
      signature: computeSignature(first.formatting, opts),
      representative: first.formatting,
      members: bucket.map((b) => b.item),
      size: bucket.length,
    });
  }
  return clusters.sort((a, b) => b.size - a.size);
}
