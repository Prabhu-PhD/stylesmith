/**
 * Alias-chain cycle + depth checking, enforced at WRITE time (PRD §10.5.2,
 * AC27.4): `a → b → a` must be rejected when the edit is made, not discovered
 * at render time. Resolution has its own runtime guard (./resolve.ts) as a
 * backstop, but the model should never persist a cycle in the first place.
 *
 * Tokens have out-degree ≤ 1 (a token has exactly one value), so an alias
 * "chain" is a simple path — cycle detection is a linear walk.
 *
 * PURE — no Office JS.
 */
import type { Token, ValueKind } from "../schema/types";
import { MAX_ALIAS_DEPTH } from "../schema/defaults";

export type WriteCheck =
  | { readonly ok: true }
  | { readonly ok: false; readonly reason: "cycle" | "depth-exceeded" };

/**
 * Would setting `targetTokenId`'s value to `newValue` create a cycle or an
 * over-deep chain? Walks the alias path starting from `newValue`; a cycle is
 * reaching `targetTokenId` again (or any repeated token). A dangling ref
 * (missing token) is not a cycle — resolution handles that via the cache.
 */
export function checkTokenValueWritable(
  tokens: ReadonlyMap<string, Token>,
  targetTokenId: string,
  newValue: ValueKind,
  maxDepth: number = MAX_ALIAS_DEPTH,
): WriteCheck {
  let current: ValueKind = newValue;
  let depth = 0;
  const visited = new Set<string>([targetTokenId]);

  while (current.kind === "tokenRef") {
    if (visited.has(current.tokenId)) return { ok: false, reason: "cycle" };
    if (depth >= maxDepth) return { ok: false, reason: "depth-exceeded" };
    visited.add(current.tokenId);
    const token = tokens.get(current.tokenId);
    if (!token) return { ok: true }; // dangling ref — not a cycle
    current = token.value;
    depth += 1;
  }
  return { ok: true };
}

/**
 * Find every token that sits on an alias cycle in the given set. Returns one
 * representative path per cycle (the token ids in traversal order). Used to
 * validate an entire document defensively (e.g. after import).
 */
export function findCycles(tokens: ReadonlyMap<string, Token>): string[][] {
  const cycles: string[][] = [];
  const seenCycleKeys = new Set<string>();

  for (const startId of tokens.keys()) {
    const path: string[] = [];
    const onPath = new Set<string>();
    let currentId: string | undefined = startId;

    while (currentId !== undefined) {
      if (onPath.has(currentId)) {
        // Found a cycle — extract it from where it first appears in the path.
        const cycle = path.slice(path.indexOf(currentId));
        const key = [...cycle].sort().join("|");
        if (!seenCycleKeys.has(key)) {
          seenCycleKeys.add(key);
          cycles.push(cycle);
        }
        break;
      }
      const token: Token | undefined = tokens.get(currentId);
      if (!token) break; // dangling
      path.push(currentId);
      onPath.add(currentId);
      currentId = token.value.kind === "tokenRef" ? token.value.tokenId : undefined;
    }
  }
  return cycles;
}
