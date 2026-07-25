/**
 * Token usage — which styles and which other tokens reference a given token.
 * Powers usage counts (S29) and the cascade preview (S27: "every affected style
 * is listed with its shape count and a correct total", AC27.1).
 *
 * "Affected" is TRANSITIVE through alias chains: editing token B affects a style
 * bound to token A when A → B. Direct references and transitive impact are both
 * exposed.
 *
 * Shape counts are NOT computed here — they require an Office scan (Phase 2+).
 * This layer stays pure and returns style-level facts only.
 *
 * PURE — no Office JS.
 */
import type { StyleSmithDocument, Token } from "../schema/types";
import { iterateStyleValues } from "../styles/model";

export interface StyleTokenUsage {
  readonly styleId: string;
  /** Property paths in this style that reference the token directly, e.g. "shape.fill", "geometry.adjustments[0]". */
  readonly properties: string[];
}

export interface TokenUsage {
  readonly tokenId: string;
  /** Styles referencing the token DIRECTLY. */
  readonly directStyles: StyleTokenUsage[];
  /** Ids of other tokens whose value references this token directly (alias parents). */
  readonly aliasedBy: string[];
  /** Styles affected if the token changes, DIRECTLY or via an alias chain. */
  readonly affectedStyleIds: string[];
  /** Count of affected styles (the number shown before an edit). */
  readonly affectedStyleCount: number;
}

/** Property path label for a value entry. */
function pathOf(layer: string, property: string, index?: number): string {
  return index === undefined ? `${layer}.${property}` : `${layer}.${property}[${index}]`;
}

/** Tokens whose own value directly references `tokenId` (one alias hop up). */
function directAliasParents(tokens: readonly Token[], tokenId: string): string[] {
  return tokens
    .filter((t) => t.value.kind === "tokenRef" && t.value.tokenId === tokenId)
    .map((t) => t.id);
}

/** The set of token ids that reach `tokenId` through alias chains (incl. itself). */
function tokensReaching(allTokens: readonly Token[], tokenId: string): Set<string> {
  const reaching = new Set<string>([tokenId]);
  let grew = true;
  while (grew) {
    grew = false;
    // Fixpoint: add any token that aliases something already in `reaching`.
    for (const t of allTokens) {
      if (reaching.has(t.id)) continue;
      if (t.value.kind === "tokenRef" && reaching.has(t.value.tokenId)) {
        reaching.add(t.id);
        grew = true;
      }
    }
  }
  return reaching;
}

/** Compute direct + transitive usage of a token across the document. */
export function computeTokenUsage(doc: StyleSmithDocument, tokenId: string): TokenUsage {
  const allTokens = doc.tokenSets.flatMap((s) => s.tokens);

  const directStyles: StyleTokenUsage[] = [];
  for (const style of doc.styles) {
    const properties: string[] = [];
    for (const entry of iterateStyleValues(style)) {
      if (entry.value.kind === "tokenRef" && entry.value.tokenId === tokenId) {
        properties.push(pathOf(entry.layer, entry.property, entry.index));
      }
    }
    if (properties.length > 0) directStyles.push({ styleId: style.id, properties });
  }

  const reaching = tokensReaching(allTokens, tokenId);
  const affected = new Set<string>();
  for (const style of doc.styles) {
    for (const entry of iterateStyleValues(style)) {
      if (entry.value.kind === "tokenRef" && reaching.has(entry.value.tokenId)) {
        affected.add(style.id);
        break;
      }
    }
  }

  const affectedStyleIds = [...affected];
  return {
    tokenId,
    directStyles,
    aliasedBy: directAliasParents(allTokens, tokenId),
    affectedStyleIds,
    affectedStyleCount: affectedStyleIds.length,
  };
}

/** Total number of styles that directly reference a token (S29 list count). */
export function directStyleUsageCount(doc: StyleSmithDocument, tokenId: string): number {
  return computeTokenUsage(doc, tokenId).directStyles.length;
}
