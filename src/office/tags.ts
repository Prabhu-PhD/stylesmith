/**
 * Linkage layer — the STYLESMITH_ID shape tag that ties a shape to a style GUID.
 *
 * ⚠️ Tag-only writes vanish on web unless the document is dirtied (CLAUDE.md).
 * The dirty guard lives in storage.ts; callers that write tags without a visual
 * change must pair the write with ensureDocumentDirty. Tags also do NOT survive
 * cross-deck paste — adoption (fingerprinting), not tags, handles pasted shapes.
 */

export const TAG_KEY = "STYLESMITH_ID";

/** Queue a linkage-tag write on a shape. Caller syncs (and dirties) at the boundary. */
export function queueTagWrite(shape: PowerPoint.Shape, styleId: string): void {
  shape.tags.add(TAG_KEY, styleId);
}

/** Queue removal of a shape's linkage tag. Caller syncs (and dirties). */
export function queueTagDelete(shape: PowerPoint.Shape): void {
  shape.tags.delete(TAG_KEY);
}

/** Read a shape's style GUID from its already-loaded tags, or null. */
export function readStyleId(shape: PowerPoint.Shape): string | null {
  try {
    const tag = shape.tags.items.find((t) => t.key === TAG_KEY);
    return tag ? tag.value : null;
  } catch {
    return null; // tags not loaded on this proxy
  }
}
