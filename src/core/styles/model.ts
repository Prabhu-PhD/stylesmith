/**
 * CRUD on the style graph, plus the value-iteration helpers the rest of core
 * builds on. Every operation is PURE and returns a NEW document (never mutates
 * input); errors are typed Results, not thrown (CLAUDE.md conventions).
 *
 * Invariants enforced here:
 *   - names are unique (styles deck-wide; tokens within their set) — AC1.3
 *   - ids are stable GUIDs; rename never touches them — AC7.1
 *   - a token value's type matches its declaration; a binding's type matches the
 *     property (AC26.3); themeColor binds only to colour properties
 *   - editing a token value can never introduce an alias cycle — AC27.4
 *
 * PURE — no Office JS.
 */
import type {
  LayerName,
  Layers,
  LiteralPrimitive,
  Result,
  Style,
  StyleSmithDocument,
  TextLayer,
  ShapeLayer,
  Token,
  TokenType,
  ValueKind,
} from "../schema/types";
import { ok, err } from "../schema/types";
import { TEXT_PROPERTIES, SHAPE_PROPERTIES } from "../schema/schemas";
import { clone, tokenTypeForProperty } from "../schema/defaults";
import { checkTokenValueWritable } from "../tokens/cycles";
import { indexTokens, resolveValue, type ResolveContext } from "../tokens/resolve";

/* ── Errors ──────────────────────────────────────────────────────────────── */

export type ModelError =
  | { readonly code: "name-conflict"; readonly name: string }
  | { readonly code: "style-not-found"; readonly id: string }
  | { readonly code: "token-not-found"; readonly id: string }
  | { readonly code: "token-set-not-found"; readonly id: string }
  | { readonly code: "unknown-property"; readonly layer: LayerName; readonly property: string }
  | { readonly code: "not-tokenisable"; readonly property: string }
  | { readonly code: "type-mismatch"; readonly expected: TokenType; readonly got: TokenType }
  | { readonly code: "themecolor-non-color"; readonly property: string }
  | { readonly code: "cycle" }
  | { readonly code: "depth-exceeded" };

/* ── Value iteration ─────────────────────────────────────────────────────── */

export interface ValueEntry {
  readonly layer: LayerName;
  readonly property: string;
  readonly index?: number;
  readonly value: ValueKind;
}

/** Yield every bound value in a style with its location. Table layer is deferred. */
export function* iterateStyleValues(style: Style): Generator<ValueEntry> {
  const { text, shape, geometry } = style.layers;
  if (text) {
    for (const property of TEXT_PROPERTIES) {
      const value = text[property];
      if (value) yield { layer: "text", property, value };
    }
  }
  if (shape) {
    for (const property of SHAPE_PROPERTIES) {
      const value = shape[property];
      if (value) yield { layer: "shape", property, value };
    }
  }
  if (geometry) {
    for (let index = 0; index < geometry.adjustments.length; index += 1) {
      const value = geometry.adjustments[index];
      if (value) yield { layer: "geometry", property: "adjustments", index, value };
    }
  }
}

/** Map every bound value in a style through `fn`, returning a new style. */
export function mapStyleValues(
  style: Style,
  fn: (value: ValueKind, entry: ValueEntry) => ValueKind,
): Style {
  const next = clone(style);
  const { text, shape, geometry } = next.layers;
  if (text) {
    for (const property of TEXT_PROPERTIES) {
      const value = text[property];
      if (value) text[property] = fn(value, { layer: "text", property, value });
    }
  }
  if (shape) {
    for (const property of SHAPE_PROPERTIES) {
      const value = shape[property];
      if (value) shape[property] = fn(value, { layer: "shape", property, value });
    }
  }
  if (geometry) {
    geometry.adjustments = geometry.adjustments.map((value, index) =>
      fn(value, { layer: "geometry", property: "adjustments", index, value }),
    );
  }
  return next;
}

/**
 * Refresh every bound value's inline cache to its currently-resolved literal, so
 * that if the token is later deleted the fallback is a recent real value
 * (AC27.5). Only values that resolve cleanly (not already degraded) are updated.
 */
export function withRefreshedCache(style: Style, ctx: ResolveContext): Style {
  return mapStyleValues(style, (value) => {
    if (value.kind === "literal") return value;
    const resolved = resolveValue(value, ctx);
    if (resolved.ok && !resolved.degraded) {
      return { ...value, cached: resolved.value };
    }
    return value;
  });
}

/* ── Lookup helpers ──────────────────────────────────────────────────────── */

function findTokenLocation(
  doc: StyleSmithDocument,
  tokenId: string,
): { setIndex: number; tokenIndex: number; token: Token } | null {
  for (let s = 0; s < doc.tokenSets.length; s += 1) {
    const set = doc.tokenSets[s];
    if (!set) continue;
    const tokenIndex = set.tokens.findIndex((t) => t.id === tokenId);
    if (tokenIndex >= 0) {
      const token = set.tokens[tokenIndex];
      if (token) return { setIndex: s, tokenIndex, token };
    }
  }
  return null;
}

/* ── Style construction from captured formatting (S1) ────────────────────── */

const litValue = (value: LiteralPrimitive): ValueKind => ({ kind: "literal", value });

/**
 * Build literal style layers from a shape's captured formatting (keyed by
 * property path, e.g. "shape.fill", "geometry.adjustments[0]"). Absent layers
 * are left out (they become null). Everything captured is a literal — promoting
 * a literal to a token is a later, explicit action (S30).
 */
export function layersFromCapture(current: Record<string, LiteralPrimitive>): Partial<Layers> {
  const result: Partial<Layers> = {};

  const text: Record<string, ValueKind> = {};
  for (const property of TEXT_PROPERTIES) {
    const v = current[`text.${property}`];
    if (v !== undefined) text[property] = litValue(v);
  }
  if (Object.keys(text).length > 0) result.text = text as TextLayer;

  const shape: Record<string, ValueKind> = {};
  for (const property of SHAPE_PROPERTIES) {
    const v = current[`shape.${property}`];
    if (v !== undefined) shape[property] = litValue(v);
  }
  if (Object.keys(shape).length > 0) result.shape = shape as ShapeLayer;

  const adjustments: ValueKind[] = [];
  for (let i = 0; ; i += 1) {
    const v = current[`geometry.adjustments[${i}]`];
    if (v === undefined) break;
    adjustments.push(litValue(v));
  }
  if (adjustments.length > 0) result.geometry = { adjustments };

  return result;
}

/* ── Style CRUD ──────────────────────────────────────────────────────────── */

function emptyLayers(): Layers {
  return { text: null, shape: null, geometry: null, table: null };
}

export interface NewStyleInput {
  readonly name: string;
  readonly layers?: Partial<Layers>;
}

export function createStyle(
  doc: StyleSmithDocument,
  input: NewStyleInput,
  makeId: () => string,
): Result<{ doc: StyleSmithDocument; styleId: string }, ModelError> {
  if (doc.styles.some((s) => s.name === input.name)) {
    return err({ code: "name-conflict", name: input.name });
  }
  const id = makeId();
  const style: Style = {
    id,
    name: input.name,
    origin: null,
    basedOn: null,
    layers: { ...emptyLayers(), ...input.layers },
  };
  const next = clone(doc);
  next.styles.push(style);
  return ok({ doc: next, styleId: id });
}

export function renameStyle(
  doc: StyleSmithDocument,
  styleId: string,
  name: string,
): Result<StyleSmithDocument, ModelError> {
  const target = doc.styles.find((s) => s.id === styleId);
  if (!target) return err({ code: "style-not-found", id: styleId });
  if (doc.styles.some((s) => s.id !== styleId && s.name === name)) {
    return err({ code: "name-conflict", name });
  }
  const next = clone(doc);
  const t = next.styles.find((s) => s.id === styleId);
  if (t) t.name = name;
  return ok(next);
}

export function deleteStyle(
  doc: StyleSmithDocument,
  styleId: string,
): Result<StyleSmithDocument, ModelError> {
  if (!doc.styles.some((s) => s.id === styleId)) {
    return err({ code: "style-not-found", id: styleId });
  }
  const next = clone(doc);
  next.styles = next.styles.filter((s) => s.id !== styleId);
  return ok(next);
}

/** Duplicate a style under a fresh id and a unique "… copy" name (S8). */
export function duplicateStyle(
  doc: StyleSmithDocument,
  styleId: string,
  makeId: () => string,
): Result<{ doc: StyleSmithDocument; styleId: string }, ModelError> {
  const src = doc.styles.find((s) => s.id === styleId);
  if (!src) return err({ code: "style-not-found", id: styleId });

  let name = `${src.name} copy`;
  let n = 2;
  while (doc.styles.some((s) => s.name === name)) name = `${src.name} copy ${n++}`;

  const id = makeId();
  const copy: Style = { ...clone(src), id, name, basedOn: null };
  const next = clone(doc);
  next.styles.push(copy);
  return ok({ doc: next, styleId: id });
}

/* ── Property binding ────────────────────────────────────────────────────── */

/** Validate that `value` may bind to `layer.property`. */
function checkBinding(
  doc: StyleSmithDocument,
  layer: LayerName,
  property: string,
  value: ValueKind,
): Result<void, ModelError> {
  const expected = tokenTypeForProperty(layer, property);
  if (expected === undefined) return err({ code: "unknown-property", layer, property });

  if (value.kind === "literal") return ok(undefined);

  if (expected === null) return err({ code: "not-tokenisable", property });

  if (value.kind === "themeColor") {
    if (expected !== "color") return err({ code: "themecolor-non-color", property });
    return ok(undefined);
  }

  // tokenRef — the referenced token's type must match the property's type.
  const token = indexTokens(doc).get(value.tokenId);
  if (token && token.type !== expected) {
    return err({ code: "type-mismatch", expected, got: token.type });
  }
  return ok(undefined);
}

/**
 * Bind a text/shape scalar property to a value (literal, tokenRef, or
 * themeColor). Creates the layer object if it was null. Geometry adjustments are
 * bound separately (they are an ordered array, not a named scalar).
 */
export function bindProperty(
  doc: StyleSmithDocument,
  styleId: string,
  layer: "text" | "shape",
  property: string,
  value: ValueKind,
): Result<StyleSmithDocument, ModelError> {
  const style = doc.styles.find((s) => s.id === styleId);
  if (!style) return err({ code: "style-not-found", id: styleId });

  const check = checkBinding(doc, layer, property, value);
  if (!check.ok) return check;

  const next = clone(doc);
  const target = next.styles.find((s) => s.id === styleId);
  if (!target) return err({ code: "style-not-found", id: styleId });

  const bag = (target.layers[layer] ??= {}) as Record<string, ValueKind>;
  bag[property] = value;
  return ok(next);
}

/** Replace a property binding with a literal, keeping the current value (S31). */
export function unbindToLiteral(
  doc: StyleSmithDocument,
  styleId: string,
  layer: "text" | "shape",
  property: string,
  value: string | number,
): Result<StyleSmithDocument, ModelError> {
  return bindProperty(doc, styleId, layer, property, { kind: "literal", value });
}

/* ── Token CRUD ──────────────────────────────────────────────────────────── */

/** Validate a token value against the token's declared type. */
function checkTokenValue(
  tokens: ReadonlyMap<string, Token>,
  ownType: TokenType,
  value: ValueKind,
): Result<void, ModelError> {
  if (value.kind === "themeColor") {
    if (ownType !== "color") return err({ code: "themecolor-non-color", property: "(token)" });
    return ok(undefined);
  }
  if (value.kind === "tokenRef") {
    const target = tokens.get(value.tokenId);
    if (target && target.type !== ownType) {
      return err({ code: "type-mismatch", expected: ownType, got: target.type });
    }
  }
  return ok(undefined);
}

export interface NewTokenInput {
  readonly name: string;
  readonly type: TokenType;
  readonly value: ValueKind;
  readonly description?: string;
}

export function createToken(
  doc: StyleSmithDocument,
  tokenSetId: string,
  input: NewTokenInput,
  makeId: () => string,
): Result<{ doc: StyleSmithDocument; tokenId: string }, ModelError> {
  const set = doc.tokenSets.find((s) => s.id === tokenSetId);
  if (!set) return err({ code: "token-set-not-found", id: tokenSetId });
  if (set.tokens.some((t) => t.name === input.name)) {
    return err({ code: "name-conflict", name: input.name });
  }

  const tokens = indexTokens(doc);
  const typeCheck = checkTokenValue(tokens, input.type, input.value);
  if (!typeCheck.ok) return typeCheck;

  const id = makeId();
  const cycle = checkTokenValueWritable(tokens, id, input.value);
  if (!cycle.ok) return err({ code: cycle.reason });

  const token: Token = {
    id,
    name: input.name,
    type: input.type,
    value: input.value,
    ...(input.description !== undefined ? { description: input.description } : {}),
  };
  const next = clone(doc);
  const targetSet = next.tokenSets.find((s) => s.id === tokenSetId);
  if (!targetSet) return err({ code: "token-set-not-found", id: tokenSetId });
  targetSet.tokens.push(token);
  return ok({ doc: next, tokenId: id });
}

export function editTokenValue(
  doc: StyleSmithDocument,
  tokenId: string,
  value: ValueKind,
): Result<StyleSmithDocument, ModelError> {
  const loc = findTokenLocation(doc, tokenId);
  if (!loc) return err({ code: "token-not-found", id: tokenId });

  const tokens = indexTokens(doc);
  const typeCheck = checkTokenValue(tokens, loc.token.type, value);
  if (!typeCheck.ok) return typeCheck;

  const cycle = checkTokenValueWritable(tokens, tokenId, value);
  if (!cycle.ok) return err({ code: cycle.reason });

  const next = clone(doc);
  const nloc = findTokenLocation(next, tokenId);
  if (nloc) {
    const set = next.tokenSets[nloc.setIndex];
    const token = set?.tokens[nloc.tokenIndex];
    if (token) token.value = value;
  }
  return ok(next);
}

export function renameToken(
  doc: StyleSmithDocument,
  tokenId: string,
  name: string,
): Result<StyleSmithDocument, ModelError> {
  const loc = findTokenLocation(doc, tokenId);
  if (!loc) return err({ code: "token-not-found", id: tokenId });
  const set = doc.tokenSets[loc.setIndex];
  if (set && set.tokens.some((t) => t.id !== tokenId && t.name === name)) {
    return err({ code: "name-conflict", name });
  }
  const next = clone(doc);
  const nloc = findTokenLocation(next, tokenId);
  if (nloc) {
    const token = next.tokenSets[nloc.setIndex]?.tokens[nloc.tokenIndex];
    if (token) token.name = name;
  }
  return ok(next);
}

/**
 * Delete a token. Styles that referenced it are NOT rewritten — their inline
 * `cached` fallback keeps them rendering (AC27.5). Flagging the now-degraded
 * styles is a read-time concern (resolve + usage), not a delete-time rewrite.
 */
export function deleteToken(
  doc: StyleSmithDocument,
  tokenId: string,
): Result<StyleSmithDocument, ModelError> {
  const loc = findTokenLocation(doc, tokenId);
  if (!loc) return err({ code: "token-not-found", id: tokenId });
  const next = clone(doc);
  const set = next.tokenSets[loc.setIndex];
  if (set) set.tokens = set.tokens.filter((t) => t.id !== tokenId);
  return ok(next);
}
