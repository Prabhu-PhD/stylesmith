import { describe, it, expect } from "vitest";
import type { StyleSmithDocument, ValueKind } from "../src/core/schema/types";
import { createEmptyDocument } from "../src/core/schema/defaults";
import {
  createStyle,
  renameStyle,
  deleteStyle,
  createToken,
  editTokenValue,
  renameToken,
  deleteToken,
  bindProperty,
  withRefreshedCache,
  iterateStyleValues,
} from "../src/core/styles/model";
import { indexTokens } from "../src/core/tokens/resolve";

const lit = (value: string | number): ValueKind => ({ kind: "literal", value });
const ref = (tokenId: string): ValueKind => ({ kind: "tokenRef", tokenId });

/** A doc with one set + a colour token "brand"; returns the doc and ids. */
function seeded(): { doc: StyleSmithDocument; setId: string } {
  const doc = createEmptyDocument(() => "set");
  return { doc, setId: "set" };
}

describe("style CRUD", () => {
  it("creates a style and returns its id; input is not mutated", () => {
    const { doc } = seeded();
    const before = JSON.stringify(doc);
    const r = createStyle(doc, { name: "Takeaway" }, () => "s1");
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.value.styleId).toBe("s1");
      expect(r.value.doc.styles).toHaveLength(1);
    }
    expect(JSON.stringify(doc)).toBe(before); // purity
  });

  it("rejects a duplicate style name", () => {
    let doc = createEmptyDocument(() => "set");
    doc = (createStyle(doc, { name: "A" }, () => "s1") as { value: { doc: StyleSmithDocument } }).value.doc;
    const r = createStyle(doc, { name: "A" }, () => "s2");
    expect(r).toEqual({ ok: false, error: { code: "name-conflict", name: "A" } });
  });

  it("renames a style, preserving its id (GUID keying)", () => {
    let doc = createEmptyDocument(() => "set");
    const created = createStyle(doc, { name: "Old" }, () => "s1");
    if (!created.ok) throw new Error("setup");
    doc = created.value.doc;
    const r = renameStyle(doc, "s1", "New");
    expect(r.ok).toBe(true);
    if (r.ok) {
      expect(r.value.styles[0]?.id).toBe("s1");
      expect(r.value.styles[0]?.name).toBe("New");
    }
  });

  it("deletes a style", () => {
    let doc = createEmptyDocument(() => "set");
    doc = (createStyle(doc, { name: "A" }, () => "s1") as { value: { doc: StyleSmithDocument } }).value.doc;
    const r = deleteStyle(doc, "s1");
    expect(r.ok && r.value.styles).toHaveLength(0);
  });
});

describe("token CRUD + type safety", () => {
  it("creates a token", () => {
    const { doc, setId } = seeded();
    const r = createToken(doc, setId, { name: "brand", type: "color", value: lit("#000") }, () => "t1");
    expect(r.ok).toBe(true);
    if (r.ok) expect(indexTokens(r.value.doc).get("t1")?.name).toBe("brand");
  });

  it("rejects a themeColor value on a non-colour token", () => {
    const { doc, setId } = seeded();
    const r = createToken(
      doc,
      setId,
      { name: "size", type: "fontSize", value: { kind: "themeColor", slot: "accent1" } },
      () => "t1",
    );
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("themecolor-non-color");
  });

  it("rejects an alias to a token of a different type", () => {
    let doc = createEmptyDocument(() => "set");
    doc = (createToken(doc, "set", { name: "c", type: "color", value: lit("#000") }, () => "c1") as { value: { doc: StyleSmithDocument } }).value.doc;
    const r = createToken(doc, "set", { name: "r", type: "radius", value: ref("c1") }, () => "r1");
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("type-mismatch");
  });

  it("rejects editing a token value into a cycle", () => {
    let doc = createEmptyDocument(() => "set");
    doc = (createToken(doc, "set", { name: "a", type: "color", value: lit("#000") }, () => "a") as { value: { doc: StyleSmithDocument } }).value.doc;
    doc = (createToken(doc, "set", { name: "b", type: "color", value: ref("a") }, () => "b") as { value: { doc: StyleSmithDocument } }).value.doc;
    // a -> b would close the loop a -> b -> a
    const r = editTokenValue(doc, "a", ref("b"));
    expect(r).toEqual({ ok: false, error: { code: "cycle" } });
  });

  it("renames a token, preserving its id", () => {
    let doc = createEmptyDocument(() => "set");
    doc = (createToken(doc, "set", { name: "old", type: "color", value: lit("#000") }, () => "t1") as { value: { doc: StyleSmithDocument } }).value.doc;
    const r = renameToken(doc, "t1", "new");
    expect(r.ok && indexTokens(r.value).get("t1")?.name).toBe("new");
  });
});

describe("property binding", () => {
  function docWithColorToken(): StyleSmithDocument {
    let doc = createEmptyDocument(() => "set");
    doc = (createToken(doc, "set", { name: "brand", type: "color", value: lit("#0057B8") }, () => "col") as { value: { doc: StyleSmithDocument } }).value.doc;
    doc = (createToken(doc, "set", { name: "radius", type: "radius", value: lit(0.15) }, () => "rad") as { value: { doc: StyleSmithDocument } }).value.doc;
    doc = (createStyle(doc, { name: "S" }, () => "s1") as { value: { doc: StyleSmithDocument } }).value.doc;
    return doc;
  }

  it("binds a colour token to shape.fill and creates the layer", () => {
    const r = bindProperty(docWithColorToken(), "s1", "shape", "fill", ref("col"));
    expect(r.ok).toBe(true);
    if (r.ok) expect(r.value.styles[0]?.layers.shape?.fill).toEqual(ref("col"));
  });

  it("rejects a token of the wrong type for the property (AC26.3)", () => {
    const r = bindProperty(docWithColorToken(), "s1", "shape", "fill", ref("rad"));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("type-mismatch");
  });

  it("rejects a themeColor on a non-colour property", () => {
    const r = bindProperty(docWithColorToken(), "s1", "text", "fontSize", { kind: "themeColor", slot: "accent1" });
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("themecolor-non-color");
  });

  it("rejects a tokenRef on a non-tokenisable property (alignment)", () => {
    const r = bindProperty(docWithColorToken(), "s1", "text", "alignment", ref("col"));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("not-tokenisable");
  });

  it("allows a literal on a non-tokenisable property", () => {
    const r = bindProperty(docWithColorToken(), "s1", "text", "alignment", lit("center"));
    expect(r.ok).toBe(true);
  });

  it("rejects an unknown property", () => {
    const r = bindProperty(docWithColorToken(), "s1", "text", "notAProp", lit("x"));
    expect(r.ok).toBe(false);
    if (!r.ok) expect(r.error.code).toBe("unknown-property");
  });
});

describe("cache lifecycle (AC27.5)", () => {
  function boundDoc(): StyleSmithDocument {
    let doc = createEmptyDocument(() => "set");
    doc = (createToken(doc, "set", { name: "brand", type: "color", value: lit("#0057B8") }, () => "col") as { value: { doc: StyleSmithDocument } }).value.doc;
    doc = (createStyle(doc, { name: "S" }, () => "s1") as { value: { doc: StyleSmithDocument } }).value.doc;
    doc = (bindProperty(doc, "s1", "shape", "fill", ref("col")) as { value: StyleSmithDocument }).value;
    return doc;
  }

  it("withRefreshedCache stamps the resolved literal onto the binding", () => {
    const doc = boundDoc();
    const style = doc.styles[0];
    if (!style) throw new Error("setup");
    const refreshed = withRefreshedCache(style, { tokens: indexTokens(doc) });
    const fill = refreshed.layers.shape?.fill;
    expect(fill?.kind).toBe("tokenRef");
    if (fill?.kind === "tokenRef") expect(fill.cached).toBe("#0057B8");
  });

  it("deleting a token leaves referencing styles intact (they keep the cached fallback)", () => {
    let doc = boundDoc();
    // Refresh cache first so a real fallback exists, then delete the token.
    const style = doc.styles[0];
    if (!style) throw new Error("setup");
    doc = { ...doc, styles: [withRefreshedCache(style, { tokens: indexTokens(doc) })] };
    const afterDelete = deleteToken(doc, "col");
    expect(afterDelete.ok).toBe(true);
    if (afterDelete.ok) {
      const fill = afterDelete.value.styles[0]?.layers.shape?.fill;
      expect(fill?.kind).toBe("tokenRef"); // binding preserved, not blanked
      if (fill?.kind === "tokenRef") expect(fill.cached).toBe("#0057B8");
    }
  });
});

describe("iterateStyleValues", () => {
  it("walks text, shape and geometry bindings", () => {
    let doc = createEmptyDocument(() => "set");
    doc = (createStyle(
      doc,
      {
        name: "S",
        layers: {
          text: { color: lit("#fff") },
          shape: { fill: lit("#000") },
          geometry: { adjustments: [lit(0.1), lit(0.2)] },
          table: null,
        },
      },
      () => "s1",
    ) as { value: { doc: StyleSmithDocument } }).value.doc;

    const style = doc.styles[0];
    if (!style) throw new Error("setup");
    const paths = [...iterateStyleValues(style)].map((e) =>
      e.index === undefined ? `${e.layer}.${e.property}` : `${e.layer}.${e.property}[${e.index}]`,
    );
    expect(paths).toEqual(["text.color", "shape.fill", "geometry.adjustments[0]", "geometry.adjustments[1]"]);
  });
});
