/**
 * PHASE 0 SPIKES — live-host probes for the three open questions in CLAUDE.md.
 *
 * These require a running PowerPoint host and, by their nature, manual steps
 * (group some shapes, close/reopen the deck, press Ctrl+D). They cannot be run
 * headlessly — the entire API contract in CLAUDE.md was established this way.
 *
 * Each function returns a SpikeResult (never throws): a boolean plus a human
 * log the SpikePanel renders verbatim. All Office JS stays in this file.
 *
 * Open questions probed here:
 *   1. Group traversal    — can a group's children be enumerated?
 *   2. Deck-level storage  — PowerPoint has NO document-level customXmlParts
 *                            (that Common API is Word/Excel only). So we probe
 *                            what IS available and test whether a write to
 *                            document.settings survives reopen, and whether it
 *                            needs an explicit dirtying change to do so.
 *   3. Ctrl+D tag survival — does a within-deck duplicate carry its tags?
 */
import { runPowerPoint, isHostAvailable } from "./context";
import { TAG_KEY } from "./tags";

export interface SpikeResult {
  ok: boolean;
  lines: string[];
}

const SPIKE2_SETTING_KEY = "STYLESMITH_SPIKE2";
const SPIKE2_MARKER = "STYLESMITH_SPIKE2_VALUE";
const SPIKE3_TAG_KEY = "STYLESMITH_SPIKE3";

function errText(e: unknown): string {
  if (e instanceof Error) return e.message;
  if (typeof e === "object" && e !== null) {
    const oe = e as { code?: string; message?: string };
    if (oe.code || oe.message) return `${oe.code ?? "Error"}: ${oe.message ?? ""}`.trim();
  }
  return String(e);
}

/** Wrap a spike body so it always resolves to a SpikeResult with a readable log. */
async function guard(body: () => Promise<string[]>): Promise<SpikeResult> {
  if (!isHostAvailable()) {
    return {
      ok: false,
      lines: ["Not running inside PowerPoint. Sideload the add-in and open a deck, then rerun."],
    };
  }
  try {
    return { ok: true, lines: await body() };
  } catch (e) {
    return { ok: false, lines: [`Spike failed: ${errText(e)}`] };
  }
}

/* ── Common-API promisifier (custom XML parts use the classic callback API) ── */
function promisify<T>(
  op: (cb: (res: Office.AsyncResult<T>) => void) => void,
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    op((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error);
    });
  });
}

/* ────────────────────────────────────────────────────────────────────────
   SPIKE 1 — Group traversal
   Precondition: at least one group on some slide (select 2+ shapes, Ctrl+G).
   Reads every shape, finds groups by type, and tries to enumerate children.
   ──────────────────────────────────────────────────────────────────────── */
export function spikeGroupTraversal(): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    await runPowerPoint(async (ctx) => {
      // Batched read (3 syncs), per the verified enumeration pattern.
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();

      const cols = slides.items.map((s) => {
        s.shapes.load("items/name,items/id,items/type");
        return s.shapes;
      });
      await ctx.sync();

      const groups: PowerPoint.Shape[] = [];
      cols.forEach((c) =>
        c.items.forEach((sh) => {
          if (sh.type === PowerPoint.ShapeType.group) groups.push(sh);
        }),
      );

      lines.push(`Scanned ${cols.reduce((n, c) => n + c.items.length, 0)} top-level shapes across ${slides.items.length} slide(s).`);
      lines.push(`Group shapes found: ${groups.length}.`);
      if (groups.length === 0) {
        lines.push("PRECONDITION UNMET: create a group (select 2+ shapes, Ctrl+G) on any slide and rerun.");
        return;
      }

      // Attempt to reach children via Shape.group.shapes. Queue the loads, sync once.
      const childCols = groups.map((g) => {
        const grp = g.group; // ShapeGroup
        grp.shapes.load("items/name,items/type,items/id");
        return grp.shapes;
      });
      await ctx.sync();

      groups.forEach((g, i) => {
        const items = childCols[i]?.items ?? [];
        lines.push(
          `  "${g.name}": ${items.length} child shape(s)` +
            (items.length ? ` — types: ${items.map((c) => c.type).join(", ")}` : ""),
        );
      });

      const anyChildren = childCols.some((c) => (c?.items.length ?? 0) > 0);
      lines.push(
        anyChildren
          ? "RESULT: group children ARE enumerable via Shape.group.shapes."
          : "RESULT: group present but children NOT enumerated — sweeps must report skipped-in-group counts.",
      );
    });
    return lines;
  });
}

/* ────────────────────────────────────────────────────────────────────────
   SPIKE 2 — Deck-level storage + the dirty flag.

   FINDING (23 Jul re-run): Office.context.document.customXmlParts is undefined
   in PowerPoint — that Common API exists only in Word/Excel. The deck-level
   store that IS available is Office.context.document.settings, so we probe the
   surface and test settings persistence instead.

     probe   — report which storage APIs actually exist in this host.
     2a      — write a marked setting + saveAsync, NO dirtying change.
     2a′     — same, but also dirty the document (shape nudge).
     2b      — after close+reopen, is the setting still there?
   Comparing 2a vs 2a′ across reopen isolates whether saveAsync alone persists.
   ──────────────────────────────────────────────────────────────────────── */

/** Report which deck-level storage surfaces this host actually exposes. */
export function spikeProbeStorage(): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    const doc = Office.context.document;
    lines.push(`document.settings:       ${doc.settings ? "available ✓" : "MISSING"}`);
    lines.push(
      `document.customXmlParts: ${doc.customXmlParts ? "available" : "MISSING — Word/Excel only, not PowerPoint"}`,
    );
    lines.push(
      `requirement set "CustomXmlParts": ${Office.context.requirements.isSetSupported("CustomXmlParts")}`,
    );
    lines.push(
      "NOTE: deck-level custom XML parts are unavailable in PowerPoint. document.settings is the deck-level store; per-shape tags remain the linkage layer.",
    );
    return lines;
  });
}

export function spikeWriteSettings(dirty: boolean): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    const settings = Office.context.document.settings;
    settings.set(SPIKE2_SETTING_KEY, { marker: SPIKE2_MARKER, dirtied: dirty });
    await promisify<void>((cb) => settings.saveAsync(cb));
    lines.push(`Wrote setting "${SPIKE2_SETTING_KEY}" and called saveAsync.`);

    if (dirty) {
      // Nudge a shape to its own value — dirties the doc without a visible change.
      await runPowerPoint(async (ctx) => {
        const target = await firstTargetShape(ctx);
        if (!target) {
          lines.push("Could not find a shape to nudge — doc may not be dirtied.");
          return;
        }
        target.load("left,name");
        await ctx.sync();
        target.left = target.left;
        await ctx.sync();
        lines.push(`Dirtied the document by nudging shape "${target.name}".`);
      });
    } else {
      lines.push("Did NOT dirty the document (settings.saveAsync only).");
    }

    lines.push("NEXT: close the deck, reopen it, then run '2b check survived'.");
    return lines;
  });
}

export function spikeCheckSettings(): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    const raw = Office.context.document.settings.get(SPIKE2_SETTING_KEY) as
      | { marker?: string; dirtied?: boolean }
      | null;
    if (!raw) {
      lines.push(
        "RESULT: setting NOT found. If this followed the no-dirty write, settings.saveAsync alone does NOT persist across reopen — the central dirty guard must accompany every save.",
      );
      return lines;
    }
    lines.push(`Found setting: marker=${raw.marker} dirtied=${String(raw.dirtied)}`);
    lines.push(raw.marker === SPIKE2_MARKER ? "marker matches ✓" : "marker MISMATCH");
    lines.push(
      "RESULT: setting survived reopen. If this followed the no-dirty write, settings.saveAsync persists on its own.",
    );
    return lines;
  });
}

/** Clear the spike-2 setting between runs so results stay unambiguous. */
export function spikeClearSettings(): Promise<SpikeResult> {
  return guard(async () => {
    const settings = Office.context.document.settings;
    settings.remove(SPIKE2_SETTING_KEY);
    await promisify<void>((cb) => settings.saveAsync(cb));
    return [`Removed setting "${SPIKE2_SETTING_KEY}" and saved.`];
  });
}

/* ────────────────────────────────────────────────────────────────────────
   SPIKE 3 — Do tags survive a within-deck Ctrl+D duplicate?
     3a tagSelectedShapes() — tag the current selection, then user presses Ctrl+D.
     3b scanSpikeTags()     — list every shape carrying the spike tag.
   A duplicate carrying the tag => tags survive Ctrl+D within one deck.
   (Ctrl+D dirties the doc, so no explicit dirtying is needed here.)
   ──────────────────────────────────────────────────────────────────────── */
export function spikeTagSelection(): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    await runPowerPoint(async (ctx) => {
      const sel = ctx.presentation.getSelectedShapes();
      sel.load("items/name,items/id");
      await ctx.sync();
      if (sel.items.length === 0) {
        lines.push("No shapes selected. Select one or more shapes and rerun 'Spike 3a'.");
        return;
      }
      sel.items.forEach((sh, i) => sh.tags.add(SPIKE3_TAG_KEY, `orig-${i}-${sh.name}`));
      await ctx.sync();
      lines.push(`Tagged ${sel.items.length} selected shape(s) with ${SPIKE3_TAG_KEY}.`);
      lines.push("NEXT: with those shapes still selected, press Ctrl+D to duplicate, then run 'Spike 3b: scan tags'.");
    });
    return lines;
  });
}

export function spikeScanTags(): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    await runPowerPoint(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();

      const cols = slides.items.map((s) => {
        s.shapes.load("items/name,items/id");
        return s.shapes;
      });
      await ctx.sync();

      const all = cols.flatMap((c) => c.items);
      all.forEach((s) => s.tags.load("items/key,items/value"));
      await ctx.sync();

      const tagged = all.filter((s) => s.tags.items.some((t) => t.key === SPIKE3_TAG_KEY));
      lines.push(`Shapes carrying ${SPIKE3_TAG_KEY}: ${tagged.length}.`);
      tagged.forEach((s) => {
        const v = s.tags.items.find((t) => t.key === SPIKE3_TAG_KEY)?.value ?? "";
        lines.push(`  "${s.name}" (id ${s.id}) = ${v}`);
      });
      lines.push(
        tagged.length > 1
          ? "RESULT: more than the originally-tagged shape carries the tag — tags SURVIVE Ctrl+D within the deck."
          : "RESULT: only the original carries the tag — duplicates do NOT inherit tags (re-run 3a first if you haven't duplicated yet).",
      );
    });
    return lines;
  });
}

/**
 * PHASE 3 GATE HELPER — tag every shape in the deck with a style GUID so the
 * apply sweep has 400+ linked shapes to run against. Tag writes are chunked and
 * the document is dirtied so they survive.
 */
export function spikeLinkAllShapes(styleId: string): Promise<SpikeResult> {
  return guard(async () => {
    const lines: string[] = [];
    let count = 0;
    await runPowerPoint(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load("items");
      await ctx.sync();
      const cols = slides.items.map((s) => {
        s.shapes.load("items/id");
        return s.shapes;
      });
      await ctx.sync();
      const all = cols.flatMap((c) => c.items);

      const CHUNK = 100;
      for (let i = 0; i < all.length; i += CHUNK) {
        all.slice(i, i + CHUNK).forEach((sh) => sh.tags.add(TAG_KEY, styleId));
        await ctx.sync();
        count += Math.min(CHUNK, all.length - i);
      }
      // Dirty so the tag-only writes persist.
      const first = all[0];
      if (first) {
        first.load("left");
        await ctx.sync();
        first.left = first.left;
        await ctx.sync();
      }
    });
    lines.push(`Tagged ${count} shapes with ${TAG_KEY}=${styleId}.`);
    lines.push("NEXT: run 'apply to all (sweep)'.");
    return lines;
  });
}

/** First selected shape, else the first shape on the first slide. */
async function firstTargetShape(ctx: PowerPoint.RequestContext): Promise<PowerPoint.Shape | null> {
  const sel = ctx.presentation.getSelectedShapes();
  sel.load("items/name");
  await ctx.sync();
  if (sel.items.length > 0) return sel.items[0] ?? null;

  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  const first = slides.items[0];
  if (!first) return null;
  first.shapes.load("items/name");
  await ctx.sync();
  return first.shapes.items[0] ?? null;
}
