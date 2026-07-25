import { useState } from "react";
import {
  Button,
  Divider,
  Subtitle2,
  Body1,
  Caption1,
  makeStyles,
  tokens,
  shorthands,
} from "@fluentui/react-components";
import {
  spikeGroupTraversal,
  spikeProbeStorage,
  spikeWriteSettings,
  spikeCheckSettings,
  spikeClearSettings,
  spikeTagSelection,
  spikeScanTags,
  spikeLinkAllShapes,
  type SpikeResult,
} from "../../office/spikes";
import { officeBridge } from "../../office/bridge";
import type { StyleSmithDocument } from "../../core/schema/types";
import { createEmptyDocument } from "../../core/schema/defaults";
import { createStyle, createToken, bindProperty } from "../../core/styles/model";

/* ── Phase 3 apply-sweep gate harness ─────────────────────────────────────── */

const SWEEP_STYLE_ID = "sty-sweep";

/** A sample style with a fill + one geometry adjustment, for the perf gate. */
function buildSweepSampleDoc(): StyleSmithDocument {
  let doc = createEmptyDocument(() => "set-1");
  const style = createStyle(
    doc,
    {
      name: "SweepSample",
      layers: {
        text: null,
        shape: { fill: { kind: "literal", value: "#0057B8" } },
        geometry: { adjustments: [{ kind: "literal", value: 0.2 }] },
        table: null,
      },
    },
    () => SWEEP_STYLE_ID,
  );
  if (style.ok) doc = style.value.doc;
  return doc;
}

async function runPrepSweep(): Promise<SpikeResult> {
  const save = await officeBridge.saveDocument(buildSweepSampleDoc());
  if (!save.ok) return { ok: false, lines: [`save failed: ${JSON.stringify(save.error)}`] };
  return spikeLinkAllShapes(SWEEP_STYLE_ID);
}

async function runSweepApply(): Promise<SpikeResult> {
  const progress: number[] = [];
  const r = await officeBridge.applyStyle(
    SWEEP_STYLE_ID,
    { kind: "deck" },
    { preserveOverrides: false, chunkSize: 50, onProgress: (done) => progress.push(done) },
  );
  if (!r.ok) return { ok: false, lines: [`apply failed: ${JSON.stringify(r.error)}`] };
  const s = r.value;
  return {
    ok: !s.cancelled,
    lines: [
      `matched ${s.matched}, applied ${s.applied}, geometry-skipped ${s.skippedGeometryShapes}`,
      `elapsed ${s.elapsedMs.toFixed(0)}ms · throughput ${s.throughput.toFixed(0)} shapes/s · chunk ${s.chunkSize}`,
      `progress callbacks: ${progress.length} · cancelled: ${s.cancelled}`,
      s.unresolvedPaths.length ? `unresolved: ${s.unresolvedPaths.join(", ")}` : "unresolved: none",
      s.deferredPaths.length ? `deferred→Phase4: ${s.deferredPaths.join(", ")}` : "deferred: none",
      "GATE: completed without hanging, with accurate progress. Record the throughput above.",
    ],
  };
}

async function runSweepCancel(): Promise<SpikeResult> {
  const controller = new AbortController();
  let armed = true;
  const r = await officeBridge.applyStyle(
    SWEEP_STYLE_ID,
    { kind: "deck" },
    {
      preserveOverrides: false,
      chunkSize: 50,
      signal: controller.signal,
      onProgress: () => {
        if (armed) {
          armed = false;
          controller.abort(); // cancel after the first chunk boundary
        }
      },
    },
  );
  if (!r.ok) return { ok: false, lines: [`apply failed: ${JSON.stringify(r.error)}`] };
  const s = r.value;
  return {
    ok: s.cancelled,
    lines: [
      `cancelled: ${s.cancelled} (expected true)`,
      `applied before stop: ${s.applied} of matched ${s.matched} — stopped at a chunk boundary`,
    ],
  };
}

/**
 * A deterministic sample document (fixed ids) so a read-back can be compared for
 * exact identity — including across close/reopen, where in-memory state is gone.
 */
function buildSampleDocument(): StyleSmithDocument {
  let doc = createEmptyDocument(() => "set-1");
  const token = createToken(
    doc,
    "set-1",
    { name: "brand-primary", type: "color", value: { kind: "themeColor", slot: "accent1" } },
    () => "tok-1",
  );
  if (token.ok) doc = token.value.doc;
  const style = createStyle(doc, { name: "Takeaway" }, () => "sty-1");
  if (style.ok) doc = style.value.doc;
  const bound = bindProperty(doc, "sty-1", "shape", "fill", {
    kind: "tokenRef",
    tokenId: "tok-1",
    cached: "#0057B8",
  });
  if (bound.ok) doc = bound.value;
  return doc;
}

async function runWriteSample(): Promise<SpikeResult> {
  const r = await officeBridge.saveDocument(buildSampleDocument());
  if (!r.ok) return { ok: false, lines: [`save failed: ${JSON.stringify(r.error)}`] };
  return {
    ok: true,
    lines: [
      "Saved sample document to document.settings.",
      "NEXT: optionally close & reopen the deck, then 'read back + compare'.",
    ],
  };
}

async function runReadCompare(): Promise<SpikeResult> {
  const r = await officeBridge.loadDocument();
  if (!r.ok) return { ok: false, lines: [`load failed: ${JSON.stringify(r.error)}`] };
  if (r.value === null) return { ok: false, lines: ["No document stored — run 'write sample' first."] };

  const actual = JSON.stringify(r.value);
  const expected = JSON.stringify(buildSampleDocument());
  if (actual === expected) {
    return {
      ok: true,
      lines: [
        "ROUND-TRIP OK — loaded document is byte-identical to what was written.",
        `styles: ${r.value.styles.length}, tokenSets: ${r.value.tokenSets.length}`,
      ],
    };
  }
  return {
    ok: false,
    lines: ["MISMATCH — loaded document differs from expected.", `expected: ${expected}`, `actual:   ${actual}`],
  };
}

async function runClearDocument(): Promise<SpikeResult> {
  const r = await officeBridge.clearDocument();
  return r.ok
    ? { ok: true, lines: ["Cleared stored document."] }
    : { ok: false, lines: [`clear failed: ${JSON.stringify(r.error)}`] };
}

const useStyles = makeStyles({
  section: {
    display: "flex",
    flexDirection: "column",
    ...shorthands.gap(tokens.spacingVerticalXS),
    marginBottom: tokens.spacingVerticalM,
  },
  buttons: {
    display: "flex",
    flexWrap: "wrap",
    ...shorthands.gap(tokens.spacingHorizontalXS),
  },
  log: {
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalS),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3,
    color: tokens.colorNeutralForeground1,
    fontFamily: tokens.fontFamilyMonospace,
    fontSize: tokens.fontSizeBase200,
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
    marginTop: tokens.spacingVerticalXS,
  },
  okBorder: { ...shorthands.border("1px", "solid", tokens.colorPaletteGreenBorder2) },
  errBorder: { ...shorthands.border("1px", "solid", tokens.colorPaletteRedBorder2) },
});

type Runner = () => Promise<SpikeResult>;

interface SpikeSectionProps {
  title: string;
  hint: string;
  actions: { label: string; run: Runner; primary?: boolean }[];
}

function SpikeSection({ title, hint, actions }: SpikeSectionProps): JSX.Element {
  const s = useStyles();
  const [result, setResult] = useState<SpikeResult | null>(null);
  const [busy, setBusy] = useState(false);

  async function invoke(run: Runner): Promise<void> {
    setBusy(true);
    try {
      setResult(await run());
    } catch (e) {
      setResult({ ok: false, lines: [e instanceof Error ? e.message : String(e)] });
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className={s.section}>
      <Subtitle2>{title}</Subtitle2>
      <Caption1>{hint}</Caption1>
      <div className={s.buttons}>
        {actions.map((a) => (
          <Button
            key={a.label}
            size="small"
            appearance={a.primary ? "primary" : "secondary"}
            disabled={busy}
            onClick={() => void invoke(a.run)}
          >
            {a.label}
          </Button>
        ))}
      </div>
      {result && (
        <div className={`${s.log} ${result.ok ? s.okBorder : s.errBorder}`}>
          {result.lines.join("\n")}
        </div>
      )}
    </div>
  );
}

/**
 * Phase 0 debug harness. This exists only to run the three open-question spikes
 * against a live host; it is not part of the shipping UI and will be removed
 * once Phase 0's gate is signed off. Office JS stays in office/spikes.ts — this
 * component only calls those async functions and renders their logs.
 */
export function SpikePanel(): JSX.Element {
  return (
    <div>
      <Divider>Phase 0 spikes</Divider>
      <Body1 as="p">Run against a sideloaded deck. Some steps are manual (grouping, close/reopen, Ctrl+D) — each result log tells you the next step.</Body1>

      <SpikeSection
        title="1 · Group traversal"
        hint="Group 2+ shapes (Ctrl+G) on a slide, then run."
        actions={[{ label: "Scan groups", run: spikeGroupTraversal, primary: true }]}
      />

      <SpikeSection
        title="2 · Deck storage + dirty flag"
        hint="PowerPoint has no doc-level custom XML — probing document.settings. Probe, write, close+reopen, check. Compare no-dirty vs dirty."
        actions={[
          { label: "probe storage", run: spikeProbeStorage, primary: true },
          { label: "2a write (no dirty)", run: () => spikeWriteSettings(false) },
          { label: "2a′ write (dirty)", run: () => spikeWriteSettings(true) },
          { label: "2b check survived", run: spikeCheckSettings },
          { label: "clear", run: spikeClearSettings },
        ]}
      />

      <SpikeSection
        title="3 · Ctrl+D tag survival"
        hint="Select shape(s) → 3a → press Ctrl+D → 3b."
        actions={[
          { label: "3a tag selection", run: spikeTagSelection, primary: true },
          { label: "3b scan tags", run: spikeScanTags },
        ]}
      />

      <Divider>Phase 2 storage</Divider>
      <SpikeSection
        title="Storage round-trip (gate)"
        hint="Write a sample doc to document.settings, optionally close+reopen, then read back and compare for exact identity."
        actions={[
          { label: "write sample", run: runWriteSample, primary: true },
          { label: "read back + compare", run: runReadCompare },
          { label: "clear", run: runClearDocument },
        ]}
      />

      <Divider>Phase 3 apply sweep</Divider>
      <SpikeSection
        title="Apply sweep (gate)"
        hint="Prep tags all shapes with a sample style (fill + geometry). Apply sweeps deck-wide with progress; cancel stops at a chunk boundary."
        actions={[
          { label: "prep: link all shapes", run: runPrepSweep, primary: true },
          { label: "apply to all (sweep)", run: runSweepApply },
          { label: "cancel test", run: runSweepCancel },
        ]}
      />
    </div>
  );
}
