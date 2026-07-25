import { useState } from "react";
import {
  makeStyles,
  tokens,
  shorthands,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Radio,
  RadioGroup,
  Checkbox,
  Caption1,
  Body1,
  Divider,
} from "@fluentui/react-components";
import type { LayerName } from "../../core/schema/types";
import { useStore, countShapesForStyle, slideIndicesForSelection } from "../state/store";
import type { SweepScope } from "../../office/bridge";

const useStyles = makeStyles({
  section: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalXS), marginBottom: tokens.spacingVerticalM },
  heading: { color: tokens.colorNeutralForeground3 },
  layers: { display: "flex", flexWrap: "wrap", ...shorthands.gap(tokens.spacingHorizontalM) },
  overrideBox: {
    ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM),
    ...shorthands.borderRadius(tokens.borderRadiusMedium),
    backgroundColor: tokens.colorNeutralBackground3,
    marginBottom: tokens.spacingVerticalM,
  },
  summary: { fontWeight: tokens.fontWeightSemibold, marginTop: tokens.spacingVerticalS },
});

type ScopeKind = "deck" | "slides" | "selection";

export function ApplyModal(): JSX.Element {
  const s = useStyles();
  const styleId = useStore((st) => st.applyStyleId);
  const doc = useStore((st) => st.doc);
  const shapes = useStore((st) => st.shapes);
  const selectionIds = useStore((st) => st.selectionIds);
  const apply = useStore((st) => st.apply);
  const closeApply = useStore((st) => st.closeApply);

  const style = doc?.styles.find((st) => st.id === styleId);

  const definedLayers = (["text", "shape", "geometry"] as LayerName[]).filter(
    (l) => style?.layers[l] != null,
  );

  const [scope, setScope] = useState<ScopeKind>("deck");
  const [layerOn, setLayerOn] = useState<Record<string, boolean>>({});
  const [preserve, setPreserve] = useState(true);

  if (!styleId || !style || !doc) return <></>;

  // Initialise layer checkboxes to the style's defined layers on first render.
  if (Object.keys(layerOn).length === 0 && definedLayers.length > 0) {
    setLayerOn(Object.fromEntries(definedLayers.map((l) => [l, true])));
  }

  const styleShapes = shapes.filter((sh) => sh.styleId === styleId);
  const selSlideIndices = slideIndicesForSelection(shapes, selectionIds);
  const selIdSet = new Set(selectionIds);

  const scopeShapes = {
    deck: styleShapes,
    slides: styleShapes.filter((sh) => selSlideIndices.includes(sh.slideIndex)),
    selection: styleShapes.filter((sh) => selIdSet.has(sh.id)),
  } as const;

  const affected = scopeShapes[scope];
  const affectedSlides = new Set(affected.map((sh) => sh.slideIndex)).size;

  const selectedLayers = definedLayers.filter((l) => layerOn[l]);
  const deckCount = countShapesForStyle(shapes, styleId);

  function buildScope(): SweepScope {
    if (scope === "slides") return { kind: "slides", slideIndices: selSlideIndices };
    if (scope === "selection") return { kind: "shapes", shapeIds: affected.map((sh) => sh.id) };
    return { kind: "deck" };
  }

  const canApply = affected.length > 0 && selectedLayers.length > 0;

  return (
    <Dialog open onOpenChange={(_, d) => { if (!d.open) closeApply(); }}>
      <DialogSurface>
        <DialogBody>
          <DialogTitle>Apply "{style.name}"</DialogTitle>
          <DialogContent>
            <div className={s.section}>
              <Caption1 className={s.heading}>SCOPE</Caption1>
              <RadioGroup value={scope} onChange={(_, d) => setScope(d.value as ScopeKind)}>
                <Radio value="deck" label={`Entire deck — ${deckCount} ${deckCount === 1 ? "shape" : "shapes"}`} />
                <Radio value="slides" label={`Selected slide(s) — ${scopeShapes.slides.length}`} disabled={selSlideIndices.length === 0} />
                <Radio value="selection" label={`Selection — ${scopeShapes.selection.length}`} disabled={selectionIds.length === 0} />
              </RadioGroup>
            </div>

            <div className={s.section}>
              <Caption1 className={s.heading}>LAYERS</Caption1>
              <div className={s.layers}>
                {definedLayers.map((l) => (
                  <Checkbox
                    key={l}
                    checked={!!layerOn[l]}
                    onChange={(_, d) => setLayerOn((prev) => ({ ...prev, [l]: !!d.checked }))}
                    label={l[0]?.toUpperCase() + l.slice(1)}
                  />
                ))}
                <Checkbox checked={false} disabled label="Table" />
              </div>
            </div>

            <div className={s.overrideBox}>
              <Caption1 className={s.heading}>LOCAL OVERRIDES</Caption1>
              <RadioGroup value={preserve ? "preserve" : "reset"} onChange={(_, d) => setPreserve(d.value === "preserve")}>
                <Radio value="preserve" label="Preserve manual overrides" />
                <Radio value="reset" label="Reset to style" />
              </RadioGroup>
            </div>

            <Divider />
            <Body1 className={s.summary}>
              {canApply
                ? `Will change ${affected.length} ${affected.length === 1 ? "shape" : "shapes"} across ${affectedSlides} ${affectedSlides === 1 ? "slide" : "slides"}`
                : "Nothing to change in this scope."}
            </Body1>
          </DialogContent>
          <DialogActions>
            <Button appearance="secondary" onClick={closeApply}>Cancel</Button>
            <Button
              appearance="primary"
              disabled={!canApply}
              onClick={() =>
                apply(styleId, buildScope(), {
                  preserveOverrides: preserve,
                  ...(selectedLayers.length < definedLayers.length ? { layers: selectedLayers } : {}),
                })
              }
            >
              Apply
            </Button>
          </DialogActions>
        </DialogBody>
      </DialogSurface>
    </Dialog>
  );
}
