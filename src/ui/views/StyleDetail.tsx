import { useState } from "react";
import {
  makeStyles,
  tokens,
  shorthands,
  Button,
  Subtitle2,
  Body1,
  Caption1,
  Divider,
  Accordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel,
  Menu,
  MenuTrigger,
  MenuPopover,
  MenuList,
  MenuItem,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Input,
  Field,
} from "@fluentui/react-components";
import type { StyleSmithDocument, ValueKind } from "../../core/schema/types";
import { TEXT_PROPERTIES, SHAPE_PROPERTIES } from "../../core/schema/schemas";
import { useStore, countShapesForStyle } from "../state/store";
import { BindingControl } from "../components/BindingControl";

const useStyles = makeStyles({
  header: { display: "flex", alignItems: "center", ...shorthands.gap(tokens.spacingHorizontalS), ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM) },
  title: { flexGrow: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  rows: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalXS), ...shorthands.padding("0", tokens.spacingHorizontalXS) },
  notDefined: { color: tokens.colorNeutralForeground3, ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM) },
  geomNote: { color: tokens.colorNeutralForeground3, ...shorthands.padding(tokens.spacingVerticalXS, tokens.spacingHorizontalXS) },
  actions: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalS), ...shorthands.padding(tokens.spacingVerticalM, tokens.spacingHorizontalM) },
});

const LABELS: Record<string, string> = {
  fontFamily: "Font",
  fontSize: "Size",
  fontWeight: "Weight",
  color: "Colour",
  lineSpacing: "Spacing",
  spaceBefore: "Space ↑",
  spaceAfter: "Space ↓",
  alignment: "Align",
  fill: "Fill",
  borderColor: "Border",
  borderWeight: "Border wt",
  borderDashStyle: "Dash",
};

function definedProps(layer: Record<string, ValueKind | undefined> | null, order: readonly string[]): string[] {
  if (!layer) return [];
  return order.filter((p) => layer[p] !== undefined);
}

export function StyleDetail({ styleId, onBack }: { styleId: string; onBack: () => void }): JSX.Element {
  const s = useStyles();
  const doc = useStore((st) => st.doc) as StyleSmithDocument | null;
  const shapes = useStore((st) => st.shapes);
  const updateFromSelection = useStore((st) => st.updateFromSelection);
  const applyToSelection = useStore((st) => st.applyToSelection);
  const openApply = useStore((st) => st.openApply);
  const rename = useStore((st) => st.rename);
  const remove = useStore((st) => st.remove);
  const duplicate = useStore((st) => st.duplicate);
  const selectStyle = useStore((st) => st.selectStyle);

  const [renameOpen, setRenameOpen] = useState(false);
  const [renameDraft, setRenameDraft] = useState("");
  const [deleteOpen, setDeleteOpen] = useState(false);

  const style = doc?.styles.find((st) => st.id === styleId);
  if (!doc || !style) {
    return (
      <div className={s.header}>
        <Button appearance="subtle" onClick={onBack}>‹ Back</Button>
      </div>
    );
  }

  const usage = countShapesForStyle(shapes, styleId);
  const textProps = definedProps(style.layers.text, TEXT_PROPERTIES);
  const shapeProps = definedProps(style.layers.shape, SHAPE_PROPERTIES);
  const geometry = style.layers.geometry;

  return (
    <div>
      <div className={s.header}>
        <Button appearance="subtle" onClick={onBack} aria-label="Back to styles">‹</Button>
        <Subtitle2 className={s.title}>{style.name}</Subtitle2>
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button appearance="subtle" aria-label="Style menu">⋯</Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem onClick={() => { setRenameDraft(style.name); setRenameOpen(true); }}>Rename</MenuItem>
              <MenuItem onClick={async () => { const id = await duplicate(styleId); if (id) selectStyle(id); }}>Duplicate</MenuItem>
              <MenuItem onClick={() => setDeleteOpen(true)}>Delete</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>

      <Accordion multiple collapsible defaultOpenItems={["text", "shape", "geometry"]}>
        <AccordionItem value="text">
          <AccordionHeader>TEXT</AccordionHeader>
          <AccordionPanel>
            {textProps.length > 0 ? (
              <div className={s.rows}>
                {textProps.map((p) => (
                  <BindingControl
                    key={p}
                    doc={doc}
                    styleId={styleId}
                    layer="text"
                    property={p}
                    label={LABELS[p] ?? p}
                    value={(style.layers.text as Record<string, ValueKind>)[p]}
                    editable
                  />
                ))}
              </div>
            ) : (
              <Body1 className={s.notDefined}>Text layer not defined.</Body1>
            )}
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="shape">
          <AccordionHeader>SHAPE</AccordionHeader>
          <AccordionPanel>
            {shapeProps.length > 0 ? (
              <div className={s.rows}>
                {shapeProps.map((p) => (
                  <BindingControl
                    key={p}
                    doc={doc}
                    styleId={styleId}
                    layer="shape"
                    property={p}
                    label={LABELS[p] ?? p}
                    value={(style.layers.shape as Record<string, ValueKind>)[p]}
                    editable
                  />
                ))}
              </div>
            ) : (
              <Body1 className={s.notDefined}>Shape layer not defined.</Body1>
            )}
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="geometry">
          <AccordionHeader>GEOMETRY</AccordionHeader>
          <AccordionPanel>
            {geometry && geometry.adjustments.length > 0 ? (
              <div className={s.rows}>
                {geometry.adjustments.map((v, i) => (
                  <BindingControl key={i} doc={doc} styleId={styleId} layer="geometry" property="adjustments" label={`⌒ [${i}]`} value={v} editable={false} />
                ))}
                <Caption1 className={s.geomNote}>ⓘ Applies to shapes with adjustment handles</Caption1>
              </div>
            ) : (
              <Body1 className={s.notDefined}>Geometry layer not defined.</Body1>
            )}
          </AccordionPanel>
        </AccordionItem>

        <AccordionItem value="table">
          <AccordionHeader>TABLE</AccordionHeader>
          <AccordionPanel>
            <Body1 className={s.notDefined}>Not defined (table styling arrives in Phase 5).</Body1>
          </AccordionPanel>
        </AccordionItem>
      </Accordion>

      <Divider />
      <div className={s.actions}>
        <Button appearance="primary" onClick={() => applyToSelection(styleId)}>Apply to selection</Button>
        <Button onClick={() => openApply(styleId)}>Apply to all…</Button>
        <Button appearance="subtle" onClick={() => updateFromSelection(styleId)}>Update from selection</Button>
      </div>

      {/* Rename */}
      <Dialog open={renameOpen} onOpenChange={(_, d) => setRenameOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Rename style</DialogTitle>
            <DialogContent>
              <Field label="Name">
                <Input value={renameDraft} onChange={(_, d) => setRenameDraft(d.value)} autoFocus />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setRenameOpen(false)}>Cancel</Button>
              <Button
                appearance="primary"
                disabled={renameDraft.trim() === ""}
                onClick={async () => { await rename(styleId, renameDraft.trim()); setRenameOpen(false); }}
              >
                Rename
              </Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>

      {/* Delete confirm */}
      <Dialog open={deleteOpen} onOpenChange={(_, d) => setDeleteOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Delete "{style.name}"?</DialogTitle>
            <DialogContent>
              {usage > 0
                ? `${usage} ${usage === 1 ? "shape carries" : "shapes carry"} this style. They will keep their formatting but become unlinked.`
                : "This style is not used by any shape."}
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setDeleteOpen(false)}>Cancel</Button>
              <Button appearance="primary" onClick={async () => { setDeleteOpen(false); onBack(); await remove(styleId); }}>Delete</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}
