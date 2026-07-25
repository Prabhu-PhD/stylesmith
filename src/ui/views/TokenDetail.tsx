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
  Radio,
  RadioGroup,
  Input,
  Select,
  Field,
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
} from "@fluentui/react-components";
import type { ThemeSlot, ValueKind } from "../../core/schema/types";
import { THEME_SLOTS } from "../../core/schema/schemas";
import { useStore, tokenCascade } from "../state/store";

const useStyles = makeStyles({
  header: { display: "flex", alignItems: "center", ...shorthands.gap(tokens.spacingHorizontalS), ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM) },
  title: { flexGrow: 1, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  body: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalM), ...shorthands.padding("0", tokens.spacingHorizontalM, tokens.spacingVerticalM) },
  valueRow: { display: "flex", alignItems: "center", ...shorthands.gap(tokens.spacingHorizontalS) },
  cascade: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalXS) },
  styleLine: { display: "flex", justifyContent: "space-between" },
  total: { display: "flex", justifyContent: "space-between", fontWeight: tokens.fontWeightSemibold, ...shorthands.borderTop("1px", "solid", tokens.colorNeutralStroke2), paddingTop: tokens.spacingVerticalXS, marginTop: tokens.spacingVerticalXS },
  actions: { display: "flex", ...shorthands.gap(tokens.spacingHorizontalS), justifyContent: "flex-end" },
  heading: { color: tokens.colorNeutralForeground3 },
});

function parseLiteral(raw: string): string | number {
  const t = raw.trim();
  if (t !== "" && !Number.isNaN(Number(t))) return Number(t);
  return raw;
}

type Kind = "literal" | "theme" | "token";

export function TokenDetail({ tokenId, onBack }: { tokenId: string; onBack: () => void }): JSX.Element {
  const s = useStyles();
  const doc = useStore((st) => st.doc);
  const shapes = useStore((st) => st.shapes);
  const setTokenValue = useStore((st) => st.setTokenValue);
  const applyTokenCascade = useStore((st) => st.applyTokenCascade);
  const renameTokenAction = useStore((st) => st.renameTokenAction);
  const deleteTokenAction = useStore((st) => st.deleteTokenAction);

  const token = doc?.tokenSets.flatMap((set) => set.tokens).find((t) => t.id === tokenId);

  const [kind, setKind] = useState<Kind>(() =>
    token?.value.kind === "themeColor" ? "theme" : token?.value.kind === "tokenRef" ? "token" : "literal",
  );
  const [literal, setLiteral] = useState(() => (token?.value.kind === "literal" ? String(token.value.value) : ""));
  const [slot, setSlot] = useState<ThemeSlot>(() => (token?.value.kind === "themeColor" ? token.value.slot : "accent1"));
  const [aliasId, setAliasId] = useState(() => (token?.value.kind === "tokenRef" ? token.value.tokenId : ""));
  const [renameOpen, setRenameOpen] = useState(false);
  const [renameDraft, setRenameDraft] = useState("");

  if (!doc || !token) {
    return <div className={s.header}><Button appearance="subtle" onClick={onBack}>‹ Back</Button></div>;
  }

  const isColour = token.type === "color";
  const aliasOptions = doc.tokenSets
    .flatMap((set) => set.tokens)
    .filter((t) => t.id !== tokenId && t.type === token.type);

  function buildValue(): ValueKind {
    if (kind === "theme") return { kind: "themeColor", slot };
    if (kind === "token") return { kind: "tokenRef", tokenId: aliasId };
    return { kind: "literal", value: parseLiteral(literal) };
  }

  const cascade = tokenCascade(doc, shapes, tokenId);

  async function apply(): Promise<void> {
    if (kind === "token" && !aliasId) return;
    const committed = await setTokenValue(tokenId, buildValue());
    if (!committed) return; // cycle/type error surfaced via notice
    await applyTokenCascade(tokenId);
  }

  return (
    <div>
      <div className={s.header}>
        <Button appearance="subtle" onClick={onBack} aria-label="Back to tokens">‹</Button>
        <Subtitle2 className={s.title}>{token.name}</Subtitle2>
        <Menu>
          <MenuTrigger disableButtonEnhancement>
            <Button appearance="subtle" aria-label="Token menu">⋯</Button>
          </MenuTrigger>
          <MenuPopover>
            <MenuList>
              <MenuItem onClick={() => { setRenameDraft(token.name); setRenameOpen(true); }}>Rename</MenuItem>
              <MenuItem onClick={async () => { onBack(); await deleteTokenAction(tokenId); }}>Delete</MenuItem>
            </MenuList>
          </MenuPopover>
        </Menu>
      </div>

      <div className={s.body}>
        <Field label="Value">
          <RadioGroup value={kind} onChange={(_, d) => setKind(d.value as Kind)}>
            <div className={s.valueRow}>
              <Radio value="literal" label="Literal" />
              {kind === "literal" && (
                <Input size="small" value={literal} onChange={(_, d) => setLiteral(d.value)} placeholder="#1A3A6B / 14 / text" />
              )}
            </div>
            {isColour && (
              <div className={s.valueRow}>
                <Radio value="theme" label="Theme colour 🎨" />
                {kind === "theme" && (
                  <Select size="small" value={slot} onChange={(_, d) => setSlot(d.value as ThemeSlot)}>
                    {THEME_SLOTS.map((sl) => <option key={sl} value={sl}>{sl}</option>)}
                  </Select>
                )}
              </div>
            )}
            <div className={s.valueRow}>
              <Radio value="token" label="Another token" disabled={aliasOptions.length === 0} />
              {kind === "token" && (
                <Select size="small" value={aliasId} onChange={(_, d) => setAliasId(d.value)}>
                  <option value="">Choose…</option>
                  {aliasOptions.map((t) => <option key={t.id} value={t.id}>{t.name}</option>)}
                </Select>
              )}
            </div>
          </RadioGroup>
        </Field>

        <Divider />

        <div className={s.cascade}>
          <Caption1 className={s.heading}>CHANGING THIS AFFECTS</Caption1>
          {cascade.styleCount === 0 ? (
            <Body1 style={{ color: tokens.colorNeutralForeground3 }}>No styles use this token yet.</Body1>
          ) : (
            <>
              {cascade.styles.map((st) => (
                <div key={st.styleId} className={s.styleLine}>
                  <Body1>{st.name}</Body1>
                  <Caption1>{st.shapeCount} {st.shapeCount === 1 ? "shape" : "shapes"}</Caption1>
                </div>
              ))}
              <div className={s.total}>
                <Body1>{cascade.shapeCount} shapes</Body1>
                <Body1>across {cascade.slideCount} {cascade.slideCount === 1 ? "slide" : "slides"}</Body1>
              </div>
            </>
          )}
        </div>

        <div className={s.actions}>
          <Button appearance="secondary" onClick={onBack}>Cancel</Button>
          <Button appearance="primary" onClick={() => void apply()}>Apply</Button>
        </div>
      </div>

      <Dialog open={renameOpen} onOpenChange={(_, d) => setRenameOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>Rename token</DialogTitle>
            <DialogContent>
              <Field label="Name"><Input value={renameDraft} onChange={(_, d) => setRenameDraft(d.value)} autoFocus /></Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setRenameOpen(false)}>Cancel</Button>
              <Button appearance="primary" disabled={renameDraft.trim() === ""} onClick={async () => { await renameTokenAction(tokenId, renameDraft.trim()); setRenameOpen(false); }}>Rename</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </div>
  );
}
