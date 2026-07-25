import { useState } from "react";
import {
  makeStyles,
  tokens,
  shorthands,
  Body1,
  Caption1,
  Button,
  Input,
  Divider,
  Popover,
  PopoverTrigger,
  PopoverSurface,
} from "@fluentui/react-components";
import type { LayerName, StyleSmithDocument, ValueKind } from "../../core/schema/types";
import { tokenTypeForProperty } from "../../core/schema/defaults";
import { resolveLiteral, tokenFor } from "../state/resolve";
import { useStore } from "../state/store";

const useStyles = makeStyles({
  root: { display: "flex", alignItems: "center", justifyContent: "space-between", minHeight: "28px", ...shorthands.gap(tokens.spacingHorizontalS) },
  label: { color: tokens.colorNeutralForeground2, flexShrink: 0, width: "84px" },
  value: { flexGrow: 1, display: "flex", alignItems: "center", justifyContent: "flex-end", ...shorthands.gap(tokens.spacingHorizontalXS), minWidth: 0 },
  chip: {
    display: "inline-flex", alignItems: "center", ...shorthands.gap("4px"),
    ...shorthands.padding("2px", tokens.spacingHorizontalS), ...shorthands.borderRadius(tokens.borderRadiusCircular),
    backgroundColor: tokens.colorBrandBackground2, color: tokens.colorBrandForeground2,
    fontSize: tokens.fontSizeBase200, maxWidth: "150px", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
  },
  resolved: { color: tokens.colorNeutralForeground3, fontFamily: tokens.fontFamilyMonospace },
  trigger: { fontFamily: tokens.fontFamilyMonospace, maxWidth: "160px" },
  picker: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalS), minWidth: "220px" },
  heading: { color: tokens.colorNeutralForeground3 },
  tokenList: { display: "flex", flexWrap: "wrap", ...shorthands.gap(tokens.spacingHorizontalXS) },
  inline: { display: "flex", ...shorthands.gap(tokens.spacingHorizontalXS) },
});

function parseLiteral(raw: string): string | number {
  const t = raw.trim();
  if (t !== "" && !Number.isNaN(Number(t))) return Number(t);
  return raw;
}

/**
 * The binding control (UX §4.2). 🔗/🎨 chips mean token/theme-bound; a plain
 * value is a literal. The picker binds to an existing token, promotes the value
 * to a new token (S30), edits the literal, or unbinds to a literal (S31).
 */
export function BindingControl({
  doc,
  styleId,
  layer,
  property,
  label,
  value,
  editable,
}: {
  doc: StyleSmithDocument;
  styleId: string;
  layer: LayerName;
  property: string;
  label: string;
  value: ValueKind | undefined;
  editable: boolean;
}): JSX.Element {
  const s = useStyles();
  const setProperty = useStore((st) => st.setProperty);
  const bindToToken = useStore((st) => st.bindToToken);
  const createTokenFromValue = useStore((st) => st.createTokenFromValue);
  const unbindToLiteralAction = useStore((st) => st.unbindToLiteralAction);

  const [open, setOpen] = useState(false);
  const [literalDraft, setLiteralDraft] = useState("");
  const [tokenName, setTokenName] = useState("");

  const tokenType = layer === "text" || layer === "shape" ? tokenTypeForProperty(layer, property) : null;
  const canPick = editable && (layer === "text" || layer === "shape");
  const resolved = value ? resolveLiteral(doc, value) : null;
  const isBound = value?.kind === "tokenRef" || value?.kind === "themeColor";
  const isLiteral = value?.kind === "literal";

  const matchingTokens = tokenType
    ? doc.tokenSets.flatMap((set) => set.tokens).filter((t) => t.type === tokenType)
    : [];

  function label_(): string {
    if (value?.kind === "literal") return String(value.value);
    return "—";
  }

  function chip(): JSX.Element | null {
    if (value?.kind === "tokenRef") {
      const t = tokenFor(doc, value);
      return <span className={s.chip}>🔗 {t?.name ?? "token"}</span>;
    }
    if (value?.kind === "themeColor") return <span className={s.chip}>🎨 {value.slot}</span>;
    return null;
  }

  const layerTS = layer === "text" || layer === "shape" ? layer : null;

  const trigger = canPick ? (
    <Popover open={open} onOpenChange={(_, d) => setOpen(d.open)} trapFocus>
      <PopoverTrigger disableButtonEnhancement>
        <Button size="small" appearance="subtle" className={s.trigger} onClick={() => { setLiteralDraft(isLiteral ? String(value.value) : ""); setTokenName(""); }}>
          {chip() ?? label_()}
        </Button>
      </PopoverTrigger>
      <PopoverSurface>
        <div className={s.picker}>
          {tokenType && matchingTokens.length > 0 && layerTS && (
            <>
              <Caption1 className={s.heading}>BIND TO TOKEN</Caption1>
              <div className={s.tokenList}>
                {matchingTokens.map((t) => (
                  <Button key={t.id} size="small" onClick={() => { void bindToToken(styleId, layerTS, property, t.id); setOpen(false); }}>
                    🔗 {t.name}
                  </Button>
                ))}
              </div>
              <Divider />
            </>
          )}

          {layerTS && (
            <>
              <Caption1 className={s.heading}>LITERAL</Caption1>
              <div className={s.inline}>
                <Input size="small" value={literalDraft} onChange={(_, d) => setLiteralDraft(d.value)} placeholder="value" />
                <Button size="small" appearance="primary" onClick={() => { void setProperty(styleId, layerTS, property, { kind: "literal", value: parseLiteral(literalDraft) }); setOpen(false); }}>
                  Set
                </Button>
              </div>
            </>
          )}

          {isLiteral && tokenType && layerTS && (
            <>
              <Divider />
              <Caption1 className={s.heading}>CREATE TOKEN FROM VALUE</Caption1>
              <div className={s.inline}>
                <Input size="small" value={tokenName} onChange={(_, d) => setTokenName(d.value)} placeholder="token name" />
                <Button size="small" disabled={tokenName.trim() === ""} onClick={() => { void createTokenFromValue(styleId, layerTS, property, tokenName.trim()); setOpen(false); }}>
                  Create
                </Button>
              </div>
            </>
          )}

          {isBound && layerTS && (
            <>
              <Divider />
              <Button size="small" onClick={() => { void unbindToLiteralAction(styleId, layerTS, property); setOpen(false); }}>
                Unbind to literal ({resolved !== null ? String(resolved) : "…"})
              </Button>
            </>
          )}
        </div>
      </PopoverSurface>
    </Popover>
  ) : (
    chip() ?? <Body1 className={s.resolved}>{label_()}</Body1>
  );

  return (
    <div className={s.root}>
      <Caption1 className={s.label}>{label}</Caption1>
      <div className={s.value}>
        {trigger}
        {value && value.kind !== "literal" && resolved !== null && (
          <Caption1 className={s.resolved}>{String(resolved)}</Caption1>
        )}
      </div>
    </div>
  );
}
