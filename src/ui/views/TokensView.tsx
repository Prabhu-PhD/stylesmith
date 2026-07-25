import { useState } from "react";
import {
  makeStyles,
  tokens,
  shorthands,
  Subtitle2,
  Body1,
  Caption1,
  Button,
  Dialog,
  DialogSurface,
  DialogBody,
  DialogTitle,
  DialogContent,
  DialogActions,
  Field,
  Input,
  Select,
} from "@fluentui/react-components";
import type { Token, TokenType } from "../../core/schema/types";
import { TOKEN_TYPES } from "../../core/schema/schemas";
import { directStyleUsageCount } from "../../core/tokens/usage";
import { useStore } from "../state/store";
import { TokenRow } from "../components/TokenRow";
import { CATEGORY_ORDER, categoryOf } from "../state/tokens-ui";

const useStyles = makeStyles({
  scroll: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalM), ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM), overflowY: "auto" },
  group: { display: "flex", flexDirection: "column", ...shorthands.gap(tokens.spacingVerticalXS) },
  heading: { color: tokens.colorNeutralForeground3, letterSpacing: "0.04em" },
  empty: { display: "flex", flexDirection: "column", alignItems: "center", textAlign: "center", ...shorthands.gap(tokens.spacingVerticalM), ...shorthands.padding(tokens.spacingVerticalXXL, tokens.spacingHorizontalL) },
  footer: { ...shorthands.padding(tokens.spacingVerticalS, tokens.spacingHorizontalM), ...shorthands.borderTop("1px", "solid", tokens.colorNeutralStroke2) },
});

function parseLiteral(raw: string): string | number {
  const t = raw.trim();
  if (t !== "" && !Number.isNaN(Number(t))) return Number(t);
  return raw;
}

export function TokensView(): JSX.Element {
  const s = useStyles();
  const doc = useStore((st) => st.doc);
  const addToken = useStore((st) => st.addToken);
  const selectToken = useStore((st) => st.selectToken);

  const [createOpen, setCreateOpen] = useState(false);
  const [name, setName] = useState("");
  const [type, setType] = useState<TokenType>("color");
  const [value, setValue] = useState("");

  if (!doc) return <></>;

  const allTokens: Token[] = doc.tokenSets.flatMap((set) => set.tokens);

  async function create(): Promise<void> {
    const id = await addToken(name.trim(), type, { kind: "literal", value: parseLiteral(value) });
    setCreateOpen(false);
    setName("");
    setValue("");
    if (id) selectToken(id);
  }

  return (
    <>
      {allTokens.length === 0 ? (
        <div className={s.empty}>
          <div aria-hidden style={{ fontSize: 28 }}>🎨</div>
          <Subtitle2>No tokens yet</Subtitle2>
          <Body1 style={{ color: tokens.colorNeutralForeground2, maxWidth: 240 }}>
            Create a token here, or promote a value to a token from any style property.
          </Body1>
        </div>
      ) : (
        <div className={s.scroll}>
          {CATEGORY_ORDER.map((category) => {
            const group = allTokens.filter((t) => categoryOf(t.type) === category);
            if (group.length === 0) return null;
            return (
              <div key={category} className={s.group}>
                <Caption1 className={s.heading}>{category.toUpperCase()}</Caption1>
                {group.map((token) => (
                  <TokenRow
                    key={token.id}
                    doc={doc}
                    token={token}
                    styleCount={directStyleUsageCount(doc, token.id)}
                    onOpen={() => selectToken(token.id)}
                  />
                ))}
              </div>
            );
          })}
        </div>
      )}

      <div className={s.footer}>
        <Button appearance="primary" style={{ width: "100%" }} onClick={() => setCreateOpen(true)}>+ New token</Button>
      </div>

      <Dialog open={createOpen} onOpenChange={(_, d) => setCreateOpen(d.open)}>
        <DialogSurface>
          <DialogBody>
            <DialogTitle>New token</DialogTitle>
            <DialogContent>
              <Field label="Name"><Input value={name} onChange={(_, d) => setName(d.value)} autoFocus /></Field>
              <Field label="Type">
                <Select value={type} onChange={(_, d) => setType(d.value as TokenType)}>
                  {TOKEN_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                </Select>
              </Field>
              <Field label="Value (literal)" hint="A hex colour, a number, or text.">
                <Input value={value} onChange={(_, d) => setValue(d.value)} />
              </Field>
            </DialogContent>
            <DialogActions>
              <Button appearance="secondary" onClick={() => setCreateOpen(false)}>Cancel</Button>
              <Button appearance="primary" disabled={name.trim() === "" || value.trim() === ""} onClick={() => void create()}>Create</Button>
            </DialogActions>
          </DialogBody>
        </DialogSurface>
      </Dialog>
    </>
  );
}
