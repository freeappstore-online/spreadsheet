import { useSpreadsheet } from "./hooks/useSpreadsheet";
import { TopBar } from "./components/TopBar";
import { FormulaBar } from "./components/FormulaBar";
import { FormattingToolbar } from "./components/FormattingToolbar";
import { ColorPicker } from "./components/ColorPicker";
import { Grid } from "./components/Grid";
import { FindBar } from "./components/FindBar";
import { SheetTabs } from "./components/SheetTabs";
import { StatusBar } from "./components/StatusBar";
import { ContextMenu } from "./components/ContextMenu";
import { HelpDialog } from "./components/HelpDialog";
import { ConditionalFormatDialog } from "./components/ConditionalFormatDialog";

export function App() {
  const s = useSpreadsheet();
  return (
    <div style={{ display: "flex", flexDirection: "column", height: "100dvh" }}>
      <TopBar s={s} />
      <FormulaBar s={s} />
      <FormattingToolbar s={s} />
      <ColorPicker s={s} />
      <Grid s={s} />
      {s.findBar.open && <FindBar s={s} />}
      <SheetTabs s={s} />
      <StatusBar s={s} />
      {s.contextMenu && <ContextMenu s={s} />}
      {s.showHelp && <HelpDialog s={s} />}
      {s.condFormatDialog && <ConditionalFormatDialog s={s} />}
    </div>
  );
}

export default App;
