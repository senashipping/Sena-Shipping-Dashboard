import React from "react";
import { HotTable } from "@handsontable/react";
import "handsontable/styles/handsontable.css";
import "handsontable/styles/ht-theme-main.css";
import { registerAllModules } from "handsontable/registry";
import { Button } from "../ui/button";

registerAllModules();

type SheetData = {
  name: string;
  grid: string[][];
  mergeCells?: Array<{ row: number; col: number; rowspan: number; colspan: number }>;
  cellMeta?: Array<{ row: number; col: number; className?: string; readOnly?: boolean }>;
};

interface HandsontableWorkbookProps {
  data: { sheets: SheetData[] };
  onChange: (next: { sheets: SheetData[] }) => void;
  readOnly?: boolean;
}

const HandsontableWorkbook: React.FC<HandsontableWorkbookProps> = ({
  data,
  onChange,
  readOnly = false,
}) => {
  const [activeSheetIndex, setActiveSheetIndex] = React.useState(0);
  const [renaming, setRenaming] = React.useState(false);
  const [renameValue, setRenameValue] = React.useState("");
  const hotRef = React.useRef<any>(null);
  const safeSheets =
    Array.isArray(data?.sheets) && data.sheets.length > 0
      ? data.sheets
      : [{ name: "Sheet1", grid: [[""]] }];
  const activeSheet = safeSheets[Math.min(activeSheetIndex, safeSheets.length - 1)];
  const safeGrid =
    Array.isArray(activeSheet?.grid) && activeSheet.grid.length > 0 ? activeSheet.grid : [[""]];

  const emitWorkbook = (nextSheet: SheetData) => {
    const nextSheets = safeSheets.map((sheet, index) =>
      index === activeSheetIndex ? nextSheet : sheet
    );
    onChange({ sheets: nextSheets });
  };

  const collectAndEmitMeta = (nextGrid: string[][]) => {
    const hot = hotRef.current?.hotInstance;
    const mergeCells =
      hot?.getPlugin?.("mergeCells")?.mergedCellsCollection?.mergedCells?.map((cell: any) => ({
        row: cell.row,
        col: cell.col,
        rowspan: cell.rowspan,
        colspan: cell.colspan,
      })) || [];
    const cellMeta: Array<{ row: number; col: number; className?: string; readOnly?: boolean }> =
      [];
    for (let r = 0; r < nextGrid.length; r++) {
      for (let c = 0; c < (nextGrid[r]?.length || 0); c++) {
        const meta = hot?.getCellMeta?.(r, c);
        if (meta?.className || meta?.readOnly) {
          cellMeta.push({
            row: r,
            col: c,
            className: meta.className,
            readOnly: meta.readOnly,
          });
        }
      }
    }
    emitWorkbook({
      ...activeSheet,
      grid: nextGrid,
      mergeCells,
      cellMeta,
    });
  };

  const syncFromHot = () => {
    const hot = hotRef.current?.hotInstance;
    if (!hot) return;
    const next = (hot.getData?.() || safeGrid).map((row: any[]) =>
      row.map((cell) => (cell == null ? "" : String(cell)))
    );
    collectAndEmitMeta(next);
  };

  return (
    <div className="space-y-2">
      <div className="flex flex-wrap items-center gap-2">
        {safeSheets.map((sheet, index) => (
          <Button
            key={`${sheet.name}-${index}`}
            type="button"
            variant={index === activeSheetIndex ? "default" : "outline"}
            size="sm"
            onClick={() => setActiveSheetIndex(index)}
          >
            {sheet.name}
          </Button>
        ))}
        {!readOnly && (
          <>
          {renaming ? (
            <div className="flex items-center gap-1">
              <input
                value={renameValue}
                onChange={(e) => setRenameValue(e.target.value)}
                className="h-8 px-2 text-sm border rounded"
                placeholder="Sheet name"
              />
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => {
                  const nextName = renameValue.trim();
                  if (!nextName) return;
                  const nextSheets = safeSheets.map((sheet, index) =>
                    index === activeSheetIndex ? { ...sheet, name: nextName } : sheet
                  );
                  onChange({ sheets: nextSheets });
                  setRenaming(false);
                }}
              >
                Save
              </Button>
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => setRenaming(false)}
              >
                Cancel
              </Button>
            </div>
          ) : (
            <Button
              type="button"
              variant="outline"
              size="sm"
              className="text-yellow-900 border-yellow-500 bg-yellow-300 hover:bg-yellow-400 hover:border-yellow-600"
              onClick={() => {
                setRenameValue(activeSheet?.name || "");
                setRenaming(true);
              }}
            >
              Rename Sheet
            </Button>
          )}
          <Button
            type="button"
            variant="outline"
            size="sm"
            className="text-green-900 border-green-500 bg-green-300 hover:bg-green-400 hover:border-green-600"
            onClick={() => {
              const nextSheets = [
                ...safeSheets,
                { name: `Sheet${safeSheets.length + 1}`, grid: [[""]] },
              ];
              onChange({ sheets: nextSheets });
              setActiveSheetIndex(nextSheets.length - 1);
            }}
          >
            + Add Sheet
          </Button>
          </>
        )}
      </div>
      <div className="overflow-hidden border rounded-md">
      <HotTable
        ref={hotRef}
        data={safeGrid}
        themeName="ht-theme-main"
        rowHeaders
        colHeaders
        licenseKey="non-commercial-and-evaluation"
        readOnly={readOnly}
        width="100%"
        stretchH="all"
        height={320}
        mergeCells={activeSheet.mergeCells || true}
        contextMenu={
          readOnly
            ? false
            : {
                items: {
                  row_above: {},
                  row_below: {},
                  col_left: {},
                  col_right: {},
                  hsep1: "---------",
                  remove_row: {},
                  remove_col: {},
                  hsep2: "---------",
                  mergeCells: {},
                  hsep3: "---------",
                  alignment: {},
                  hsep4: "---------",
                  undo: {},
                  redo: {},
                },
              }
        }
        className="ht-theme-main"
        manualRowResize
        manualColumnResize
        wordWrap
        autoWrapRow
        autoWrapCol
        cell={activeSheet.cellMeta as any}
        afterChange={(changes, source) => {
          if (!changes || source === "loadData" || readOnly) return;
          syncFromHot();
        }}
        afterMergeCells={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterUnmergeCells={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterCreateRow={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterCreateCol={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterRemoveRow={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterRemoveCol={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterSetCellMeta={() => {
          if (readOnly) return;
          syncFromHot();
        }}
        afterContextMenuExecute={() => {
          if (readOnly) return;
          syncFromHot();
        }}
      />
    </div>
    </div>
  );
};

export default HandsontableWorkbook;
