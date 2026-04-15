import React from "react";
import { SheetData } from "./workbookTypes";

type SelectionRange = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export const useWorkbookSelection = ({
  workbookRef,
  activeSheetIndexRef,
}: {
  workbookRef: React.MutableRefObject<{ sheets: SheetData[] }>;
  activeSheetIndexRef: React.MutableRefObject<number>;
}) => {
  const lastSelectionRef = React.useRef<SelectionRange>({
    startRow: 0,
    endRow: 0,
    startCol: 0,
    endCol: 0,
  });

  const sheetSelectionRef = React.useRef<Record<number, SelectionRange>>({});

  const getToolbarActionRange = React.useCallback((hot: any) => {
    if (!hot) return null;
    const idx = activeSheetIndexRef.current;
    const sheet = workbookRef.current.sheets[idx];
    const rowCount = Math.max(1, sheet?.grid?.length || 1);
    const colCount = Math.max(
      1,
      ...(sheet?.grid || []).map((row) => (Array.isArray(row) ? row.length : 0)),
    );

    const clamp = (range: SelectionRange): SelectionRange => ({
      startRow: Math.max(
        0,
        Math.min(rowCount - 1, Math.min(range.startRow, range.endRow)),
      ),
      endRow: Math.max(
        0,
        Math.min(rowCount - 1, Math.max(range.startRow, range.endRow)),
      ),
      startCol: Math.max(
        0,
        Math.min(colCount - 1, Math.min(range.startCol, range.endCol)),
      ),
      endCol: Math.max(
        0,
        Math.min(colCount - 1, Math.max(range.startCol, range.endCol)),
      ),
    });

    const sel =
      typeof hot.getSelectedRangeLast === "function"
        ? hot.getSelectedRangeLast()
        : null;
    if (sel?.from != null && sel?.to != null) {
      const r = clamp({
        startRow: sel.from.row,
        endRow: sel.to.row,
        startCol: sel.from.col,
        endCol: sel.to.col,
      });
      lastSelectionRef.current = r;
      sheetSelectionRef.current[idx] = r;
      return r;
    }

    const last = hot.getSelectedLast?.();
    if (
      last &&
      last.length >= 4 &&
      last.every((v: any) => Number.isInteger(v))
    ) {
      const r = clamp({
        startRow: last[0],
        endRow: last[2],
        startCol: last[1],
        endCol: last[3],
      });
      lastSelectionRef.current = r;
      sheetSelectionRef.current[idx] = r;
      return r;
    }

    const cached = sheetSelectionRef.current[idx] ?? lastSelectionRef.current;
    return clamp(cached);
  }, [activeSheetIndexRef, workbookRef]);

  const restoreHotRange = React.useCallback(
    (hot: any, range: SelectionRange | null) => {
      if (!hot || !range) return;
      hot.render();
      hot.selectCell(
        range.startRow,
        range.startCol,
        range.endRow,
        range.endCol,
        false,
        false,
      );
    },
    [],
  );

  return {
    lastSelectionRef,
    sheetSelectionRef,
    getToolbarActionRange,
    restoreHotRange,
  };
};

