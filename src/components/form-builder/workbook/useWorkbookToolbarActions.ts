import React from "react";
import { SheetData } from "./workbookTypes";
import {
  SINGLE_CHECKBOX_CLASS,
  YES_NO_PAIR_TOKEN_PREFIX,
  buildYesNoOppositeMap,
} from "./workbookUtils";

type Range = {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
};

export const useWorkbookToolbarActions = ({
  hotRef,
  readOnly,
  activeSheetIndexRef,
  workbookRef,
  yesNoOppositeCellMapRef,
  getToolbarActionRange,
  collectCurrentSheetFromHot,
  scheduleUndoRedoRefresh,
  restoreHotRange,
}: {
  hotRef: React.MutableRefObject<any>;
  readOnly: boolean;
  activeSheetIndexRef: React.MutableRefObject<number>;
  workbookRef: React.MutableRefObject<{ sheets: SheetData[] }>;
  yesNoOppositeCellMapRef: React.MutableRefObject<Map<string, { row: number; col: number }>>;
  getToolbarActionRange: (hot: any) => Range | null;
  collectCurrentSheetFromHot: (includeMeta: boolean, sheetIndex?: number) => void;
  scheduleUndoRedoRefresh: () => void;
  restoreHotRange: (hot: any, range: Range | null) => void;
}) => {
  const applyCheckboxMetaToSelection = React.useCallback(
    (
      options:
        | { kind: "yesno"; checkedTemplate: "YES"; uncheckedTemplate: "NO" }
        | { kind: "checkbox"; checkedTemplate: "true"; uncheckedTemplate: "" },
    ) => {
      const hot = hotRef.current?.hotInstance;
      if (!hot || readOnly) return;
      const range = getToolbarActionRange(hot);
      if (!range) return;
      if (options.kind === "yesno" && range.endCol - range.startCol + 1 !== 2) return;
      const pairEpoch = Date.now();

      const apply = () => {
        if (options.kind === "yesno") {
          for (let r = range.startRow; r <= range.endRow; r++) {
            const leftCol = range.startCol;
            const rightCol = range.startCol + 1;
            const pairToken = `${YES_NO_PAIR_TOKEN_PREFIX}${pairEpoch}-${r}`;
            const coords: Array<[number, number]> = [
              [r, leftCol],
              [r, rightCol],
            ];
            for (const [rowIndex, colIndex] of coords) {
              const currentTokens = String(hot.getCellMeta(rowIndex, colIndex)?.className || "")
                .split(/\s+/)
                .filter(Boolean)
                .filter(
                  (token: string) =>
                    token !== SINGLE_CHECKBOX_CLASS &&
                    !token.startsWith(YES_NO_PAIR_TOKEN_PREFIX),
                );
              currentTokens.push(pairToken);
              hot.setCellMeta(rowIndex, colIndex, "className", currentTokens.join(" "));
              hot.setCellMeta(rowIndex, colIndex, "type", "checkbox");
              hot.setCellMeta(rowIndex, colIndex, "checkedTemplate", options.checkedTemplate);
              hot.setCellMeta(rowIndex, colIndex, "uncheckedTemplate", options.uncheckedTemplate);
            }
          }
          return;
        }
        for (let r = range.startRow; r <= range.endRow; r++) {
          for (let c = range.startCol; c <= range.endCol; c++) {
            const currentTokens = String(hot.getCellMeta(r, c)?.className || "")
              .split(/\s+/)
              .filter(Boolean)
              .filter((token: string) => !token.startsWith(YES_NO_PAIR_TOKEN_PREFIX));
            if (!currentTokens.includes(SINGLE_CHECKBOX_CLASS))
              currentTokens.push(SINGLE_CHECKBOX_CLASS);
            hot.setCellMeta(r, c, "className", currentTokens.join(" "));
            hot.setCellMeta(r, c, "type", "checkbox");
            hot.setCellMeta(r, c, "checkedTemplate", options.checkedTemplate);
            hot.setCellMeta(r, c, "uncheckedTemplate", options.uncheckedTemplate);
          }
        }
      };

      if (typeof hot.batch === "function") hot.batch(apply);
      else apply();

      collectCurrentSheetFromHot(true);
      const idx = activeSheetIndexRef.current;
      const sheet = workbookRef.current.sheets[idx];
      yesNoOppositeCellMapRef.current = buildYesNoOppositeMap(sheet?.cellMeta);
      scheduleUndoRedoRefresh();
      hot.render();
      restoreHotRange(hot, range);
    },
    [
      activeSheetIndexRef,
      collectCurrentSheetFromHot,
      getToolbarActionRange,
      hotRef,
      readOnly,
      restoreHotRange,
      scheduleUndoRedoRefresh,
      workbookRef,
      yesNoOppositeCellMapRef,
    ],
  );

  const setYesNoToggle = React.useCallback(() => {
    applyCheckboxMetaToSelection({
      kind: "yesno",
      checkedTemplate: "YES",
      uncheckedTemplate: "NO",
    });
  }, [applyCheckboxMetaToSelection]);

  const setSingleCheckbox = React.useCallback(() => {
    applyCheckboxMetaToSelection({
      kind: "checkbox",
      checkedTemplate: "true",
      uncheckedTemplate: "",
    });
  }, [applyCheckboxMetaToSelection]);

  return {
    setYesNoToggle,
    setSingleCheckbox,
  };
};

