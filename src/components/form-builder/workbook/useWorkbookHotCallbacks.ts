import React from "react";
import { SheetData } from "./workbookTypes";
import { cellCoordKey, toCheckboxChecked } from "./workbookUtils";

export const useWorkbookHotCallbacks = ({
  hotRef,
  yesNoOppositeCellMapRef,
  readOnly,
  scheduleUndoRedoRefresh,
  activeSheetIndexRef,
  workbookRef,
  readOnlyPreviewDirtyRef,
  isEditingRef,
  pendingReadOnlyEmitRef,
  onReadOnlyEdit,
  onCellChanges,
}: {
  hotRef: React.MutableRefObject<any>;
  yesNoOppositeCellMapRef: React.MutableRefObject<
    Map<string, { row: number; col: number }>
  >;
  readOnly: boolean;
  scheduleUndoRedoRefresh: () => void;
  activeSheetIndexRef: React.MutableRefObject<number>;
  workbookRef: React.MutableRefObject<{ sheets: SheetData[] }>;
  readOnlyPreviewDirtyRef: React.MutableRefObject<boolean>;
  isEditingRef: React.MutableRefObject<boolean>;
  pendingReadOnlyEmitRef: React.MutableRefObject<boolean>;
  onReadOnlyEdit?: () => void;
  onCellChanges?: (
    changes: [number, number, unknown, unknown][],
    source: string,
  ) => void;
}) => {
  const afterChange = React.useCallback(
    (changes: any, source: string) => {
      setTimeout(() => {
        if (source === "loadData") return;
        const hot = hotRef.current?.hotInstance;

        if (
          Array.isArray(changes) &&
          changes.length > 0 &&
          source !== "updateData" &&
          String(source) !== "yesNoSync"
        ) {
          onCellChanges?.(
            changes as [number, number, unknown, unknown][],
            source,
          );
          if (hot) {
            const oppositeCellByKey = yesNoOppositeCellMapRef.current;
            for (const [row, col, oldValue, newValue] of changes as [
              number,
              number,
              unknown,
              unknown,
            ][]) {
              if (!Number.isFinite(row) || !Number.isFinite(col)) continue;
              if (Object.is(oldValue, newValue)) continue;
              const key = cellCoordKey(row, col);
              const opposite = oppositeCellByKey.get(key);
              if (!opposite || !toCheckboxChecked(newValue)) continue;
              setTimeout(() => {
                const oppositeCurrentValue = hot.getDataAtCell(
                  opposite.row,
                  opposite.col,
                );
                if (!toCheckboxChecked(oppositeCurrentValue)) return;
                hot.setDataAtCell(opposite.row, opposite.col, "", "yesNoSync");
              }, 0);
            }
          }
        }

        if (!readOnly && Array.isArray(changes) && changes.length > 0) {
          scheduleUndoRedoRefresh();
        }
        if (readOnly && changes && source !== "updateData") {
          const idx = activeSheetIndexRef.current;
          const sheet = workbookRef.current.sheets[idx];
          if (!sheet) return;
          const baseGrid = sheet.grid;
          let newGrid: string[][] = baseGrid;
          const clonedRows = new Set<number>();
          for (const [row, col, , newValue] of changes as [
            number,
            number,
            unknown,
            unknown,
          ][]) {
            if (
              typeof row !== "number" ||
              typeof col !== "number" ||
              !Array.isArray(baseGrid[row])
            )
              continue;
            if (newGrid === baseGrid) newGrid = [...baseGrid];
            if (!clonedRows.has(row)) {
              newGrid[row] = [...baseGrid[row]];
              clonedRows.add(row);
            }
            const normalizedNewValue = newValue == null ? "" : String(newValue);
            const currentValue = String(baseGrid[row][col] ?? "");
            if (currentValue === normalizedNewValue) continue;
            newGrid[row][col] = normalizedNewValue;
          }
          if (newGrid !== baseGrid) {
            workbookRef.current.sheets[idx] = {
              ...sheet,
              grid: newGrid,
            };
            readOnlyPreviewDirtyRef.current = true;
            if (source === "edit") {
              // Keep local workbook state in sync during typing, but do not emit
              // to parent until editing has ended.
              pendingReadOnlyEmitRef.current = true;
              return;
            }
            const isEditorStillOpen =
              typeof hot?.isEditorOpened === "function" &&
              hot.isEditorOpened();
            // Defer all parent syncs until a dedicated "editing finished" phase.
            // This prevents focus/caret jumps when imported templates emit mixed
            // change sources while users are still typing.
            pendingReadOnlyEmitRef.current = true;
            if (!isEditorStillOpen && !isEditingRef.current) {
              onReadOnlyEdit?.();
            }
          }
        }
      }, 0);
    },
    [
      activeSheetIndexRef,
      hotRef,
      readOnly,
      readOnlyPreviewDirtyRef,
      isEditingRef,
      pendingReadOnlyEmitRef,
      onReadOnlyEdit,
      onCellChanges,
      scheduleUndoRedoRefresh,
      workbookRef,
      yesNoOppositeCellMapRef,
    ],
  );

  return { afterChange };
};
