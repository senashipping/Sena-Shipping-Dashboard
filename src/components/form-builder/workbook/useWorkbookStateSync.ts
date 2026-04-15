import React from "react";
import { SheetData } from "./workbookTypes";
import { deepCloneSheet, workbookSignature } from "./workbookUtils";

export const useWorkbookStateSync = ({
  workbookRef,
  lastIncomingSignatureRef,
  suppressNextHotReloadRef,
  onChange,
}: {
  workbookRef: React.MutableRefObject<{ sheets: SheetData[] }>;
  lastIncomingSignatureRef: React.MutableRefObject<string>;
  suppressNextHotReloadRef: React.MutableRefObject<boolean>;
  onChange: (next: { sheets: SheetData[] }) => void;
}) => {
  // Keep a ref so emitWorkbookToParent is always stable even when the parent
  // passes a new onChange function reference on every render. Without this, a
  // new onChange reference would cascade through useCallback deps all the way
  // into hotTableSettings, causing HotTable to call hot.updateSettings() and
  // close any open text editor mid-typing.
  const onChangeRef = React.useRef(onChange);
  onChangeRef.current = onChange;

  const emitWorkbookToParent = React.useCallback(() => {
    const nextSheets = workbookRef.current.sheets.map(deepCloneSheet);
    const snapshot = { sheets: nextSheets };
    lastIncomingSignatureRef.current = workbookSignature(nextSheets);
    suppressNextHotReloadRef.current = true;
    onChangeRef.current(snapshot);
  }, [lastIncomingSignatureRef, suppressNextHotReloadRef, workbookRef]);

  return {
    emitWorkbookToParent,
  };
};

