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
  const emitWorkbookToParent = React.useCallback(() => {
    const nextSheets = workbookRef.current.sheets.map(deepCloneSheet);
    const snapshot = { sheets: nextSheets };
    lastIncomingSignatureRef.current = workbookSignature(nextSheets);
    suppressNextHotReloadRef.current = true;
    onChange(snapshot);
  }, [lastIncomingSignatureRef, onChange, suppressNextHotReloadRef, workbookRef]);

  return {
    emitWorkbookToParent,
  };
};

