import React from "react";
import { Card, CardContent, CardHeader, CardTitle } from "../ui/card";
import { Input } from "../ui/input";
import { Textarea } from "../ui/textarea";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../ui/select";
import { Checkbox } from "../ui/checkbox";
import { RadioGroup, RadioGroupItem } from "../ui/radio-group";
import { Label } from "../ui/label";
import { Button } from "../ui/button";
import { Plus, Trash2, Upload, X } from "lucide-react";
import { FormField, FormSection, TableConfig } from "../../types";
import { Alert, AlertDescription } from "../ui/alert";
import {
  EXCEL_PREVIEW_SHEET_FRAME_CLASS,
  getExcelPreviewHotHeightPx,
} from "./excelSheetPreviewLayout";
import HandsontableWorkbook, {
  HandsontableWorkbookRef,
} from "./HandsontableWorkbook";

/**
 * Read-only preview renders a truncated workbook for performance. When HOT emits
 * edits, `next` only contains that truncated shape, so replacing local state with
 * it would drop off-screen rows/metadata (including checkbox cell meta).
 * Merge visible cell value edits back into the full workbook instead.
 */
function mergePreviewEditsIntoWorkbook(
  baseWorkbook: { sheets: any[] },
  editedPreviewWorkbook: { sheets: any[] },
): { sheets: any[] } {
  const baseSheets = Array.isArray(baseWorkbook?.sheets) ? baseWorkbook.sheets : [];
  const editedSheets = Array.isArray(editedPreviewWorkbook?.sheets)
    ? editedPreviewWorkbook.sheets
    : [];

  const mergedSheets = baseSheets.map((baseSheet: any, sheetIndex: number) => {
    const editedSheet = editedSheets[sheetIndex];
    if (!editedSheet) return baseSheet;

    const baseGrid = Array.isArray(baseSheet?.grid) ? baseSheet.grid : [];
    const editedGrid = Array.isArray(editedSheet?.grid) ? editedSheet.grid : [];
    if (!editedGrid.length) return baseSheet;

    const nextGrid = [...baseGrid];
    for (let r = 0; r < editedGrid.length; r++) {
      const editedRow = Array.isArray(editedGrid[r]) ? editedGrid[r] : [];
      if (!Array.isArray(nextGrid[r])) {
        nextGrid[r] = [];
      } else {
        nextGrid[r] = [...nextGrid[r]];
      }
      for (let c = 0; c < editedRow.length; c++) {
        nextGrid[r][c] = editedRow[c];
      }
    }

    return {
      ...baseSheet,
      grid: nextGrid,
    };
  });

  return {
    ...baseWorkbook,
    sheets: mergedSheets,
  };
}

/** Memoized so opening preview does not rebuild a truncated workbook object on every parent render. */
const EmbeddedExcelHandsontableBlock: React.FC<{
  fieldName: string;
  workbook: { sheets: any[] };
  excelReadOnly: boolean;
  useLocalExcelState: boolean;
  setLocalExcelState: React.Dispatch<React.SetStateAction<Record<string, any>>>;
  onFieldChange: (fieldName: string, value: any) => void;
  workbookComponentRef?: React.RefObject<HandsontableWorkbookRef | null>;
}> = React.memo(function EmbeddedExcelHandsontableBlock({
  fieldName,
  workbook,
  excelReadOnly,
  useLocalExcelState,
  setLocalExcelState,
  onFieldChange,
  workbookComponentRef,
}) {
  const data = workbook;

  const readOnlyHotHeight = React.useMemo(
    () => (excelReadOnly ? getExcelPreviewHotHeightPx() : undefined),
    [excelReadOnly],
  );

  // Keep a ref to `workbook` so the memoized onChange below never goes stale
  // while also not needing `workbook` as a dep (which would recreate on every edit).
  const workbookRef = React.useRef(workbook);
  workbookRef.current = workbook;

  // Stable onChange: does not recreate when parent re-renders with a new
  // onFieldChange reference, preventing unnecessary HandsontableWorkbook
  // re-renders and hotTableSettings churn that would close the active editor.
  const handleWorkbookChange = React.useCallback(
    (next: { sheets: any[] }) => {
      if (useLocalExcelState) {
        setLocalExcelState((prev) => {
          const currentWorkbook =
            prev[fieldName]?.sheets?.length
              ? prev[fieldName]
              : workbookRef.current;
          const mergedWorkbook =
            excelReadOnly && currentWorkbook?.sheets?.length
              ? mergePreviewEditsIntoWorkbook(currentWorkbook, next)
              : next;
          return { ...prev, [fieldName]: mergedWorkbook };
        });
        return;
      }
      onFieldChange(fieldName, next);
    },
    [useLocalExcelState, setLocalExcelState, fieldName, excelReadOnly, onFieldChange],
  );

  return (
    <div className="space-y-2">
      <div className={EXCEL_PREVIEW_SHEET_FRAME_CLASS}>
        <HandsontableWorkbook
          ref={workbookComponentRef}
          data={data}
          readOnly={excelReadOnly}
          readOnlyHotHeight={readOnlyHotHeight}
          onChange={handleWorkbookChange}
        />
      </div>
    </div>
  );
});

interface SharedFormRendererProps {
  formState: {
    formType: "regular" | "table" | "mixed";
    title: string;
    description?: string;
    fields: FormField[];
    sections: FormSection[];
    tableConfig?: TableConfig;
  };
  formData: Record<string, any>;
  tableData: any[];

  onFieldChange: (fieldName: string, value: any) => void;
  onTableChange?: (rowIndex: number, columnName: string, value: any) => void;
  onMixedTableChange?: (
    sectionId: string,
    rowIndex: number,
    columnName: string,
    value: any,
  ) => void;
  onAddTableRow?: (tableId?: string) => void;
  onRemoveTableRow?: (rowIndex: number, tableId?: string) => void;
  submitButton?: React.ReactNode;
  lightweightExcelPreview?: boolean;
  useLocalExcelState?: boolean;
  excelReadOnly?: boolean;
  onResolvedFormDataChange?: (data: Record<string, any>) => void;
}

export type SharedFormRendererRef = {
  getResolvedFormData: () => Record<string, any>;
};

const SharedFormRenderer = React.forwardRef<
  SharedFormRendererRef,
  SharedFormRendererProps
>(function SharedFormRenderer({
  formState,
  formData,
  tableData,

  onFieldChange,
  onTableChange,
  onMixedTableChange,
  onAddTableRow,
  onRemoveTableRow,
  submitButton,
  lightweightExcelPreview = false,
  useLocalExcelState = false,
  excelReadOnly = false,
  onResolvedFormDataChange,
}, ref) {
  const [signatureErrors, setSignatureErrors] = React.useState<
    Record<string, string>
  >({});
  const [excelPreviewSheetIndex, setExcelPreviewSheetIndex] = React.useState<
    Record<string, number>
  >({});
  const [localExcelState, setLocalExcelState] = React.useState<
    Record<string, any>
  >({});
  const workbookRefs = React.useRef<
    Record<string, React.RefObject<HandsontableWorkbookRef | null>>
  >({});

  const getWorkbookRef = React.useCallback((fieldName: string) => {
    if (!workbookRefs.current[fieldName]) {
      workbookRefs.current[fieldName] =
        React.createRef<HandsontableWorkbookRef | null>();
    }
    return workbookRefs.current[fieldName];
  }, []);

  const resolvedFormData = React.useMemo(
    () => ({
      ...formData,
      ...localExcelState,
    }),
    [formData, localExcelState],
  );

  React.useEffect(() => {
    onResolvedFormDataChange?.(resolvedFormData);
  }, [onResolvedFormDataChange, resolvedFormData]);

  const getResolvedFormData = React.useCallback(() => {
    const nextData = {
      ...formData,
      ...localExcelState,
    };

    const syncWorkbookSnapshot = (field: FormField) => {
      if (field.type !== "embedded_excel") return;
      const workbookSnapshot = workbookRefs.current[
        field.name
      ]?.current?.getWorkbookSnapshot();
      if (workbookSnapshot?.sheets?.length) {
        nextData[field.name] = workbookSnapshot;
      }
    };

    for (const field of formState.fields || []) {
      syncWorkbookSnapshot(field);
    }
    for (const section of formState.sections || []) {
      for (const field of section.fields || []) {
        syncWorkbookSnapshot(field);
      }
    }

    return nextData;
  }, [formData, formState.fields, formState.sections, localExcelState]);

  React.useImperativeHandle(
    ref,
    () => ({
      getResolvedFormData,
    }),
    [getResolvedFormData],
  );

  const getFilePreview = (
    rawValue: unknown,
  ): { name: string; url?: string } | null => {
    if (!rawValue) return null;

    if (typeof rawValue === "string") {
      const trimmed = rawValue.trim();
      if (!trimmed) return null;
      const safeName = trimmed.split("/").pop() || trimmed;
      const isUrl = /^https?:\/\//i.test(trimmed) || trimmed.startsWith("/");
      return { name: safeName, url: isUrl ? trimmed : undefined };
    }

    if (rawValue instanceof FileList) {
      const first = rawValue[0];
      return first ? { name: first.name } : null;
    }

    if (rawValue instanceof File) {
      return { name: rawValue.name };
    }

    if (Array.isArray(rawValue) && rawValue.length > 0) {
      return getFilePreview(rawValue[0]);
    }

    if (typeof rawValue === "object" && rawValue !== null) {
      const candidate = rawValue as Record<string, unknown>;
      const name =
        (typeof candidate.originalName === "string" &&
          candidate.originalName) ||
        (typeof candidate.filename === "string" && candidate.filename) ||
        (typeof candidate.name === "string" && candidate.name) ||
        (typeof candidate.path === "string" &&
          candidate.path.split("/").pop()) ||
        "";
      const url =
        (typeof candidate.url === "string" && candidate.url) ||
        (typeof candidate.path === "string" && candidate.path) ||
        undefined;
      if (!name) return null;
      return { name, url };
    }

    return null;
  };

  // Helper functions for pre-filled data
  const getInitialCellValue = (
    config: TableConfig,
    rowIndex: number,
    columnName: string,
    userValue?: any,
  ) => {
    // Check if there's a pre-filled value
    const preFilledCell = config.preFilledData?.find(
      (cell) => cell.rowIndex === rowIndex && cell.columnName === columnName,
    );

    if (preFilledCell && preFilledCell.isReadOnly) {
      return preFilledCell.value;
    }

    return userValue || "";
  };

  const isCellReadOnly = (
    config: TableConfig,
    rowIndex: number,
    columnName: string,
  ) => {
    const preFilledCell = config.preFilledData?.find(
      (cell) => cell.rowIndex === rowIndex && cell.columnName === columnName,
    );
    return preFilledCell?.isReadOnly || false;
  };

  const handleSignatureUpload = async (
    fieldName: string,
    file: File | null,
  ) => {
    if (!file) {
      onFieldChange(fieldName, "");
      setSignatureErrors((prev) => ({ ...prev, [fieldName]: "" }));
      return;
    }

    // Validate file type
    const validTypes = [
      "image/jpeg",
      "image/jpg",
      "image/png",
      "image/gif",
      "image/webp",
    ];
    if (!validTypes.includes(file.type)) {
      setSignatureErrors((prev) => ({
        ...prev,
        [fieldName]:
          "Please upload a valid image file (JPEG, PNG, GIF, or WebP)",
      }));
      return;
    }

    // Validate file size (2MB = 2 * 1024 * 1024 bytes)
    const maxSize = 2 * 1024 * 1024;
    if (file.size > maxSize) {
      setSignatureErrors((prev) => ({
        ...prev,
        [fieldName]: "Image size must be less than 2MB",
      }));
      return;
    }

    // Convert to base64
    try {
      const base64 = await fileToBase64(file);
      onFieldChange(fieldName, base64);
      setSignatureErrors((prev) => ({ ...prev, [fieldName]: "" }));
    } catch (error) {
      setSignatureErrors((prev) => ({
        ...prev,
        [fieldName]: "Failed to process image. Please try again.",
      }));
    }
  };

  const fileToBase64 = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onload = () => resolve(reader.result as string);
      reader.onerror = (error) => reject(error);
    });
  };

  const getFieldClassName = (field: FormField) => {
    switch (field.layout?.width) {
      case "half":
        return "col-span-6";
      case "third":
        return "col-span-4";
      case "quarter":
        return "col-span-3";
      case "auto":
        return "col-span-auto";
      case "full":
      default:
        return "col-span-12";
    }
  };

  const renderField = (field: FormField) => {
    const value = formData[field.name] || "";

    switch (field.type) {
      case "text":
      case "email":
      case "number":
      case "date":
      case "datetime-local":
      case "time":
      case "phone":
      case "url":
        return (
          <Input
            type={field.type}
            value={value}
            onChange={(e) => onFieldChange(field.name, e.target.value)}
            placeholder={field.placeholder}
            required={field.required}
          />
        );

      case "file":
        const filePreview = getFilePreview(value);
        return (
          <div className="space-y-2">
            <Input
              type="file"
              onChange={(e) => onFieldChange(field.name, e.target.files)}
              required={field.required}
              className="file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-violet-50 file:text-violet-700 hover:file:bg-violet-100"
            />
            {filePreview && (
              <p className="text-xs text-gray-600">
                Current file:{" "}
                {filePreview.url ? (
                  <a
                    href={filePreview.url}
                    target="_blank"
                    rel="noreferrer"
                    className="text-blue-600 underline hover:text-blue-700"
                  >
                    {filePreview.name}
                  </a>
                ) : (
                  <span className="font-medium">{filePreview.name}</span>
                )}
              </p>
            )}
            {field.description && (
              <p className="text-xs text-gray-500">{field.description}</p>
            )}
          </div>
        );

      case "embedded_excel": {
        const templateWorkbook = field.excelTemplate;
        const currentWorkbook = formData[field.name];
        const localWorkbook = localExcelState[field.name];
        const workbook = localWorkbook?.sheets?.length
          ? localWorkbook
          : currentWorkbook?.sheets?.length
            ? currentWorkbook
            : templateWorkbook;
        if (!workbook?.sheets?.length) {
          return (
            <Alert>
              <AlertDescription className="text-sm">
                No spreadsheet template created yet. In the form builder
                Properties, click "Create Sheet".
              </AlertDescription>
            </Alert>
          );
        }
        if (lightweightExcelPreview) {
          const sheetIndex = Math.min(
            excelPreviewSheetIndex[field.name] || 0,
            Math.max(0, workbook.sheets.length - 1),
          );
          const activeSheet = workbook.sheets[sheetIndex];
          const grid = Array.isArray(activeSheet?.grid) ? activeSheet.grid : [];
          const rows = grid.slice(0, 60);
          let widest = 0;
          for (const r of rows) {
            if (Array.isArray(r) && r.length > widest) widest = r.length;
          }
          const maxCols = Math.min(24, widest);
          return (
            <div className="space-y-2">
              <p className="text-xs text-muted-foreground">
                Snapshot preview (first 60 rows × 24 columns). Fillable
                spreadsheet cells work on the live submission page.
              </p>
              <div className="flex flex-wrap gap-2">
                {workbook.sheets.map((sheet: any, idx: number) => (
                  <Button
                    key={`${field.name}-sheet-${idx}`}
                    type="button"
                    size="sm"
                    variant={idx === sheetIndex ? "default" : "outline"}
                    onClick={() =>
                      setExcelPreviewSheetIndex((prev) => ({
                        ...prev,
                        [field.name]: idx,
                      }))
                    }
                  >
                    {sheet.name || `Sheet${idx + 1}`}
                  </Button>
                ))}
              </div>
              <div
                className={`overflow-auto border rounded-md bg-background ${EXCEL_PREVIEW_SHEET_FRAME_CLASS}`}
                style={{ maxHeight: getExcelPreviewHotHeightPx() }}
              >
                <table className="w-full border-collapse text-xs">
                  <tbody>
                    {rows.map((row: any[], rIdx: number) => (
                      <tr key={`${field.name}-row-${rIdx}`}>
                        {Array.from({ length: maxCols }).map((_, cIdx) => (
                          <td
                            key={`${field.name}-cell-${rIdx}-${cIdx}`}
                            className="p-1 border min-w-[72px]"
                          >
                            {String(row?.[cIdx] ?? "")}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {(grid.length > 60 || maxCols >= 24) && (
                <p className="text-xs text-muted-foreground">
                  Preview limited to first 60 rows and 24 columns.
                </p>
              )}
            </div>
          );
        }
        return (
          <EmbeddedExcelHandsontableBlock
            fieldName={field.name}
            workbook={workbook}
            excelReadOnly={excelReadOnly}
            useLocalExcelState={useLocalExcelState}
            setLocalExcelState={setLocalExcelState}
            onFieldChange={onFieldChange}
            workbookComponentRef={getWorkbookRef(field.name)}
          />
        );
      }

      case "signature":
        return (
          <div className="space-y-2">
            <div className="flex flex-col items-center justify-center w-full p-4 border-2 border-dashed rounded-lg bg-gray-50">
              {value ? (
                <div className="relative w-full">
                  <img
                    src={value}
                    alt="Signature preview"
                    className="max-w-full max-h-40 mx-auto border border-gray-300 rounded"
                  />
                  <Button
                    type="button"
                    variant="destructive"
                    size="sm"
                    className="absolute top-2 right-2"
                    onClick={() => {
                      onFieldChange(field.name, "");
                      setSignatureErrors((prev) => ({
                        ...prev,
                        [field.name]: "",
                      }));
                    }}
                  >
                    <X className="w-4 h-4" />
                  </Button>
                </div>
              ) : (
                <div className="text-center">
                  <Upload className="w-8 h-8 mx-auto mb-2 text-gray-400" />
                  <Label
                    htmlFor={`signature-${field.name}`}
                    className="cursor-pointer text-sm font-medium text-blue-600 hover:text-blue-700"
                  >
                    Click to upload signature
                  </Label>
                  <p className="text-xs text-gray-500 mt-1">
                    Images only (JPEG, PNG, GIF, WebP) • Max 2MB
                  </p>
                  <Input
                    id={`signature-${field.name}`}
                    type="file"
                    accept="image/jpeg,image/jpg,image/png,image/gif,image/webp"
                    className="hidden"
                    onChange={(e) => {
                      const file = e.target.files?.[0] || null;
                      handleSignatureUpload(field.name, file);
                    }}
                    required={field.required}
                  />
                </div>
              )}
            </div>
            {signatureErrors[field.name] && (
              <Alert variant="destructive" className="py-2">
                <AlertDescription className="text-sm">
                  {signatureErrors[field.name]}
                </AlertDescription>
              </Alert>
            )}
            {field.description && !signatureErrors[field.name] && (
              <p className="text-xs text-gray-500">{field.description}</p>
            )}
          </div>
        );

      case "textarea":
        return (
          <Textarea
            value={value}
            onChange={(e) => onFieldChange(field.name, e.target.value)}
            placeholder={field.placeholder}
            required={field.required}
            rows={4}
          />
        );

      case "select":
        return (
          <Select
            value={value}
            onValueChange={(value) => onFieldChange(field.name, value)}
            required={field.required}
          >
            <SelectTrigger>
              <SelectValue placeholder="Select an option" />
            </SelectTrigger>
            <SelectContent>
              {field.options?.map((option: any) => (
                <SelectItem key={option.value} value={option.value}>
                  {option.label}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        );

      case "checkbox":
        return (
          <div className="space-y-2">
            {field.options?.map((option: any) => {
              const isChecked = Array.isArray(value)
                ? value.includes(option.value)
                : false;
              return (
                <div key={option.value} className="flex items-center space-x-2">
                  <Checkbox
                    id={`${field.name}-${option.value}`}
                    checked={isChecked}
                    onCheckedChange={(checked) => {
                      const currentValues = Array.isArray(value) ? value : [];
                      if (checked) {
                        // Add value to array if checked
                        onFieldChange(field.name, [
                          ...currentValues,
                          option.value,
                        ]);
                      } else {
                        // Remove value from array if unchecked
                        onFieldChange(
                          field.name,
                          currentValues.filter((v: any) => v !== option.value),
                        );
                      }
                    }}
                  />
                  <Label
                    htmlFor={`${field.name}-${option.value}`}
                    className="text-sm"
                  >
                    {option.label}
                  </Label>
                </div>
              );
            })}
          </div>
        );

      case "radio":
        return (
          <RadioGroup
            value={value}
            onValueChange={(value) => onFieldChange(field.name, value)}
          >
            {field.options?.map((option: any) => (
              <div key={option.value} className="flex items-center space-x-2">
                <RadioGroupItem
                  value={option.value}
                  id={`${field.name}-${option.value}`}
                />
                <Label htmlFor={`${field.name}-${option.value}`}>
                  {option.label}
                </Label>
              </div>
            ))}
          </RadioGroup>
        );

      default:
        return (
          <Input
            value={value}
            onChange={(e) => onFieldChange(field.name, e.target.value)}
            placeholder={field.placeholder}
            required={field.required}
          />
        );
    }
  };

  const renderTableInput = (
    column: any,
    rowIndex: number,
    value: any,
    sectionId?: string,
    tableConfig?: TableConfig,
  ) => {
    // Get the actual cell value (pre-filled or user input)
    const cellValue = tableConfig
      ? getInitialCellValue(tableConfig, rowIndex, column.name, value)
      : value || "";
    const isReadOnly = tableConfig
      ? isCellReadOnly(tableConfig, rowIndex, column.name)
      : false;

    const handleChange = (newValue: any) => {
      // Don't allow changes to read-only cells
      if (isReadOnly) return;

      if (sectionId && onMixedTableChange) {
        onMixedTableChange(sectionId, rowIndex, column.name, newValue);
      } else if (onTableChange) {
        onTableChange(rowIndex, column.name, newValue);
      }
    };

    const inputClass = `${isReadOnly ? "bg-blue-50 border-blue-200 cursor-not-allowed" : "bg-white"}`;

    switch (column.type) {
      case "text":
      case "number":
      case "date":
      case "email":
      case "phone":
        return (
          <Input
            type={column.type}
            value={cellValue}
            onChange={(e) => handleChange(e.target.value)}
            required={column.required && !isReadOnly}
            disabled={isReadOnly}
            className={inputClass}
          />
        );

      case "select":
        return (
          <Select
            value={cellValue}
            onValueChange={handleChange}
            required={column.required && !isReadOnly}
            disabled={isReadOnly}
          >
            <SelectTrigger className={inputClass}>
              <SelectValue
                placeholder={isReadOnly ? cellValue : "Select option"}
              />
            </SelectTrigger>
            {!isReadOnly && (
              <SelectContent>
                {column.options?.map((option: string) => (
                  <SelectItem key={option} value={option}>
                    {option}
                  </SelectItem>
                ))}
              </SelectContent>
            )}
          </Select>
        );

      case "checkbox":
        return (
          <Checkbox
            checked={!!cellValue}
            onCheckedChange={handleChange}
            disabled={isReadOnly}
            className={inputClass}
          />
        );

      default:
        return (
          <Input
            value={cellValue}
            onChange={(e) => handleChange(e.target.value)}
            required={column.required && !isReadOnly}
            disabled={isReadOnly}
            className={inputClass}
          />
        );
    }
  };

  const renderTable = (config: TableConfig, sectionId?: string) => {
    const rows = sectionId ? formData[`table_${sectionId}`] || [{}] : tableData;

    // Ensure we have at least the minimum required rows
    const minRowsNeeded = Math.max(config.defaultRows, config.minRows);
    while (rows.length < minRowsNeeded) {
      rows.push({});
    }

    const handleAddRow = () => {
      if (config.allowAddRows && onAddTableRow) {
        onAddTableRow(sectionId ? sectionId : config.id);
      }
    };

    const handleRemoveRow = (rowIndex: number) => {
      if (config.allowDeleteRows && onRemoveTableRow) {
        onRemoveTableRow(rowIndex, sectionId ? sectionId : config.id);
      }
    };

    return (
      <div className="space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <h4 className="font-medium text-md">{config.title}</h4>
          {config.allowAddRows && (
            <Button type="button" onClick={handleAddRow} size="sm">
              <Plus className="w-4 h-4 mr-2" />
              Add Row
            </Button>
          )}
        </div>

        <div className="overflow-x-auto">
          <table className="w-full border border-collapse">
            <thead>
              <tr className="">
                {config.columns.map((column) => (
                  <th key={column.name} className="p-2 text-left border">
                    {column.label}
                    {column.required && (
                      <span className="ml-1 text-destructive">*</span>
                    )}
                  </th>
                ))}
                {config.allowDeleteRows && (
                  <th
                    scope="col"
                    className="w-17 min-w-17 max-w-17 p-2 border text-center align-middle whitespace-nowrap"
                  >
                    Actions
                  </th>
                )}
              </tr>
            </thead>
            <tbody>
              {rows.map((row: any, rowIndex: number) => (
                <tr key={rowIndex}>
                  {config.columns.map((column) => (
                    <td key={column.name} className="p-2 border">
                      {renderTableInput(
                        column,
                        rowIndex,
                        row[column.name],
                        sectionId,
                        config,
                      )}
                    </td>
                  ))}
                  {config.allowDeleteRows && (
                    <td className="w-12 min-w-12 max-w-12 p-1 border text-center align-middle whitespace-nowrap">
                      <Button
                        type="button"
                        variant="destructive"
                        size="icon"
                        className="h-8 w-8 shrink-0"
                        onClick={() => handleRemoveRow(rowIndex)}
                        aria-label="Delete row"
                      >
                        <Trash2 className="w-4 h-4" />
                      </Button>
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {config.preFilledData && config.preFilledData.length > 0 && (
          <div className="text-xs text-gray-600 bg-blue-50 p-2 rounded">
            ℹ️ Blue highlighted cells contain pre-filled data and cannot be
            modified.
          </div>
        )}
      </div>
    );
  };

  const showFormCardHeader =
    Boolean(formState.title?.trim()) || Boolean(formState.description?.trim());

  return (
    <div className="space-y-6">
      <Card>
        {showFormCardHeader && (
          <CardHeader>
            {formState.title?.trim() ? (
              <CardTitle>{formState.title}</CardTitle>
            ) : null}
            {formState.description?.trim() ? (
              <p className="text-sm text-gray-600">{formState.description}</p>
            ) : null}
          </CardHeader>
        )}
        <CardContent className={!showFormCardHeader ? "pt-6" : undefined}>
          <div className="space-y-6">
            {formState.formType === "regular" && (
              <div className="grid grid-cols-12 gap-4">
                {formState.fields.map((field) => (
                  <div
                    key={field.id}
                    className={`space-y-2 ${getFieldClassName(field)}`}
                  >
                    <Label
                      htmlFor={field.name}
                      className={field.required ? "form-field-required" : ""}
                    >
                      {field.label}
                      {field.required && (
                        <span className="text-red-500 ml-1">*</span>
                      )}
                    </Label>
                    {renderField(field)}
                  </div>
                ))}
              </div>
            )}

            {formState.formType === "mixed" &&
              formState.sections.map((section) => (
                <div key={section.id} className="space-y-4">
                  <h3 className="pb-2 text-lg font-semibold border-b">
                    {section.title}
                  </h3>
                  {section.description && (
                    <p className="text-sm text-gray-600">
                      {section.description}
                    </p>
                  )}

                  {section.type === "fields" && section.fields && (
                    <div className="grid grid-cols-12 gap-4">
                      {section.fields.map((field) => (
                        <div
                          key={field.id}
                          className={`space-y-2 ${getFieldClassName(field)}`}
                        >
                          <Label
                            htmlFor={field.name}
                            className={
                              field.required ? "form-field-required" : ""
                            }
                          >
                            {field.label}
                            {field.required && (
                              <span className="text-red-500 ml-1">*</span>
                            )}
                          </Label>
                          {renderField(field)}
                        </div>
                      ))}
                    </div>
                  )}

                  {section.type === "table" &&
                    section.tableConfig &&
                    renderTable(section.tableConfig, section.id)}
                </div>
              ))}

            {formState.formType === "table" &&
              formState.tableConfig &&
              renderTable(formState.tableConfig)}

            {submitButton && (
              <div className="flex items-center space-x-4">{submitButton}</div>
            )}
          </div>
        </CardContent>
      </Card>
    </div>
  );
});

export default SharedFormRenderer;
