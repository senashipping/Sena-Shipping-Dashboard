import React from "react";
import { Card, CardContent, CardHeader, CardTitle } from "../ui/card";
import { Input } from "../ui/input";
import { Textarea } from "../ui/textarea";
import { Label } from "../ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../ui/select";
import { Checkbox } from "../ui/checkbox";
import { Button } from "../ui/button";
import { Plus, Trash2, X } from "lucide-react";
import {
  Dialog,
  DialogContent,
  DialogDescription,
  DialogHeader,
  DialogTitle,
} from "../ui/dialog";
import { FormField, FormSection, TableConfig, TableColumn } from "../../types";
import { useToast } from "../ui/toast";
import * as XLSX from "xlsx";
import type { HandsontableWorkbookRef } from "./workbook/workbookTypes";
import HandsontableWorkbook from "./HandsontableWorkbook";

interface PropertiesPanelProps {
  selectedItem: {
    id: string;
    type: "field" | "section" | "table";
    data: FormField | FormSection | TableConfig;
  } | null;
  onUpdate: (updates: any) => void;
  className?: string;
}

const PropertiesPanel: React.FC<PropertiesPanelProps> = ({
  selectedItem,
  onUpdate,
  className = "",
}) => {
  const { toast } = useToast();
  const [sheetRows, setSheetRows] = React.useState(10);
  const [sheetCols, setSheetCols] = React.useState(6);
  const [sheetCount, setSheetCount] = React.useState(1);
  const [isImportingTemplate, setIsImportingTemplate] = React.useState(false);
  const [isWorkbookEditorOpen, setIsWorkbookEditorOpen] = React.useState(false);
  const [workbookDraft, setWorkbookDraft] = React.useState<any | null>(null);
  const [applyImportedLogoToAllSheets, setApplyImportedLogoToAllSheets] = React.useState(true);
  const workbookEditorRef = React.useRef<HandsontableWorkbookRef>(null);
  const importTemplateInputRef = React.useRef<HTMLInputElement | null>(null);
  React.useEffect(() => {
    setIsWorkbookEditorOpen(false);
    setWorkbookDraft(null);
  }, [selectedItem?.id]);


  const MAX_IMPORT_ROWS = 650;
  const MAX_IMPORT_COLS = 120;
  const MAX_STYLE_SCAN_CELLS = 26000;
  const MAX_IMPORT_IMAGES = 40;

  const normalizeHexColor = (raw?: string): string | null => {
    if (!raw) return null;
    const clean = raw.replace("#", "").trim();
    if (clean.length === 6) return clean.toLowerCase();
    if (clean.length === 8) return clean.slice(2).toLowerCase();
    return null;
  };

  const getXlsxColorHex = (color: any): string | null => {
    if (!color) return null;
    if (typeof color.rgb === "string") return normalizeHexColor(color.rgb);
    if (typeof color.argb === "string") return normalizeHexColor(color.argb);
    return null;
  };

  const toBase64 = (bytes: Uint8Array): string => {
    let binary = "";
    const chunk = 0x8000;
    for (let i = 0; i < bytes.length; i += chunk) {
      binary += String.fromCharCode(...bytes.subarray(i, i + chunk));
    }
    return btoa(binary);
  };

  const ensureWorkbookLogoOnEverySheet = (
    sheets: Array<{
      name: string;
      grid: string[][];
      images?: Array<{ row: number; col: number; rowspan?: number; colspan?: number; dataUrl: string }>;
    }>,
  ) => {
    const firstLogo =
      sheets.flatMap((sheet) => sheet.images || []).find((img) => typeof img?.dataUrl === "string" && img.dataUrl.length > 0) || null;
    if (!firstLogo) {
      return { sheets, patchedCount: 0 };
    }

    let patchedCount = 0;
    const nextSheets = sheets.map((sheet) => {
      const hasLogo = Array.isArray(sheet.images) && sheet.images.some((img) => typeof img?.dataUrl === "string" && img.dataUrl.length > 0);
      if (hasLogo) return sheet;

      patchedCount += 1;
      const fallbackRow = Math.max(0, Math.min(firstLogo.row || 0, Math.max(0, (sheet.grid?.length || 1) - 1)));
      const fallbackCol = Math.max(0, Math.min(firstLogo.col || 0, Math.max(0, (sheet.grid?.[0]?.length || 1) - 1)));
      return {
        ...sheet,
        images: [
          {
            row: fallbackRow,
            col: fallbackCol,
            rowspan: firstLogo.rowspan || 1,
            colspan: firstLogo.colspan || 1,
            dataUrl: firstLogo.dataUrl,
          },
        ],
      };
    });

    return { sheets: nextSheets, patchedCount };
  };

  if (!selectedItem) {
    return (
      <div className={`p-6 ${className}`}>
        <div className="text-center text-gray-500">
          <h3 className="text-lg font-medium mb-2">Properties</h3>
          <p className="text-sm">Select an element to edit its properties</p>
        </div>
      </div>
    );
  }

  const importWorkbookTemplate = async (file: File) => {
    if (isImportingTemplate) return;
    setIsImportingTemplate(true);
    try {
      if (!/\.xlsx$/i.test(file.name)) {
        toast({
          title: "Import failed",
          description: "Only .xlsx files are supported.",
          variant: "destructive",
        });
        return;
      }

      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, {
        type: "array",
        cellDates: true,
        cellFormula: true,
        cellNF: true,
        cellStyles: true,
      });
      const imageMapBySheet = new Map<string, Array<{ row: number; col: number; rowspan?: number; colspan?: number; dataUrl: string }>>();
      const excelJsSheetMap = new Map<string, any>();
      let excelJsBooleanValueType: unknown = undefined;
      try {
        const { default: ExcelJS } = await import("exceljs");
        excelJsBooleanValueType = (ExcelJS as any)?.ValueType?.Boolean;
        const excelJsWorkbook = new ExcelJS.Workbook();
        await excelJsWorkbook.xlsx.load(buffer as ArrayBuffer);
        for (const ws of excelJsWorkbook.worksheets) {
          excelJsSheetMap.set(ws.name, ws);
          const items: Array<{ row: number; col: number; rowspan?: number; colspan?: number; dataUrl: string }> = [];
          const images = ws.getImages().slice(0, MAX_IMPORT_IMAGES);
          for (const item of images) {
            const image = excelJsWorkbook.getImage(Number(item.imageId)) as any;
            if (!image?.buffer || !image?.extension) continue;
            const range = item.range as any;
            const tl = range?.tl;
            if (!tl) continue;
            const row = Math.max(0, Math.floor(Number(tl.nativeRow || 0)));
            const col = Math.max(0, Math.floor(Number(tl.nativeCol || 0)));
            const brRow = Math.max(row + 1, Math.ceil(Number(range?.br?.nativeRow || row + 1)));
            const brCol = Math.max(col + 1, Math.ceil(Number(range?.br?.nativeCol || col + 1)));
            const rawBuffer = image.buffer instanceof Uint8Array ? image.buffer : new Uint8Array(image.buffer);
            const dataUrl = `data:image/${String(image.extension).toLowerCase()};base64,${toBase64(rawBuffer)}`;
            items.push({
              row,
              col,
              rowspan: Math.max(1, brRow - row),
              colspan: Math.max(1, brCol - col),
              dataUrl,
            });
          }
          imageMapBySheet.set(ws.name, items);
        }
      } catch {
        // Keep import robust even if image extraction fails.
      }

      const importedSheets = workbook.SheetNames.map((name, index) => {
        const ws = workbook.Sheets[name];
        const ref = ws?.["!ref"] || "A1";
        const range = XLSX.utils.decode_range(ref);
        // Excel often keeps `!ref` tight to used values while `!merges` extends past it.
        // Importing only the ref box leaves merges mis-sized vs the grid (in-app workbooks
        // never hit this). Grow the imported rectangle to cover all merges (then cap).
        let maxR = range.e.r;
        let maxC = range.e.c;
        if (Array.isArray(ws?.["!merges"])) {
          for (const m of ws["!merges"] as { s?: { r: number; c: number }; e?: { r: number; c: number } }[]) {
            if (m?.s && m?.e) {
              maxR = Math.max(maxR, m.s.r, m.e.r);
              maxC = Math.max(maxC, m.s.c, m.e.c);
            }
          }
        }
        const rows = Math.min(MAX_IMPORT_ROWS, Math.max(1, maxR - range.s.r + 1));
        const cols = Math.min(MAX_IMPORT_COLS, Math.max(1, maxC - range.s.c + 1));
        const excelJsSheet = excelJsSheetMap.get(name);
        const checkboxCoordSet = new Set<string>();
        const toBooleanString = (value: unknown): "true" | "false" | null => {
          if (value === true) return "true";
          if (value === false) return "false";
          if (typeof value === "string") {
            const normalized = value.trim().toLowerCase();
            if (normalized === "true" || normalized === "false") return normalized;
          }
          return null;
        };
        const getExcelJsComparableValue = (value: unknown): unknown => {
          if (!value || typeof value !== "object") return value;
          if ("result" in (value as any)) return (value as any).result;
          return value;
        };

        const formulaByCoord = new Map<
          string,
          { formula: string; formulaCachedValue: string }
        >();
        const grid = Array.from({ length: rows }, (_, r) =>
          Array.from({ length: cols }, (_, c) => {
            const addr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
            const cell = ws?.[addr] as { w?: unknown; v?: unknown; f?: unknown } | undefined;
            const excelJsCell = excelJsSheet?.getCell(addr);
            const excelJsComparableValue = getExcelJsComparableValue(excelJsCell?.value);
            const isExcelBooleanCell =
              (excelJsBooleanValueType != null && excelJsCell?.type === excelJsBooleanValueType) ||
              excelJsComparableValue === true ||
              excelJsComparableValue === false;
            const booleanStringValue =
              toBooleanString(excelJsComparableValue) ??
              toBooleanString(cell?.v) ??
              toBooleanString(cell?.w);
            if (isExcelBooleanCell || booleanStringValue) {
              checkboxCoordSet.add(`${r}:${c}`);
              return booleanStringValue ?? String(excelJsComparableValue);
            }
            if (!cell) return "";
            if (typeof cell.f === "string" && cell.f.trim().length > 0) {
              const formula = `=${cell.f}`;
              const cached = cell.v == null ? "" : String(cell.v);
              formulaByCoord.set(`${r}:${c}`, {
                formula,
                formulaCachedValue: cached,
              });
              return cached;
            }
            if (cell.w != null) return String(cell.w);
            if (cell.v == null) return "";
            return String(cell.v);
          })
        );

        const merges = Array.isArray(ws?.["!merges"])
          ? ws["!merges"].map((m: any) => ({
              row: m.s.r - range.s.r,
              col: m.s.c - range.s.c,
              rowspan: m.e.r - m.s.r + 1,
              colspan: m.e.c - m.s.c + 1,
            }))
          : [];

        const cellMeta: Array<{
          row: number;
          col: number;
          className?: string;
          type?: string;
          formula?: string;
          formulaCachedValue?: string;
        }> = [];
        const cellMetaByCoord = new Map<
          string,
          {
            row: number;
            col: number;
            className?: string;
            type?: string;
            formula?: string;
            formulaCachedValue?: string;
          }
        >();
        if (rows * cols <= MAX_STYLE_SCAN_CELLS) {
          for (let r = 0; r < rows; r++) {
            for (let c = 0; c < cols; c++) {
              const addr = XLSX.utils.encode_cell({ r: r + range.s.r, c: c + range.s.c });
              const cell = ws?.[addr] as any;
              const style = cell?.s;
              if (!style) continue;
              const classes: string[] = [];

              if (style.font?.bold) classes.push("meta-bold");
              if (style.font?.italic) classes.push("meta-italic");
              if (style.font?.underline) classes.push("meta-underline");
              if (style.font?.strike) classes.push("meta-strike");

              const fontName = typeof style.font?.name === "string" ? style.font.name.trim() : "";
              if (fontName) classes.push(`meta-font-${fontName.replace(/\s+/g, "_")}`);
              // Excel stores font size in points; renderer expects pixels.
              const fontSizePt = Number(style.font?.sz);
              if (!Number.isNaN(fontSizePt) && fontSizePt > 0) {
                const fontSizePx = Math.round((fontSizePt * 96) / 72);
                classes.push(`meta-size-${Math.max(8, fontSizePx)}`);
              }

              const textColor = getXlsxColorHex(style.font?.color);
              if (textColor) classes.push(`meta-color-${textColor}`);
              const fillColor = getXlsxColorHex(style.fill?.fgColor) || getXlsxColorHex(style.fill?.bgColor);
              if (fillColor) classes.push(`meta-fill-${fillColor}`);

              const hAlign = String(style.alignment?.horizontal || "").toLowerCase();
              if (hAlign === "left" || hAlign === "center" || hAlign === "right" || hAlign === "justify") {
                classes.push(`meta-align-${hAlign}`);
              }
              const vAlign = String(style.alignment?.vertical || "").toLowerCase();
              if (vAlign === "top") classes.push("meta-valign-top");
              if (vAlign === "center") classes.push("meta-valign-middle");
              if (vAlign === "bottom") classes.push("meta-valign-bottom");
              if (style.alignment?.wrapText) classes.push("meta-wrap");

              if (classes.length > 0) {
                const entry = {
                  row: r,
                  col: c,
                  className: classes.join(" "),
                };
                cellMeta.push(entry);
                cellMetaByCoord.set(`${r}:${c}`, entry);
              }
            }
          }
        }
        for (const coord of checkboxCoordSet) {
          const existing = cellMetaByCoord.get(coord);
          const [row, col] = coord.split(":").map((v) => Number(v));
          if (existing) {
            existing.type = "checkbox";
            if (existing.className == null) existing.className = "";
          } else {
            const entry = { row, col, type: "checkbox", className: "" };
            cellMeta.push(entry);
            cellMetaByCoord.set(coord, entry);
          }
        }
        for (const [coord, formulaMeta] of formulaByCoord.entries()) {
          const existing = cellMetaByCoord.get(coord);
          if (existing) {
            existing.formula = formulaMeta.formula;
            existing.formulaCachedValue = formulaMeta.formulaCachedValue;
          } else {
            const [row, col] = coord.split(":").map(Number);
            const entry = {
              row,
              col,
              formula: formulaMeta.formula,
              formulaCachedValue: formulaMeta.formulaCachedValue,
            };
            cellMeta.push(entry);
            cellMetaByCoord.set(coord, entry);
          }
        }

        const colWidthsPx = Array.from({ length: cols }, (_, i) => {
          const col = (ws?.["!cols"] || [])[i + range.s.c] as any;
          if (col?.wpx) return Math.round(col.wpx);
          if (col?.wch) return Math.round(col.wch * 7 + 8);
          return 80;
        });

        const rowHeightsPx = Array.from({ length: rows }, (_, i) => {
          const row = (ws?.["!rows"] || [])[i + range.s.r] as any;
          if (row?.hpx) return Math.round(row.hpx);
          if (row?.hpt) return Math.round((row.hpt * 96) / 72);
          return 24;
        });

        const importedImages = (imageMapBySheet.get(name) || [])
          .map((img) => ({
            row: img.row - range.s.r,
            col: img.col - range.s.c,
            rowspan: img.rowspan,
            colspan: img.colspan,
            dataUrl: img.dataUrl,
          }))
          .filter((img) => img.row >= 0 && img.col >= 0 && img.row < rows && img.col < cols);

        return {
          name: name || `Sheet${index + 1}`,
          grid,
          mergeCells: merges,
          images: importedImages,
          colWidthsPx,
          rowHeightsPx,
          cellMeta,
        };
      });

      const { sheets, patchedCount } = applyImportedLogoToAllSheets
        ? ensureWorkbookLogoOnEverySheet(importedSheets)
        : { sheets: importedSheets, patchedCount: 0 };

      onUpdate({
        excelTemplate: {
          sheets: sheets.length ? sheets : [{ name: "Sheet1", grid: [[""]] }],
        },
        excelFileDataUrl: "",
        excelDisplayName: "",
        excelFileUrl: "",
      });

      toast({
        title: "Workbook imported",
        description:
          patchedCount > 0
            ? `${sheets.length || 1} sheet(s) loaded. Logo was applied to ${patchedCount} sheet(s) without images.`
            : `${sheets.length || 1} sheet(s) loaded into template`,
        variant: "success",
      });
    } catch (error: any) {
      toast({
        title: "Import failed",
        description: typeof error?.message === "string" ? error.message : "Could not read workbook.",
        variant: "destructive",
      });
    } finally {
      setIsImportingTemplate(false);
    }
  };

  const renderFieldProperties = (field: FormField) => (
    <div className="space-y-4">
      <div>
        <Label htmlFor="field-label">Label</Label>
        <Input
          id="field-label"
          value={field.label}
          onChange={(e) => onUpdate({ label: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="field-name">Field Name</Label>
        <Input
          id="field-name"
          value={field.name}
          onChange={(e) => onUpdate({ name: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="field-placeholder">Placeholder</Label>
        <Input
          id="field-placeholder"
          value={field.placeholder || ""}
          onChange={(e) => onUpdate({ placeholder: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="field-description">Description</Label>
        <Textarea
          id="field-description"
          value={field.description || ""}
          onChange={(e) => onUpdate({ description: e.target.value })}
          rows={2}
        />
      </div>

      <div className="flex items-center space-x-2">
        <Checkbox
          id="field-required"
          checked={field.required}
          onCheckedChange={(checked) => onUpdate({ required: checked })}
        />
        <Label htmlFor="field-required">Required Field</Label>
      </div>

      <div>
        <Label htmlFor="field-width">Width</Label>
        <Select
          value={field.layout?.width || "full"}
          onValueChange={(value) => onUpdate({
            layout: { ...field.layout, width: value }
          })}
        >
          <SelectTrigger>
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="full">Full Width</SelectItem>
            <SelectItem value="half">Half Width</SelectItem>
            <SelectItem value="third">Third Width</SelectItem>
            <SelectItem value="quarter">Quarter Width</SelectItem>
            <SelectItem value="auto">Auto Width</SelectItem>
          </SelectContent>
        </Select>
      </div>

      {field.type === "embedded_excel" && (
        <div className="space-y-3">
          <div>
            <Label>Spreadsheet Template</Label>
            <input
              ref={importTemplateInputRef}
              type="file"
              accept=".xlsx"
              className="hidden"
              onChange={(e) => {
                const file = e.target.files?.[0];
                if (file) {
                  void importWorkbookTemplate(file);
                }
                e.currentTarget.value = "";
              }}
            />
            <div className="flex gap-2 mt-2">
              <Button
                type="button"
                variant="outline"
                size="sm"
                disabled={isImportingTemplate}
                onClick={() => importTemplateInputRef.current?.click()}
              >
                {isImportingTemplate ? "Importing..." : "Import .xlsx Template"}
              </Button>
            </div>
            <div className="flex items-center space-x-2 mt-2">
              <Checkbox
                id="apply-logo-all-sheets"
                checked={applyImportedLogoToAllSheets}
                onCheckedChange={(checked) => setApplyImportedLogoToAllSheets(Boolean(checked))}
              />
              <Label htmlFor="apply-logo-all-sheets" className="text-sm">
                When importing, copy workbook logo to every sheet that has no image
              </Label>
            </div>
            <div className="grid grid-cols-2 gap-2 mt-2">
              <div>
                <Label htmlFor="sheet-rows" className="text-sm">Rows</Label>
                <Input
                  id="sheet-rows"
                  type="number"
                  min={1}
                  max={200}
                  value={sheetRows}
                  onChange={(e) => setSheetRows(Math.max(1, Math.min(200, Number(e.target.value) || 1)))}
                />
              </div>
              <div>
                <Label htmlFor="sheet-cols" className="text-sm">Columns</Label>
                <Input
                  id="sheet-cols"
                  type="number"
                  min={1}
                  max={50}
                  value={sheetCols}
                  onChange={(e) => setSheetCols(Math.max(1, Math.min(50, Number(e.target.value) || 1)))}
                />
              </div>
              <div>
                <Label htmlFor="sheet-count" className="text-sm">Sheets</Label>
                <Input
                  id="sheet-count"
                  type="number"
                  min={1}
                  max={20}
                  value={sheetCount}
                  onChange={(e) => setSheetCount(Math.max(1, Math.min(20, Number(e.target.value) || 1)))}
                />
              </div>
            </div>
            <div className="flex gap-2 mt-2">
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => {
                  const grid = Array.from({ length: sheetRows }, () =>
                    Array.from({ length: sheetCols }, () => "")
                  );
                  const sheets = Array.from({ length: sheetCount }, (_, index) => ({
                    name: `Sheet${index + 1}`,
                    grid: grid.map((row) => [...row]),
                  }));
                  onUpdate({
                    excelTemplate: {
                      sheets,
                    },
                    excelFileDataUrl: "",
                    excelDisplayName: "",
                    excelFileUrl: "",
                  });
                  toast({
                    title: "Workbook created",
                    description: `${sheetCount} sheet(s), ${sheetRows} rows x ${sheetCols} columns`,
                    variant: "success",
                  });
                }}
              >
                Create Workbook
              </Button>
              <Button
                type="button"
                variant="outline"
                size="sm"
                onClick={() => {
                  onUpdate({
                    excelTemplate: undefined,
                    excelFileDataUrl: "",
                    excelDisplayName: "",
                    excelFileUrl: "",
                  });
                }}
              >
                Clear
              </Button>
            </div>
            <p className="text-xs text-muted-foreground mt-1">
              Spreadsheet is created and edited directly inside the app (Handsontable).
            </p>
            {field.excelTemplate?.sheets?.[0]?.grid?.length ? (
              <div className="flex flex-wrap items-center gap-2 mt-2">
                <span className="text-xs rounded-md border bg-muted px-2 py-1 font-mono truncate max-w-full">
                  Using: {field.excelTemplate.sheets[0].grid.length} rows x {field.excelTemplate.sheets[0].grid[0]?.length || 0} columns
                </span>
                <Button
                  type="button"
                  size="sm"
                  onClick={() => {
                    const cloned = JSON.parse(JSON.stringify(field.excelTemplate));
                    setWorkbookDraft(cloned);
                    setIsWorkbookEditorOpen(true);
                  }}
                >
                  Edit Workbook
                </Button>
              </div>
            ) : null}
          </div>
        </div>
      )}

      {(field.type === "select" || field.type === "radio" || field.type === "checkbox") && (
        <div>
          <Label>Options</Label>
          <div className="space-y-2">
            {field.options?.map((option, index) => (
              <div key={index} className="flex items-center space-x-2">
                <Input
                  value={option.label}
                  onChange={(e) => {
                    const updatedOptions = [...(field.options || [])];
                    updatedOptions[index] = { ...option, label: e.target.value };
                    onUpdate({ options: updatedOptions });
                  }}
                  placeholder="Option label"
                  className="flex-1"
                />
                <Input
                  value={option.value}
                  onChange={(e) => {
                    const updatedOptions = [...(field.options || [])];
                    updatedOptions[index] = { ...option, value: e.target.value };
                    onUpdate({ options: updatedOptions });
                  }}
                  placeholder="Option value"
                  className="flex-1"
                />
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => {
                    const updatedOptions = field.options?.filter((_, i) => i !== index);
                    onUpdate({ options: updatedOptions });
                  }}
                >
                  <Trash2 className="w-4 h-4" />
                </Button>
              </div>
            ))}
            <Button
              variant="outline"
              size="sm"
              onClick={() => {
                const newOption = { label: "New Option", value: `option${(field.options?.length || 0) + 1}` };
                onUpdate({ options: [...(field.options || []), newOption] });
              }}
              className="w-full"
            >
              <Plus className="w-4 h-4 mr-2" />
              Add Option
            </Button>
          </div>
        </div>
      )}

      {(field.type === "number" || field.type === "text") && (
        <div className="space-y-3">
          <Label>Validation</Label>
          <div className="grid grid-cols-2 gap-2">
            <div>
              <Label htmlFor="min-value" className="text-sm">Min Value</Label>
              <Input
                id="min-value"
                type="number"
                value={field.validation?.min || ""}
                onChange={(e) => onUpdate({
                  validation: { ...field.validation, min: e.target.value ? Number(e.target.value) : undefined }
                })}
              />
            </div>
            <div>
              <Label htmlFor="max-value" className="text-sm">Max Value</Label>
              <Input
                id="max-value"
                type="number"
                value={field.validation?.max || ""}
                onChange={(e) => onUpdate({
                  validation: { ...field.validation, max: e.target.value ? Number(e.target.value) : undefined }
                })}
              />
            </div>
          </div>
        </div>
      )}
      {field.type === "embedded_excel" && (
        <Dialog open={isWorkbookEditorOpen} onOpenChange={(open) => setIsWorkbookEditorOpen(open)}>
          <DialogContent className="max-w-[min(96vw,1200px)] max-h-[90vh] overflow-visible">
            <DialogHeader>
              <DialogTitle>Edit Workbook</DialogTitle>
              <DialogDescription>
                Save Changes copies the live grid from the editor (cells, column widths, merges)
                into the form field. The toolbar Save Workbook button is optional.
              </DialogDescription>
            </DialogHeader>
            <div className="max-h-[65vh] overflow-auto">
              {workbookDraft?.sheets?.length ? (
                <React.Suspense
                  fallback={<div className="p-3 text-sm text-muted-foreground">Loading workbook editor...</div>}
                >
                  <HandsontableWorkbook
                    ref={workbookEditorRef as any}
                    data={workbookDraft}
                    onChange={(next) => setWorkbookDraft(next)}
                    readOnly={false}
                  />
                </React.Suspense>
              ) : (
                <div className="p-3 text-sm text-muted-foreground border rounded-md bg-muted/30">
                  No workbook data available. Create or import a workbook first.
                </div>
              )}
            </div>
            <div className="flex items-center justify-end gap-2">
              <Button
                type="button"
                variant="outline"
                onClick={() => {
                  setIsWorkbookEditorOpen(false);
                  setWorkbookDraft(null);
                }}
              >
                Cancel
              </Button>
              <Button
                type="button"
                disabled={!workbookDraft?.sheets?.length}
                onClick={() => {
                  if (!workbookDraft?.sheets?.length) return;
                  const snap = workbookEditorRef.current?.getWorkbookSnapshot();
                  const template = snap?.sheets?.length
                    ? { sheets: snap.sheets }
                    : workbookDraft;
                  if (!template?.sheets?.length) return;
                  onUpdate({ excelTemplate: template });
                  setIsWorkbookEditorOpen(false);
                  setWorkbookDraft(null);
                  toast({
                    title: "Workbook saved",
                    description: "Template changes applied to this field.",
                    variant: "success",
                  });
                }}
              >
                Save Changes
              </Button>
            </div>
          </DialogContent>
        </Dialog>
      )}
    </div>
  );

  const renderSectionProperties = (section: FormSection) => (
    <div className="space-y-4">
      <div>
        <Label htmlFor="section-title">Section Title</Label>
        <Input
          id="section-title"
          value={section.title}
          onChange={(e) => onUpdate({ title: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="section-name">Section Name</Label>
        <Input
          id="section-name"
          value={section.name}
          onChange={(e) => onUpdate({ name: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="section-description">Description</Label>
        <Textarea
          id="section-description"
          value={section.description || ""}
          onChange={(e) => onUpdate({ description: e.target.value })}
          rows={2}
        />
      </div>

      <div>
        <Label htmlFor="columns-per-row">Columns Per Row</Label>
        <Select
          value={(section.layout?.columnsPerRow || 1).toString()}
          onValueChange={(value) => onUpdate({
            layout: { ...section.layout, columnsPerRow: Number(value) }
          })}
        >
          <SelectTrigger>
            <SelectValue />
          </SelectTrigger>
          <SelectContent>
            <SelectItem value="1">1 Column</SelectItem>
            <SelectItem value="2">2 Columns</SelectItem>
            <SelectItem value="3">3 Columns</SelectItem>
            <SelectItem value="4">4 Columns</SelectItem>
          </SelectContent>
        </Select>
      </div>
    </div>
  );

  const renderTableProperties = (table: TableConfig) => (
    <div className="space-y-4">
      <div>
        <Label htmlFor="table-title">Table Title</Label>
        <Input
          id="table-title"
          value={table.title}
          onChange={(e) => onUpdate({ title: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="table-name">Table Name</Label>
        <Input
          id="table-name"
          value={table.name}
          onChange={(e) => onUpdate({ name: e.target.value })}
        />
      </div>

      <div>
        <Label htmlFor="table-description">Description</Label>
        <Textarea
          id="table-description"
          value={table.description || ""}
          onChange={(e) => onUpdate({ description: e.target.value })}
          rows={2}
        />
      </div>

      <div className="grid grid-cols-2 gap-2">
        <div>
          <Label htmlFor="min-rows">Min Rows</Label>
          <Input
            id="min-rows"
            type="number"
            value={table.minRows}
            onChange={(e) => onUpdate({ minRows: Number(e.target.value) })}
            min={1}
          />
        </div>
        <div>
          <Label htmlFor="max-rows">Max Rows</Label>
          <Input
            id="max-rows"
            type="number"
            value={table.maxRows}
            onChange={(e) => onUpdate({ maxRows: Number(e.target.value) })}
            min={1}
          />
        </div>
      </div>

      <div>
        <Label htmlFor="default-rows">Default Rows</Label>
        <Input
          id="default-rows"
          type="number"
          value={table.defaultRows}
          onChange={(e) => onUpdate({ defaultRows: Number(e.target.value) })}
          min={1}
        />
      </div>

      <div className="space-y-2">
        <div className="flex items-center space-x-2">
          <Checkbox
            id="allow-add-rows"
            checked={table.allowAddRows}
            onCheckedChange={(checked) => onUpdate({ allowAddRows: checked })}
          />
          <Label htmlFor="allow-add-rows">Allow Add Rows</Label>
        </div>
        <div className="flex items-center space-x-2">
          <Checkbox
            id="allow-delete-rows"
            checked={table.allowDeleteRows}
            onCheckedChange={(checked) => onUpdate({ allowDeleteRows: checked })}
          />
          <Label htmlFor="allow-delete-rows">Allow Delete Rows</Label>
        </div>
      </div>

      {/* Pre-filled Data Summary */}
      {table.preFilledData && table.preFilledData.length > 0 && (
        <div className="space-y-2">
          <Label>Pre-filled Cells ({table.preFilledData.length})</Label>
          <div className="max-h-32 overflow-y-auto space-y-1">
            {table.preFilledData.map((cell, index) => (
              <div key={index} className="text-xs p-2 bg-blue-50 rounded border">
                Row {cell.rowIndex + 1}, Column: {cell.columnName} = "{String(cell.value)}"
                <Button
                  variant="ghost"
                  size="sm"
                  onClick={() => {
                    const updatedData = table.preFilledData?.filter((_, i) => i !== index) || [];
                    onUpdate({ preFilledData: updatedData });
                  }}
                  className="ml-2 h-4 w-4 p-0"
                >
                  <X className="w-3 h-3" />
                </Button>
              </div>
            ))}
          </div>
          <Button
            variant="outline"
            size="sm"
            onClick={() => onUpdate({ preFilledData: [] })}
            className="w-full"
          >
            Clear All Pre-filled Data
          </Button>
        </div>
      )}

      <div>
        <Label>Columns</Label>
        <div className="space-y-3">
          {table.columns.map((column, index) => (
            <Card key={column.id} className="p-3">
              <div className="space-y-2">
                <div className="flex items-center justify-between">
                  <Label className="font-medium">Column {index + 1}</Label>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => {
                      const updatedColumns = table.columns.filter((_, i) => i !== index);
                      onUpdate({ columns: updatedColumns });
                    }}
                  >
                    <Trash2 className="w-4 h-4" />
                  </Button>
                </div>
                <Input
                  value={column.label}
                  onChange={(e) => {
                    const updatedColumns = [...table.columns];
                    updatedColumns[index] = { ...column, label: e.target.value };
                    onUpdate({ columns: updatedColumns });
                  }}
                  placeholder="Column label"
                />
                <Select
                  value={column.type}
                  onValueChange={(value) => {
                    const updatedColumns = [...table.columns];
                    updatedColumns[index] = { ...column, type: value as TableColumn["type"] };
                    onUpdate({ columns: updatedColumns });
                  }}
                >
                  <SelectTrigger>
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="text">Text</SelectItem>
                    <SelectItem value="number">Number</SelectItem>
                    <SelectItem value="date">Date</SelectItem>
                    <SelectItem value="email">Email</SelectItem>
                    <SelectItem value="phone">Phone</SelectItem>
                    <SelectItem value="select">Select</SelectItem>
                    <SelectItem value="checkbox">Checkbox</SelectItem>
                  </SelectContent>
                </Select>
                <div className="flex items-center space-x-2">
                  <Checkbox
                    checked={column.required}
                    onCheckedChange={(checked) => {
                      const updatedColumns = [...table.columns];
                      updatedColumns[index] = { ...column, required: Boolean(checked) };
                      onUpdate({ columns: updatedColumns });
                    }}
                  />
                  <Label className="text-sm">Required</Label>
                </div>
              </div>
            </Card>
          ))}
          <Button
            variant="outline"
            onClick={() => {
              const newColumn: TableColumn = {
                id: `col-${Date.now()}`,
                name: `column${table.columns.length + 1}`,
                label: `Column ${table.columns.length + 1}`,
                type: "text",
                required: false,
                order: table.columns.length,
              };
              onUpdate({ columns: [...table.columns, newColumn] });
            }}
            className="w-full"
          >
            <Plus className="w-4 h-4 mr-2" />
            Add Column
          </Button>
        </div>
      </div>
    </div>
  );

  const { type, data } = selectedItem;

  return (
    <div className={`${className}`}>
      <Card>
        <CardHeader>
          <CardTitle className="text-lg">
            {type === "field" && "Field Properties"}
            {type === "section" && "Section Properties"}
            {type === "table" && "Table Properties"}
          </CardTitle>
        </CardHeader>
        <CardContent>
          {type === "field" && renderFieldProperties(data as FormField)}
          {type === "section" && renderSectionProperties(data as FormSection)}
          {type === "table" && renderTableProperties(data as TableConfig)}
        </CardContent>
      </Card>
    </div>
  );
};

export default PropertiesPanel;