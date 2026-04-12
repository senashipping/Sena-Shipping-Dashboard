import { useMemo, useState, type FC } from "react";
import { Button } from "../ui/button";
import { Label } from "../ui/label";
import { Checkbox } from "../ui/checkbox";
import SharedFormRenderer from "./SharedFormRenderer";
import { FormField, FormSection, TableConfig } from "../../types";

interface PreviewFormProps {
  formState: {
    formType: "regular" | "table" | "mixed";
    title: string;
    description?: string;
    fields: FormField[];
    sections: FormSection[];
    tableConfig?: TableConfig;
  };
  onSubmit?: () => void;
  isSubmitting?: boolean;
}

function formHasEmbeddedExcel(formState: PreviewFormProps["formState"]): boolean {
  const inFields = (fields?: FormField[]) =>
    Boolean(fields?.some((f) => f.type === "embedded_excel"));
  if (inFields(formState.fields)) return true;
  if (formState.formType === "mixed") {
    return formState.sections.some(
      (s) => s.type === "fields" && inFields(s.fields),
    );
  }
  return false;
}

const PreviewForm: FC<PreviewFormProps> = ({ formState, onSubmit, isSubmitting = false }) => {
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [tableData, setTableData] = useState<any[]>([]);
  /** Full Handsontable in-dialog is heavy; default to a simple grid preview (opt in for fillable testing). */
  const [interactiveExcelPreview, setInteractiveExcelPreview] = useState(false);
  const hasEmbeddedExcel = useMemo(
    () => formHasEmbeddedExcel(formState),
    [formState],
  );

  const handleFieldChange = (fieldName: string, value: any) => {
    setFormData((prev) => ({
      ...prev,
      [fieldName]: value,
    }));
  };

  const handleTableChange = (rowIndex: number, columnName: string, value: any) => {
    setTableData((prev) => {
      const newData = [...prev];
      if (!newData[rowIndex]) newData[rowIndex] = {};
      newData[rowIndex][columnName] = value;
      return newData;
    });
  };

  const handleMixedTableChange = (sectionId: string, rowIndex: number, columnName: string, value: any) => {
    const tableKey = `table_${sectionId}`;
    setFormData((prev) => {
      const tableDataInner = prev[tableKey] || [];
      const newTableData = [...tableDataInner];
      if (!newTableData[rowIndex]) {
        newTableData[rowIndex] = {};
      }
      newTableData[rowIndex][columnName] = value;
      return {
        ...prev,
        [tableKey]: newTableData,
      };
    });
  };

  const handleAddTableRow = (tableId?: string) => {
    if (tableId) {
      const tableKey = `table_${tableId}`;
      setFormData((prev) => ({
        ...prev,
        [tableKey]: [...(prev[tableKey] || [{}]), {}],
      }));
    } else {
      setTableData((prev) => [...prev, {}]);
    }
  };

  const handleRemoveTableRow = (rowIndex: number, tableId?: string) => {
    if (tableId) {
      const tableKey = `table_${tableId}`;
      setFormData((prev) => ({
        ...prev,
        [tableKey]: (prev[tableKey] || []).filter((_: any, i: number) => i !== rowIndex),
      }));
    } else {
      setTableData((prev) => prev.filter((_, i) => i !== rowIndex));
    }
  };

  const submitButton = (
    <Button type="button" className="w-full" onClick={onSubmit} disabled={!onSubmit || isSubmitting}>
      {isSubmitting ? "Submitting..." : "Submit Form"}
    </Button>
  );

  return (
    <div className="space-y-4">
      {hasEmbeddedExcel && (
        <div className="flex items-start gap-3 rounded-md border border-amber-200 bg-amber-50/80 p-3 text-sm text-amber-950 dark:border-amber-900 dark:bg-amber-950/40 dark:text-amber-100">
          <Checkbox
            id="preview-interactive-excel"
            checked={interactiveExcelPreview}
            onCheckedChange={(v) => setInteractiveExcelPreview(v === true)}
          />
          <div className="space-y-1 leading-snug">
            <Label htmlFor="preview-interactive-excel" className="font-medium cursor-pointer">
              Interactive spreadsheet preview
            </Label>
            <p className="text-xs opacity-90">
              Leave off for a fast layout preview. Turn on only if you need to try fillable cells here; large
              templates can stress the browser.
            </p>
          </div>
        </div>
      )}
      <SharedFormRenderer
        formState={formState}
        formData={formData}
        tableData={tableData}
        onFieldChange={handleFieldChange}
        onTableChange={handleTableChange}
        onMixedTableChange={handleMixedTableChange}
        onAddTableRow={handleAddTableRow}
        onRemoveTableRow={handleRemoveTableRow}
        useLocalExcelState
        excelReadOnly
        lightweightExcelPreview={!interactiveExcelPreview}
        submitButton={submitButton}
      />
    </div>
  );
};

export default PreviewForm;
