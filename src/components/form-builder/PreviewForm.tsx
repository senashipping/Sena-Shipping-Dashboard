import { useState, useImperativeHandle, forwardRef } from "react";
import { Button } from "../ui/button";
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
}

export type PreviewFormHandle = {
  getValues: () => { formData: Record<string, any>; tableData: any[] };
};

const PreviewForm = forwardRef<PreviewFormHandle, PreviewFormProps>(({ formState }, ref) => {
  const [formData, setFormData] = useState<Record<string, any>>({});
  const [tableData, setTableData] = useState<any[]>([]);

  useImperativeHandle(
    ref,
    () => ({
      getValues: () => ({ formData, tableData }),
    }),
    [formData, tableData]
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
    <Button type="button" className="w-full">
      Submit Form
    </Button>
  );

  return (
    <SharedFormRenderer
      formState={formState}
      formData={formData}
      tableData={tableData}
      onFieldChange={handleFieldChange}
      onTableChange={handleTableChange}
      onMixedTableChange={handleMixedTableChange}
      onAddTableRow={handleAddTableRow}
      onRemoveTableRow={handleRemoveTableRow}
      submitButton={submitButton}
    />
  );
});

PreviewForm.displayName = "PreviewForm";

export default PreviewForm;