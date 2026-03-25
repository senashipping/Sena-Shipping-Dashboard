import React from "react";
import { useDroppable } from "@dnd-kit/core";
import { SortableContext, verticalListSortingStrategy } from "@dnd-kit/sortable";
import { Card, CardContent, CardHeader } from "../ui/card";
import { Button } from "../ui/button";
import { Plus } from "lucide-react";
import { FormField, FormSection, TableConfig } from "../../types";
import FieldRenderer from "./FieldRenderer";
import SectionRenderer from "./SectionRenderer";

interface FormCanvasProps {
  formType: "regular" | "table" | "mixed";
  fields?: FormField[];
  sections?: FormSection[];
  tableConfig?: TableConfig;
  onUpdateField: (fieldId: string, updates: Partial<FormField>) => void;
  onDeleteField: (fieldId: string) => void;
  onUpdateSection: (sectionId: string, updates: Partial<FormSection>) => void;
  onDeleteSection: (sectionId: string) => void;
  onUpdateTable: (updates: Partial<TableConfig>) => void;
  onSelectItem: (itemId: string, itemType: "field" | "section" | "table") => void;
  selectedItemId?: string;
  onMoveField?: (fieldId: string, direction: "up" | "down", context?: { sectionId?: string }) => void;
}

const FormCanvas: React.FC<FormCanvasProps> = ({
  formType,
  fields = [],
  sections = [],
  tableConfig,
  onUpdateField,
  onDeleteField,
  onUpdateSection,
  onDeleteSection,
  onUpdateTable,
  onSelectItem,
  selectedItemId,
  onMoveField,
}) => {
  const { setNodeRef, isOver } = useDroppable({
    id: "form-canvas",
    data: { 
      accepts: ["field", "table", "section"],
      type: "canvas"
    },
  });

  // Helper function to get pre-filled value for a cell
  const getPreFilledValue = (tableConfig: TableConfig, rowIndex: number, columnName: string) => {
    const preFilledCell = tableConfig.preFilledData?.find(
      cell => cell.rowIndex === rowIndex && cell.columnName === columnName
    );
    return preFilledCell?.value || '';
  };

  // Helper function to check if cell is read-only
  const isCellReadOnly = (tableConfig: TableConfig, rowIndex: number, columnName: string) => {
    const preFilledCell = tableConfig.preFilledData?.find(
      cell => cell.rowIndex === rowIndex && cell.columnName === columnName
    );
    return preFilledCell?.isReadOnly || false;
  };

  // Handle cell value changes in form builder
  const handleCellChange = (tableConfig: TableConfig, rowIndex: number, columnName: string, value: any, isReadOnly: boolean = true) => {
    const updatedPreFilledData = [...(tableConfig.preFilledData || [])];
    
    // Remove existing entry if value is empty
    if (!value || value === '') {
      const filteredData = updatedPreFilledData.filter(
        cell => !(cell.rowIndex === rowIndex && cell.columnName === columnName)
      );
      onUpdateTable({ preFilledData: filteredData });
      return;
    }

    // Find existing entry
    const existingIndex = updatedPreFilledData.findIndex(
      cell => cell.rowIndex === rowIndex && cell.columnName === columnName
    );

    const newCell: { rowIndex: number; columnName: string; value: any; isReadOnly: boolean } = {
      rowIndex,
      columnName,
      value,
      isReadOnly
    };

    if (existingIndex >= 0) {
      updatedPreFilledData[existingIndex] = newCell;
    } else {
      updatedPreFilledData.push(newCell);
    }

    onUpdateTable({ preFilledData: updatedPreFilledData });
  };

  const renderRegularForm = () => (
    <SortableContext items={fields.map(f => f.id)} strategy={verticalListSortingStrategy}>
      <div className="space-y-4">
        {fields.map((field) => (
          <FieldRenderer
            key={field.id}
            field={field}
            isSelected={selectedItemId === field.id}
            onUpdate={(updates: Partial<FormField>) => onUpdateField(field.id, updates)}
            onDelete={() => onDeleteField(field.id)}
            onSelect={() => onSelectItem(field.id, "field")}
            onMoveUp={() => onMoveField?.(field.id, "up")}
            onMoveDown={() => onMoveField?.(field.id, "down")}
          />
        ))}
      </div>
    </SortableContext>
  );

  const renderTableForm = () => (
    <div>
      {tableConfig && (
        <Card 
          className={`cursor-pointer transition-all ${
            selectedItemId === tableConfig.id ? "ring-2 ring-blue-500" : ""
          }`}
          onClick={() => onSelectItem(tableConfig.id, "table")}
        >
          <CardHeader className="pb-3">
                        <div className="flex items-center justify-between">
              <h3 
                className="text-lg font-semibold cursor-pointer hover:text-blue-600" 
                onClick={() => onSelectItem("table", "table")}
              >
                {tableConfig.title}
              </h3>
              <Button
                variant="ghost"
                size="sm"
                onClick={() => onUpdateTable({ title: "Updated Table" })}
                className="text-blue-500 hover:text-blue-700"
              >
                Edit
              </Button>
            </div>
            {tableConfig.description && (
              <p className="text-sm text-gray-600">{tableConfig.description}</p>
            )}
          </CardHeader>
          <CardContent>
            <div className="border rounded-lg overflow-hidden">
              <table className="w-full text-sm">
                <thead className="bg-gray-50">
                  <tr>
                    {tableConfig.columns.map((column) => (
                      <th
                        key={column.id}
                        className="px-3 py-2 text-left font-medium text-gray-700 border-r last:border-r-0"
                        style={{ width: column.width }}
                      >
                        {column.label}
                        {column.required && <span className="text-red-500 ml-1">*</span>}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: Math.min(5, tableConfig.defaultRows) }).map((_, rowIndex) => (
                    <tr key={rowIndex} className="border-t">
                      {tableConfig.columns.map((column) => {
                        const cellValue = getPreFilledValue(tableConfig, rowIndex, column.name);
                        const isReadOnly = isCellReadOnly(tableConfig, rowIndex, column.name);
                        
                        return (
                          <td
                            key={column.id}
                            className="px-3 py-2 border-r last:border-r-0"
                          >
                            {column.type === "select" ? (
                              <select 
                                value={cellValue}
                                onChange={(e) => handleCellChange(tableConfig, rowIndex, column.name, e.target.value)}
                                className={`w-full text-sm bg-transparent border rounded px-2 py-1 ${
                                  isReadOnly && cellValue ? 'bg-blue-50 border-blue-200' : 'bg-white'
                                }`}
                              >
                                <option value="">Select...</option>
                                {column.options?.map(option => (
                                  <option key={option} value={option}>{option}</option>
                                ))}
                              </select>
                            ) : column.type === "checkbox" ? (
                              <input 
                                type="checkbox" 
                                checked={!!cellValue}
                                onChange={(e) => handleCellChange(tableConfig, rowIndex, column.name, e.target.checked)}
                                className={`rounded ${isReadOnly && cellValue ? 'bg-blue-50' : ''}`}
                              />
                            ) : (
                              <input
                                type={column.type}
                                value={cellValue}
                                onChange={(e) => handleCellChange(tableConfig, rowIndex, column.name, e.target.value)}
                                placeholder={cellValue ? '' : `Sample ${column.type}`}
                                className={`w-full text-sm bg-transparent border rounded px-2 py-1 ${
                                  isReadOnly && cellValue ? 'bg-blue-50 border-blue-200' : 'bg-white'
                                }`}
                              />
                            )}
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
              <div className="p-3 bg-gray-50 text-xs text-gray-600">
                💡 Pre-fill cells by typing values above. Blue cells will be read-only for users.
              </div>
            </div>
          </CardContent>
        </Card>
      )}
    </div>
  );

  const renderMixedForm = () => (
    <SortableContext items={sections.map(s => s.id)} strategy={verticalListSortingStrategy}>
      <div className="space-y-6">
        {sections.map((section) => (
          <SectionRenderer
            key={section.id}
            section={section}
            isSelected={selectedItemId === section.id}
            onUpdate={(updates: Partial<FormSection>) => onUpdateSection(section.id, updates)}
            onDelete={() => onDeleteSection(section.id)}
            onSelect={() => onSelectItem(section.id, "section")}
            onFieldSelect={(fieldId: string) => onSelectItem(fieldId, "field")}
            onFieldUpdate={(fieldId: string, updates: Partial<FormField>) => {
              // Find the field within sections and update it
              const updatedSections = sections.map(s => {
                if (s.fields?.some(f => f.id === fieldId)) {
                  return {
                    ...s,
                    fields: s.fields?.map(f => f.id === fieldId ? { ...f, ...updates } : f)
                  };
                }
                return s;
              });
              // Call onUpdateSection for the section containing this field
              const containingSection = sections.find(s => s.fields?.some(f => f.id === fieldId));
              if (containingSection) {
                const updatedSection = updatedSections.find(s => s.id === containingSection.id);
                if (updatedSection) {
                  onUpdateSection(containingSection.id, updatedSection);
                }
              }
            }}
            onFieldDelete={(fieldId: string) => {
              // Remove field from the section it belongs to
              const containingSection = sections.find(s => s.fields?.some(f => f.id === fieldId));
              if (containingSection) {
                const updatedFields = containingSection.fields?.filter(f => f.id !== fieldId);
                onUpdateSection(containingSection.id, { fields: updatedFields });
              }
            }}
            onTableSelect={(tableId: string) => onSelectItem(tableId, "table")}
            onFieldMove={(fieldId: string, direction: "up" | "down") => onMoveField?.(fieldId, direction, { sectionId: section.id })}
          />
        ))}
      </div>
    </SortableContext>
  );

  const isEmpty = formType === "regular" 
    ? fields.length === 0
    : formType === "table" 
    ? !tableConfig
    : sections.length === 0;

  return (
    <div
      ref={setNodeRef}
      className={`min-h-96 p-6 border-2 border-dashed rounded-lg transition-all ${
        isOver ? "border-blue-500 bg-blue-50 border-solid" : "border-gray-300"
      } ${isEmpty ? "" : "min-h-fit"}`}
    >
      {isEmpty ? (
        <div className="flex flex-col items-center justify-center h-64 text-gray-500">
          <Plus className="w-12 h-12 mb-4 text-gray-400" />
          <h3 className="text-lg font-medium mb-2">Start Building Your Form</h3>
          <p className="text-center">
            Drag and drop elements from the left panel to create your form
          </p>
        </div>
      ) : (
        <div>
          {formType === "regular" && renderRegularForm()}
          {formType === "table" && renderTableForm()}
          {formType === "mixed" && renderMixedForm()}
        </div>
      )}
    </div>
  );
};

export default FormCanvas;