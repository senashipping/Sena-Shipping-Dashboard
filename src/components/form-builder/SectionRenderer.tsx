import React from "react";
import { useSortable } from "@dnd-kit/sortable";
import { useDroppable } from "@dnd-kit/core";
import { CSS } from "@dnd-kit/utilities";
import { Card, CardContent, CardHeader, CardTitle } from "../ui/card";
import { Button } from "../ui/button";
import { GripVertical, Trash2, Settings } from "lucide-react";
import { FormSection, FormField } from "../../types";
import FieldRenderer from "./FieldRenderer";

interface SectionRendererProps {
  section: FormSection;
  isSelected: boolean;
  onUpdate: (updates: Partial<FormSection>) => void;
  onDelete: () => void;
  onSelect: () => void;
  onFieldSelect?: (fieldId: string) => void;
  onFieldUpdate?: (fieldId: string, updates: Partial<FormField>) => void;
  onFieldDelete?: (fieldId: string) => void;
  onTableSelect?: (tableId: string) => void;
  onFieldMove?: (fieldId: string, direction: "up" | "down") => void;
}

const SectionRenderer: React.FC<SectionRendererProps> = ({
  section,
  isSelected,
  onUpdate,
  onDelete,
  onSelect,
  onFieldSelect,
  onFieldUpdate,
  onFieldDelete,
  onTableSelect,
  onFieldMove,
}) => {
  const {
    attributes,
    listeners,
    setNodeRef: setSortableNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: section.id });

  const { setNodeRef: setDroppableNodeRef, isOver: isDropOver } = useDroppable({
    id: section.id,
    data: {
      accepts: ["field", "table"],
      type: "section"
    }
  });

  // Combine refs
  const setNodeRef = (element: HTMLElement | null) => {
    setSortableNodeRef(element);
    setDroppableNodeRef(element);
  };

  // Handle cell changes for table sections
  const handleTableCellChange = (rowIndex: number, columnName: string, value: any, isReadOnly: boolean = true) => {
    if (section.type === "table" && section.tableConfig) {
      const updatedPreFilledData = [...(section.tableConfig.preFilledData || [])];
      
      if (!value || value === '') {
        const filteredData = updatedPreFilledData.filter(
          cell => !(cell.rowIndex === rowIndex && cell.columnName === columnName)
        );
        onUpdate({ 
          tableConfig: { 
            ...section.tableConfig, 
            preFilledData: filteredData 
          } 
        });
        return;
      }

      const existingIndex = updatedPreFilledData.findIndex(
        cell => cell.rowIndex === rowIndex && cell.columnName === columnName
      );

      const newCell = { rowIndex, columnName, value, isReadOnly };

      if (existingIndex >= 0) {
        updatedPreFilledData[existingIndex] = newCell;
      } else {
        updatedPreFilledData.push(newCell);
      }

      onUpdate({ 
        tableConfig: { 
          ...section.tableConfig, 
          preFilledData: updatedPreFilledData 
        } 
      });
    }
  };

  const getPreFilledValue = (rowIndex: number, columnName: string) => {
    if (section.type === "table" && section.tableConfig) {
      const preFilledCell = section.tableConfig.preFilledData?.find(
        cell => cell.rowIndex === rowIndex && cell.columnName === columnName
      );
      return preFilledCell?.value || '';
    }
    return '';
  };

  const isCellReadOnly = (rowIndex: number, columnName: string) => {
    if (section.type === "table" && section.tableConfig) {
      const preFilledCell = section.tableConfig.preFilledData?.find(
        cell => cell.rowIndex === rowIndex && cell.columnName === columnName
      );
      return preFilledCell?.isReadOnly || false;
    }
    return false;
  };

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  const renderSectionContent = () => {
    if (section.type === "fields" && section.fields) {
      return (
        <div className="space-y-4">
          {section.fields.map((field) => (
            <FieldRenderer
              key={field.id}
              field={field}
              isSelected={false}
              onUpdate={(updates) => {
                if (onFieldUpdate) {
                  onFieldUpdate(field.id, updates);
                } else {
                  // Fallback to section update
                  const updatedFields = section.fields?.map(f => 
                    f.id === field.id ? { ...f, ...updates } : f
                  );
                  onUpdate({ fields: updatedFields });
                }
              }}
              onDelete={() => {
                if (onFieldDelete) {
                  onFieldDelete(field.id);
                } else {
                  // Fallback to section update
                  const updatedFields = section.fields?.filter(f => f.id !== field.id);
                  onUpdate({ fields: updatedFields });
                }
              }}
              onSelect={() => {
                if (onFieldSelect) {
                  onFieldSelect(field.id);
                }
              }}
              onMoveUp={() => onFieldMove?.(field.id, "up")}
              onMoveDown={() => onFieldMove?.(field.id, "down")}
            />
          ))}
        </div>
      );
    }

    if (section.type === "table" && section.tableConfig) {
      const tableConfig = section.tableConfig;
      return (
        <div 
          className="overflow-hidden transition-colors border rounded-lg cursor-pointer hover:border-blue-300" 
          onClick={(e) => {
            e.stopPropagation(); // Prevent section selection
            if (onTableSelect) {
              onTableSelect(tableConfig.id);
            }
          }}
        >
          <div className="flex items-center justify-between p-3 border-b bg-gray-50">
            <h4 className="font-medium text-gray-700">{tableConfig.title}</h4>
            <Button
              variant="ghost"
              size="sm"
              onClick={(e) => {
                e.stopPropagation();
                if (onTableSelect) {
                  onTableSelect(tableConfig.id);
                }
              }}
              className="text-blue-500 hover:text-blue-700"
            >
              <Settings className="w-4 h-4" />
            </Button>
          </div>
          <table className="w-full text-sm">
            <thead className="bg-gray-50">
              <tr>
                {tableConfig.columns.map((column) => (
                  <th
                    key={column.id}
                    className="px-3 py-2 font-medium text-left text-gray-700 border-r cursor-pointer last:border-r-0 hover:bg-gray-100"
                    style={{ width: column.width }}
                    onClick={(e) => {
                      e.stopPropagation();
                      if (onTableSelect) {
                        onTableSelect(tableConfig.id);
                      }
                    }}
                  >
                    {column.label}
                    {column.required && <span className="ml-1 text-red-500">*</span>}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {Array.from({ length: tableConfig.defaultRows }).map((_, rowIndex) => (
                <tr key={rowIndex} className="border-t">
                  {tableConfig.columns.map((column) => {
                    const cellValue = getPreFilledValue(rowIndex, column.name);
                    const isReadOnly = isCellReadOnly(rowIndex, column.name);
                    
                    return (
                      <td
                        key={column.id}
                        className="px-3 py-2 border-r last:border-r-0"
                      >
                        {column.type === "select" ? (
                          <select 
                            value={cellValue}
                            onChange={(e) => handleTableCellChange(rowIndex, column.name, e.target.value)}
                            className={`w-full text-sm bg-transparent border rounded px-1 py-1 ${
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
                            onChange={(e) => handleTableCellChange(rowIndex, column.name, e.target.checked)}
                            className={`rounded ${isReadOnly && cellValue ? 'bg-blue-50' : ''}`}
                          />
                        ) : (
                          <input
                            type={column.type}
                            value={cellValue}
                            onChange={(e) => handleTableCellChange(rowIndex, column.name, e.target.value)}
                            placeholder={cellValue ? '' : `Sample ${column.type}`}
                            className={`w-full text-sm bg-transparent border rounded px-1 py-1 ${
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
          <div className="p-2 text-xs text-gray-600 bg-gray-50">
            💡 Pre-fill cells by typing values. Blue cells will be read-only for users.
          </div>
        </div>
      );
    }

    return null;
  };

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={isDragging ? "opacity-50" : ""}
    >
      <Card
        className={`cursor-pointer transition-all hover:shadow-md ${
          isSelected ? "ring-2 ring-blue-500" : ""
        } ${isDropOver ? "ring-2 ring-green-400 bg-green-50" : ""}`}
        onClick={onSelect}
      >
        <CardHeader className="pb-3">
          <div className="flex items-center justify-between">
            <CardTitle className="text-lg">{section.title}</CardTitle>
            <div className="flex items-center space-x-1">
              <Button
                variant="ghost"
                size="sm"
                className="drag-handle cursor-grab"
                {...attributes}
                {...listeners}
              >
                <GripVertical className="w-4 h-4" />
              </Button>
              <Button
                variant="ghost"
                size="sm"
                onClick={(e) => {
                  e.stopPropagation();
                  onSelect();
                }}
              >
                <Settings className="w-4 h-4" />
              </Button>
              <Button
                variant="ghost"
                size="sm"
                onClick={(e) => {
                  e.stopPropagation();
                  onDelete();
                }}
                className="text-red-500 hover:text-red-700"
              >
                <Trash2 className="w-4 h-4" />
              </Button>
            </div>
          </div>
          {section.description && (
            <p className="text-sm text-gray-600">{section.description}</p>
          )}
        </CardHeader>
        <CardContent>
          {renderSectionContent()}
        </CardContent>
      </Card>
    </div>
  );
};

export default SectionRenderer;