import React from "react";
import { useSortable } from "@dnd-kit/sortable";
import { useDroppable } from "@dnd-kit/core";
import { CSS } from "@dnd-kit/utilities";
import { Card, CardContent } from "../ui/card";
import { Button } from "../ui/button";
import { Input } from "../ui/input";
import { Textarea } from "../ui/textarea";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../ui/select";
import { Checkbox } from "../ui/checkbox";
import { RadioGroup, RadioGroupItem } from "../ui/radio-group";
import { Label } from "../ui/label";
import { Trash2, Settings, ArrowUp, ArrowDown } from "lucide-react";
import { FormField } from "../../types";

interface FieldRendererProps {
  field: FormField;
  isSelected: boolean;
  onUpdate: (updates: Partial<FormField>) => void;
  onDelete: () => void;
  onSelect: () => void;
  onMoveUp?: () => void;
  onMoveDown?: () => void;
}

const FieldRenderer: React.FC<FieldRendererProps> = ({
  field,
  isSelected,
  onDelete,
  onSelect,
  onMoveUp,
  onMoveDown,
}) => {
  const {
    setNodeRef: setSortableNodeRef,
    transform,
    transition,
    isDragging,
  } = useSortable({ id: field.id });

  const { setNodeRef: setDroppableNodeRef, isOver: isDropOver } = useDroppable({
    id: field.id,
    data: {
      accepts: ["field"],
      type: "field"
    }
  });

  // Combine refs
  const setNodeRef = (element: HTMLElement | null) => {
    setSortableNodeRef(element);
    setDroppableNodeRef(element);
  };

  const style = {
    transform: CSS.Transform.toString(transform),
    transition,
  };

  const renderFieldInput = () => {
    const commonProps = {
      placeholder: field.placeholder || `Enter ${field.label.toLowerCase()}...`,
      disabled: true,
      className: "w-full",
    };

    switch (field.type) {
      case "text":
      case "email":
      case "number":
      case "date":
      case "datetime-local":
      case "time":
      case "phone":
      case "url":
        return <Input type={field.type} {...commonProps} />;
      
      case "textarea":
        return <Textarea {...commonProps} rows={3} />;
      
      case "select":
        return (
          <Select disabled>
            <SelectTrigger className="w-full">
              <SelectValue placeholder="Select an option..." />
            </SelectTrigger>
            <SelectContent>
              {field.options?.map((option, index) => (
                <SelectItem key={index} value={option.value}>
                  {option.label}
                </SelectItem>
              ))}
            </SelectContent>
          </Select>
        );
      
      case "checkbox":
        return (
          <div className="space-y-2">
            {field.options?.map((option, index) => (
              <div key={index} className="flex items-center space-x-2">
                <Checkbox id={`${field.id}-${index}`} disabled />
                <Label htmlFor={`${field.id}-${index}`} className="text-sm">
                  {option.label}
                </Label>
              </div>
            )) || (
              <div className="flex items-center space-x-2">
                <Checkbox id={field.id} disabled />
                <Label htmlFor={field.id} className="text-sm">
                  {field.label}
                </Label>
              </div>
            )}
          </div>
        );
      
      case "radio":
        return (
          <RadioGroup disabled>
            {field.options?.map((option, index) => (
              <div key={index} className="flex items-center space-x-2">
                <RadioGroupItem value={option.value} id={`${field.id}-${index}`} disabled />
                <Label htmlFor={`${field.id}-${index}`} className="text-sm">
                  {option.label}
                </Label>
              </div>
            ))}
          </RadioGroup>
        );
      
      case "file":
        return (
          <Input type="file" disabled className="w-full" />
        );
      
      case "signature":
        return (
          <div className="space-y-2">
            <div className="flex items-center justify-center w-full h-32 border-2 border-dashed rounded-lg bg-gray-50">
              <div className="text-center">
                <p className="text-sm text-gray-500">Signature Upload Field</p>
                <p className="text-xs text-gray-400">Images only • Max 2MB</p>
              </div>
            </div>
          </div>
        );
      
      default:
        return <Input {...commonProps} />;
    }
  };

  const widthClass = {
    full: "w-full",
    half: "w-4/5",
    third: "w-2/3",
    quarter: "w-3/5",
    auto: "w-auto",
  }[field.layout?.width || "full"];

  return (
    <div
      ref={setNodeRef}
      style={style}
      className={`${widthClass} ${isDragging ? "opacity-50" : ""}`}
    >
      <Card
        className={`cursor-pointer transition-all hover:shadow-md ${
          isSelected ? "ring-2 ring-blue-500" : ""
        } ${isDropOver ? "ring-2 ring-green-400 bg-green-50" : ""}`}
        onClick={onSelect}
      >
        <CardContent className="p-4">
          <div className="flex items-start justify-between mb-3">
            <div className="flex-1">
              <div className="flex items-center mb-1 space-x-2">
                <Label className="text-sm font-medium">
                  {field.label}
                  {field.required && <span className="ml-1 text-red-500">*</span>}
                </Label>
              </div>
              {field.description && (
                <p className="mb-2 text-xs text-gray-600">{field.description}</p>
              )}
            </div>
            <div className="flex items-center space-x-1">
              {/* Drag handle removed as requested */}
              <Button
                variant="ghost"
                size="sm"
                onClick={(e) => {
                  e.stopPropagation();
                  onMoveUp?.();
                }}
                title="Move up"
              >
                <ArrowUp className="w-4 h-4" />
              </Button>
              <Button
                variant="ghost"
                size="sm"
                onClick={(e) => {
                  e.stopPropagation();
                  onMoveDown?.();
                }}
                title="Move down"
              >
                <ArrowDown className="w-4 h-4" />
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
          <div className="space-y-2">
            {renderFieldInput()}
          </div>
        </CardContent>
      </Card>
    </div>
  );
};

export default FieldRenderer;