import React from "react";
import { useDraggable } from "@dnd-kit/core";
import { Card, CardContent } from "../ui/card";
import { Button } from "../ui/button";
import { FIELD_TEMPLATES, ICON_MAP } from "./FieldTemplates";
import { FieldTemplate } from "../../types";
import { Plus } from "lucide-react";

interface DraggableFieldProps {
  template: FieldTemplate;
  onAdd?: (template: FieldTemplate) => void;
}

const DraggableField: React.FC<DraggableFieldProps> = ({ template, onAdd }) => {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: template.id,
    data: { type: "field", template },
  });

  const IconComponent = ICON_MAP[template.icon as keyof typeof ICON_MAP];

  return (
    <Card
      ref={setNodeRef}
      {...attributes}
      {...listeners}
      className={`cursor-grab transition-all hover:shadow-md ${
        isDragging ? "opacity-50 scale-95" : ""
      }`}
    >
      <CardContent className="p-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-2">
            <IconComponent className="w-4 h-4 text-blue-500" />
            <div>
              <div className="text-sm font-medium">{template.label}</div>
              <div className="text-xs text-gray-500">{template.description}</div>
            </div>
          </div>
          <Button
            variant="ghost"
            size="sm"
            onClick={(e) => {
              e.stopPropagation();
              onAdd?.(template);
            }}
            title="Add to form"
          >
            <Plus className="w-4 h-4" />
          </Button>
        </div>
      </CardContent>
    </Card>
  );
};

interface FieldPaletteProps {
  className?: string;
  onAddTemplate?: (template: FieldTemplate) => void;
  onAddTableSection?: () => void;
}

const FieldPalette: React.FC<FieldPaletteProps> = ({ className = "", onAddTemplate, onAddTableSection }) => {
  return (
    <div className={`space-y-4 ${className}`}>
      <div>
        <h3 className="font-semibold text-lg mb-3">Form Elements</h3>
        <div className="space-y-2">
          <div>
            <h4 className="text-sm font-medium text-gray-600 mb-2">Basic Inputs</h4>
            <div className="grid grid-cols-1 gap-2">
              {FIELD_TEMPLATES.filter(t => 
                ["text", "email", "number", "phone", "url"].includes(t.id)
              ).map((template) => (
                <DraggableField key={template.id} template={template} onAdd={onAddTemplate} />
              ))}
            </div>
          </div>

          <div>
            <h4 className="text-sm font-medium text-gray-600 mb-2">Date & Time</h4>
            <div className="grid grid-cols-1 gap-2">
              {FIELD_TEMPLATES.filter(t => 
                ["date", "datetime-local", "time"].includes(t.id)
              ).map((template) => (
                <DraggableField key={template.id} template={template} onAdd={onAddTemplate} />
              ))}
            </div>
          </div>

          <div>
            <h4 className="text-sm font-medium text-gray-600 mb-2">Selection</h4>
            <div className="grid grid-cols-1 gap-2">
              {FIELD_TEMPLATES.filter(t => 
                ["select", "checkbox", "radio"].includes(t.id)
              ).map((template) => (
                <DraggableField key={template.id} template={template} onAdd={onAddTemplate} />
              ))}
            </div>
          </div>

          <div>
            <h4 className="text-sm font-medium text-gray-600 mb-2">Other</h4>
            <div className="grid grid-cols-1 gap-2">
              {FIELD_TEMPLATES.filter(t => 
                ["textarea", "file", "signature"].includes(t.id)
              ).map((template) => (
                <DraggableField key={template.id} template={template} onAdd={onAddTemplate} />
              ))}
            </div>
          </div>
        </div>
      </div>

      <div>
        <h4 className="text-sm font-medium text-gray-600 mb-2">Advanced</h4>
        <div className="grid grid-cols-1 gap-2">
          <DraggableTableSection onAdd={onAddTableSection} />
          {FIELD_TEMPLATES.filter((t) => t.id === "embedded_excel").map((template) => (
            <DraggableField key={template.id} template={template} onAdd={onAddTemplate} />
          ))}
        </div>
      </div>
    </div>
  );
};

const DraggableTableSection: React.FC<{ onAdd?: () => void }> = ({ onAdd }) => {
  const { attributes, listeners, setNodeRef, isDragging } = useDraggable({
    id: "table-section",
    data: { type: "table", template: null },
  });

  return (
    <Card
      ref={setNodeRef}
      {...attributes}
      {...listeners}
      className={`cursor-grab transition-all hover:shadow-md ${
        isDragging ? "opacity-50 scale-95" : ""
      }`}
    >
      <CardContent className="p-3">
        <div className="flex items-center justify-between">
          <div className="flex items-center space-x-2">
            <ICON_MAP.Table className="w-4 h-4 text-green-500" />
            <div>
              <div className="text-sm font-medium">Table Section</div>
              <div className="text-xs text-gray-500">Excel-like data table</div>
            </div>
          </div>
          <Button
            variant="ghost"
            size="sm"
            onClick={(e) => {
              e.stopPropagation();
              onAdd?.();
            }}
            title="Add table"
          >
            <Plus className="w-4 h-4" />
          </Button>
        </div>
      </CardContent>
    </Card>
  );
};

export default FieldPalette;