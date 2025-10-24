import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Input } from "@/components/ui/input";
import { List, Search, MousePointer, Trash2, Eye } from "lucide-react";
import { ContentControlInfo } from "../types";

interface ContentControlsListProps {
  controls: ContentControlInfo[];
  isLoading: boolean;
  onRefresh: () => void;
  onSelect: (controlId: number) => void;
  onDelete: (controlId: number) => void;
  onEdit: (control: ContentControlInfo) => void;
}

export const ContentControlsList: React.FC<ContentControlsListProps> = ({
  controls,
  isLoading,
  onRefresh,
  onSelect,
  onDelete,
  onEdit,
}) => {
  const [searchQuery, setSearchQuery] = React.useState("");

  const filteredControls = controls.filter(
    (control) =>
      control.title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      control.tag.toLowerCase().includes(searchQuery.toLowerCase()) ||
      control.type.toLowerCase().includes(searchQuery.toLowerCase())
  );

  const getTypeColor = (type: string) => {
    const colors: Record<string, string> = {
      RichText: "bg-blue-100 text-blue-800",
      PlainText: "bg-gray-100 text-gray-800",
      CheckBox: "bg-green-100 text-green-800",
      ComboBox: "bg-purple-100 text-purple-800",
      DropDownList: "bg-purple-100 text-purple-800",
      DatePicker: "bg-pink-100 text-pink-800",
      Picture: "bg-orange-100 text-orange-800",
      BuildingBlockGallery: "bg-indigo-100 text-indigo-800",
      Paragraph: "bg-teal-100 text-teal-800",
    };
    return colors[type] || "bg-gray-100 text-gray-800";
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <List className="w-5 h-5" />
          Content Controls du Document
        </CardTitle>
        <CardDescription>
          Gérez tous les Content Controls présents dans votre document
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="flex gap-2">
          <Button onClick={onRefresh} disabled={isLoading} variant="outline" className="flex-1">
            <List className="w-4 h-4 mr-2" />
            {isLoading ? "Chargement..." : "Actualiser la liste"}
          </Button>
        </div>

        {controls.length > 0 && (
          <div className="relative">
            <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 w-4 h-4 text-gray-400" />
            <Input
              type="text"
              placeholder="Rechercher par titre, tag ou type..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="pl-9"
            />
          </div>
        )}

        {controls.length === 0 ? (
          <div className="text-center py-8 text-gray-500">
            <List className="w-12 h-12 mx-auto mb-3 text-gray-300" />
            <p className="text-sm">Aucun Content Control trouvé</p>
            <p className="text-xs mt-1">Créez-en un à partir d&apos;une sélection de texte</p>
          </div>
        ) : filteredControls.length === 0 ? (
          <div className="text-center py-8 text-gray-500">
            <Search className="w-12 h-12 mx-auto mb-3 text-gray-300" />
            <p className="text-sm">Aucun résultat pour &quot;{searchQuery}&quot;</p>
          </div>
        ) : (
          <div className="space-y-2 max-h-96 overflow-y-auto">
            {filteredControls.map((control) => (
              <div
                key={control.id}
                className="border rounded-lg p-3 hover:bg-gray-50 transition-colors"
              >
                <div className="flex items-start justify-between gap-2">
                  <div className="flex-1 min-w-0">
                    <div className="flex items-center gap-2 mb-1">
                      <h4 className="font-semibold text-sm truncate">
                        {control.title || "(Sans titre)"}
                      </h4>
                      <Badge variant="secondary" className="text-xs">
                        ID: {control.id}
                      </Badge>
                    </div>

                    {control.tag && (
                      <p className="text-xs text-gray-600 mb-1">Tag: {control.tag}</p>
                    )}

                    {control.text && (
                      <p className="text-xs text-gray-500 italic mb-2 line-clamp-1">
                        {control.text}
                      </p>
                    )}

                    <div className="flex flex-wrap gap-1">
                      <Badge className={getTypeColor(control.type)}>{control.type}</Badge>
                      <Badge variant="outline" className="text-xs">
                        {control.appearance}
                      </Badge>
                      {control.cannotDelete && (
                        <Badge variant="destructive" className="text-xs">
                          🔒 Cannot Delete
                        </Badge>
                      )}
                      {control.cannotEdit && (
                        <Badge variant="destructive" className="text-xs">
                          🔒 Cannot Edit
                        </Badge>
                      )}
                    </div>
                  </div>

                  <div className="flex gap-1 flex-shrink-0">
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => onSelect(control.id)}
                      title="Sélectionner dans le document"
                    >
                      <MousePointer className="w-3 h-3" />
                    </Button>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => onEdit(control)}
                      title="Voir les détails"
                    >
                      <Eye className="w-3 h-3" />
                    </Button>
                    <Button
                      size="sm"
                      variant="outline"
                      onClick={() => {
                        if (window.confirm(`Supprimer le Content Control "${control.title}" ?`)) {
                          onDelete(control.id);
                        }
                      }}
                      title="Supprimer"
                      className="text-red-600 hover:text-red-700 hover:bg-red-50"
                      disabled={control.cannotDelete}
                    >
                      <Trash2 className="w-3 h-3" />
                    </Button>
                  </div>
                </div>
              </div>
            ))}
          </div>
        )}

        {controls.length > 0 && (
          <div className="text-xs text-gray-500 text-center pt-2 border-t">
            {filteredControls.length} content control(s) affiché(s) sur {controls.length} au total
          </div>
        )}
      </CardContent>
    </Card>
  );
};
