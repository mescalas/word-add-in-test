import * as React from "react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Info, Save } from "lucide-react";
import { ContentControlInfo } from "../types";

interface ContentControlDetailsProps {
  control: ContentControlInfo | null;
  onUpdateText: (controlId: number, newText: string) => void;
}

export const ContentControlDetails: React.FC<ContentControlDetailsProps> = ({
  control,
  onUpdateText,
}) => {
  const [newText, setNewText] = React.useState("");

  React.useEffect(() => {
    if (control) {
      setNewText(control.text);
    }
  }, [control]);

  if (!control) {
    return (
      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2">
            <Info className="w-5 h-5" />
            Détails du Content Control
          </CardTitle>
          <CardDescription>
            Sélectionnez un Content Control dans la liste pour voir ses détails
          </CardDescription>
        </CardHeader>
        <CardContent>
          <div className="text-center py-8 text-gray-500">
            <Info className="w-12 h-12 mx-auto mb-3 text-gray-300" />
            <p className="text-sm">Aucun Content Control sélectionné</p>
          </div>
        </CardContent>
      </Card>
    );
  }

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Info className="w-5 h-5" />
          Détails du Content Control
        </CardTitle>
        <CardDescription>ID: {control.id}</CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="space-y-2">
          <Label className="text-sm font-semibold">Titre</Label>
          <p className="text-sm">{control.title || "(Sans titre)"}</p>
        </div>

        <div className="space-y-2">
          <Label className="text-sm font-semibold">Tag</Label>
          <p className="text-sm">{control.tag || "(Aucun)"}</p>
        </div>

        <div className="space-y-2">
          <Label className="text-sm font-semibold">Type</Label>
          <Badge variant="secondary">{control.type}</Badge>
        </div>

        <div className="space-y-2">
          <Label className="text-sm font-semibold">Apparence</Label>
          <Badge variant="outline">{control.appearance}</Badge>
        </div>

        <div className="space-y-2">
          <Label className="text-sm font-semibold">Couleur</Label>
          <div className="flex items-center gap-2">
            <div
              className="w-8 h-8 rounded border"
              style={{ backgroundColor: control.color }}
            />
            <span className="text-sm font-mono">{control.color}</span>
          </div>
        </div>

        {control.placeholderText && (
          <div className="space-y-2">
            <Label className="text-sm font-semibold">Placeholder</Label>
            <p className="text-sm italic text-gray-600">{control.placeholderText}</p>
          </div>
        )}

        <div className="space-y-2">
          <Label className="text-sm font-semibold">Restrictions</Label>
          <div className="flex gap-2">
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
            {!control.cannotDelete && !control.cannotEdit && (
              <span className="text-sm text-gray-500">Aucune restriction</span>
            )}
          </div>
        </div>

        <div className="border-t pt-4 space-y-3">
          <Label className="text-sm font-semibold">Modifier le contenu</Label>
          <Input
            type="text"
            placeholder="Nouveau texte..."
            value={newText}
            onChange={(e) => setNewText(e.target.value)}
            disabled={control.cannotEdit}
          />
          <Button
            onClick={() => onUpdateText(control.id, newText)}
            disabled={control.cannotEdit || !newText.trim()}
            className="w-full"
            size="sm"
          >
            <Save className="w-4 h-4 mr-2" />
            Enregistrer le nouveau texte
          </Button>
          {control.cannotEdit && (
            <p className="text-xs text-red-600 text-center">
              Ce Content Control est verrouillé en édition
            </p>
          )}
        </div>
      </CardContent>
    </Card>
  );
};
