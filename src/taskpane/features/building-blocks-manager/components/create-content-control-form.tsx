import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select } from "@/components/ui/select";
import { Checkbox } from "@/components/ui/checkbox";
import { Plus } from "lucide-react";
import { CreateContentControlParams, CONTENT_CONTROL_TYPES } from "../types";

interface CreateContentControlFormProps {
  onCreate: (params: CreateContentControlParams) => void;
}

export const CreateContentControlForm: React.FC<CreateContentControlFormProps> = ({ onCreate }) => {
  const [title, setTitle] = React.useState("");
  const [tag, setTag] = React.useState("");
  const [type, setType] = React.useState<string>("RichText");
  const [appearance, setAppearance] = React.useState<string>("BoundingBox");
  const [color, setColor] = React.useState("#3b82f6");
  const [placeholderText, setPlaceholderText] = React.useState("");
  const [cannotDelete, setCannotDelete] = React.useState(false);
  const [cannotEdit, setCannotEdit] = React.useState(false);

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    if (!title.trim()) {
      alert("Veuillez entrer un titre pour le Content Control");
      return;
    }

    onCreate({
      title: title.trim(),
      tag: tag.trim(),
      type: Word.ContentControlType[type as keyof typeof Word.ContentControlType],
      appearance:
        Word.ContentControlAppearance[appearance as keyof typeof Word.ContentControlAppearance],
      color,
      placeholderText: placeholderText.trim(),
      cannotDelete,
      cannotEdit,
    });

    setTitle("");
    setTag("");
    setType("RichText");
    setAppearance("BoundingBox");
    setColor("#3b82f6");
    setPlaceholderText("");
    setCannotDelete(false);
    setCannotEdit(false);
  };

  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Plus className="w-5 h-5" />
          Créer un Content Control
        </CardTitle>
        <CardDescription>
          Sélectionnez du texte et créez un Content Control avec des options personnalisées
        </CardDescription>
      </CardHeader>
      <CardContent>
        <form onSubmit={handleSubmit} className="space-y-4">
          <div className="space-y-2">
            <Label htmlFor="title">Titre *</Label>
            <Input
              id="title"
              type="text"
              placeholder="Ex: Nom du client"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              required
            />
          </div>

          <div className="space-y-2">
            <Label htmlFor="tag">Tag (Identifiant)</Label>
            <Input
              id="tag"
              type="text"
              placeholder="Ex: customer_name"
              value={tag}
              onChange={(e) => setTag(e.target.value)}
            />
          </div>

          <div className="space-y-2">
            <Label htmlFor="type">Type</Label>
            <Select id="type" value={type} onChange={(e) => setType(e.target.value)}>
              {CONTENT_CONTROL_TYPES.map((t) => (
                <option key={t.value} value={t.value}>
                  {t.label}
                </option>
              ))}
            </Select>
          </div>

          <div className="space-y-2">
            <Label htmlFor="appearance">Apparence</Label>
            <Select
              id="appearance"
              value={appearance}
              onChange={(e) => setAppearance(e.target.value)}
            >
              <option value="BoundingBox">Bounding Box</option>
              <option value="Tags">Tags</option>
              <option value="Hidden">Hidden</option>
            </Select>
          </div>
          <div className="space-y-2">
            <Label htmlFor="color">Couleur</Label>
            <div className="flex gap-2">
              <Input
                id="color"
                type="color"
                value={color}
                onChange={(e) => setColor(e.target.value)}
                className="w-20 h-10"
              />
              <Input
                type="text"
                value={color}
                onChange={(e) => setColor(e.target.value)}
                placeholder="#3b82f6"
              />
            </div>
          </div>

          <div className="space-y-2">
            <Label htmlFor="placeholder">Texte de placeholder</Label>
            <Input
              id="placeholder"
              type="text"
              placeholder="Entrez le nom du client ici..."
              value={placeholderText}
              onChange={(e) => setPlaceholderText(e.target.value)}
            />
          </div>

          <div className="space-y-3">
            <div className="flex items-center space-x-2">
              <Checkbox
                id="cannotDelete"
                checked={cannotDelete}
                onCheckedChange={(checked) => setCannotDelete(checked as boolean)}
              />
              <Label htmlFor="cannotDelete" className="cursor-pointer">
                Verrouiller la suppression
              </Label>
            </div>

            <div className="flex items-center space-x-2">
              <Checkbox
                id="cannotEdit"
                checked={cannotEdit}
                onCheckedChange={(checked) => setCannotEdit(checked as boolean)}
              />
              <Label htmlFor="cannotEdit" className="cursor-pointer">
                Verrouiller l&apos;édition
              </Label>
            </div>
          </div>

          <div className="pt-2">
            <Button type="submit" className="w-full">
              <Plus className="w-4 h-4 mr-2" />
              Créer depuis la sélection
            </Button>
          </div>

          <p className="text-xs text-gray-500 text-center">
            Sélectionnez du texte dans le document avant de créer le Content Control
          </p>
        </form>
      </CardContent>
    </Card>
  );
};
