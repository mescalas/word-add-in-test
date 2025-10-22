import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Switch } from "@/components/ui/switch";
import { Zap, Search } from "lucide-react";

interface PatternConfigCardProps {
  startPattern: string;
  endPattern: string;
  savePatterns: boolean;
  onStartPatternChange: (value: string) => void;
  onEndPatternChange: (value: string) => void;
  onSavePatternsChange: (value: boolean) => void;
  onSetPreset: (start: string, end: string) => void;
  onExtract: () => void;
}

export const PatternConfigCard: React.FC<PatternConfigCardProps> = ({
  startPattern,
  endPattern,
  savePatterns,
  onStartPatternChange,
  onEndPatternChange,
  onSavePatternsChange,
  onSetPreset,
  onExtract,
}) => {
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Zap className="w-5 h-5" />
          Configuration des Patterns
        </CardTitle>
        <CardDescription>
          Définissez comment vos variables sont délimitées dans le document
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="grid grid-cols-2 gap-4">
          <div className="space-y-2">
            <Label htmlFor="start-pattern">Pattern de début</Label>
            <Input
              id="start-pattern"
              value={startPattern}
              onChange={(e) => onStartPatternChange(e.target.value)}
              placeholder="Ex: ["
            />
          </div>
          <div className="space-y-2">
            <Label htmlFor="end-pattern">Pattern de fin</Label>
            <Input
              id="end-pattern"
              value={endPattern}
              onChange={(e) => onEndPatternChange(e.target.value)}
              placeholder="Ex: ]"
            />
          </div>
        </div>

        <div className="space-y-2">
          <Label>Patterns prédéfinis</Label>
          <div className="grid grid-cols-2 gap-2">
            <Button variant="outline" size="sm" onClick={() => onSetPreset("[", "]")}>
              [ ] Crochets
            </Button>
            <Button variant="outline" size="sm" onClick={() => onSetPreset("{{", "}}")}>
              {"{{ }}"} Double accolades
            </Button>
            <Button variant="outline" size="sm" onClick={() => onSetPreset("{", "}")}>
              {"{ }"} Accolades
            </Button>
            <Button variant="outline" size="sm" onClick={() => onSetPreset("<", ">")}>
              {"< >"} Chevrons
            </Button>
          </div>
        </div>

        <div className="flex items-center justify-between">
          <Label htmlFor="save-patterns" className="cursor-pointer">
            Sauvegarder les patterns automatiquement
          </Label>
          <Switch
            id="save-patterns"
            checked={savePatterns}
            onCheckedChange={onSavePatternsChange}
          />
        </div>

        <Button onClick={onExtract} className="w-full" size="lg">
          <Search className="w-4 h-4 mr-2" />
          Extraire les Variables
        </Button>
      </CardContent>
    </Card>
  );
};
