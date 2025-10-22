import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { Select } from "@/components/ui/select";
import { RefreshCw } from "lucide-react";
import { ExtractedVariable } from "../types";

interface ReplaceVariableCardProps {
  variables: ExtractedVariable[];
  startPattern: string;
  endPattern: string;
  selectedVariable: string;
  replacementValue: string;
  onSelectedVariableChange: (value: string) => void;
  onReplacementValueChange: (value: string) => void;
  onReplace: () => void;
}

export const ReplaceVariableCard: React.FC<ReplaceVariableCardProps> = ({
  variables,
  startPattern,
  endPattern,
  selectedVariable,
  replacementValue,
  onSelectedVariableChange,
  onReplacementValueChange,
  onReplace,
}) => {
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <RefreshCw className="w-5 h-5" />
          Remplacer une Variable
        </CardTitle>
        <CardDescription>
          Sélectionnez une variable et définissez sa nouvelle valeur
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="space-y-2">
          <Label htmlFor="variable-select">Variable à remplacer</Label>
          <Select
            id="variable-select"
            value={selectedVariable}
            onChange={(e) => onSelectedVariableChange(e.target.value)}
          >
            <option value="">-- Choisir une variable --</option>
            {variables.map((variable, index) => (
              <option key={index} value={variable.name}>
                {startPattern}
                {variable.name}
                {endPattern} ({variable.count}x)
              </option>
            ))}
          </Select>
        </div>

        <div className="space-y-2">
          <Label htmlFor="replacement-value">Nouvelle valeur</Label>
          <Input
            id="replacement-value"
            value={replacementValue}
            onChange={(e) => onReplacementValueChange(e.target.value)}
            placeholder="Entrez la valeur de remplacement"
          />
        </div>

        <Button onClick={onReplace} className="w-full">
          <RefreshCw className="w-4 h-4 mr-2" />
          Remplacer Tout
        </Button>
      </CardContent>
    </Card>
  );
};
