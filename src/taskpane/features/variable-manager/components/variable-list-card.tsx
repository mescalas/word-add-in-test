import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Variable, Search, Highlighter, RefreshCw, Copy, Download } from "lucide-react";
import { ExtractedVariable } from "../types";

interface VariableListCardProps {
  variables: ExtractedVariable[];
  startPattern: string;
  endPattern: string;
  onSearch: (varName: string) => void;
  onHighlight: (varName: string) => void;
  onHighlightAll: () => void;
  onRemoveHighlights: () => void;
  onCopyList: () => void;
  onExport: () => void;
}

export const VariableListCard: React.FC<VariableListCardProps> = ({
  variables,
  startPattern,
  endPattern,
  onSearch,
  onHighlight,
  onHighlightAll,
  onRemoveHighlights,
  onCopyList,
  onExport,
}) => {
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Variable className="w-5 h-5" />
          Variables Trouvées
        </CardTitle>
        <CardDescription>
          {variables.length} variable(s) unique(s) -{" "}
          {variables.reduce((sum, v) => sum + v.count, 0)} occurrence(s) totale(s)
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="flex gap-2">
          <Button onClick={onHighlightAll} variant="secondary" className="flex-1">
            <Highlighter className="w-4 h-4 mr-2" />
            Tout Surligner
          </Button>
          <Button onClick={onRemoveHighlights} variant="outline" className="flex-1">
            <RefreshCw className="w-4 h-4 mr-2" />
            Retirer
          </Button>
        </div>

        <ScrollArea className="h-64 rounded-md border p-4">
          <div className="space-y-2">
            {variables.map((variable, index) => (
              <div
                key={index}
                className="flex items-center justify-between p-3 rounded-lg bg-white hover:bg-gray-50 transition-colors"
              >
                <div className="flex items-center gap-3 flex-1">
                  <Badge variant="secondary">{variable.count}x</Badge>
                  <code className="text-sm font-mono">
                    {startPattern}
                    {variable.name}
                    {endPattern}
                  </code>
                </div>
                <div className="flex gap-1">
                  <Button
                    size="sm"
                    variant="ghost"
                    onClick={() => onSearch(variable.name)}
                    title="Rechercher"
                  >
                    <Search className="w-4 h-4" />
                  </Button>
                  <Button
                    size="sm"
                    variant="ghost"
                    onClick={() => onHighlight(variable.name)}
                    title="Surligner"
                  >
                    <Highlighter className="w-4 h-4" />
                  </Button>
                </div>
              </div>
            ))}
          </div>
        </ScrollArea>

        <div className="flex gap-2">
          <Button onClick={onCopyList} variant="outline" className="flex-1">
            <Copy className="w-4 h-4 mr-2" />
            Copier la Liste
          </Button>
          <Button onClick={onExport} variant="outline" className="flex-1">
            <Download className="w-4 h-4 mr-2" />
            Export JSON
          </Button>
        </div>
      </CardContent>
    </Card>
  );
};
