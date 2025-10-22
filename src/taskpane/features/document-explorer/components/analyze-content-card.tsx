import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Eye, Type, FileText, List, Table, Palette } from "lucide-react";

interface AnalyzeContentCardProps {
  onReadSelectedText: () => void;
  onReadAllContent: () => void;
  onReadParagraphs: () => void;
  onReadTables: () => void;
  onReadMetadata: () => void;
}

export const AnalyzeContentCard: React.FC<AnalyzeContentCardProps> = ({
  onReadSelectedText,
  onReadAllContent,
  onReadParagraphs,
  onReadTables,
  onReadMetadata,
}) => {
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Eye className="w-5 h-5" />
          Analyser le Contenu
        </CardTitle>
        <CardDescription>
          Explorez différentes façons de lire et analyser le contenu de votre document (résultats
          dans la console)
        </CardDescription>
      </CardHeader>
      <CardContent className="grid grid-cols-1 md:grid-cols-2 gap-3">
        <Button onClick={onReadSelectedText} variant="outline">
          <Type className="w-4 h-4 mr-2" />
          Texte Sélectionné
        </Button>
        <Button onClick={onReadAllContent} variant="outline">
          <FileText className="w-4 h-4 mr-2" />
          Document Complet
        </Button>
        <Button onClick={onReadParagraphs} variant="outline">
          <List className="w-4 h-4 mr-2" />
          Paragraphes
        </Button>
        <Button onClick={onReadTables} variant="outline">
          <Table className="w-4 h-4 mr-2" />
          Tableaux
        </Button>
        <Button onClick={onReadMetadata} variant="outline" className="md:col-span-2">
          <Palette className="w-4 h-4 mr-2" />
          Métadonnées
        </Button>
      </CardContent>
    </Card>
  );
};
