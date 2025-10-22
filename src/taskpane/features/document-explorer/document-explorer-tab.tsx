import * as React from "react";
import { Card, CardContent, CardHeader, CardTitle } from "@/components/ui/card";
import { useDocumentExplorer } from "./hooks/use-document-explorer";
import { CreateDemoCard } from "./components/create-demo-card";
import { AnalyzeContentCard } from "./components/analyze-content-card";

interface DocumentExplorerTabProps {
  onStatusChange: (status: string) => void;
}

export const DocumentExplorerTab: React.FC<DocumentExplorerTabProps> = ({ onStatusChange }) => {
  const {
    createDemoDocument,
    readSelectedText,
    readAllContent,
    readParagraphs,
    readTables,
    readMetadata,
  } = useDocumentExplorer(onStatusChange);

  return (
    <div className="space-y-6">
      <CreateDemoCard onCreate={createDemoDocument} />

      <AnalyzeContentCard
        onReadSelectedText={readSelectedText}
        onReadAllContent={readAllContent}
        onReadParagraphs={readParagraphs}
        onReadTables={readTables}
        onReadMetadata={readMetadata}
      />

      <Card className="bg-blue-50 border-blue-200">
        <CardHeader>
          <CardTitle className="text-blue-900">Instructions</CardTitle>
        </CardHeader>
        <CardContent className="text-sm text-blue-800 space-y-2">
          <p>
            1. Cliquez sur <strong>"Générer le Document"</strong> pour créer un document de
            démonstration
          </p>
          <p>
            2. Ouvrez la <strong>Console du navigateur</strong> (F12) pour voir les résultats
            détaillés
          </p>
          <p>3. Utilisez les boutons d'analyse pour explorer le contenu via console.log</p>
          <p>
            4. Sélectionnez du texte dans le document et cliquez sur{" "}
            <strong>"Texte Sélectionné"</strong> pour voir ses propriétés
          </p>
        </CardContent>
      </Card>
    </div>
  );
};
