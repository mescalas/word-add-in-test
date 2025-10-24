import * as React from "react";
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
    </div>
  );
};
