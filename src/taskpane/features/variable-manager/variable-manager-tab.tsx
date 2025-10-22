import * as React from "react";
import {
  Empty,
  EmptyContent,
  EmptyDescription,
  EmptyHeader,
  EmptyMedia,
  EmptyTitle,
} from "@/components/ui/empty";
import { AlertCircle } from "lucide-react";
import { useVariableManager } from "./hooks/use-variable-manager";
import { PatternConfigCard } from "./components/pattern-config-card";
import { VariableListCard } from "./components/variable-list-card";
import { ReplaceVariableCard } from "./components/replace-variable-card";

interface VariableManagerTabProps {
  onStatusChange: (status: string) => void;
}

export const VariableManagerTab: React.FC<VariableManagerTabProps> = ({ onStatusChange }) => {
  const {
    startPattern,
    setStartPattern,
    endPattern,
    setEndPattern,
    variables,
    savePatterns,
    setSavePatterns,
    selectedVariable,
    setSelectedVariable,
    replacementValue,
    setReplacementValue,
    setPresetPattern,
    extractVariables,
    searchVariable,
    highlightVariable,
    highlightAllVariables,
    removeHighlights,
    replaceVariable,
    copyVariablesList,
    exportToJSON,
  } = useVariableManager(onStatusChange);

  return (
    <div className="space-y-6">
      <PatternConfigCard
        startPattern={startPattern}
        endPattern={endPattern}
        savePatterns={savePatterns}
        onStartPatternChange={setStartPattern}
        onEndPatternChange={setEndPattern}
        onSavePatternsChange={setSavePatterns}
        onSetPreset={setPresetPattern}
        onExtract={extractVariables}
      />

      {variables.length > 0 && (
        <>
          <VariableListCard
            variables={variables}
            startPattern={startPattern}
            endPattern={endPattern}
            onSearch={searchVariable}
            onHighlight={highlightVariable}
            onHighlightAll={highlightAllVariables}
            onRemoveHighlights={removeHighlights}
            onCopyList={copyVariablesList}
            onExport={exportToJSON}
          />

          <ReplaceVariableCard
            variables={variables}
            startPattern={startPattern}
            endPattern={endPattern}
            selectedVariable={selectedVariable}
            replacementValue={replacementValue}
            onSelectedVariableChange={setSelectedVariable}
            onReplacementValueChange={setReplacementValue}
            onReplace={replaceVariable}
          />
        </>
      )}

      {variables.length === 0 && (
        <Empty className="border border-dashed">
          <EmptyHeader>
            <EmptyMedia variant="icon">
              <AlertCircle />
            </EmptyMedia>
            <EmptyTitle>Aucune variable trouvée</EmptyTitle>
            <EmptyDescription className="text-sm text-gray-500">
              Configurez vos patterns ci-dessus et cliquez sur "Extraire les Variables" pour
              analyser votre document. Exemple : avec les patterns [ et ], la variable [nom] sera
              détectée.
            </EmptyDescription>
          </EmptyHeader>
        </Empty>
      )}
    </div>
  );
};
