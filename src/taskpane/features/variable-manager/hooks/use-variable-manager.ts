import * as React from "react";
import { ExtractedVariable, STORAGE_KEY } from "../types";

export const useVariableManager = (
  onStatusChange: (status: string) => void
) => {
  const [startPattern, setStartPattern] = React.useState<string>("[");
  const [endPattern, setEndPattern] = React.useState<string>("]");
  const [variables, setVariables] = React.useState<ExtractedVariable[]>([]);
  const [savePatterns, setSavePatterns] = React.useState<boolean>(true);
  const [selectedVariable, setSelectedVariable] = React.useState<string>("");
  const [replacementValue, setReplacementValue] = React.useState<string>("");

  React.useEffect(() => {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (saved) {
      try {
        const { start, end } = JSON.parse(saved);
        setStartPattern(start);
        setEndPattern(end);
      } catch (e) {
        console.error("Erreur lors du chargement des patterns:", e);
      }
    }
  }, []);

  React.useEffect(() => {
    if (savePatterns) {
      localStorage.setItem(
        STORAGE_KEY,
        JSON.stringify({
          start: startPattern,
          end: endPattern,
        })
      );
    }
  }, [startPattern, endPattern, savePatterns]);

  const setPresetPattern = (start: string, end: string) => {
    setStartPattern(start);
    setEndPattern(end);
    onStatusChange(`Pattern défini: ${start}...${end}`);
  };

  const escapeRegex = (str: string) => {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  };

  const extractVariables = async () => {
    try {
      if (!startPattern || !endPattern) {
        onStatusChange("❌ Veuillez définir les patterns de début et fin");
        return;
      }

      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");
        await context.sync();

        const text = body.text;
        const escapedStart = escapeRegex(startPattern);
        const escapedEnd = escapeRegex(endPattern);
        const regex = new RegExp(`${escapedStart}([^${escapedEnd}]+)${escapedEnd}`, "g");

        const variableMap = new Map<string, number>();
        let match: RegExpExecArray | null;

        while ((match = regex.exec(text)) !== null) {
          const varName = match[1].trim();
          variableMap.set(varName, (variableMap.get(varName) || 0) + 1);
        }

        const extractedVars: ExtractedVariable[] = Array.from(variableMap.entries())
          .map(([name, count]) => ({ name, count }))
          .sort((a, b) => b.count - a.count);

        setVariables(extractedVars);
        onStatusChange(`✅ ${extractedVars.length} variable(s) unique(s) trouvée(s)`);

        console.log("═══════════════════════════════════");
        console.log("📋 VARIABLES EXTRAITES");
        console.log("═══════════════════════════════════");
        extractedVars.forEach((v) => {
          console.log(`${startPattern}${v.name}${endPattern} - ${v.count} occurrence(s)`);
        });
        console.log("═══════════════════════════════════");
      });
    } catch (error) {
      console.error("❌ Erreur lors de l'extraction:", error);
      onStatusChange("❌ Erreur lors de l'extraction des variables");
    }
  };

  const searchVariable = async (varName: string) => {
    try {
      await Word.run(async (context) => {
        const searchText = `${startPattern}${varName}${endPattern}`;
        const results = context.document.body.search(searchText);
        results.load("items");
        await context.sync();

        if (results.items.length > 0) {
          results.items[0].select();
          await context.sync();
          onStatusChange(`✅ Variable "${varName}" trouvée et sélectionnée`);
        } else {
          onStatusChange(`❌ Variable "${varName}" non trouvée`);
        }
      });
    } catch (error) {
      console.error("❌ Erreur lors de la recherche:", error);
      onStatusChange("❌ Erreur lors de la recherche");
    }
  };

  const highlightVariable = async (varName: string) => {
    try {
      await Word.run(async (context) => {
        const searchText = `${startPattern}${varName}${endPattern}`;
        const results = context.document.body.search(searchText);
        results.load("items");
        await context.sync();

        results.items.forEach((item) => {
          item.font.highlightColor = "#FFFF00";
        });

        await context.sync();
        onStatusChange(`✅ ${results.items.length} occurrence(s) de "${varName}" surlignée(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors du surlignage:", error);
      onStatusChange("❌ Erreur lors du surlignage");
    }
  };

  const highlightAllVariables = async () => {
    try {
      let totalHighlighted = 0;

      await Word.run(async (context) => {
        for (const variable of variables) {
          const searchText = `${startPattern}${variable.name}${endPattern}`;
          const results = context.document.body.search(searchText);
          results.load("items");
          await context.sync();

          results.items.forEach((item) => {
            item.font.highlightColor = "#FFFF00";
          });

          totalHighlighted += results.items.length;
          await context.sync();
        }

        onStatusChange(`✅ ${totalHighlighted} variable(s) surlignée(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors du surlignage:", error);
      onStatusChange("❌ Erreur lors du surlignage");
    }
  };

  const removeHighlights = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.font.highlightColor = null;
        await context.sync();
        onStatusChange("✅ Surlignage retiré");
      });
    } catch (error) {
      console.error("❌ Erreur:", error);
      onStatusChange("❌ Erreur lors du retrait du surlignage");
    }
  };

  const replaceVariable = async () => {
    try {
      if (!selectedVariable || !replacementValue) {
        onStatusChange("❌ Veuillez sélectionner une variable et entrer une valeur");
        return;
      }

      await Word.run(async (context) => {
        const searchText = `${startPattern}${selectedVariable}${endPattern}`;
        const results = context.document.body.search(searchText, { matchCase: false });
        results.load("items");
        await context.sync();

        const count = results.items.length;

        results.items.forEach((item) => {
          item.insertText(replacementValue, Word.InsertLocation.replace);
        });

        await context.sync();
        onStatusChange(
          `✅ ${count} occurrence(s) de "${selectedVariable}" remplacée(s) par "${replacementValue}"`
        );

        await extractVariables();
      });
    } catch (error) {
      console.error("❌ Erreur lors du remplacement:", error);
      onStatusChange("❌ Erreur lors du remplacement");
    }
  };

  const copyVariablesList = () => {
    const list = variables.map((v) => `${startPattern}${v.name}${endPattern}`).join("\n");
    navigator.clipboard.writeText(list);
    onStatusChange("✅ Liste copiée dans le presse-papier");
  };

  const exportToJSON = () => {
    const data = {
      patterns: { start: startPattern, end: endPattern },
      variables: variables,
      totalUnique: variables.length,
      totalOccurrences: variables.reduce((sum, v) => sum + v.count, 0),
    };

    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "variables-export.json";
    a.click();
    URL.revokeObjectURL(url);

    onStatusChange("✅ Export JSON téléchargé");
  };

  return {
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
  };
};
