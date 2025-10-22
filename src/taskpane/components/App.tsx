import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Separator } from "@/components/ui/separator";
import { Badge } from "@/components/ui/badge";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { ScrollArea } from "@/components/ui/scroll-area";
import { Select } from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import {
  Empty,
  EmptyContent,
  EmptyDescription,
  EmptyHeader,
  EmptyMedia,
  EmptyTitle,
} from "@/components/ui/empty";
import {
  FileText,
  Eye,
  Palette,
  Table,
  List,
  Type,
  Variable,
  Search,
  Highlighter,
  Copy,
  Download,
  RefreshCw,
  Zap,
  AlertCircle,
} from "lucide-react";

interface ExtractedVariable {
  name: string;
  count: number;
}

const STORAGE_KEY = "word-addin-variable-patterns";

const App: React.FC = () => {
  const [status, setStatus] = React.useState<string>("");
  const [activeTab, setActiveTab] = React.useState<string>("explorer");

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
    setStatus(`Pattern défini: ${start}...${end}`);
  };

  const escapeRegex = (str: string) => {
    return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  };

  const extractVariables = async () => {
    try {
      if (!startPattern || !endPattern) {
        setStatus("❌ Veuillez définir les patterns de début et fin");
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
        setStatus(`✅ ${extractedVars.length} variable(s) unique(s) trouvée(s)`);

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
      setStatus("❌ Erreur lors de l'extraction des variables");
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
          setStatus(`✅ Variable "${varName}" trouvée et sélectionnée`);
        } else {
          setStatus(`❌ Variable "${varName}" non trouvée`);
        }
      });
    } catch (error) {
      console.error("❌ Erreur lors de la recherche:", error);
      setStatus("❌ Erreur lors de la recherche");
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
          item.font.highlightColor = "#FFFF00"; // Jaune
        });

        await context.sync();
        setStatus(`✅ ${results.items.length} occurrence(s) de "${varName}" surlignée(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors du surlignage:", error);
      setStatus("❌ Erreur lors du surlignage");
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

        setStatus(`✅ ${totalHighlighted} variable(s) surlignée(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors du surlignage:", error);
      setStatus("❌ Erreur lors du surlignage");
    }
  };

  const removeHighlights = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.font.highlightColor = null;
        await context.sync();
        setStatus("✅ Surlignage retiré");
      });
    } catch (error) {
      console.error("❌ Erreur:", error);
      setStatus("❌ Erreur lors du retrait du surlignage");
    }
  };

  const replaceVariable = async () => {
    try {
      if (!selectedVariable || !replacementValue) {
        setStatus("❌ Veuillez sélectionner une variable et entrer une valeur");
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
        setStatus(
          `✅ ${count} occurrence(s) de "${selectedVariable}" remplacée(s) par "${replacementValue}"`
        );

        await extractVariables();
      });
    } catch (error) {
      console.error("❌ Erreur lors du remplacement:", error);
      setStatus("❌ Erreur lors du remplacement");
    }
  };

  const copyVariablesList = () => {
    const list = variables.map((v) => `${startPattern}${v.name}${endPattern}`).join("\n");
    navigator.clipboard.writeText(list);
    setStatus("✅ Liste copiée dans le presse-papier");
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

    setStatus("✅ Export JSON téléchargé");
  };

  // ========== Fonctions de l'onglet Explorateur (code existant) ==========

  const createDemoDocument = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();

        const title = body.insertParagraph(
          "Guide de Découverte de l'API Office JS pour Word",
          Word.InsertLocation.start
        );
        title.styleBuiltIn = Word.BuiltInStyleName.title;
        title.font.color = "#2563eb";
        title.font.size = 28;

        const subtitle = body.insertParagraph(
          "Explorez les capacités de formatage et de manipulation de texte",
          Word.InsertLocation.end
        );
        subtitle.styleBuiltIn = Word.BuiltInStyleName.subtitle;
        subtitle.font.color = "#64748b";

        const heading1 = body.insertParagraph("1. Formatage de Texte", Word.InsertLocation.end);
        heading1.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading1.font.color = "#1e40af";

        const para1 = body.insertParagraph(
          "Voici un exemple de texte avec différents styles : ",
          Word.InsertLocation.end
        );

        const boldText = para1.insertText("texte en gras", Word.InsertLocation.end);
        boldText.font.bold = true;

        para1.insertText(", ", Word.InsertLocation.end);

        const italicText = para1.insertText("texte en italique", Word.InsertLocation.end);
        italicText.font.italic = true;

        para1.insertText(", ", Word.InsertLocation.end);

        const underlineText = para1.insertText("texte souligné", Word.InsertLocation.end);
        underlineText.font.underline = Word.UnderlineType.single;

        para1.insertText(", et ", Word.InsertLocation.end);

        const colorText = para1.insertText("texte coloré", Word.InsertLocation.end);
        colorText.font.color = "#dc2626";
        colorText.font.bold = true;

        const heading2 = body.insertParagraph(
          "2. Listes à Puces et Numérotées",
          Word.InsertLocation.end
        );
        heading2.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading2.font.color = "#1e40af";

        body.insertParagraph("Liste à puces des fonctionnalités :", Word.InsertLocation.end);

        const bullet1 = body.insertParagraph("Insertion de texte", Word.InsertLocation.end);
        const bulletList = bullet1.startNewList();
        bulletList.load("id");
        await context.sync();

        const bullet2 = body.insertParagraph("Formatage de paragraphes", Word.InsertLocation.end);
        bullet2.attachToList(bulletList.id, 0);

        const bullet3 = body.insertParagraph("Gestion des styles", Word.InsertLocation.end);
        bullet3.attachToList(bulletList.id, 0);

        body.insertParagraph("", Word.InsertLocation.end);

        body.insertParagraph("Liste numérotée des étapes :", Word.InsertLocation.end);

        const num1 = body.insertParagraph("Ouvrir Word", Word.InsertLocation.end);
        const numberedList = num1.startNewList();
        numberedList.load("id");
        await context.sync();

        const num2 = body.insertParagraph("Lancer le Add-in", Word.InsertLocation.end);
        num2.attachToList(numberedList.id, 0);

        const num3 = body.insertParagraph("Explorer les fonctionnalités", Word.InsertLocation.end);
        num3.attachToList(numberedList.id, 0);

        const heading3 = body.insertParagraph("3. Tableaux", Word.InsertLocation.end);
        heading3.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading3.font.color = "#1e40af";

        body.insertParagraph("Voici un exemple de tableau :", Word.InsertLocation.end);

        const table = body.insertTable(4, 3, Word.InsertLocation.end, [
          ["Fonctionnalité", "Type", "Complexité"],
          ["Formatage texte", "Basique", "Facile"],
          ["Tableaux", "Intermédiaire", "Moyenne"],
          ["Images", "Avancé", "Difficile"],
        ]);

        table.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
        table.headerRowCount = 1;

        const headerRow = table.rows.getFirst();
        headerRow.font.bold = true;
        headerRow.font.color = "#ffffff";
        headerRow.shadingColor = "#2563eb";

        const heading4 = body.insertParagraph("4. Tailles de Police", Word.InsertLocation.end);
        heading4.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading4.font.color = "#1e40af";

        const size12 = body.insertParagraph("Texte en taille 12", Word.InsertLocation.end);
        size12.font.size = 12;

        const size16 = body.insertParagraph("Texte en taille 16", Word.InsertLocation.end);
        size16.font.size = 16;

        const size20 = body.insertParagraph("Texte en taille 20", Word.InsertLocation.end);
        size20.font.size = 20;

        const heading5 = body.insertParagraph("5. Palette de Couleurs", Word.InsertLocation.end);
        heading5.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading5.font.color = "#1e40af";

        const colors = [
          { name: "Rouge", color: "#ef4444" },
          { name: "Bleu", color: "#3b82f6" },
          { name: "Vert", color: "#10b981" },
          { name: "Violet", color: "#8b5cf6" },
          { name: "Orange", color: "#f97316" },
        ];

        colors.forEach(({ name, color }) => {
          const colorPara = body.insertParagraph(`${name} - `, Word.InsertLocation.end);
          const coloredText = colorPara.insertText(
            "Exemple de texte coloré",
            Word.InsertLocation.end
          );
          coloredText.font.color = color;
          coloredText.font.bold = true;
        });

        // Section 6: Variables et Templates
        const heading6 = body.insertParagraph("6. Système de Variables", Word.InsertLocation.end);
        heading6.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading6.font.color = "#1e40af";

        body.insertParagraph(
          "Ce document contient des exemples de variables que vous pouvez extraire et manipuler via l'onglet 'Variables' :",
          Word.InsertLocation.end
        );

        // Exemple avec crochets []
        const varPara1 = body.insertParagraph(
          "Bonjour [prenom] [nom], bienvenue dans notre système !",
          Word.InsertLocation.end
        );
        varPara1.font.size = 14;

        // Exemple de lettre type
        const letterHeading = body.insertParagraph(
          "\nExemple de Lettre Type :",
          Word.InsertLocation.end
        );
        letterHeading.font.bold = true;
        letterHeading.font.size = 13;

        body.insertParagraph("Cher/Chère [titre] [nom],", Word.InsertLocation.end);
        body.insertParagraph(
          "Nous vous écrivons concernant votre commande [numero_commande] passée le [date_commande].",
          Word.InsertLocation.end
        );
        body.insertParagraph(
          "Le montant total s'élève à [montant] euros et sera livré à l'adresse [adresse].",
          Word.InsertLocation.end
        );
        body.insertParagraph(
          "Pour toute question, contactez-nous au [telephone] ou par email à [email].",
          Word.InsertLocation.end
        );

        // Exemples avec d'autres patterns
        body.insertParagraph("", Word.InsertLocation.end);
        const otherPatterns = body.insertParagraph(
          "\nAutres formats de variables :",
          Word.InsertLocation.end
        );
        otherPatterns.font.bold = true;
        otherPatterns.font.size = 13;

        body.insertParagraph(
          "Double accolades : Bonjour {{prenom}} {{nom}}",
          Word.InsertLocation.end
        );
        body.insertParagraph(
          "Accolades simples : Commande {ref} du {date}",
          Word.InsertLocation.end
        );
        body.insertParagraph("Chevrons : Template <type> pour <client>", Word.InsertLocation.end);

        // Exemple de facture
        body.insertParagraph("", Word.InsertLocation.end);
        const invoiceHeading = body.insertParagraph(
          "Exemple de Facture :",
          Word.InsertLocation.end
        );
        invoiceHeading.font.bold = true;
        invoiceHeading.font.size = 13;

        const invoiceTable = body.insertTable(4, 2, Word.InsertLocation.end, [
          ["Champ", "Valeur"],
          ["Client", "[nom_client]"],
          ["Facture N°", "[numero_facture]"],
          ["Total", "[montant_total] €"],
        ]);

        invoiceTable.styleBuiltIn = Word.BuiltInStyleName.gridTable1Light;
        invoiceTable.headerRowCount = 1;

        const invoiceHeader = invoiceTable.rows.getFirst();
        invoiceHeader.font.bold = true;
        invoiceHeader.font.color = "#ffffff";
        invoiceHeader.shadingColor = "#059669";

        // Instructions
        body.insertParagraph("", Word.InsertLocation.end);
        const instructionsPara = body.insertParagraph(
          "💡 Astuce : Allez dans l'onglet 'Variables' pour extraire et manipuler toutes ces variables automatiquement !",
          Word.InsertLocation.end
        );
        instructionsPara.font.italic = true;
        instructionsPara.font.color = "#6366f1";

        const conclusion = body.insertParagraph(
          "\n🎉 Félicitations ! Vous venez de créer un document riche avec l'API Office JS !",
          Word.InsertLocation.end
        );
        conclusion.alignment = Word.Alignment.centered;
        conclusion.font.size = 16;
        conclusion.font.color = "#059669";
        conclusion.font.bold = true;

        await context.sync();
        setStatus("✅ Document de démonstration créé avec succès !");
        console.log("✅ Document de démonstration créé avec succès !");
      });
    } catch (error) {
      console.error("❌ Erreur lors de la création du document:", error);
      setStatus("❌ Erreur lors de la création du document");
    }
  };

  const readSelectedText = async () => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text,font,style");

        await context.sync();

        console.log("═══════════════════════════════════");
        console.log("📝 TEXTE SÉLECTIONNÉ");
        console.log("═══════════════════════════════════");
        console.log("Texte:", selection.text);
        console.log("Police:", selection.font.name);
        console.log("Taille:", selection.font.size);
        console.log("Couleur:", selection.font.color);
        console.log("Gras:", selection.font.bold);
        console.log("Italique:", selection.font.italic);
        console.log("Souligné:", selection.font.underline);
        console.log("Style:", selection.style);
        console.log("═══════════════════════════════════");

        setStatus(
          `✅ Texte sélectionné: "${selection.text.substring(0, 50)}${selection.text.length > 50 ? "..." : ""}"`
        );
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture du texte sélectionné:", error);
      setStatus("❌ Erreur - Veuillez sélectionner du texte");
    }
  };

  const readAllContent = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.load("text");

        await context.sync();

        console.log("═══════════════════════════════════");
        console.log("📄 CONTENU COMPLET DU DOCUMENT");
        console.log("═══════════════════════════════════");
        console.log("Texte complet:", body.text);
        console.log("Longueur:", body.text.length, "caractères");
        console.log("═══════════════════════════════════");

        setStatus(`✅ Document lu: ${body.text.length} caractères`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture du document:", error);
      setStatus("❌ Erreur lors de la lecture");
    }
  };

  const readParagraphs = async () => {
    try {
      await Word.run(async (context) => {
        const paragraphs = context.document.body.paragraphs;
        paragraphs.load("items,text,style,alignment,font");

        await context.sync();

        console.log("═══════════════════════════════════");
        console.log("📑 STRUCTURE DES PARAGRAPHES");
        console.log("═══════════════════════════════════");
        console.log("Nombre de paragraphes:", paragraphs.items.length);

        paragraphs.items.forEach((paragraph, index) => {
          console.log(`\n--- Paragraphe ${index + 1} ---`);
          console.log("Texte:", paragraph.text.substring(0, 100));
          console.log("Style:", paragraph.style);
          console.log("Alignement:", paragraph.alignment);
          console.log("Taille de police:", paragraph.font.size);
          console.log("Couleur:", paragraph.font.color);
        });

        console.log("═══════════════════════════════════");

        setStatus(`✅ ${paragraphs.items.length} paragraphes analysés`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture des paragraphes:", error);
      setStatus("❌ Erreur lors de la lecture des paragraphes");
    }
  };

  const readTables = async () => {
    try {
      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items,rowCount,values");

        await context.sync();

        console.log("═══════════════════════════════════");
        console.log("📊 TABLEAUX DU DOCUMENT");
        console.log("═══════════════════════════════════");
        console.log("Nombre de tableaux:", tables.items.length);

        tables.items.forEach((table, index) => {
          console.log(`\n--- Tableau ${index + 1} ---`);
          console.log("Lignes:", table.rowCount);
          console.log("Contenu:", table.values);
        });

        console.log("═══════════════════════════════════");

        setStatus(`✅ ${tables.items.length} tableau(x) analysé(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture des tableaux:", error);
      setStatus("❌ Erreur lors de la lecture des tableaux");
    }
  };

  const readMetadata = async () => {
    try {
      await Word.run(async (context) => {
        const properties = context.document.properties;
        properties.load(
          "title,subject,author,keywords,comments,creationDate,lastAuthor,lastPrintDate,lastSaveTime,revisionNumber"
        );

        await context.sync();

        console.log("═══════════════════════════════════");
        console.log("ℹ️  MÉTADONNÉES DU DOCUMENT");
        console.log("═══════════════════════════════════");
        console.log("Titre:", properties.title);
        console.log("Sujet:", properties.subject);
        console.log("Auteur:", properties.author);
        console.log("Mots-clés:", properties.keywords);
        console.log("Commentaires:", properties.comments);
        console.log("Date de création:", properties.creationDate);
        console.log("Dernier auteur:", properties.lastAuthor);
        console.log("Dernière impression:", properties.lastPrintDate);
        console.log("Dernière sauvegarde:", properties.lastSaveTime);
        console.log("Numéro de révision:", properties.revisionNumber);
        console.log("═══════════════════════════════════");

        setStatus("✅ Métadonnées lues avec succès");
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture des métadonnées:", error);
      setStatus("❌ Erreur lors de la lecture des métadonnées");
    }
  };

  return (
    <div className="min-h-screen p-6">
      <div className="max-w-4xl mx-auto space-y-6">
        <div className="text-center space-y-2">
          <h1 className="text-4xl font-bold text-gray-900">Explorateur API Office JS</h1>
          <p className="text-gray-600">
            Découvrez les capacités de l'API Word pour créer et analyser des documents
          </p>
          {status && (
            <Badge variant="secondary" className="mt-2 text-sm">
              {status}
            </Badge>
          )}
        </div>
        <Separator className="bg-gray-500" />
        <Tabs value={activeTab} onValueChange={setActiveTab}>
          <TabsList className="grid w-full grid-cols-2 bg-gray-100 rounded-md">
            <TabsTrigger value="explorer" className="data-[state=active]:bg-white">
              <Eye className="w-4 h-4 mr-2" />
              Explorateur API
            </TabsTrigger>
            <TabsTrigger value="variables" className="data-[state=active]:bg-white">
              <Variable className="w-4 h-4 mr-2" />
              Variables
            </TabsTrigger>
          </TabsList>

          {/* Onglet Explorateur */}
          <TabsContent value="explorer" className="space-y-6">
            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <FileText className="w-5 h-5" />
                  Créer un Document de Démonstration
                </CardTitle>
                <CardDescription>
                  Générez automatiquement un document Word riche avec différents styles, tableaux et
                  listes
                </CardDescription>
              </CardHeader>
              <CardContent>
                <Button onClick={createDemoDocument} className="w-full" size="lg">
                  <FileText className="w-4 h-4 mr-2" />
                  Générer le Document
                </Button>
              </CardContent>
            </Card>

            <Card>
              <CardHeader>
                <CardTitle className="flex items-center gap-2">
                  <Eye className="w-5 h-5" />
                  Analyser le Contenu
                </CardTitle>
                <CardDescription>
                  Explorez différentes façons de lire et analyser le contenu de votre document
                  (résultats dans la console)
                </CardDescription>
              </CardHeader>
              <CardContent className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <Button onClick={readSelectedText} variant="outline">
                  <Type className="w-4 h-4 mr-2" />
                  Texte Sélectionné
                </Button>
                <Button onClick={readAllContent} variant="outline">
                  <FileText className="w-4 h-4 mr-2" />
                  Document Complet
                </Button>
                <Button onClick={readParagraphs} variant="outline">
                  <List className="w-4 h-4 mr-2" />
                  Paragraphes
                </Button>
                <Button onClick={readTables} variant="outline">
                  <Table className="w-4 h-4 mr-2" />
                  Tableaux
                </Button>
                <Button onClick={readMetadata} variant="outline" className="md:col-span-2">
                  <Palette className="w-4 h-4 mr-2" />
                  Métadonnées
                </Button>
              </CardContent>
            </Card>

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
          </TabsContent>

          {/* Onglet Variables */}
          <TabsContent value="variables" className="space-y-6">
            {/* Configuration des patterns */}
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
                      onChange={(e) => setStartPattern(e.target.value)}
                      placeholder="Ex: ["
                    />
                  </div>
                  <div className="space-y-2">
                    <Label htmlFor="end-pattern">Pattern de fin</Label>
                    <Input
                      id="end-pattern"
                      value={endPattern}
                      onChange={(e) => setEndPattern(e.target.value)}
                      placeholder="Ex: ]"
                    />
                  </div>
                </div>

                <div className="space-y-2">
                  <Label>Patterns prédéfinis</Label>
                  <div className="grid grid-cols-2 gap-2">
                    <Button variant="outline" size="sm" onClick={() => setPresetPattern("[", "]")}>
                      [ ] Crochets
                    </Button>
                    <Button
                      variant="outline"
                      size="sm"
                      onClick={() => setPresetPattern("{{", "}}")}
                    >
                      {"{{ }}"} Double accolades
                    </Button>
                    <Button variant="outline" size="sm" onClick={() => setPresetPattern("{", "}")}>
                      {"{ }"} Accolades
                    </Button>
                    <Button variant="outline" size="sm" onClick={() => setPresetPattern("<", ">")}>
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
                    onCheckedChange={setSavePatterns}
                  />
                </div>

                <Button onClick={extractVariables} className="w-full" size="lg">
                  <Search className="w-4 h-4 mr-2" />
                  Extraire les Variables
                </Button>
              </CardContent>
            </Card>

            {/* Résultats */}
            {variables.length > 0 && (
              <>
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
                      <Button
                        onClick={highlightAllVariables}
                        variant="secondary"
                        className="flex-1"
                      >
                        <Highlighter className="w-4 h-4 mr-2" />
                        Tout Surligner
                      </Button>
                      <Button onClick={removeHighlights} variant="outline" className="flex-1">
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
                                onClick={() => searchVariable(variable.name)}
                                title="Rechercher"
                              >
                                <Search className="w-4 h-4" />
                              </Button>
                              <Button
                                size="sm"
                                variant="ghost"
                                onClick={() => highlightVariable(variable.name)}
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
                      <Button onClick={copyVariablesList} variant="outline" className="flex-1">
                        <Copy className="w-4 h-4 mr-2" />
                        Copier la Liste
                      </Button>
                      <Button onClick={exportToJSON} variant="outline" className="flex-1">
                        <Download className="w-4 h-4 mr-2" />
                        Export JSON
                      </Button>
                    </div>
                  </CardContent>
                </Card>

                {/* Remplacement de variables */}
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
                        onChange={(e) => setSelectedVariable(e.target.value)}
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
                        onChange={(e) => setReplacementValue(e.target.value)}
                        placeholder="Entrez la valeur de remplacement"
                      />
                    </div>

                    <Button onClick={replaceVariable} className="w-full">
                      <RefreshCw className="w-4 h-4 mr-2" />
                      Remplacer Tout
                    </Button>
                  </CardContent>
                </Card>
              </>
            )}

            {/* Message si pas de variables */}
            {variables.length === 0 && (
              <Empty className="border border-dashed">
                <EmptyHeader>
                  <EmptyMedia variant="icon">
                    <AlertCircle />
                  </EmptyMedia>
                  <EmptyTitle>Aucune variable trouvée</EmptyTitle>
                  <EmptyDescription className="text-sm text-gray-500">
                    Configurez vos patterns ci-dessus et cliquez sur "Extraire les Variables" pour
                    analyser votre document. Exemple : avec les patterns [ et ], la variable [nom]
                    sera détectée.
                  </EmptyDescription>
                </EmptyHeader>
              </Empty>
            )}
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
};

export default App;
