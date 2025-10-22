import * as React from "react";
import { Button } from "../../components/ui/button";
import {
  Card,
  CardContent,
  CardDescription,
  CardHeader,
  CardTitle,
} from "../../components/ui/card";
import { Separator } from "../../components/ui/separator";
import { Badge } from "../../components/ui/badge";
import { FileText, Eye, Palette, Table, List, Type } from "lucide-react";

const App: React.FC = () => {
  const [status, setStatus] = React.useState<string>("");

  // Fonction pour créer un document de démonstration riche
  const createDemoDocument = async () => {
    try {
      await Word.run(async (context) => {
        const body = context.document.body;
        body.clear();

        // Titre principal
        const title = body.insertParagraph(
          "Guide de Découverte de l'API Office JS pour Word",
          Word.InsertLocation.start
        );
        title.styleBuiltIn = Word.BuiltInStyleName.title;
        title.font.color = "#2563eb";
        title.font.size = 28;

        // Sous-titre
        const subtitle = body.insertParagraph(
          "Explorez les capacités de formatage et de manipulation de texte",
          Word.InsertLocation.end
        );
        subtitle.styleBuiltIn = Word.BuiltInStyleName.subtitle;
        subtitle.font.color = "#64748b";

        // Section 1: Formatage de texte
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

        // Section 2: Listes
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

        body.insertParagraph("", Word.InsertLocation.end); // Espace

        body.insertParagraph("Liste numérotée des étapes :", Word.InsertLocation.end);

        const num1 = body.insertParagraph("Ouvrir Word", Word.InsertLocation.end);
        const numberedList = num1.startNewList();
        numberedList.load("id");
        await context.sync();

        const num2 = body.insertParagraph("Lancer le Add-in", Word.InsertLocation.end);
        num2.attachToList(numberedList.id, 0);

        const num3 = body.insertParagraph("Explorer les fonctionnalités", Word.InsertLocation.end);
        num3.attachToList(numberedList.id, 0);

        // Section 3: Tableaux
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

        // Colorer l'en-tête
        const headerRow = table.rows.getFirst();
        headerRow.font.bold = true;
        headerRow.font.color = "#ffffff";
        headerRow.shadingColor = "#2563eb";

        // Section 4: Différentes tailles de police
        const heading4 = body.insertParagraph("4. Tailles de Police", Word.InsertLocation.end);
        heading4.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading4.font.color = "#1e40af";

        const size12 = body.insertParagraph("Texte en taille 12", Word.InsertLocation.end);
        size12.font.size = 12;

        const size16 = body.insertParagraph("Texte en taille 16", Word.InsertLocation.end);
        size16.font.size = 16;

        const size20 = body.insertParagraph("Texte en taille 20", Word.InsertLocation.end);
        size20.font.size = 20;

        // Section 5: Couleurs
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

        // Conclusion
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

  // Fonction pour lire le texte sélectionné
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

  // Fonction pour lire tout le contenu du document
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

  // Fonction pour lire les paragraphes
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

  // Fonction pour lire les tableaux
  const readTables = async () => {
    try {
      await Word.run(async (context) => {
        const tables = context.document.body.tables;
        tables.load("items,rowCount,columnCount,values");

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

  // Fonction pour lire les métadonnées
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
    <div className="min-h-screen bg-gradient-to-br from-blue-50 to-indigo-100 p-6">
      <div className="max-w-4xl mx-auto space-y-6">
        {/* En-tête */}
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

        <Separator />

        {/* Section Création */}
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

        {/* Section Lecture */}
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

        {/* Instructions */}
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
    </div>
  );
};

export default App;
