import * as React from "react";

export const useDocumentExplorer = (onStatusChange: (status: string) => void) => {
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

        const heading6 = body.insertParagraph("6. Système de Variables", Word.InsertLocation.end);
        heading6.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading6.font.color = "#1e40af";

        body.insertParagraph(
          "Ce document contient des exemples de variables que vous pouvez extraire et manipuler via l'onglet 'Variables' :",
          Word.InsertLocation.end
        );

        const varPara1 = body.insertParagraph(
          "Bonjour [prenom] [nom], bienvenue dans notre système !",
          Word.InsertLocation.end
        );
        varPara1.font.size = 14;

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

        body.insertParagraph("", Word.InsertLocation.end);
        const instructionsPara = body.insertParagraph(
          "💡 Astuce : Allez dans l'onglet 'Variables' pour extraire et manipuler toutes ces variables automatiquement !",
          Word.InsertLocation.end
        );
        instructionsPara.font.italic = true;
        instructionsPara.font.color = "#6366f1";

        // Section 7: Content Controls
        body.insertParagraph("", Word.InsertLocation.end);
        const heading7 = body.insertParagraph(
          "7. Content Controls Interactifs",
          Word.InsertLocation.end
        );
        heading7.styleBuiltIn = Word.BuiltInStyleName.heading1;
        heading7.font.color = "#1e40af";

        body.insertParagraph(
          "Les Content Controls permettent de créer des zones de contenu structurées et contrôlées. Voici des exemples :",
          Word.InsertLocation.end
        );

        // Exemple 1: Rich Text Content Control
        body.insertParagraph("", Word.InsertLocation.end);
        const richTextLabel = body.insertParagraph(
          "Exemple 1 - Rich Text Control :",
          Word.InsertLocation.end
        );
        richTextLabel.font.bold = true;
        richTextLabel.font.size = 12;

        const richTextPara = body.insertParagraph(
          "Cliquez ici pour entrer du texte enrichi avec formatage",
          Word.InsertLocation.end
        );
        const richTextRange = richTextPara.getRange();
        const richTextCC = richTextRange.insertContentControl(Word.ContentControlType.richText);
        richTextCC.title = "Description du produit";
        richTextCC.tag = "product_description";
        richTextCC.appearance = Word.ContentControlAppearance.tags;
        richTextCC.color = "#3b82f6";
        richTextCC.placeholderText = "Entrez une description détaillée du produit...";
        await context.sync();

        // Exemple 2: Plain Text Content Control
        body.insertParagraph("", Word.InsertLocation.end);
        const plainTextLabel = body.insertParagraph(
          "Exemple 2 - Plain Text Control (sans formatage) :",
          Word.InsertLocation.end
        );
        plainTextLabel.font.bold = true;
        plainTextLabel.font.size = 12;

        const plainTextPara = body.insertParagraph("Nom du client", Word.InsertLocation.end);
        const plainTextRange = plainTextPara.getRange();
        const plainTextCC = plainTextRange.insertContentControl(Word.ContentControlType.richText);
        plainTextCC.title = "Nom du client";
        plainTextCC.tag = "customer_name";
        plainTextCC.appearance = Word.ContentControlAppearance.tags;
        plainTextCC.color = "#10b981";
        plainTextCC.placeholderText = "Entrez le nom complet du client";
        await context.sync();

        // Exemple 3: Content Control verrouillé
        body.insertParagraph("", Word.InsertLocation.end);
        const lockedLabel = body.insertParagraph(
          "Exemple 3 - Content Control protégé (ne peut pas être supprimé) :",
          Word.InsertLocation.end
        );
        lockedLabel.font.bold = true;
        lockedLabel.font.size = 12;

        const lockedPara = body.insertParagraph(
          "Ce contenu est protégé et ne peut pas être supprimé",
          Word.InsertLocation.end
        );
        const lockedRange = lockedPara.getRange();
        const lockedCC = lockedRange.insertContentControl(Word.ContentControlType.richText);
        lockedCC.title = "Clause légale";
        lockedCC.tag = "legal_clause";
        lockedCC.appearance = Word.ContentControlAppearance.boundingBox;
        lockedCC.color = "#ef4444";
        lockedCC.cannotDelete = true;
        lockedCC.placeholderText = "Texte de la clause légale";
        await context.sync();

        // Exemple 4: Combo Box avec données
        body.insertParagraph("", Word.InsertLocation.end);
        const comboLabel = body.insertParagraph(
          "Exemple 4 - Combo Box (avec choix prédéfinis) :",
          Word.InsertLocation.end
        );
        comboLabel.font.bold = true;
        comboLabel.font.size = 12;

        const comboPara = body.insertParagraph("Sélectionnez une priorité", Word.InsertLocation.end);
        const comboRange = comboPara.getRange();
        const comboCC = comboRange.insertContentControl(Word.ContentControlType.comboBox);
        comboCC.title = "Niveau de priorité";
        comboCC.tag = "priority_level";
        comboCC.appearance = Word.ContentControlAppearance.boundingBox;
        comboCC.color = "#f59e0b";
        await context.sync();

        // Exemple 5: Date Picker
        body.insertParagraph("", Word.InsertLocation.end);
        const dateLabel = body.insertParagraph(
          "Exemple 5 - Date Picker :",
          Word.InsertLocation.end
        );
        dateLabel.font.bold = true;
        dateLabel.font.size = 12;

        const datePara = body.insertParagraph(
          "Sélectionnez une date d'échéance",
          Word.InsertLocation.end
        );
        const dateRange = datePara.getRange();
        const dateCC = dateRange.insertContentControl(Word.ContentControlType.datePicker);
        dateCC.title = "Date d'échéance";
        dateCC.tag = "due_date";
        dateCC.appearance = Word.ContentControlAppearance.boundingBox;
        dateCC.color = "#8b5cf6";
        await context.sync();

        // Exemple 6: CheckBox
        body.insertParagraph("", Word.InsertLocation.end);
        const checkboxLabel = body.insertParagraph(
          "Exemple 6 - Check Box :",
          Word.InsertLocation.end
        );
        checkboxLabel.font.bold = true;
        checkboxLabel.font.size = 12;

        const checkboxPara = body.insertParagraph(
          "J'accepte les conditions générales",
          Word.InsertLocation.end
        );
        const checkboxRange = checkboxPara.getRange();
        const checkboxCC = checkboxRange.insertContentControl(Word.ContentControlType.checkBox);
        checkboxCC.title = "Acceptation CGV";
        checkboxCC.tag = "terms_accepted";
        checkboxCC.appearance = Word.ContentControlAppearance.hidden;
        await context.sync();

        // Instructions pour les Content Controls
        body.insertParagraph("", Word.InsertLocation.end);
        const ccInstructionsPara = body.insertParagraph(
          "💡 Astuce : Allez dans l'onglet 'Content Controls' pour voir la liste complète des Content Controls de ce document et les gérer !",
          Word.InsertLocation.end
        );
        ccInstructionsPara.font.italic = true;
        ccInstructionsPara.font.color = "#8b5cf6";

        await context.sync();
        onStatusChange("✅ Document de démonstration créé avec succès !");
        console.log("✅ Document de démonstration créé avec succès !");
      });
    } catch (error) {
      console.error("❌ Erreur lors de la création du document:", error);
      onStatusChange("❌ Erreur lors de la création du document");
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

        onStatusChange(
          `✅ Texte sélectionné: "${selection.text.substring(0, 50)}${selection.text.length > 50 ? "..." : ""}"`
        );
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture du texte sélectionné:", error);
      onStatusChange("❌ Erreur - Veuillez sélectionner du texte");
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

        onStatusChange(`✅ Document lu: ${body.text.length} caractères`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture du document:", error);
      onStatusChange("❌ Erreur lors de la lecture");
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

        onStatusChange(`✅ ${paragraphs.items.length} paragraphes analysés`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture des paragraphes:", error);
      onStatusChange("❌ Erreur lors de la lecture des paragraphes");
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

        onStatusChange(`✅ ${tables.items.length} tableau(x) analysé(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture des tableaux:", error);
      onStatusChange("❌ Erreur lors de la lecture des tableaux");
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

        onStatusChange("✅ Métadonnées lues avec succès");
      });
    } catch (error) {
      console.error("❌ Erreur lors de la lecture des métadonnées:", error);
      onStatusChange("❌ Erreur lors de la lecture des métadonnées");
    }
  };

  return {
    createDemoDocument,
    readSelectedText,
    readAllContent,
    readParagraphs,
    readTables,
    readMetadata,
  };
};
