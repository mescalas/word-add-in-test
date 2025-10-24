import * as React from "react";
import { ContentControlInfo, CreateContentControlParams } from "../types";

export const useContentControlsManager = (onStatusChange: (status: string) => void) => {
  const [contentControls, setContentControls] = React.useState<ContentControlInfo[]>([]);
  const [isLoading, setIsLoading] = React.useState(false);

  /**
   * Liste tous les Content Controls dans le document
   */
  const listContentControls = async () => {
    setIsLoading(true);
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.getContentControls();
        console.warn('document', context.document);
        await context.sync();
        contentControls.load("items");
        await context.sync();

        console.log(`[DEBUG] Nombre de Content Controls trouvés: ${contentControls.items.length}`);

        const loadPromises = contentControls.items.map((cc, index) => {
          console.log(`[DEBUG] Chargement du Content Control ${index + 1}`);
          cc.load(
            "id,type,title,tag,appearance,cannotDelete,cannotEdit,color,placeholderText,text"
          );
        });
        await context.sync();

        console.log(`[DEBUG] Propriétés chargées pour ${contentControls.items.length} Content Controls`);

        const controlInfos: ContentControlInfo[] = [];

        contentControls.items.forEach((cc, index) => {
          try {
            const info: ContentControlInfo = {
              id: cc.id,
              title: cc.title || "(Sans titre)",
              tag: cc.tag || "",
              type: cc.type,
              appearance: cc.appearance,
              cannotDelete: cc.cannotDelete,
              cannotEdit: cc.cannotEdit,
              color: cc.color || "#000000",
              placeholderText: cc.placeholderText || "",
              text: (cc.text || "").substring(0, 100),
            };
            controlInfos.push(info);
            console.log(`[DEBUG] Content Control ${index + 1} ajouté: ${info.title} (${info.type})`);
          } catch (err) {
            console.error(`[DEBUG] Erreur lors du traitement du Content Control ${index + 1}:`, err);
          }
        });

        setContentControls(controlInfos);

        console.log("═══════════════════════════════════");
        console.log("📋 CONTENT CONTROLS");
        console.log("═══════════════════════════════════");
        console.log(`Nombre total: ${controlInfos.length}`);
        controlInfos.forEach((control, index) => {
          console.log(`\n--- Content Control ${index + 1} ---`);
          console.log(`ID: ${control.id}`);
          console.log(`Titre: ${control.title}`);
          console.log(`Tag: ${control.tag}`);
          console.log(`Type: ${control.type}`);
          console.log(`Appearance: ${control.appearance}`);
          console.log(`Cannot Delete: ${control.cannotDelete}`);
          console.log(`Cannot Edit: ${control.cannotEdit}`);
          console.log(`Color: ${control.color}`);
          console.log(`Placeholder: ${control.placeholderText}`);
          console.log(`Text: ${control.text}`);
        });
        console.log("═══════════════════════════════════");

        onStatusChange(`✅ ${controlInfos.length} Content Control(s) trouvé(s)`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la liste des content controls:", error);
      onStatusChange("❌ Erreur lors de la recherche");
      setContentControls([]);
    } finally {
      setIsLoading(false);
    }
  };

  /**
   * Crée un nouveau Content Control à partir de la sélection
   */
  const createContentControl = async (params: CreateContentControlParams) => {
    try {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();

        if (!selection.text || selection.text.trim() === "") {
          onStatusChange("❌ Veuillez sélectionner du texte pour créer un Content Control");
          return;
        }

        const contentControl = selection.insertContentControl(params.type as any);
        contentControl.title = params.title;
        contentControl.tag = params.tag;

        if (params.appearance) {
          contentControl.appearance = params.appearance;
        }

        if (params.color) {
          contentControl.color = params.color;
        }

        if (params.placeholderText) {
          contentControl.placeholderText = params.placeholderText;
        }

        if (params.cannotDelete !== undefined) {
          contentControl.cannotDelete = params.cannotDelete;
        }

        if (params.cannotEdit !== undefined) {
          contentControl.cannotEdit = params.cannotEdit;
        }

        await context.sync();

        console.log(`✅ Content Control "${params.title}" créé avec succès`);
        onStatusChange(`✅ Content Control "${params.title}" créé`);

        await listContentControls();
      });
    } catch (error) {
      console.error("❌ Erreur lors de la création du Content Control:", error);
      onStatusChange("❌ Erreur lors de la création");
    }
  };

  /**
   * Supprime un Content Control par son ID
   */
  const deleteContentControl = async (controlId: number) => {
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        await context.sync();

        contentControls.items.forEach((cc) => {
          cc.load("id");
        });
        await context.sync();

        const targetControl = contentControls.items.find((cc) => cc.id === controlId);

        if (!targetControl) {
          onStatusChange(`❌ Content Control ID ${controlId} introuvable`);
          return;
        }

        targetControl.delete(true);
        await context.sync();

        console.log(`✅ Content Control ID ${controlId} supprimé avec succès`);
        onStatusChange(`✅ Content Control supprimé`);

        await listContentControls();
      });
    } catch (error) {
      console.error("❌ Erreur lors de la suppression:", error);
      onStatusChange("❌ Erreur lors de la suppression");
    }
  };

  /**
   * Met à jour les propriétés d'un Content Control
   */
  const updateContentControl = async (
    controlId: number,
    updates: Partial<CreateContentControlParams>
  ) => {
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        await context.sync();

        contentControls.items.forEach((cc) => {
          cc.load("id");
        });
        await context.sync();

        const targetControl = contentControls.items.find((cc) => cc.id === controlId);

        if (!targetControl) {
          onStatusChange(`❌ Content Control ID ${controlId} introuvable`);
          return;
        }

        if (updates.title !== undefined) {
          targetControl.title = updates.title;
        }
        if (updates.tag !== undefined) {
          targetControl.tag = updates.tag;
        }
        if (updates.appearance !== undefined) {
          targetControl.appearance = updates.appearance;
        }
        if (updates.color !== undefined) {
          targetControl.color = updates.color;
        }
        if (updates.placeholderText !== undefined) {
          targetControl.placeholderText = updates.placeholderText;
        }
        if (updates.cannotDelete !== undefined) {
          targetControl.cannotDelete = updates.cannotDelete;
        }
        if (updates.cannotEdit !== undefined) {
          targetControl.cannotEdit = updates.cannotEdit;
        }

        await context.sync();

        console.log(`✅ Content Control ID ${controlId} mis à jour avec succès`);
        onStatusChange(`✅ Content Control mis à jour`);

        await listContentControls();
      });
    } catch (error) {
      console.error("❌ Erreur lors de la mise à jour:", error);
      onStatusChange("❌ Erreur lors de la mise à jour");
    }
  };

  /**
   * Sélectionne un Content Control dans le document
   */
  const selectContentControl = async (controlId: number) => {
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        await context.sync();

        contentControls.items.forEach((cc) => {
          cc.load("id");
        });
        await context.sync();

        const targetControl = contentControls.items.find((cc) => cc.id === controlId);

        if (!targetControl) {
          onStatusChange(`❌ Content Control ID ${controlId} introuvable`);
          return;
        }

        targetControl.select(Word.SelectionMode.select);
        await context.sync();

        console.log(`✅ Content Control ID ${controlId} sélectionné`);
        onStatusChange(`✅ Content Control sélectionné dans le document`);
      });
    } catch (error) {
      console.error("❌ Erreur lors de la sélection:", error);
      onStatusChange("❌ Erreur lors de la sélection");
    }
  };

  /**
   * Change le contenu d'un Content Control
   */
  const setContentControlText = async (controlId: number, newText: string) => {
    try {
      await Word.run(async (context) => {
        const contentControls = context.document.contentControls;
        contentControls.load("items");
        await context.sync();

        contentControls.items.forEach((cc) => {
          cc.load("id");
        });
        await context.sync();

        const targetControl = contentControls.items.find((cc) => cc.id === controlId);

        if (!targetControl) {
          onStatusChange(`❌ Content Control ID ${controlId} introuvable`);
          return;
        }

        targetControl.insertText(newText, Word.InsertLocation.replace);
        await context.sync();

        console.log(`✅ Texte du Content Control ID ${controlId} modifié`);
        onStatusChange(`✅ Texte modifié avec succès`);

        await listContentControls();
      });
    } catch (error) {
      console.error("❌ Erreur lors de la modification du texte:", error);
      onStatusChange("❌ Erreur lors de la modification");
    }
  };

  return {
    contentControls,
    isLoading,
    listContentControls,
    createContentControl,
    deleteContentControl,
    updateContentControl,
    selectContentControl,
    setContentControlText,
  };
};
