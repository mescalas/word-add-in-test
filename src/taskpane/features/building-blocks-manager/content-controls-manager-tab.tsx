import * as React from "react";
import { useContentControlsManager } from "./hooks/use-content-controls-manager";
import { ContentControlsList } from "./components/content-controls-list";
import { CreateContentControlForm } from "./components/create-content-control-form";
import { ContentControlDetails } from "./components/content-control-details";
import { Alert, AlertDescription, AlertTitle } from "@/components/ui/alert";
import { Info } from "lucide-react";
import { ContentControlInfo } from "./types";

interface ContentControlsManagerTabProps {
  onStatusChange: (status: string) => void;
}

export const ContentControlsManagerTab: React.FC<ContentControlsManagerTabProps> = ({
  onStatusChange,
}) => {
  const {
    contentControls,
    isLoading,
    listContentControls,
    createContentControl,
    deleteContentControl,
    selectContentControl,
    setContentControlText,
  } = useContentControlsManager(onStatusChange);

  const [selectedControl, setSelectedControl] = React.useState<ContentControlInfo | null>(null);

  React.useEffect(() => {
    const timer = setTimeout(() => {
      listContentControls();
    }, 300);
    return () => clearTimeout(timer);
  }, []);

  const handleEdit = (control: ContentControlInfo) => {
    setSelectedControl(control);
  };

  const handleUpdateText = async (controlId: number, newText: string) => {
    await setContentControlText(controlId, newText);
    setSelectedControl(null);
  };

  return (
    <div className="space-y-6">
      <Alert variant="default" className="bg-blue-50 border-blue-200">
        <Info className="h-4 w-4 text-blue-600" />
        <AlertTitle className="text-blue-900">Content Controls API</AlertTitle>
        <AlertDescription className="text-blue-800 text-sm">
          Les Content Controls permettent de créer des zones de contenu structurées et réutilisables
          dans vos documents Word. Ils supportent différents types : texte riche, texte simple,
          cases à cocher, listes déroulantes, sélecteurs de date, et bien plus !
        </AlertDescription>
      </Alert>

      <ContentControlsList
        controls={contentControls}
        isLoading={isLoading}
        onRefresh={listContentControls}
        onSelect={selectContentControl}
        onDelete={deleteContentControl}
        onEdit={handleEdit}
      />

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <CreateContentControlForm onCreate={createContentControl} />
        <ContentControlDetails control={selectedControl} onUpdateText={handleUpdateText} />
      </div>
    </div>
  );
};
