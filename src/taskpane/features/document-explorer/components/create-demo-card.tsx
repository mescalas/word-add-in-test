import * as React from "react";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { FileText } from "lucide-react";

interface CreateDemoCardProps {
  onCreate: () => void;
}

export const CreateDemoCard: React.FC<CreateDemoCardProps> = ({ onCreate }) => {
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <FileText className="w-5 h-5" />
          Créer un Document de Démonstration
        </CardTitle>
        <CardDescription>
          Générez automatiquement un document Word riche avec différents styles, tableaux et listes
        </CardDescription>
      </CardHeader>
      <CardContent>
        <Button onClick={onCreate} className="w-full" size="lg">
          <FileText className="w-4 h-4 mr-2" />
          Générer le Document
        </Button>
      </CardContent>
    </Card>
  );
};
