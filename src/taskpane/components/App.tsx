import * as React from "react";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { Tabs, TabsContent, TabsList, TabsTrigger } from "@/components/ui/tabs";
import { Eye, Variable } from "lucide-react";
import { DocumentExplorerTab } from "../features/document-explorer";
import { VariableManagerTab } from "../features/variable-manager";

const App: React.FC = () => {
  const [status, setStatus] = React.useState<string>("");
  const [activeTab, setActiveTab] = React.useState<string>("explorer");

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

          <TabsContent value="explorer" className="space-y-6">
            <DocumentExplorerTab onStatusChange={setStatus} />
          </TabsContent>

          <TabsContent value="variables" className="space-y-6">
            <VariableManagerTab onStatusChange={setStatus} />
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
};

export default App;
