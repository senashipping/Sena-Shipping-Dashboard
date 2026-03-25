import React, { useState } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Input } from "../../components/ui/input";
import { Label } from "../../components/ui/label";
import { Alert, AlertDescription } from "../../components/ui/alert";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../../components/ui/table";
import { Badge } from "../../components/ui/badge";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "../../components/ui/dialog";
import { ConfirmationDialog } from "../../components/ui/confirmation-dialog";
import { Plus, Search, Edit, Trash2 } from "lucide-react";
import { useClientSearch } from "../../hooks/useDebounce";
import { useAuth } from "../../contexts/AuthContext";

const AdminShips: React.FC = () => {
  const { isSuperAdmin } = useAuth();
  const [isCreateDialogOpen, setIsCreateDialogOpen] = useState(false);
  const [editingShip, setEditingShip] = useState<any>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    isOpen: boolean;
    ship: any | null;
  }>({ isOpen: false, ship: null });

  const { data: shipsData, isLoading, refetch } = useQuery({
    queryKey: ["ships"],
    queryFn: () => api.getShips({}),
  });

  const createShipMutation = useMutation({
    mutationFn: (shipData: any) => api.createShip(shipData),
    onSuccess: () => {
      setIsCreateDialogOpen(false);
      setEditingShip(null);
      refetch();
    },
    onError: (error) => {
      console.error('Create ship error:', error);
    },
  });

  const updateShipMutation = useMutation({
    mutationFn: (data: any) => api.updateShip(editingShip._id, data),
    onSuccess: () => {
      setIsCreateDialogOpen(false);
      setEditingShip(null);
      refetch();
    },
    onError: (error) => {
      console.error('Update ship error:', error);
    },
  });

  const deleteShipMutation = useMutation({
    mutationFn: (id: string) => api.deleteShip(id),
    onSuccess: () => {
      refetch();
      setDeleteConfirmation({ isOpen: false, ship: null });
    },
  });

  const handleCreateShip = async (e: React.FormEvent) => {
    e.preventDefault();
    const formData = new FormData(e.target as HTMLFormElement);
    const shipData = {
      name: formData.get("name"),
      imoNumber: formData.get("imoNumber"),
      vesselType: formData.get("vesselType"),
      flag: formData.get("flag"),
      grossTonnage: Number(formData.get("grossTonnage")),
      buildYear: Number(formData.get("buildYear")),
      owner: formData.get("owner"),
      operator: formData.get("operator"),
    };

    createShipMutation.mutate(shipData);
  };

  const handleUpdateShip = async (e: React.FormEvent) => {
    e.preventDefault();
    const formData = new FormData(e.target as HTMLFormElement);
    const shipData = {
      name: formData.get("name"),
      imoNumber: formData.get("imoNumber"),
      vesselType: formData.get("vesselType"),
      flag: formData.get("flag"),
      grossTonnage: Number(formData.get("grossTonnage")),
      buildYear: Number(formData.get("buildYear")),
      owner: formData.get("owner"),
      operator: formData.get("operator"),
    };

    updateShipMutation.mutate(shipData);
  };

  const handleDeleteShip = (ship: any) => {
    setDeleteConfirmation({ isOpen: true, ship });
  };

  const confirmDeleteShip = () => {
    if (deleteConfirmation.ship) {
      deleteShipMutation.mutate(deleteConfirmation.ship._id);
    }
  };

  const rawShips = shipsData?.data?.data || [];
  
  // Apply client-side search including user name and email
  const searchFields = ['name', 'imoNumber', 'user.name', 'user.email'];
  const filteredShips = useClientSearch(rawShips, searchTerm, searchFields);

  return (
    <div className="space-y-4 sm:space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <h1 className="text-xl font-bold sm:text-3xl text-sena-navy dark:text-white">Ship Management</h1>
        <Button onClick={() => {
          setEditingShip(null);
          setIsCreateDialogOpen(true);
        }} className="self-start bg-sena-gold hover:bg-sena-gold/90">
          <Plus className="w-4 h-4 mr-2" />
          <span className="hidden sm:inline">Add Ship</span>
          <span className="sm:hidden">Add</span>
        </Button>
      </div>

            <Card className="border-sena-lightBlue/20">
        <CardHeader>
          <CardTitle className="text-sena-navy">Ships</CardTitle>
          <CardDescription className="text-sena-lightBlue">
            Manage ship information and assignments</CardDescription>
        </CardHeader>
        <CardContent className="p-0 sm:p-6">
          <div className="flex items-center mb-4 space-x-2">
            <div className="relative flex-1">
              <Search className="absolute w-4 h-4 left-3 top-3" />
              <Input
                placeholder="Search ships by name, IMO, or assigned user..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="pl-9"
              />
            </div>
          </div>

          {isLoading ? (
            <div className="flex items-center justify-center h-32">
              <div className="w-8 h-8 border-b-2 rounded-full animate-spin border-primary"></div>
            </div>
          ) : (
            <Table>
              <TableHeader>
                <TableRow>
                  <TableHead>Name</TableHead>
                  <TableHead>IMO Number</TableHead>
                  <TableHead>Vessel Type</TableHead>
                  <TableHead>Flag</TableHead>
                  <TableHead>Owner</TableHead>
                  <TableHead>Assigned User</TableHead>
                  <TableHead>Status</TableHead>
                  {isSuperAdmin() && <TableHead>Actions</TableHead>}
                </TableRow>
              </TableHeader>
              <TableBody>
                {filteredShips.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={isSuperAdmin() ? 8 : 7} className="py-8 text-center text-gray-500">
                      No ships found
                    </TableCell>
                  </TableRow>
                ) : (
                  filteredShips.map((ship: any) => (
                  <TableRow key={ship._id}>
                    <TableCell className="font-medium">{ship.name}</TableCell>
                    <TableCell>{ship.imoNumber}</TableCell>
                    <TableCell>{ship.vesselType || "-"}</TableCell>
                    <TableCell>{ship.flag || "-"}</TableCell>
                    <TableCell>{ship.owner || "-"}</TableCell>
                    <TableCell>
                      {ship.user ? (
                        <div className="flex flex-col">
                          <span className="text-sm font-medium">{ship.user.name}</span>
                          <span className="text-xs">{ship.user.email}</span>
                        </div>
                      ) : (
                        <Badge variant="outline" className="text-xs">
                          Unassigned
                        </Badge>
                      )}
                    </TableCell>
                    <TableCell>
                      <Badge variant={ship.isActive ? "default" : "destructive"}>
                        {ship.isActive ? "Active" : "Inactive"}
                      </Badge>
                    </TableCell>
                    {isSuperAdmin() && (
                      <TableCell>
                        <div className="flex space-x-2">
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => {
                              setEditingShip(ship);
                              setIsCreateDialogOpen(true);
                            }}
                          >
                            <Edit className="w-4 h-4" />
                          </Button>
                          <Button
                            variant="outline"
                            size="sm"
                            onClick={() => handleDeleteShip(ship)}
                          >
                            <Trash2 className="w-4 h-4" />
                          </Button>
                        </div>
                      </TableCell>
                    )}
                  </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
          )}
        </CardContent>
      </Card>

      <Dialog open={isCreateDialogOpen} onOpenChange={setIsCreateDialogOpen}>
        <DialogContent className="max-w-2xl max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>{editingShip ? "Edit Ship" : "Add New Ship"}</DialogTitle>
          </DialogHeader>
          <form onSubmit={editingShip ? handleUpdateShip : handleCreateShip} className="space-y-4">
            {createShipMutation.error ? (
              <Alert variant="destructive">
                <AlertDescription>
                  {String(createShipMutation.error instanceof Error 
                    ? createShipMutation.error.message 
                    : "Failed to create ship")}
                </AlertDescription>
              </Alert>
            ) : null}
            
            {updateShipMutation.error ? (
              <Alert variant="destructive">
                <AlertDescription>
                  {String(updateShipMutation.error instanceof Error 
                    ? updateShipMutation.error.message 
                    : "Failed to update ship")}
                </AlertDescription>
              </Alert>
            ) : null}
            
            <div className="grid grid-cols-1 gap-4 md:grid-cols-2">
              <div className="space-y-2">
                <Label htmlFor="name">Ship Name</Label>
                <Input
                  id="name"
                  name="name"
                  defaultValue={editingShip?.name || ""}
                  required
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="imoNumber">IMO Number</Label>
                <Input
                  id="imoNumber"
                  name="imoNumber"
                  defaultValue={editingShip?.imoNumber || ""}
                  required
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="vesselType">Vessel Type</Label>
                <Input
                  id="vesselType"
                  name="vesselType"
                  defaultValue={editingShip?.vesselType || ""}
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="flag">Flag</Label>
                <Input
                  id="flag"
                  name="flag"
                  defaultValue={editingShip?.flag || ""}
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="grossTonnage">Gross Tonnage</Label>
                <Input
                  id="grossTonnage"
                  name="grossTonnage"
                  type="number"
                  defaultValue={editingShip?.grossTonnage || ""}
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="buildYear">Build Year</Label>
                <Input
                  id="buildYear"
                  name="buildYear"
                  type="number"
                  defaultValue={editingShip?.buildYear || ""}
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="owner">Owner</Label>
                <Input
                  id="owner"
                  name="owner"
                  defaultValue={editingShip?.owner || ""}
                />
              </div>
              
              <div className="space-y-2">
                <Label htmlFor="operator">Operator</Label>
                <Input
                  id="operator"
                  name="operator"
                  defaultValue={editingShip?.operator || ""}
                />
              </div>
            </div>
            
            <Button type="submit" className="w-full">
              {editingShip ? "Update Ship" : "Create Ship"}
            </Button>
          </form>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <ConfirmationDialog
        isOpen={deleteConfirmation.isOpen}
        onClose={() => setDeleteConfirmation({ isOpen: false, ship: null })}
        onConfirm={confirmDeleteShip}
        title="Delete Ship"
        description={`Are you sure you want to delete "${deleteConfirmation.ship?.name}" (IMO: ${deleteConfirmation.ship?.imoNumber})? This action cannot be undone. The ship will be permanently removed${deleteConfirmation.ship?.user ? ' and the assigned user will lose their ship assignment' : ''}.`}
        confirmText="Delete Ship"
        isLoading={deleteShipMutation.isPending}
      />
    </div>
  );
};

export default AdminShips;