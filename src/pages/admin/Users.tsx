import React, { useState, useMemo, useEffect } from "react";
import { useQuery, useMutation } from "@tanstack/react-query";
import { Link } from "react-router-dom";
import api from "../../api";
import { Button } from "../../components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "../../components/ui/card";
import { Input } from "../../components/ui/input";
import { Label } from "../../components/ui/label";
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from "../../components/ui/select";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "../../components/ui/table";
import { Badge } from "../../components/ui/badge";
import { Dialog, DialogContent, DialogHeader, DialogTitle } from "../../components/ui/dialog";
import { ConfirmationDialog } from "../../components/ui/confirmation-dialog";
import { Plus, Search, Edit, Trash2 } from "lucide-react";
import { useClientSearch } from "../../hooks/useDebounce";
import { useAuth } from "../../contexts/AuthContext";
import { useToast } from "../../components/ui/toast";
import { getApiErrorMessage } from "../../lib/utils";

const AdminUsers: React.FC = () => {
  const { isSuperAdmin } = useAuth();
  const { toast } = useToast();
  const [isCreateDialogOpen, setIsCreateDialogOpen] = useState(false);
  const [editingUser, setEditingUser] = useState<any>(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [selectedRole, setSelectedRole] = useState<string>("user");
  const [selectedUserType, setSelectedUserType] = useState<string>("deck");
  const [selectedShip, setSelectedShip] = useState<string>("none");
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    isOpen: boolean;
    user: any | null;
  }>({ isOpen: false, user: null });

  const { data: usersData, isLoading, refetch } = useQuery({
    queryKey: ["users"],
    queryFn: () => api.getUsers({}),
  });

  // Fetch ships data when dialog is open
  const { data: shipsData, isLoading: shipsLoading } = useQuery({
    queryKey: ["ships"],
    queryFn: () => api.getShips({}),
    enabled: isCreateDialogOpen,
    retry: 1,
    retryOnMount: false,
  });

  const createUserMutation = useMutation({
    mutationFn: (userData: any) => api.createUser(userData),
    onSuccess: () => {
      setIsCreateDialogOpen(false);
      setEditingUser(null);
      refetch();
      toast({ title: "User created", variant: "success" });
    },
    onError: (error) => {
      toast({
        title: "Failed to create user",
        description: getApiErrorMessage(error, "Failed to create user"),
        variant: "destructive",
      });
    },
  });

  const updateUserMutation = useMutation({
    mutationFn: (data: any) => api.updateUser(editingUser._id, data),
    onSuccess: () => {
      setIsCreateDialogOpen(false);
      setEditingUser(null);
      refetch();
      toast({ title: "User updated", variant: "success" });
    },
    onError: (error) => {
      toast({
        title: "Failed to update user",
        description: getApiErrorMessage(error, "Failed to update user"),
        variant: "destructive",
      });
    },
  });

  const deleteUserMutation = useMutation({
    mutationFn: (id: string) => api.deleteUser(id),
    onSuccess: () => {
      refetch();
      setDeleteConfirmation({ isOpen: false, user: null });
      toast({ title: "User deleted", variant: "success" });
    },
    onError: (error) => {
      toast({
        title: "Failed to delete user",
        description: getApiErrorMessage(error, "Failed to delete user"),
        variant: "destructive",
      });
    },
  });

  const handleDeleteUser = (user: any) => {
    setDeleteConfirmation({ isOpen: true, user });
  };

  const confirmDeleteUser = () => {
    if (deleteConfirmation.user) {
      deleteUserMutation.mutate(deleteConfirmation.user._id);
    }
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const formData = new FormData(e.target as HTMLFormElement);
    const userData = Object.fromEntries(formData.entries());
    
    // Build user data explicitly using state values to ensure consistency
    const processedUserData: any = {
      name: userData.name,
      email: userData.email,
      role: selectedRole
    };

    // Add password for new users, or for existing users if super admin provides one
    if (!editingUser && userData.password) {
      processedUserData.password = userData.password;
    } else if (editingUser && userData.password) {
      const passwordValue = String(userData.password);
      if (passwordValue && passwordValue.trim() !== "") {
        // Only include password if it's provided and not empty for editing users
        processedUserData.password = passwordValue;
      }
    }

    // Add userType and shipId for regular users
    if (selectedRole === "user") {
      processedUserData.userType = selectedUserType;
      
      if (selectedShip && selectedShip !== "" && selectedShip !== "none") {
        processedUserData.shipId = selectedShip;
      }
    }
    
    // Submitting user data
    
    if (editingUser) {
      updateUserMutation.mutate(processedUserData);
    } else {
      createUserMutation.mutate(processedUserData);
    }
  };

  // Reset form state when dialog opens or editing user changes
  useEffect(() => {
    if (isCreateDialogOpen) {
      setSelectedRole(editingUser?.role || "user");
      // For existing users without userType, default to "deck"
      // For new users, also default to "deck"
      setSelectedUserType(editingUser?.userType || "deck");
      setSelectedShip(editingUser?.ship?._id ? editingUser.ship._id : "none");
    }
  }, [isCreateDialogOpen, editingUser]);

  // Clear ship assignment when role changes to admin or super_admin
  useEffect(() => {
    if (selectedRole === "admin" || selectedRole === "super_admin") {
      setSelectedShip("none");
    }
  }, [selectedRole]);

  // Reset ship selection when userType changes to ensure available ships are updated
  useEffect(() => {
    if (selectedRole === "user" && !editingUser) {
      // Only reset for new users, not when editing existing users
      setSelectedShip("none");
    }
  }, [selectedUserType, selectedRole, editingUser]);

  const rawUsers = usersData?.data?.data || [];
  const rawShips = shipsData?.data?.data || [];
  
  // Get available ships for the selected user type
  const availableShips = useMemo(() => {
    if (!rawShips || !Array.isArray(rawShips) || rawShips.length === 0) return [];
    if (!rawUsers || !Array.isArray(rawUsers)) return rawShips;
    
    // Group users by ship and userType
    const shipUserCounts = rawUsers.reduce((acc: any, user: any) => {
      if (user?.ship?._id && user.userType && user.role === "user") {
        const shipId = user.ship._id;
        if (!acc[shipId]) {
          acc[shipId] = { deck: false, engine: false };
        }
        acc[shipId][user.userType] = true;
      }
      return acc;
    }, {});
    
    // If editing a user, allow their current ship
    const currentUserShipId = editingUser?.ship?._id;
    
    return rawShips.filter((ship: any) => {
      if (!ship?._id) return false;
      
      const shipId = ship._id;
      const counts = shipUserCounts[shipId] || { deck: false, engine: false };
      
      // If editing current user's ship, it's always available
      if (shipId === currentUserShipId) return true;
      
      // Check if this userType slot is available for this ship
      return !counts[selectedUserType];
    });
  }, [rawShips, rawUsers, editingUser, selectedUserType]);
  
  // Apply client-side search
  const searchFields = ['name', 'email', 'ship.name'];
  const users = useClientSearch(rawUsers, searchTerm, searchFields);

  return (
    <div className="space-y-4 sm:space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <h1 className="text-xl font-bold sm:text-3xl text-sena-navy dark:text-white">User Management</h1>
        <Button onClick={() => { setEditingUser(null); setIsCreateDialogOpen(true); }} className="self-start bg-sena-gold hover:bg-sena-gold/90">
          <Plus className="w-4 h-4 mr-2" /> 
          <span className="hidden sm:inline">Add User</span>
          <span className="sm:hidden">Add</span>
        </Button>
      </div>

            <Card className="border-sena-lightBlue/20">
        <CardHeader>
          <CardTitle className="text-sena-navy">Users</CardTitle>
          <CardDescription className="text-sena-lightBlue">
            Manage user accounts and permissions</CardDescription>
        </CardHeader>
        <CardContent className="p-0 sm:p-6">
          <div className="p-4 sm:p-0">
            <div className="flex items-center mb-4 space-x-2">
              <div className="relative flex-1">
                <Search className="absolute w-4 h-4 left-3 top-3 " />
                <Input
                  placeholder="Search users..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="pl-9 border-sena-lightBlue/30 focus:border-sena-gold focus:ring-sena-gold/20"
                />
              </div>
            </div>

            {isLoading ? (
              <div className="p-8 text-center text-sena-navy">Loading...</div>
            ) : (
              <>
                {/* Mobile Cards View */}
                <div className="block lg:hidden">
                  <div className="space-y-4">
                    {users.length === 0 ? (
                      <Card className="border-sena-lightBlue/20">
                        <CardContent className="p-8 text-center">
                          <p className="text-sena-lightBlue">No users found</p>
                        </CardContent>
                      </Card>
                    ) : (
                      users.map((user: any) => (
                        <Card key={user._id} className="border-sena-lightBlue/20">
                          <CardContent className="p-4">
                            <div className="space-y-3">
                              <div className="flex items-start justify-between gap-2">
                                <div className="flex-1 min-w-0">
                                  <h3 className="font-medium truncate text-sena-navy">
                                    {user.name}
                                  </h3>
                                  <p className="text-sm truncate text-sena-lightBlue">
                                    {user.email}
                                  </p>
                                </div>
                                <Badge 
                                  variant={user.role === 'admin' ? 'default' : user.role === 'super_admin' ? 'destructive' : 'secondary'} 
                                  className="shrink-0"
                                >
                                  {user.role === 'super_admin' ? 'Super Admin' : user.role}
                                </Badge>
                              </div>
                              
                              <div className="space-y-2 text-sm">
                                {user.role === "user" && user.userType && (
                                  <div>
                                    <span className="text-gray-600 dark:text-gray-300">User Type:</span>
                                    <div className="inline-block ml-2">
                                      <Badge variant={user.userType === "deck" ? "outline" : "secondary"} className="text-xs">
                                        {user.userType === "deck" ? "Deck Officer" : "Engine Officer"}
                                      </Badge>
                                    </div>
                                  </div>
                                )}
                                <div>
                                  <span className="text-gray-600 dark:text-gray-300">Ship:</span>
                                  <div className="font-medium text-sena-navy dark:text-white">
                                    {user.ship?.name || "N/A"}
                                  </div>
                                </div>
                                <div>
                                  <span className="text-gray-600 dark:text-gray-300">Status:</span>
                                  <div className="inline-block ml-2">
                                    <Badge variant={user.isActive ? "default" : "destructive"} className="text-xs">
                                      {user.isActive ? "Active" : "Inactive"}
                                    </Badge>
                                  </div>
                                </div>
                              </div>
                              
                              <div className="flex space-x-2">
                                {isSuperAdmin() && (
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    onClick={() => {
                                      setEditingUser(user);
                                      setIsCreateDialogOpen(true);
                                    }}
                                    className="flex-1"
                                  >
                                    <Edit className="w-4 h-4 mr-2" />
                                    Edit
                                  </Button>
                                )}
                                {isSuperAdmin() && (
                                  <Button
                                    variant="outline"
                                    size="sm"
                                    onClick={() => handleDeleteUser(user)}
                                    className="flex-1 text-red-600 hover:text-red-700"
                                  >
                                    <Trash2 className="w-4 h-4 mr-2" />
                                    Delete
                                  </Button>
                                )}
                              </div>
                            </div>
                          </CardContent>
                        </Card>
                      ))
                    )}
                  </div>
                </div>

                {/* Desktop Table View */}
                <div className="hidden lg:block">
                  <Table>
                    <TableHeader>
                      <TableRow>
                        <TableHead className="text-sena-navy">Name</TableHead>
                        <TableHead className="text-sena-navy">Email</TableHead>
                        <TableHead className="text-sena-navy">Role</TableHead>
                        <TableHead className="text-sena-navy">User Type</TableHead>
                        <TableHead className="text-sena-navy">Ship</TableHead>
                        <TableHead className="text-sena-navy">Status</TableHead>
                        {isSuperAdmin() && <TableHead className="text-sena-navy">Actions</TableHead>}
                      </TableRow>
                    </TableHeader>
              <TableBody>
                {users.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={isSuperAdmin() ? 7 : 6} className="py-8 text-center text-gray-500">
                      No users found
                    </TableCell>
                  </TableRow>
                ) : (
                  users.map((user: any) => (
                    <TableRow key={user._id}>
                      <TableCell className="font-medium">
                        <Link to={`/admin/users/${user._id}/submissions`} className="hover:underline">
                          {user.name}
                        </Link>
                      </TableCell>
                      <TableCell>{user.email}</TableCell>
                      <TableCell>
                        <Badge 
                          variant={user.role === "admin" ? "default" : user.role === "super_admin" ? "destructive" : "secondary"}
                        >
                          {user.role === "super_admin" ? "Super Admin" : user.role}
                        </Badge>
                      </TableCell>
                      <TableCell>
                        {user.role === "user" && user.userType ? (
                          <Badge variant={user.userType === "deck" ? "outline" : "secondary"}>
                            {user.userType === "deck" ? "Deck Officer" : "Engine Officer"}
                          </Badge>
                        ) : (
                          <span className="text-gray-400">N/A</span>
                        )}
                      </TableCell>
                      <TableCell>{user.ship?.name || "N/A"}</TableCell>
                      <TableCell><Badge variant={user.isActive ? "default" : "destructive"}>{user.isActive ? "Active" : "Inactive"}</Badge></TableCell>
                      {isSuperAdmin() && (
                        <TableCell className="flex space-x-2">
                          <Button variant="outline" size="sm" onClick={() => { setEditingUser(user); setIsCreateDialogOpen(true); }}>
                            <Edit className="w-4 h-4" />
                          </Button>
                          <Button variant="destructive" size="sm" onClick={() => handleDeleteUser(user)}>
                            <Trash2 className="w-4 h-4" />
                          </Button>
                        </TableCell>
                      )}
                    </TableRow>
                  ))
                  )}
                </TableBody>
                  </Table>
                </div>
              </>
            )}
          </div>
        </CardContent>
      </Card>

      <Dialog open={isCreateDialogOpen} onOpenChange={setIsCreateDialogOpen}>
        <DialogContent className="max-w-md max-h-[90vh] overflow-y-auto">
          <DialogHeader>
            <DialogTitle>{editingUser ? "Edit User" : "Create New User"}</DialogTitle>
          </DialogHeader>
          <form onSubmit={handleSubmit} className="space-y-4">
            <div className="space-y-2">
              <Label htmlFor="name">Name</Label>
              <Input id="name" name="name" defaultValue={editingUser?.name || ""} required />
            </div>
            <div className="space-y-2">
              <Label htmlFor="email">Email</Label>
              <Input id="email" name="email" type="email" defaultValue={editingUser?.email || ""} required />
            </div>
            {(!editingUser || (editingUser && isSuperAdmin())) && (
              <div className="space-y-2">
                <Label htmlFor="password">
                  Password {editingUser ? "(Leave blank to keep current password)" : ""}
                </Label>
                <Input 
                  id="password" 
                  name="password" 
                  type="password" 
                  required={!editingUser}
                  placeholder={editingUser ? "Enter new password to change" : ""}
                />
                {editingUser && (
                  <div className="p-2 text-xs text-gray-600 bg-yellow-50 rounded-md dark:bg-yellow-900/20 dark:text-yellow-300">
                    <p><strong>Super Admin Only:</strong> You can change any user's password directly.</p>
                  </div>
                )}
              </div>
            )}
            <div className="space-y-2">
              <Label htmlFor="role">Role</Label>
              <Select name="role" value={selectedRole} onValueChange={setSelectedRole}>
                <SelectTrigger><SelectValue /></SelectTrigger>
                <SelectContent>
                  <SelectItem value="user">User</SelectItem>
                  {isSuperAdmin() && <SelectItem value="admin">Admin</SelectItem>}
                  {isSuperAdmin() && <SelectItem value="super_admin">Super Admin</SelectItem>}
                </SelectContent>
              </Select>
            </div>

            {selectedRole === "user" && (
              <div className="space-y-2">
                <Label htmlFor="userType">User Type</Label>
                <Select name="userType" value={selectedUserType} onValueChange={setSelectedUserType}>
                  <SelectTrigger><SelectValue /></SelectTrigger>
                  <SelectContent>
                    <SelectItem value="deck">Deck Officer (All Forms Except Engine)</SelectItem>
                    <SelectItem value="engine">Engine Officer (Engine Forms Only)</SelectItem>
                  </SelectContent>
                </Select>
                <div className="p-2 text-xs text-gray-600 bg-gray-50 rounded-md dark:bg-gray-800 dark:text-gray-300">
                  <p><strong>Deck Officer:</strong> Can access and submit Deck, MLC, ISPS, and Drill forms</p>
                  <p><strong>Engine Officer:</strong> Can only access and submit Engine forms</p>
                </div>
              </div>
            )}
            
            <div className="space-y-2">
              <Label htmlFor="ship">Assigned Ship</Label>
              {(selectedRole === "admin" || selectedRole === "super_admin") && (
                <div className="p-3 text-sm text-gray-600 border border-gray-200 rounded-md bg-gray-50 dark:bg-gray-800 dark:text-gray-300 dark:border-gray-700">
                  <p>⚠️ Admin and Super Admin users cannot be assigned to ships</p>
                  <p className="mt-1 text-xs">Only users with "User" role can be assigned to ships</p>
                </div>
              )}
              {selectedRole === "user" && (
                <>
                  {shipsLoading ? (
                    <div className="p-2 text-sm">Loading ships...</div>
                  ) : (
                    <Select name="ship" value={selectedShip} onValueChange={setSelectedShip} required={!editingUser}>
                      <SelectTrigger><SelectValue placeholder={editingUser ? "Select a ship (optional)" : "Select a ship (required)"} /></SelectTrigger>
                      <SelectContent>
                        {availableShips.map((ship: any) => (
                          <SelectItem key={ship._id} value={ship._id}>
                            {ship.name} - IMO: {ship.imoNumber}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>
                  )}
                  {selectedRole === "user" && (
                    <div className="p-2 text-xs text-gray-600 bg-blue-50 rounded-md dark:bg-blue-900/20 dark:text-blue-300">
                      <p><strong>Note:</strong> Each ship can have exactly 2 users:</p>
                      <p>• 1 Deck Officer (for deck, MLC, ISPS, drill forms)</p>
                      <p>• 1 Engine Officer (for engine forms only)</p>
                    </div>
                  )}
                  {availableShips.length === 0 && !shipsLoading && selectedRole === "user" && (
                    <div className="p-3 text-sm text-orange-600 border border-orange-200 rounded-md bg-orange-50 dark:bg-orange-900/20 dark:text-orange-300 dark:border-orange-700">
                      <p>⚠️ No ships available for {selectedUserType} user type</p>
                      <p className="mt-1 text-xs">All ships already have a {selectedUserType} officer assigned. Create a new ship or change the user type.</p>
                    </div>
                  )}
                </>
              )}
            </div>
            
            <Button type="submit" className="w-full" disabled={createUserMutation.isPending || updateUserMutation.isPending}>
              {editingUser ? "Update User" : "Create User"}
            </Button>
          </form>
        </DialogContent>
      </Dialog>

      {/* Delete Confirmation Dialog */}
      <ConfirmationDialog
        isOpen={deleteConfirmation.isOpen}
        onClose={() => setDeleteConfirmation({ isOpen: false, user: null })}
        onConfirm={confirmDeleteUser}
        title="Delete User"
        description={`Are you sure you want to delete "${deleteConfirmation.user?.name}"? This action cannot be undone. The user will be permanently removed and their assigned ship will become available for other users.`}
        confirmText="Delete User"
        isLoading={deleteUserMutation.isPending}
      />
    </div>
  );
};

export default AdminUsers;