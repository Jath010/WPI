Import-module Microsoft.Graph.Users
Connect-MgGraph
(Get-MgUser -UserId (get-azureaduser -objectid "badorr@wpi.edu").ObjectID).OnPremisesExtensionAttributes|format-list