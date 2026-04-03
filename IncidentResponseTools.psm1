
# was fighting dll issues for a while. tried loading newest msal first. didn't work. leaving here just in case.


# need to load module with the newest version of MSAL first
$Modules = @(
    'ImportExcel'
    'Microsoft.Graph.Authentication'
    'Microsoft.Graph.Beta.Identity.Signins'
    'Microsoft.Graph.Beta.Reports'
    'Microsoft.Graph.Applications'
    'Microsoft.Graph.DirectoryObjects'
    'Microsoft.Graph.Groups'
    'Microsoft.Graph.Identity.DirectoryManagement'
    'Microsoft.Graph.Identity.Signins'
    'Microsoft.Graph.Reports'
    'Microsoft.Graph.Users'
    'Microsoft.Graph.Users.Actions'
    'ExchangeOnlineManagement' # have exchange be the last to import
)
Initialize-Modules -ModuleNames $Modules
