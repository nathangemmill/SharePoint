#----------------------------------------------------------------
#	Type:	Script
#	Desc:	Create SharePoint Document Libraries and Security Groups
#	Author:	Nathan Gemmill
#	Ver:	1.0
#----------------------------------------------------------------

#----------------------------------------------------------------
# Variables
#----------------------------------------------------------------
$LibraryNames = Import-Csv -Path ".\DocumentLibraries.csv"
[string]$SiteAddress = Import-Csv -Path ".\DocumentLibraries.csv" | Select-Object -ExpandProperty SharePointSite
$GroupNamePrefix = Import-Csv -Path ".\DocumentLibraries.csv" | Select-Object -ExpandProperty GroupPrefix

#----------------------------------------------------------------
# Connect to Exchange Online & SharePoint Site
#----------------------------------------------------------------
#Connect-MsolService
#Connect-PnPOnline -Url $SiteAddress -UseWebLogin

#----------------------------------------------------------------
# Loop through spreadhseet and create libraries
#----------------------------------------------------------------
foreach ($Library in $LibraryNames)
{
    $Title = $Library.LibraryName 
    write-host Creating Document Library named: $Title
    New-PnPList -Title $Title -Template DocumentLibrary -OnQuickLaunch -ErrorAction SilentlyContinue
}

#----------------------------------------------------------------
# Loop through spreadhseet and create security groups
#----------------------------------------------------------------
foreach ($Library in $LibraryNames)
{
    $Title = $Library.LibraryName 
    [string]$GroupNameWrite = $GroupNamePrefix.replace('   ','') + ' - ' + $Title.replace(' ','') + ' - R&W'
    [string]$GroupNameRead = $GroupNamePrefix.replace('   ','') + ' - ' + $Title.replace(' ','') + ' - RO'
    write-host Creating security groups named: $GroupNameWrite
    New-MsolGroup -DisplayName $GroupNameWrite -ErrorAction SilentlyContinue
    write-host Creating security groups named: $GroupNameRead
    New-MsolGroup -DisplayName $GroupNameRead -ErrorAction SilentlyContinue
}

#----------------------------------------------------------------
# Loop through spreadhseet and set Document Library permissions
#----------------------------------------------------------------
foreach ($Library in $LibraryNames)
{
    $Title = $Library.LibraryName 
    [string]$GroupNameWrite = $GroupNamePrefix.replace('   ','') + ' - ' + $Title.replace(' ','') + ' - R&W'
    [string]$GroupNameRead = $GroupNamePrefix.replace('   ','') + ' - ' + $Title.replace(' ','') + ' - RO'
    Write-Host Breaking inheritance on the following document library: $Title
    Set-PnPList -Identity $Title -BreakRoleInheritance -CopyRoleAssignments
    Write-Host Setting edit permissions for the following group: $GroupNameWrite
    Set-PnPListPermission -Identity "$Title" -AddRole "Edit" -User "$GroupNameWrite"
    Write-Host Setting read permissions for the following group: $GroupNameRead
    Set-PnPListPermission -Identity "$Title" -AddRole "Read" -User "$GroupNameWrite"
}