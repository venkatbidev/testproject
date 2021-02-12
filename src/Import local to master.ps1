<#
.Synopsis
    Copies the contents of a Power BI workspace to another Power BI workspace.
.Description
    Copies the contents of a Power BI workspace to another Power BI workspace, including dashboards, reports and datasets.
    This script uses the Power BI Management module for Windows PowerShell. If this module isn't installed, install it by using the command 'Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser'.
.Parameter SourceWorkspaceName
    The name of the workspace you'd like to copy the contents from.
.Parameter TargetWorkspaceName
    The name of the workspace you'd like to copy to. You must have edit access to the specified workspace.
.Parameter CreateTargetWorkspaceIfNotExists
    A flag to indicate if the script should create the target workspace if it doesn't exist. The default is to create the target workspace.
.Example
    PS C:\> .\CopyWorkspaceSampleScript.ps1 -SourceWorkspaceName "My Workspace" -TargetWorkspaceName "Copy of My Workspace"
	Copies the contents of the current user's personal workspace to a new workspace called "Copy of My Workspace".
#>

[CmdletBinding()]
param
(
    [string] $CompanyKey,

    [string] $DBServer = "pe-0-62.pestpac.local",

    [string] $SqlDB = "PestPac7335",

    [string] $TargetWorkspaceName = "PestPac Analytics Master"
)

#region Helper Functions 

function Assert-ModuleExists([string]$ModuleName) {
    $module = Get-Module $ModuleName -ListAvailable -ErrorAction SilentlyContinue
    if (!$module) {
        Write-Host "Installing module $ModuleName ..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser
        Write-Host "Module installed"
    }
    elseif ($module.Version -ne '1.0.0' -and $module.Version -le '1.0.410') {
        Write-Host "Updating module $ModuleName ..."
        Update-Module -Name $ModuleName -Force -ErrorAction Stop
        Write-Host "Module updated"
    }
}

#endregion


# ==================================================================
# PART 1: Verify that the Power BI Management module is installed
#         and authenticate the current user.
# ==================================================================
Assert-ModuleExists -ModuleName "MicrosoftPowerBIMgmt"
Connect-PowerBIServiceAccount 

# ==================================================================
# PART 2: Getting source and target workspace
# ==================================================================
# STEP 2.1: Get the source workspace
$tarWS = Get-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction SilentlyContinue

if (!$tarWS -or $tarWS.isReadOnly -eq "True") {
    Write-Error "Invalid choice: you must have edit access to the workspace."
    Disconnect-PowerBIServiceAccount
    exit
}

# ==================================================================
# PART 5: Rebind
# ==================================================================
Write-Host "Updating parameters"

$ds = Get-PowerBIDataset -Workspace $tarWS | where -Property Name -in -Value "Pestpac Dataset"
$ds

$BODY = ""
$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/Default.TakeOver"
$to = Invoke-PowerBIRestMethod -Method POST -Url $URL -Body $BODY
$to

$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/Parameters"
$gateways = Invoke-PowerBIRestMethod -Method GET -Url $URL
$gateways

$BODY = @{updateDetails=@{name="DBServer";newValue="$DBServer"},@{name="DB";newValue="$SqlDB"} } | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/Default.UpdateParameters"
$gateways = Invoke-PowerBIRestMethod -Method POST -Url $URL -Body $BODY
$gateways

$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/Parameters"
$gateways = Invoke-PowerBIRestMethod -Method GET -Url $URL
$gateways

$BODY = ""
$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/Default.DiscoverGateways"
$gateways = Invoke-PowerBIRestMethod -Method Get -Url $URL | ConvertFrom-Json
$gateways

$BODY = @{gatewayObjectId=$gateways.value.id} | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/Default.BindToGateway"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY
Write-Host "$URL $BODY"

$BODY = @{value=@{days="Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday";times="01:00","02:00","03:00","04:00","05:00","06:00","07:00","08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00","18:00","19:00","20:00","21:00","22:00","23:00","00:00";localTimeZoneId="Eastern Standard Time";enabled="TRUE" }}| ConvertTo-Json -Depth 3
$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/refreshSchedule"
Invoke-PowerBIRestMethod -Method PATCH -Url $URL -Body $BODY
Write-Host "$URL $BODY"

$BODY = @{notifyOption="MailOnFailure"} | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$($ds.Id.Guid)/refreshes"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY
Write-Host "$URL $BODY"

Disconnect-PowerBIServiceAccount
Resolve-PowerBIError
pause
