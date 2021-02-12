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
    [string] $SourceWorkspaceName = "PestPac Analytics Master",

    [string] $CompanyKey,

    [string] $DBServer = "pe-0-62.pestpac.local",

    [string] $SqlDB = "PestPac7335",

    [string] $TargetWorkspaceName = "workspace-dev-999003"
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
$srcWS = Get-PowerBIWorkspace -Name $SourceWorkspaceName -ErrorAction SilentlyContinue

if (!$srcWS) {
    Write-Warning "Could not get a workspace with that name. Please try again, making sure to type the exact name of the workspace" 
    Disconnect-PowerBIServiceAccount
    exit 
}


# STEP 2.2: Get the target workspace

$tarWS = Get-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction SilentlyContinue

if (!$tarWS ) {
    New-PowerBIWorkspace -Name $TargetWorkspaceName 
    $tarWS = Get-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction SilentlyContinue
    Add-PowerBIWorkspaceUser -AccessRight Admin -Identifier 6080fe38-e159-499d-8cc7-f4cb7ae05bf1 -PrincipalType App -Workspace $tarWS -ErrorAction Continue
    Add-PowerBIWorkspaceUser -AccessRight Admin -Identifier dce354f9-1d6a-48c3-bb0f-021049aaa297 -PrincipalType Group -Workspace $tarWS -ErrorAction SilentlyContinue

}

if (!$tarWS -or $tarWS.isReadOnly -eq "True") {
    Write-Error "Invalid choice: you must have edit access to the workspace."
    Disconnect-PowerBIServiceAccount
    exit
}

# ==================================================================
# PART 3: Copying reports and datasets via Export/Import of 
#         reports built on PBIXes (this step creates the datasets)
# ==================================================================
$report_ID_mapping = @{ }      # mapping of old report ID to new report ID
$dataset_ID_mapping = @{ }     # mapping of old model ID to new model ID

# STEP 3.1: Create a temporary folder to export the PBIX files.
$temp_path_root = "$PSScriptRoot\temp"
$temp_dir = New-Item -Path "$temp_path_root" -ItemType Directory -ErrorAction Continue

# STEP 3.2: Get the reports from the source workspace
$reports = Get-PowerBIReport -Workspace $srcWS | where -Property Name -in -Value "Billing", "Operations", "Pestpac Dataset"

# STEP 3.3: Export the PBIX files from the source and then import them into the target workspace
Foreach ($report in $reports) {
   
    $report_id = [guid]$report.id
    $dataset_id = [guid]$report.datasetId
    $report_name = $report.name
    $newreport_name = "dataset-dev-pestpac7335"
    $temp_path = "$temp_path_root\$report_name.pbix"

    # Only export if this dataset id it hasn't already been exported already
    if ($dataset_ID_mapping -and $dataset_ID_mapping[$dataset_id]) {
        continue
    }

    Write-Host "== Exporting $report_name with id: $report_id to $temp_path"
    try {
        if ($source_workspace_ID -eq "me") {
            Export-PowerBIReport -Id $report_id -OutFile "$temp_path" -ErrorAction Stop
        }
        else {
            Export-PowerBIReport -Id $report_id -WorkspaceId $srcWS.Id.Guid -OutFile "$temp_path" -ErrorAction Stop
        }
    }
    catch {
        Write-Warning "= This report and dataset cannot be copied, skipping. This is expected for most workspaces."
        continue
    }
     
    try {
        Write-Host "== Importing $report_name to target workspace"

        $DeleteDataset = Get-PowerBIDataset -WorkspaceId $tarWS.id.guid | where -Property Name -eq -Value $newreport_name
        $URL = "groups/$($tarWS.Id.Guid)/datasets/$($DeleteDataset.Id.Guid)"
        $gateways = Invoke-PowerBIRestMethod -Method DELETE -Url $URL -ErrorAction Continue

        $new_report = New-PowerBIReport -WorkspaceId $tarWS.Id.Guid -Path $temp_path -Name $newreport_name -ConflictAction CreateOrOverwrite
                
        # Get the report again because the dataset id is not immediately available with New-PowerBIReport
        $new_report = Get-PowerBIReport -WorkspaceId $tarWS.Id.Guid -Id $new_report.id
        if ($new_report) {
            # keep track of the report and dataset IDs
            $report_id_mapping[$report_id] = $new_report.id
            $dataset_id_mapping[$dataset_id] = $new_report.datasetId
        }
    }
    catch [Exception] {
        Write-Error "== Error: failed to import PBIX"

        $exception = Resolve-PowerBIError -Last
        Write-Error "Error Description:" $exception.Message
        continue
    }
}

# STEP 3.4: Copy any remaining reports that have not been copied yet. 
$failure_log = @()  

Foreach ($report in $reports) {
    $report_name = $report.name
    $report_datasetId = [guid]$report.datasetId

    $target_dataset_Id = $dataset_id_mapping[$report_datasetId]
    if ($target_dataset_Id -and !$report_ID_mapping[$report.id]) {
        Write-Host "== Copying report $report_name"
        $report_copy =  Copy-PowerBIReport -Report $report -WorkspaceId $srcWS.Id.Guid -TargetWorkspaceId $tarWS.Id.Guid -TargetDatasetId $target_dataset_Id 

        $report_ID_mapping[$report.id] = $report_copy.id
    }
    else {
        $failure_log += $report
    }
}

# ==================================================================
# PART 4: Copy dashboards and tiles
# ==================================================================

# STEP 4.1 Get all dashboards from the source workspace
# If source is My Workspace, filter out dashboards that I don't own - e.g. those shared with me
$dashboards = "" 
if ($srcWS.Id.Guid -eq "me") {
    $dashboards = Get-PowerBIDashboard
    $dashboards_temp = @()
    Foreach ($dashboard in $dashboards) {
        if ($dashboard.isReadOnly -ne "True") {
            $dashboards_temp += $dashboard
        }
    }
    $dashboards = $dashboards_temp
}
else {
    $dashboards = Get-PowerBIDashboard -WorkspaceId $srcWS.Id.Guid
}

# STEP 4.2 Copy the dashboards and their tiles to the target workspace
Foreach ($dashboard in $dashboards) {
    $dashboard_id = $dashboard.id
    $dashboard_name = $dashboard.Name

    Write-Host "== Cloning dashboard: $dashboard_name"

    # create new dashboard in the target workspace
    $dashboard_copy = New-PowerBIDashboard -Name $dashboard_name -WorkspaceId $tarWS.Id.Guid
    $target_dashboard_id = $dashboard_copy.id

    Write-Host " = Copying the tiles..." 
    $tiles =  Get-PowerBITile -WorkspaceId $srcWS.Id.Guid -DashboardId $dashboard_id

    Foreach ($tile in $tiles) {
        try {
            $tile_id = $tile.id
            if ($tile.reportId) {
                $tile_report_Id = [GUID]($tile.reportId)
            }
            else {
                $tile_report_Id = $null
            }

            if (!$tile.datasetId) {
                Write-Warning "= Skipping tile $tile_id, no dataset id..."
                continue
            }
            else {
                $tile_dataset_Id = [GUID]($tile.datasetId)
            }

            if ($tile_report_id) { $tile_target_report_id = $report_id_mapping[$tile_report_id] }
            if ($tile_dataset_id) { $tile_target_dataset_id = $dataset_id_mapping[$tile_dataset_id] }

            # clone the tile only if a) it is not built on a dataset or b) if it is built on a report and/or dataset that we've moved
            if (!$tile_report_id -Or $dataset_id_mapping[$tile_dataset_id]) {
                $tile_copy = if ($source_workspace_ID -eq "me") { 
                    Copy-PowerBITile -DashboardId $dashboard_id -TileId $tile_id -TargetDashboardId $target_dashboard_id -TargetWorkspaceId $target_workspace_ID -TargetReportId $tile_target_report_id -TargetDatasetId $tile_target_dataset_id 
                }
                else {
                    Copy-PowerBITile -WorkspaceId $srcWS.Id.Guid -DashboardId $dashboard_id -TileId $tile_id -TargetDashboardId $target_dashboard_id -TargetWorkspaceId $tarWS.Id.Guid -TargetReportId $tile_target_report_id -TargetDatasetId $tile_target_dataset_id 
                }
                
                Write-Host "." -NoNewLine
            }
            else {
                $failure_log += $tile
            } 
           
        }
        catch [Exception] {
            Write-Error "Error: skipping tile..."
            Write-Error $_.Exception
        }
    }
    Write-Host "Done copying onto updating the parameters and gateway binding!"
}
# ==================================================================
# PART 5: Rebind
# ==================================================================
Write-Host "Updating parameters"

$BODY = @{updateDetails=@{name="DBServer";newValue="$DBServer"},@{name="DB";newValue="$SqlDB"} } | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/Default.UpdateParameters"
$gateways = Invoke-PowerBIRestMethod -Method POST -Url $URL -Body $BODY

$BODY = ""
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/Default.DiscoverGateways"
$gateways = Invoke-PowerBIRestMethod -Method Get -Url $URL | ConvertFrom-Json

$BODY = @{gatewayObjectId=$gateways.value.id} | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/Default.BindToGateway"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY

$BODY = @{value=@{days="Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday";times="01:00","02:00","03:00","04:00","05:00","06:00","07:00","08:00","09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00","17:00","18:00","19:00","20:00","21:00","22:00","23:00","00:00";localTimeZoneId="Eastern Standard Time";enabled="TRUE" }}| ConvertTo-Json -Depth 3
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/refreshSchedule"
Invoke-PowerBIRestMethod -Method PATCH -Url $URL -Body $BODY

$BODY = @{notifyOption="MailOnFailure"} | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/refreshes"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY

# ==================================================================
# PART 6: Cleanup
# ==================================================================
Write-Host "Cleaning up temporary files"
Remove-Item -path $temp_path_root -Recurse -ErrorAction Continue -Force
Disconnect-PowerBIServiceAccount
