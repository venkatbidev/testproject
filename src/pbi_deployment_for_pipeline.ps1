[CmdletBinding()]
param
(
    [string] $SourceWorkspaceName = "workspace-staging-999004",

    [string] $TargetWorkspaceName = "workspace-prod-[enter companykey]",

    [string] $TargetDBServer = "[server name]",

    [string] $TargetDB = "[database name]",

    [bool] $CreateTargetWorkspaceIfNotExists = $false
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
$User = "pbiembedded@workwave.com"
$PWord = ConvertTo-SecureString -String "[enter password here]" -AsPlainText -Force
$cred = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $User, $PWord
Connect-PowerBIServiceAccount -Credential $cred 


# ==================================================================
# PART 2: Getting source and target workspace
# ==================================================================
# STEP 2.1: Get the source workspace
$srcWS = Get-PowerBIWorkspace -Name $SourceWorkspaceName -ErrorAction SilentlyContinue

if (!$srcWS) {
    Write-Error "Could not find a workspace with source workspace name. Please try again, making sure to type the exact name of the workspace" 
    Disconnect-PowerBIServiceAccount
    exit 
}


# STEP 2.2: Get the target workspace
$tarWS = Get-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction SilentlyContinue

if (!$tarWS -and $CreateTargetWorkspaceIfNotExists -eq $true) {
    New-PowerBIWorkspace -Name $TargetWorkspaceName 
    $tarWS = Get-PowerBIWorkspace -Name $TargetWorkspaceName -ErrorAction SilentlyContinue
    Add-PowerBIWorkspaceUser -AccessRight Admin -Identifier 6080fe38-e159-499d-8cc7-f4cb7ae05bf1 -PrincipalType App -Workspace $tarWS -ErrorAction Continue            # Service Principal (application id)
    Add-PowerBIWorkspaceUser -AccessRight Admin -Identifier dce354f9-1d6a-48c3-bb0f-021049aaa297 -PrincipalType Group -Workspace $tarWS -ErrorAction SilentlyContinue  # PowerBIAdmin
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
$temp_path_root = "$PSScriptRoot\pbix_temp"
$temp_dir = New-Item -Path "$temp_path_root" -ItemType Directory -ErrorAction SilentlyContinue

# STEP 3.2: Get the reports from the source workspace
$reports = Get-PowerBIReport -Workspace $srcWS | where -Property Name -in -Value "Billing", "Operations", "Pestpac Dataset"

# STEP 3.3: Export the PBIX files from the source and then import them into the target workspace

$pbienv = ""
$ExistingDatasetName = ""

if ($TargetWorkspaceName.Contains("pestpac_")){
    # ex:  workspace name of "workspace-pestpac_prod_msc-123456" will result in environment of "pestpac_prod_msc"

    $pbienv = $TargetWorkspaceName.Replace("workspace-","")
    $pbienv = $pbienv.Substring(0,$pbienv.Length-7)
}
else { 
    switch ($TargetWorkspaceName.ToLower()) {
        {$_.contains("prod")} {$pbienv="prod"}
        {$_.contains("staging")} {$pbienv="staging"}
        {$_.contains("dev")} {$pbienv="dev"}
    }
}


if ($pbienv -eq "") {
    Write-Error "Power BI environment not found"
    Exit
}

$ExistingDatasetId = ""
$ExistingServerName = ""
$ExistingDatabaseName = ""
$ExistingGatewayId = ""

$ws_ds = Get-PowerBIDataset -WorkspaceId $tarWS.Id.Guid

Foreach ($ds in $ws_ds) {
    $ExistingDatasetName = $ds.Name
            
    if ($ExistingDatasetName.ToLower() -notlike "*dataset*") {
        $ExistingDatasetName = ""
    }
    else {   
        $ExistingDatasetId = $ds.Id
     
        Break
    }
}

$SourceServerName = ""
$SourceDatabaseName = ""
$SourceGatewayId = ""

$source_ws_ds = Get-PowerBIDataset -WorkspaceId $srcWS.Id.Guid

Foreach ($ds in $source_ws_ds) {
    $SourceDatasetName = $ds.Name
    $SourceDatasetId = $ds.Id
        
    if ($SourceDatasetName.ToLower() -notlike "dataset-*") {
        $SourceDatasetName = ""
    }
    else {
        
        $URL = "datasets/$SourceDatasetId/datasources"
        $Response = Invoke-PowerBIRestMethod -Url "datasets/$($ds.Id)/datasources" -Method Get | ConvertFrom-Json
            
        $SourceDataSource = $Response.value[0]

        $SourceServerName = $SourceDataSource.connectionDetails.server
        $SourceDatabaseName = $SourceDataSource.connectionDetails.database
        $SourceGatewayId = $SourceDataSource.gatewayId
        
        Break
    }
}

$dataset_report_id = ""
   
Foreach ($report in $reports) {
   
    $report_id = [guid]$report.id
    $dataset_id = [guid]$report.datasetId
    $report_name = $report.name
    
    $temp_path = "$temp_path_root\$report_name.pbix"

    # Only export if this dataset hasn't already been exported already (because the same dataset is used with both reports).
    if ($dataset_ID_mapping -and $dataset_ID_mapping[$dataset_id]) {
        continue
    }

    try {
        if ($source_workspace_ID -eq "me") {
            Export-PowerBIReport -Id $report_id -OutFile "$temp_path" -ErrorAction Stop
        }
        else {
            Export-PowerBIReport -Id $report_id -WorkspaceId $srcWS.Id.Guid -OutFile "$temp_path" -ErrorAction Stop
        }
    }
    catch {
        # This report and dataset cannot be copied, skipping. This is expected for most workspaces."
        continue
    }
     
    try {
        
        if ($report_name.ToLower() -like "*dataset*") {
            $report_name = "dataset-$pbienv-$TargetDB" #here
            $report_name = $report_name.ToLower()
            $dataset_report_id = $report_id
        }

        if ($ExistingDatasetId -ne "") {
            # Delete old dataset.   Will remove all old reports as well.
            $URL = "groups/$($tarWS.Id.Guid)/datasets/$ExistingDatasetId"
            Invoke-PowerBIRestMethod -Method Delete -Url $URL
        }

        $new_report = New-PowerBIReport -WorkspaceId $tarWS.Id.Guid -Path $temp_path -Name $report_name -ConflictAction CreateOrOverwrite
                
        # Get the report again because the dataset id is not immediately available with New-PowerBIReport
        $new_report = Get-PowerBIReport -WorkspaceId $tarWS.Id.Guid -Id $new_report.id
        if ($new_report) {
            # keep track of the report and dataset IDs
            $report_id_mapping[$report_id] = $new_report.id
            $dataset_id_mapping[$dataset_id] = $new_report.datasetId
        }
    }
    catch [Exception] {
        Write-Error "Failed to import PBIX"

        $exception = Resolve-PowerBIError -Last
        Write-Error "Error Description:" $exception.Message
        continue
    }
}


# STEP 3.5: Copy any remaining reports that have not been copied yet. 
$failure_log = @()  
Foreach ($report in $reports) {
    $report_name = $report.name
    $report_datasetId = [guid]$report.datasetId

    $target_dataset_Id = $dataset_id_mapping[$report_datasetId]
    if ($target_dataset_Id -and !$report_ID_mapping[$report.id]) {
        
        $report_copy =  Copy-PowerBIReport -Report $report -WorkspaceId $srcWS.Id.Guid -TargetWorkspaceId $tarWS.Id.Guid -TargetDatasetId $target_dataset_Id

        $report_ID_mapping[$report.id] = $report_copy.id
    }
    else {
        $failure_log += $report
    }
}


# ==================================================================
# PART 5: Rebind
# ==================================================================
$BODY = "{
  `"updateDetails`": [
    {
      `"datasourceSelector`": {
        `"datasourceType`": `"Sql`",
        `"connectionDetails`": {
          `"server`": `"$SourceServerName`",
          `"database`": `"$SourceDatabaseName`"
        }
      },
      `"connectionDetails`": {
        `"server`": `"$TargetDBServer`",
        `"database`": `"$TargetDB`"
      }
    }
  ]
}"

# Not necessary
#$URL = "datasets/$target_dataset_Id/Default.UpdateDatasources"
#Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY

$BODY = @{updateDetails=@{name="DBServer";newValue=$TargetDBServer},@{name="DB";newValue=$TargetDB} } | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/Default.UpdateParameters"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY

$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/Default.DiscoverGateways"

$gateways = ((Invoke-PowerBIRestMethod -Url $URL -Method Get) | ConvertFrom-Json).value

$BODY = @{gatewayObjectId=$gateways.Id} | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/Default.BindToGateway"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY

# Enable Scheduled Refresh
$BODY = @{value=@{enabled="true"} } | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/refreshSchedule"
Invoke-PowerBIRestMethod -Method Patch -Url $URL -Body $BODY

if ($pbienv -eq "staging" -or $pbienv -eq "dev") {
    $BODY = @{value=@{days="Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday";times="09:00","10:00","11:00","12:00","13:00","14:00","15:00","16:00";localTimeZoneId="Eastern Standard Time";enabled="TRUE" }}| ConvertTo-Json -Depth 3
}
else {
    ## prod refreshes once per day
    $BODY = @{value=@{days="Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday";times="04:00","00:00";localTimeZoneId="Eastern Standard Time";enabled="TRUE" }}| ConvertTo-Json -Depth 3
    $BODY = $BODY.Replace("`"00:00`"","").Replace("`"04:00`",","`"04:00`"")  ## hacky way to get the times as an array
}

$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/refreshSchedule"
Invoke-PowerBIRestMethod -Method PATCH -Url $URL -Body $BODY

## Trigger a refresh on the database.
$BODY = @{notifyOption="MailOnFailure"} | ConvertTo-Json
$URL = "groups/$($tarWS.Id.Guid)/datasets/$target_dataset_Id/refreshes"
Invoke-PowerBIRestMethod -Method Post -Url $URL -Body $BODY

# ==================================================================
# PART 6: Cleanup
# ==================================================================
Remove-Item -path $temp_path_root -Recurse -ErrorAction Continue
Disconnect-PowerBIServiceAccount
