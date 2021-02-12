Clear-Host

$SQL_Query = "
select [key] as companykey, [database] as databasename, si.InstanceName as databaseserver
from PestPacDatabases d
join SQLInstances si on si.InstanceID = d.InstanceID
where (
    exists (
	select * 
	from companyconfig cc
	join configfields cf on cf.ConfigFieldID = cc.ConfigFieldID
	where cc.DatabaseID = d.DatabaseID
	and cf.FieldName = 'UsePowerBI'
	and cc.ConfigFieldValue = 'Y'
    )
)
order by si.InstanceName, [key]
"
# test company key's above are being filtered out.   not sure why reps turned on BA for these.

$count = 0

#Assert-ModuleExists -ModuleName "MicrosoftPowerBIMgmt"
$User = ""
$PWord = ConvertTo-SecureString -String "" -AsPlainText -Force
$cred = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $User, $PWord
Connect-PowerBIServiceAccount -Credential $cred 

$timeZoneId = "Eastern Standard Time"

try {
    # AWS WEST
    #$DS = Invoke-Sqlcmd -ServerInstance "10.212.33.30" -Query $SQL_Query -As DataSet -Database "PestPacConfig"

    # AWS EAST
    $DS = Invoke-Sqlcmd -ServerInstance "10.1.30.95" -Query $SQL_Query -As DataSet -Database "PestPacConfig"
      
    $ba_customers = $DS.Tables[0]

    $total =  $ba_customers.Rows.Count

    foreach ($customer in $ba_customers) #for each separate server / database pair in $databases
    {
        $count = $count + 1

        $company_key = $customer.companykey.Trim()
        $database = $customer.databasename
        $databaseserver = $customer.databaseserver

        $tar_WorkspaceName = "workspace-prod-$company_key"

        $TargetServerName = ""
        $TargetDatabaseName = ""
        $TargetGatewayId = ""

        try {
            $tarWS = Get-PowerBIWorkspace -Name $tar_WorkspaceName
            # -ErrorAction SilentlyContinue
        }
        catch {
            Write-Error "($count of $total) $company_key --> Workspace not found!"
            Continue
        }

        $ExistingDatasetName = ""
        $respStr = ""

        if (!$tarWS) {
            Write-Error "Workspace $tarWS not found for $company_key."
        }
        else {
     
            $target_ws_ds = Get-PowerBIDataset -WorkspaceId $tarWS.Id.Guid

            $datasetFound = $false

            $refreshStatus = ""
            $refreshInfo = "Unable to locate refresh history"
            $refreshState = "error"

            Foreach ($ds in $target_ws_ds) {
                $TargetDatasetName = $ds.Name
                $TargetDatasetId = $ds.Id
                $datasetFound = $true
        
                $URL = "datasets/$TargetDatasetId/datasources"
                $Response = Invoke-PowerBIRestMethod -Url "datasets/$($ds.Id)/datasources" -Method Get | ConvertFrom-Json
                
                $TargetDataSource = $Response.value[0]


                $dataSourceType = $TargetDataSource.datasourceType
                if ($dataSourceType.ToLower() -eq "sql") {
                    $TargetServerName = $TargetDataSource.connectionDetails.server
                    $TargetDatabaseName = $TargetDataSource.connectionDetails.database
                    $TargetGatewayId = $TargetDataSource.gatewayId

                    $historyurl = "groups/$($tarWS.id)/datasets/$($ds.Id)/refreshes"
                    $history = ((Invoke-PowerBIRestMethod -Url $historyurl -Method Get) |
                        ConvertFrom-Json).value |
                        Sort-Object -Descending -Property timeStart |
                        Select-Object -First 1
                    
                    $refreshStatus = $history.status
                    
                    $outputDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($completedAt, "Greenwich Standard Time", $timeZoneId)


                    switch ($refreshStatus.ToLower()) {
                        {$_.contains("completed")} {$refreshState="ok"}
                        {$_.contains("unknown")} {$refreshState="warning"}
                    }

                    if ($refreshStatus.ToLower() -eq "completed") {
                        $completedAtStr = $history.endTime.Replace("T"," ").Replace("Z"," ").Substring(0,16)
                        $completedAt = [datetime]::parseexact($completedAtStr, 'yyyy-MM-dd hh:mm', $null)

                        $outputDateTime = [System.TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($completedAt, "Greenwich Standard Time", $timeZoneId)

                        $Now = Get-Date
                        $TimeSpan = $Now - $outputDateTime

                        $hoursSinceLastRefresh = $TimeSpan.TotalHours
                        if ($hoursSinceLastRefresh -gt 23) {
                            $refreshState = "error"
                            $refreshInfo = "Last refresh: $refreshStatus at $outputDateTime * REFRESH EXPIRED *"
                        }
                        else {
                            $refreshInfo = "Last refresh: $refreshStatus at $outputDateTime"
                        }
                        
                    }
                    elseif ($refreshStatus.ToLower() -eq "unknown") {
                        $refreshInfo = "Refresh is pending completion or unknown"
                    }
                    else {
                        $refreshInfo = "Last refresh: $refreshStatus at $completedAt"
                    }


                    Break
                }
             
            }

            if (($database.ToLower() -ne $TargetDatabaseName.ToLower()) -or ($databaseserver.ToLower() -ne $TargetServerName.ToLower())) {
                Write-Error "($count of $total) $company_key --> $databaseserver | $database vs $TargetServerName | $TargetDatabaseName ($datasetFound - $respStr)"
            }
            else {

                if ($refreshState -eq "error") {
                    Write-Error "($count of $total) $company_key --> $databaseserver | $database vs $TargetServerName | $TargetDatabaseName | $refreshInfo "                
                }
                elseif ($refreshState -eq "warning") {
                    Write-Warning "($count of $total) $company_key --> $databaseserver | $database vs $TargetServerName | $TargetDatabaseName | $refreshInfo "
                }
                else {
                    Write-Host "($count of $total) $company_key --> $databaseserver | $database vs $TargetServerName | $TargetDatabaseName | $refreshInfo "
                }
                
            }

        }

    }

}
catch {
    Write-Error "* $company_key *"
    Resolve-PowerBIError -Last    
    Exit
}

Write-Host "complete"




