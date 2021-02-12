Clear-Host

$SQL_Query = "
select [key] as companykey, [database] as databsename, si.InstanceName as databaseserver
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
order by [key]
"
# test company key's above are being filtered out.   not sure why reps turned on BA for these.

$src_Workspace = "workspace-staging-999004"

Write-Host "Source workspace: $src_Workspace"
$count = 0

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

        $company_key = $customer.companykey
        $database = $customer.databsename
        $databaseserver = $customer.databaseserver

        $tar_WorkspaceName = "workspace-prod-$company_key"

        Write-Host "Updating $count of $total workspaces ($tar_WorkspaceName) ..."

        $tar_DBServer = "$databaseserver"
    
        $tar_DB = "$database"
        
        # datasource for the test server has the IP address, not servername.
        if ($tar_DBServer.ToLower() -eq "sql-test-db") {
            $tar_DBServer = "10.1.30.107"
        }

        C:\PestPac\PowerBIEmbedded\PBI\src\mine\pbi_deployment_for_pipeline.ps1 -SourceWorkspaceName "$src_Workspace" -TargetWorkspaceName "$tar_WorkspaceName" -TargetDBServer "$tar_DBServer" -TargetDB "$tar_DB"

    }

}
catch {
    Resolve-PowerBIError -Last    
    Exit
}

Write-Host "complete"




