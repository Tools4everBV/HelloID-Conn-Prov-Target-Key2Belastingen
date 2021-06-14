$config = $configuration | ConvertFrom-Json;
$success = $False;
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];

$p = $person | ConvertFrom-Json;
$m = $manager | ConvertFrom-Json;
$aRef = $accountReference | ConvertFrom-Json;
$mRef = $managerAccountReference | ConvertFrom-Json;

# The permissionReference object contains the Identification object provided in the retrieve permissions call
$pRef = $permissionReference | ConvertFrom-json;

$DataSource = $config.dataSource
$Username = $config.username
$Password = $config.password

$OracleConnectionString = "User Id=$Username;Password=$Password;Data Source=$DataSource"

if(-Not($dryRun -eq $True)) {
    try{
		$null =[Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")

		#check correlation before create
        $OracleConnection = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
        $OracleConnection.Open()
        Write-Verbose -Verbose "Successfully connected Oracle to database '$DataSource'" 
				
        # Execute the command against the database
		
        $OracleCmd = $OracleConnection.CreateCommand()
        
	    $OracleQuery = "
            MERGE INTO menu_b_user t1
				  USING
				  	(SELECT DISTINCT '$($pRef.Id)' AS GROUP_NAME, '$($p.Accounts.OracleKey2Belasting.USERNAME)' AS USER_NAME
					 FROM menu_b_user) t2
				  ON (t1.GROUP_NAME = t2.GROUP_NAME AND t1.USER_NAME = t2.USER_NAME)
			  WHEN NOT MATCHED THEN
			  	INSERT VALUES (t2.GROUP_NAME, t2.USER_NAME)"
        
        Write-Verbose -Verbose $OracleQuery
        
        $OracleCmd.CommandText = $OracleQuery
        $OracleCmd.ExecuteNonQuery() | Out-Null
        
        Write-Verbose -Verbose "Successfully performed Oracle query."

        $success = $True;
        $auditMessage = " succesfully";   
		
    } catch {
        Write-Error $_
    }finally{
        if($OracleConnection.State -eq "Open"){
            $OracleConnection.close()
        }
        Write-Verbose -Verbose "Successfully disconnected from Oracle database '$DataSource'"
    }
}
$success = $True;
$auditLogs.Add([PSCustomObject]@{
    # Action = "GrantMembership"; Optionally specify a different action for this audit log
    Message = "Permission to $($pRef.Id) added to account $($aRef)";
    IsError = $False;
});

# Send results
$result = [PSCustomObject]@{
    Success= $success;
    AuditLogs = $auditLogs;
    Account = [PSCustomObject]@{};
};
Write-Output $result | ConvertTo-Json -Depth 10;