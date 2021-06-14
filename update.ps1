$config = $configuration | ConvertFrom-Json;
$p = $person | ConvertFrom-Json;
$m = $manager | ConvertFrom-Json;
$aRef = $accountReference | ConvertFrom-Json;
$mRef = $managerAccountReference | ConvertFrom-Json;
$success = $False;
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];


$DataSource = $config.dataSource
$Username = $config.username
$Password = $config.password

$OracleConnectionString = "User Id=$Username;Password=$Password;Data Source=$DataSource"

$displayname = ""

$prefix = ""
if(-Not([string]::IsNullOrEmpty($p.Name.FamilyNamePrefix)))
{
    $prefix = $p.Name.FamilyNamePrefix + " "
    $prefixEnd = " " + $p.Name.FamilyNamePrefix
}

$partnerprefix = ""
if(-Not([string]::IsNullOrEmpty($p.Name.FamilyNamePartnerPrefix)))
{
    $partnerprefix = $p.Name.FamilyNamePartnerPrefix + " "
    $partnerprefixEnd = " " + $p.Name.FamilyNamePartnerPrefix
}

switch($p.Name.Convention)
{
    "B" {$displayname += $p.Name.FamilyName + ", " + $p.Name.NickName + $prefixEnd}
    "P" {$displayname += $p.Name.FamilyNamePartner + ", " + $p.Name.NickName + $partnerprefixEnd}
    "BP" {$displayname += $p.Name.FamilyName + " - " + $partnerprefix + $p.Name.FamilyNamePartner + ", " + $p.Name.NickName + $prefixEnd}
    "PB" {$displayname += $p.Name.FamilyNamePartner + " - " + $prefix + $p.Name.FamilyName + ", " + $p.Name.NickName + $partnerprefixEnd}
    default {$displayname += $p.Name.FamilyName + ", " + $p.Name.NickName + $prefixEnd}
}

# Change mapping here
$account = [PSCustomObject]@{
    GEBRCODE						= $aRef;
    GEBR_OMS						= $displayname;
    EMAIL							= $p.Accounts.MicrosoftActiveDirectory.UserPrincipalName;
};

if(-Not($dryRun -eq $True)) {
    try{
		$null =[Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")

		#check correlation before create
        $OracleConnection = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
        $OracleConnection.Open()
        Write-Verbose -Verbose "Successfully connected Oracle to database '$DataSource'" 

        
		$OracleQueryUpdate = "UPDATE wms_gebrcode SET GEBR_OMS = '$($account.GEBR_OMS)', EMAIL = '$($account.EMAIL)' WHERE GEBRCODE = '$aRef'"
	
		Write-Verbose -Verbose $OracleQueryUpdate
		
        $OracleCmd = $OracleConnection.CreateCommand()
		$OracleCmd.CommandText = $OracleQueryUpdate
		$OracleCmd.ExecuteNonQuery() | Out-Null
		
		Write-Verbose -Verbose "Successfully performed Oracle update query."

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
    # Action = "CreateAccount"; Optionally specify a different action for this audit log
    Message = "Created account with username $($account.userName)";
    IsError = $False;
});

# Send results
$result = [PSCustomObject]@{
	Success= $success;
	AuditLogs = $auditLogs;
    Account = $account;
};
Write-Output $result | ConvertTo-Json -Depth 10;
