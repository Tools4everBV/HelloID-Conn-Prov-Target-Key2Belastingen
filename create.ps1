$config = $configuration | ConvertFrom-Json;
$p = $person | ConvertFrom-Json;
$m = $manager | ConvertFrom-Json;
$success = $False;
$auditLogs = New-Object Collections.Generic.List[PSCustomObject];

function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}

function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

function GenerateRandomPassword(){
    $password = Get-RandomCharacters -length 4 -characters 'abcdefghiklmnoprstuvwxyz'
    $password += Get-RandomCharacters -length 2 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += Get-RandomCharacters -length 2 -characters '1234567890'
    #$password += Get-RandomCharacters -length 2 -characters '!§$%&/()=?}][{#*+'

    $password = Scramble-String $password

    return $password
}

$oracleUsername = $p.Accounts.OracleKey2Belasting.USERNAME

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
    GEBRCODE						= $oracleUsername.subString(0,6);
    GEBR_OMS						= $displayname;
    INDACTIEF						= "J";
    GEBR_ORA						= $oracleUsername;
    SUBJECTNR						= "";
    TELNR							= "";
    FAXNR							= "";
    LOCATIE							= "";
    EMAIL							= $p.Accounts.MicrosoftActiveDirectory.UserPrincipalName;
    VRIJ_VELD						= "";
    IND_DW							= "N";
    EXE_USER						= "";
    HASHEE							= "";
    IND_JURIST						= "N";
    PASSWORD                        = GenerateRandomPassword
};

if(-Not($dryRun -eq $True)) {
    try{
		$null =[Reflection.Assembly]::LoadWithPartialName("System.Data.OracleClient")

		#check correlation before create
        $OracleConnection = New-Object System.Data.OracleClient.OracleConnection($OracleConnectionString)
        $OracleConnection.Open()
        Write-Verbose -Verbose "Successfully connected Oracle to database '$DataSource'" 

        $unique = $false
        $i=0
        $maxIterations = 10
        while(-not($unique) -and $i -lt $maxIterations)
        {
            if($i -gt 0)
            {
                $account.GEBRCODE = $account.GEBRCODE.subString(0,5) + "$i"
            }
            # Execute the command against the database
            $OracleQuery = "SELECT GEBRCODE FROM wms_gebrcode WHERE GEBRCODE = '$($account.GEBRCODE)'"
            #Write-Verbose -Verbose $OracleQuery
            $OracleCmd = $OracleConnection.CreateCommand()
            $OracleCmd.CommandText = $OracleQuery

            $OracleAdapter = New-Object System.Data.OracleClient.OracleDataAdapter($cmd)
            $OracleAdapter.SelectCommand = $OracleCmd;

            # Execute the command against the database, returning results.
            $DataSet = New-Object system.Data.DataSet
            $null = $OracleAdapter.fill($DataSet)

            $result = $DataSet.Tables[0] | Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors;

            Write-Verbose -Verbose "Successfully performed Oracle '$OracleQuery'. Returned [$($DataSet.Tables[0].Columns.Count)] columns and [$($DataSet.Tables[0].Rows.Count)] rows"
            
            $rowcount = $($DataSet.Tables[0].Rows.Count)
            
            if($rowcount -eq 0){    
                $unique = $true
            }
            $i++
        }

        if(-not($unique)){
            Write-Error "Used $maxIterations iterations and no unique GEBRCODE was found"
            $success = $False
        }
        else{            
            $OracleQueryCreate = "		
               MERGE INTO wms_gebrcode t1
				  USING
				  	(SELECT DISTINCT 
							'$($account.GEBRCODE)' AS GEBRCODE,
							'$($account.GEBR_OMS)' AS GEBR_OMS,
							'$($account.INDACTIEF)' AS INDACTIEF,
							'$($account.GEBR_ORA)' AS GEBR_ORA,
							'$($account.SUBJECTNR)' AS SUBJECTNR,
							'$($account.TELNR)' AS TELNR,
							'$($account.FAXNR)' AS FAXNR,
							'$($account.LOCATIE)' AS LOCATIE,
							'$($account.EMAIL)' AS EMAIL,
							'$($account.VRIJ_VELD)' AS VRIJ_VELD,
							'$($account.IND_DW)' AS IND_DW,
							'$($account.EXE_USER)' AS EXE_USER,
							'$($account.HASHEE)' AS HASHEE,
                            '$($account.IND_JURIST)' AS IND_JURIST
					 FROM wms_gebrcode) t2
				  ON (t1.GEBRCODE = t2.GEBRCODE AND t1.GEBR_ORA = t2.GEBR_ORA)
			  WHEN NOT MATCHED THEN
			  	INSERT VALUES (t2.GEBRCODE, t2.GEBR_OMS, t2.INDACTIEF, t2.GEBR_ORA, t2.SUBJECTNR, t2.TELNR, t2.FAXNR, t2.LOCATIE, t2.EMAIL, t2.VRIJ_VELD, t2.IND_DW, t2.EXE_USER, t2.HASHEE, t2.IND_JURIST)"
        
            Write-Verbose -Verbose $OracleQueryCreate
            
            $mduDir = $($account.GEBR_ORA).ToLower() + "\ACC\InFiles"
            $OracleQueryMDU = "INSERT INTO MDU_GEBRUIKER (MDU_USER, MDU_IN_DIR, ORA_USER) VALUES ('$($account.GEBRCODE)', '$mduDir', '$($account.GEBR_ORA)')"
        
            Write-Verbose -Verbose $OracleQueryMDU
            
            $OracleCmd.CommandText = $OracleQueryMDU
            $OracleCmd.ExecuteNonQuery() | Out-Null

            $OracleCmd.CommandText = $OracleQueryCreate
            $OracleCmd.ExecuteNonQuery() | Out-Null
			
            $OracleQueryPwdReset = "ALTER USER $($account.GEBR_ORA) IDENTIFIED BY $($account.PASSWORD)"
        
            Write-Verbose -Verbose $OracleQueryPwdReset
            
            $OracleCmd.CommandText = $OracleQueryPwdReset
            $OracleCmd.ExecuteNonQuery() | Out-Null

            Write-Verbose -Verbose "Successfully performed Oracle creation query."

            $success = $True;
            $auditMessage = " succesfully";   
        }
		
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
	AccountReference= $account.GEBRCODE;
	AuditLogs = $auditLogs;
    Account = $account;

    # Optionally return data for use in other systems
    ExportData = [PSCustomObject]@{
        GEBRCODE = $account.GEBRCODE;
        GEBR_ORA = $account.GEBR_ORA;
    };
    
};
Write-Output $result | ConvertTo-Json -Depth 10;
