#################################################
# HelloID-Conn-Prov-Target-Key2Belastingen-Create
# PowerShell V2
#
# FIT FOR PURPOSE (FFP): optional MDU_GEBRUIKER record (ManageMduUser)
# Creating an MDU_GEBRUIKER record comes from a previous implementation and is disabled by default. It is not certain
# that every implementation uses MDU_GEBRUIKER or this record layout. Review before enabling it.
# See README.md, section "Fit For Purpose (FFP)".
#################################################

# Enable TLS1.2
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor [System.Net.SecurityProtocolType]::Tls12

#region functions
function New-OracleConnection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $ConnectionString,

        [Parameter()]
        [string]
        $Username,

        [Parameter()]
        [string]
        $Password
    )
    try {
        $oracleAssembly = [Reflection.Assembly]::LoadWithPartialName('System.Data.OracleClient')
        if ($null -eq $oracleAssembly) {
            throw 'System.Data.OracleClient could not be loaded. Verify that the action runs in Windows PowerShell 5.1 and the Oracle Client matches the PowerShell architecture.'
        }
        if (-not [string]::IsNullOrEmpty($Username) -and -not [string]::IsNullOrEmpty($Password)) {
            $oracleConnectionString = "$ConnectionString;User Id=$Username;Password=$Password;"
        }
        elseif ([string]::IsNullOrEmpty($Username) -and [string]::IsNullOrEmpty($Password)) {
            if ($ConnectionString -notmatch '(?i)(^|;)\s*Integrated Security\s*=\s*(yes|true)\s*(;|$)') {
                throw "Configure both Username and Password, or add 'Integrated Security=yes' to the configured connection string to use the Windows account running the HelloID Agent."
            }
            $oracleConnectionString = $ConnectionString
        }
        else {
            throw "Configure both Username and Password, or leave both empty and add 'Integrated Security=yes' to the configured connection string."
        }
        $connection = [System.Data.OracleClient.OracleConnection]::new($oracleConnectionString)
        $connection.Open()
        Write-Verbose 'Successfully connected to Oracle database'
        Write-Output $connection
    }
    catch {
        $PSCmdlet.ThrowTerminatingError($_)
    }
}

function Invoke-OracleQuery {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [System.Data.OracleClient.OracleConnection]
        $Connection,

        [Parameter(Mandatory)]
        [string]
        $Query,

        [Parameter(Mandatory)]
        [bool]
        $NonQuery
    )
    $command = $Connection.CreateCommand()
    $command.CommandText = $Query
    if ($NonQuery) {
        Write-Output $command.ExecuteNonQuery()
    }
    else {
        $adapter = [System.Data.OracleClient.OracleDataAdapter]::new($command)
        $dataSet = [System.Data.DataSet]::new()
        [void]$adapter.Fill($dataSet)
        Write-Output ($dataSet.Tables[0] | Select-Object -Property * -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors)
    }
}

function ConvertTo-OracleSqlLiteral {
    [CmdletBinding()]
    param (
        [Parameter()]
        [AllowNull()]
        $Value
    )

    if ($null -eq $Value -or [string]::IsNullOrEmpty([string]$Value)) {
        Write-Output 'NULL'
    }
    else {
        Write-Output "'$(([string]$Value).Replace("'", "''"))'"
    }
}

function Resolve-OracleError {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [object]
        $ErrorObject
    )
    process {
        $oracleErrorObj = [PSCustomObject]@{
            ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
            Line             = $ErrorObject.InvocationInfo.Line
            ErrorDetails     = $ErrorObject.Exception.Message
            FriendlyMessage  = $ErrorObject.Exception.Message
        }
        if ($ErrorObject.Exception.InnerException) {
            $oracleErrorObj.FriendlyMessage = $ErrorObject.Exception.InnerException.Message
        }
        Write-Output $oracleErrorObj
    }
}
#endregion functions

try {
    # Initial Assignments
    $outputContext.AccountReference = 'Currently not available'

    $connectionString = "Data Source=$($actionContext.Configuration.DataSource)"
    $splatNewOracleConnection = @{
        ConnectionString = $connectionString
        Username         = $actionContext.Configuration.Username
        Password         = $actionContext.Configuration.Password
    }
    $actionMessage = 'opening Oracle connection'
    $connection = New-OracleConnection @splatNewOracleConnection

    $outputFields = @($outputContext.Data.PSObject.Properties.Name | Where-Object { $_ })

    # Validate correlation configuration
    # Correlation field is expected to be GEBR_ORA, matched against the Oracle username created by the
    # Key2BelastingenOracle connector (that account must already exist - see README).
    $actionMessage = 'validating correlation configuration'
    if ($actionContext.CorrelationConfiguration.Enabled) {
        $correlationField = $actionContext.CorrelationConfiguration.AccountField
        $correlationValue = $actionContext.CorrelationConfiguration.PersonFieldValue

        if ([string]::IsNullOrEmpty($($correlationField))) {
            throw 'Correlation is enabled but not configured correctly'
        }
        if ([string]::IsNullOrEmpty($($correlationValue))) {
            throw 'Correlation is enabled but [accountFieldValue] is empty. Please make sure it is correctly mapped. This most likely means the Key2BelastingenOracle account does not (yet) exist for this person.'
        }

        $actionMessage = "querying WMS_GEBRCODE where $($correlationField) = [$($correlationValue)]"
        $queryCorrelateAccount = "
        SELECT
            $((@('GEBRCODE') + $outputFields | Select-Object -Unique) -join ', ')
        FROM WMS_GEBRCODE
        WHERE $correlationField = '$($correlationValue.Replace("'", "''"))'
        "
        $splatQueryCorrelateAccount = @{
            Connection = $connection
            Query      = $queryCorrelateAccount
            NonQuery   = $false
        }
        $correlatedAccount = Invoke-OracleQuery @splatQueryCorrelateAccount
        Write-Information "Queried WMS_GEBRCODE where $($correlationField) = [$($correlationValue)]. Result count: $(($correlatedAccount | Measure-Object).Count)"
    }

    # Determine action
    $actionMessage = 'determining action'
    if (($correlatedAccount | Measure-Object).Count -eq 0) {
        $action = 'CreateAccount'
    }
    elseif (($correlatedAccount | Measure-Object).Count -eq 1) {
        $action = 'CorrelateAccount'
    }
    else {
        $action = 'MultipleFound'
    }
    Write-Information "Determined action: [$action]"

    switch ($action) {
        'CreateAccount' {
            $actionMessage = "creating WMS_GEBRCODE record for Oracle user [$($actionContext.Data.GEBR_ORA)]"

            if ([string]::IsNullOrEmpty($actionContext.Data.GEBRCODE)) {
                throw 'The mapped GEBRCODE is empty. Verify the field mapping and the dependent Oracle account.'
            }
            if ($actionContext.Data.GEBRCODE.Length -gt 6) {
                throw "The mapped GEBRCODE [$($actionContext.Data.GEBRCODE)] exceeds the maximum length of 6 characters."
            }

            $gebrCode = $actionContext.Data.GEBRCODE.ToUpper()
            $actionContext.Data.GEBRCODE = $gebrCode
            $accountColumns = $actionContext.Data.PSObject.Properties.Name -join ', '
            $accountValues = ($actionContext.Data.PSObject.Properties.Value | ForEach-Object { ConvertTo-OracleSqlLiteral -Value $_ }) -join ', '
            $queryCreateAccount = "
            INSERT INTO WMS_GEBRCODE
                ($accountColumns)
            VALUES
                ($accountValues)
            "
            $splatQueryCreateAccount = @{
                Connection = $connection
                Query      = $queryCreateAccount
                NonQuery   = $true
            }

            if (-not ($actionContext.DryRun -eq $true)) {
                [void](Invoke-OracleQuery @splatQueryCreateAccount)

                # FIT FOR PURPOSE (FFP): the MDU_GEBRUIKER record layout (MDU_USER, MDU_IN_DIR, ORA_USER) comes from a previous implementation; review before enabling ManageMduUser.
                if ($actionContext.Configuration.ManageMduUser -eq $true) {
                    $mduUser = ConvertTo-OracleSqlLiteral -Value $gebrCode
                    $mduInDirectory = ConvertTo-OracleSqlLiteral -Value "$($actionContext.Data.GEBR_ORA.ToLower())$($actionContext.Configuration.MduInDirectorySuffix)"
                    $mduOracleUser = ConvertTo-OracleSqlLiteral -Value $actionContext.Data.GEBR_ORA
                    $queryCreateMduUser = "
                    INSERT INTO MDU_GEBRUIKER
                        (MDU_USER, MDU_IN_DIR, ORA_USER)
                    VALUES
                        ($mduUser, $mduInDirectory, $mduOracleUser)
                    "
                    $splatQueryCreateMduUser = @{
                        Connection = $connection
                        Query      = $queryCreateMduUser
                        NonQuery   = $true
                    }
                    [void](Invoke-OracleQuery @splatQueryCreateMduUser)
                }

                $outputContext.AccountReference = $gebrCode

                $actionMessage = "querying created WMS_GEBRCODE record [$gebrCode]"
                $queryGetCreatedAccount = "
                SELECT
                    $((@('GEBRCODE') + $outputFields | Select-Object -Unique) -join ', ')
                FROM WMS_GEBRCODE
                WHERE GEBRCODE = '$gebrCode'
                "
                $splatQueryGetCreatedAccount = @{
                    Connection = $connection
                    Query      = $queryGetCreatedAccount
                    NonQuery   = $false
                }
                $createdAccount = Invoke-OracleQuery @splatQueryGetCreatedAccount
                if (($createdAccount | Measure-Object).Count -ne 1) {
                    throw "Could not retrieve the created WMS_GEBRCODE record [$gebrCode]."
                }
                $outputContext.Data = ($createdAccount | Select-Object -Property $outputFields | ConvertTo-Json -Depth 10 | ConvertFrom-Json)

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = 'CreateAccount'
                        Message = "Created WMS_GEBRCODE record [$gebrCode] for Oracle user [$($actionContext.Data.GEBR_ORA)]. AccountReference is: [$($outputContext.AccountReference)]"
                        IsError = $false
                    })
            }
            else {
                Write-Information "[DryRun] Would create WMS_GEBRCODE record [$gebrCode] for Oracle user [$($actionContext.Data.GEBR_ORA)]"
            }
            break
        }

        'CorrelateAccount' {
            $actionMessage = "correlating to WMS_GEBRCODE record on field: [$($correlationField)] with value: [$($correlationValue)]"
            $outputContext.AccountReference = $correlatedAccount.GEBRCODE
            $outputContext.Data = ($correlatedAccount | Select-Object -Property $outputFields | ConvertTo-Json -Depth 10 | ConvertFrom-Json)

            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = 'CorrelateAccount'
                    Message = "Correlated to WMS_GEBRCODE record [$($correlatedAccount.GEBRCODE)] on field: [$($correlationField)] with value: [$($correlationValue)]"
                    IsError = $false
                })
            $outputContext.AccountCorrelated = $true
            break
        }

        'MultipleFound' {
            throw "Multiple WMS_GEBRCODE records found where $($correlationField) = [$($correlationValue)]. Please correct this so the accounts are unique."
        }
    }

    $outputContext.Success = $true
}
catch {
    $outputContext.Success = $false
    $ex = $PSItem
    if ($ex.Exception.GetBaseException().GetType().FullName -eq 'System.Data.OracleClient.OracleException') {
        $errorObj = Resolve-OracleError -ErrorObject $ex
        $warningMessage = "Error at Line '$($errorObj.ScriptLineNumber)': $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
        $errorMessage = "Error $($actionMessage). Error: $($errorObj.FriendlyMessage)"
    }
    else {
        $warningMessage = "Error at Line '$($ex.InvocationInfo.ScriptLineNumber)': $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
        $errorMessage = "Error $($actionMessage). Error: $($ex.Exception.Message)"
    }
    Write-Warning $warningMessage

    $outputContext.AuditLogs.Add([PSCustomObject]@{
            Message = $errorMessage
            IsError = $true
        })
}
finally {
    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
        Write-Verbose 'Successfully disconnected from Oracle database'
    }
    if ($connection) {
        $connection.Dispose()
    }
}
