#################################################
# HelloID-Conn-Prov-Target-Key2Belastingen-Delete
# PowerShell V2
#
# FIT FOR PURPOSE (FFP): optional delete stored procedure (DeleteStoredProcedureName)
# Calling a customer-specific stored procedure on delete was built for one implementation. Not every implementation
# has such a procedure, and what it does differs per customer. Review the procedure before configuring it.
# Without a procedure, delete performs a generic soft-delete via the field mapping.
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
    $actionMessage = 'verifying account reference'
    if ([string]::IsNullOrEmpty($($actionContext.References.Account))) {
        throw 'The account reference could not be found'
    }

    $connectionString = "Data Source=$($actionContext.Configuration.DataSource)"
    $splatNewOracleConnection = @{
        ConnectionString = $connectionString
        Username         = $actionContext.Configuration.Username
        Password         = $actionContext.Configuration.Password
    }
    $actionMessage = 'opening Oracle connection'
    $connection = New-OracleConnection @splatNewOracleConnection

    # FIT FOR PURPOSE (FFP): DeleteStoredProcedureName is an optional extension, built for the specific requirements of one implementation that uses a customer-specific stored procedure.
    # It is empty by default: delete then performs a soft-delete by updating the fields mapped for the Delete action (e.g. INDACTIEF = N).
    # A delete stored procedure is not part of Key2 Belastingen itself and it is not certain that this approach applies to every implementation.
    # Before configuring one, the consultant must review what the procedure does and adjust this script where needed.
    $storedProcedureName = "$($actionContext.Configuration.DeleteStoredProcedureName)".Trim()
    $useStoredProcedure = -not [string]::IsNullOrEmpty($storedProcedureName)

    $outputFields = @($outputContext.Data.PSObject.Properties.Name | Where-Object { $_ })
    # Governance reconciliation resolutions run without person context, so the Delete field mapping is not available.
    # Without a stored procedure, apply the same value as the supplied field mapping: INDACTIEF = N.
    if ($actionContext.ReconciliationOrigin -eq 'reconciliation') {
        if (-not $useStoredProcedure) {
            $actionContext.Data = [PSCustomObject]@{
                INDACTIEF = 'N'
            }
        }
        if (($outputFields | Measure-Object).Count -eq 0) {
            $outputFields = @('GEBRCODE', 'GEBR_ORA', 'INDACTIEF', 'EMAIL')
        }
        Write-Information "Reconciliation mode: delete (stored procedure: [$useStoredProcedure], output fields: $($outputFields -join ', '))"
    }

    $actionMessage = 'verifying mapped fields'
    $actionFields = @($actionContext.Data.PSObject.Properties.Name | Where-Object { $_ -and $_ -notin @('GEBRCODE', 'GEBR_ORA') })
    if (-not $useStoredProcedure) {
        if (($actionFields | Measure-Object).Count -eq 0) {
            throw 'No fields are mapped for the Delete action. Map at least INDACTIEF with value N, or configure DeleteStoredProcedureName.'
        }
        $outputFields = @(@('GEBRCODE') + $outputFields + $actionFields | Select-Object -Unique)
    }

    $actionMessage = "querying WMS_GEBRCODE where GEBRCODE = [$($actionContext.References.Account)]"
    $queryGetAccount = "
    SELECT
        $((@('GEBRCODE', 'GEBR_ORA') + $outputFields | Select-Object -Unique) -join ', ')
    FROM WMS_GEBRCODE
    WHERE GEBRCODE = '$($actionContext.References.Account)'
    "
    $splatQueryGetAccount = @{
        Connection = $connection
        Query      = $queryGetAccount
        NonQuery   = $false
    }
    $correlatedAccount = Invoke-OracleQuery @splatQueryGetAccount

    $actionMessage = 'determining action'
    if (($correlatedAccount | Measure-Object).Count -eq 1) {
        $outputContext.PreviousData = ($correlatedAccount | Select-Object -Property $outputFields | ConvertTo-Json -Depth 10 | ConvertFrom-Json)
        $outputContext.Data = ($correlatedAccount | Select-Object -Property $outputFields | ConvertTo-Json -Depth 10 | ConvertFrom-Json)

        if ($useStoredProcedure) {
            # The effect of the stored procedure is unknown to HelloID, so it is always executed (also for an already inactive account)
            $action = 'DeleteAccountByStoredProcedure'
        }
        else {
            $desiredProperties = foreach ($fieldName in $actionFields) {
                [PSCustomObject]@{
                    Name  = $fieldName
                    Value = if ([string]::IsNullOrEmpty([string]$actionContext.Data.$fieldName)) { '' } else { [string]$actionContext.Data.$fieldName }
                }
                $outputContext.Data.$fieldName = $actionContext.Data.$fieldName
            }
            $currentProperties = foreach ($fieldName in $actionFields) {
                [PSCustomObject]@{
                    Name  = $fieldName
                    Value = if ([string]::IsNullOrEmpty([string]$correlatedAccount.$fieldName)) { '' } else { [string]$correlatedAccount.$fieldName }
                }
            }
            $propertiesChanged = @(Compare-Object -ReferenceObject $currentProperties -DifferenceObject $desiredProperties -Property Name, Value -PassThru | Where-Object SideIndicator -eq '=>' | Select-Object -ExpandProperty Name)

            if (($propertiesChanged | Measure-Object).Count -gt 0) {
                $action = 'DeleteAccount'
            }
            else {
                $action = 'NoChanges'
            }
        }
    }
    elseif (($correlatedAccount | Measure-Object).Count -gt 1) {
        $action = 'MultipleFound'
    }
    else {
        $action = 'NotFound'
    }
    Write-Information "Determined action: [$action]"

    switch ($action) {
        'DeleteAccount' {
            $actionMessage = "deleting (soft-delete) WMS_GEBRCODE record [$($actionContext.References.Account)]. Properties changed: [$($propertiesChanged -join ', ')]"
            $setClauses = foreach ($fieldName in $propertiesChanged) {
                $sqlValue = ConvertTo-OracleSqlLiteral -Value $actionContext.Data.$fieldName
                "$fieldName = $sqlValue"
            }
            $queryDeleteAccount = "
            UPDATE WMS_GEBRCODE
            SET $($setClauses -join ', ')
            WHERE GEBRCODE = '$($actionContext.References.Account)'
            "
            $splatQueryDeleteAccount = @{
                Connection = $connection
                Query      = $queryDeleteAccount
                NonQuery   = $true
            }

            if (-not ($actionContext.DryRun -eq $true)) {
                [void](Invoke-OracleQuery @splatQueryDeleteAccount)

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = 'DeleteAccount'
                        Message = "Deleted (soft-delete) WMS_GEBRCODE record [$($actionContext.References.Account)]. Properties changed: [$($propertiesChanged -join ', ')]"
                        IsError = $false
                    })
            }
            else {
                Write-Information "[DryRun] Would delete (soft-delete) WMS_GEBRCODE record [$($actionContext.References.Account)]. Properties changed: [$($propertiesChanged -join ', ')]"
            }
            break
        }

        'DeleteAccountByStoredProcedure' {
            $actionMessage = "validating prerequisites for stored procedure [$storedProcedureName] for WMS_GEBRCODE record [$($actionContext.References.Account)]"
            $gebrOra = "$($correlatedAccount.GEBR_ORA)"
            if ([string]::IsNullOrEmpty($gebrOra)) {
                throw "Cannot execute stored procedure [$storedProcedureName] for WMS_GEBRCODE record [$($actionContext.References.Account)] because GEBR_ORA is empty."
            }

            $actionMessage = "executing stored procedure [$storedProcedureName] for Oracle user [$gebrOra]"
            $queryDeleteAccount = "
            BEGIN
                $storedProcedureName('$($gebrOra.Replace("'", "''"))');
            END;
            "
            $splatQueryDeleteAccount = @{
                Connection = $connection
                Query      = $queryDeleteAccount
                NonQuery   = $true
            }

            if (-not ($actionContext.DryRun -eq $true)) {
                [void](Invoke-OracleQuery @splatQueryDeleteAccount)

                $outputContext.AuditLogs.Add([PSCustomObject]@{
                        Action  = 'DeleteAccount'
                        Message = "Executed stored procedure [$storedProcedureName] for Oracle user [$gebrOra]. The stored procedure does not return its own logging to HelloID; check the logging of the stored procedure in the database for the actions it performed."
                        IsError = $false
                    })

                # Re-query only to return the current account data; the result is not validated, because the effect of the stored procedure is unknown to HelloID
                $actionMessage = "querying WMS_GEBRCODE record [$($actionContext.References.Account)] after executing stored procedure [$storedProcedureName]"
                $deletedAccount = Invoke-OracleQuery @splatQueryGetAccount
                if (($deletedAccount | Measure-Object).Count -eq 1) {
                    $outputContext.Data = ($deletedAccount | Select-Object -Property $outputFields | ConvertTo-Json -Depth 10 | ConvertFrom-Json)
                }
            }
            else {
                Write-Information "[DryRun] Would execute stored procedure [$storedProcedureName] for Oracle user [$gebrOra]"
            }
            break
        }

        'NoChanges' {
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = 'DeleteAccount'
                    Message = "Skipped deleting WMS_GEBRCODE record [$($actionContext.References.Account)]. Reason: Already deleted (no changes)."
                    IsError = $false
                })
            break
        }

        'MultipleFound' {
            throw "Multiple WMS_GEBRCODE records found with GEBRCODE: [$($actionContext.References.Account)]. Please correct this so the accounts are unique."
        }

        'NotFound' {
            $outputContext.AuditLogs.Add([PSCustomObject]@{
                    Action  = 'DeleteAccount'
                    Message = "Skipped deleting WMS_GEBRCODE record [$($actionContext.References.Account)]. Reason: Account does not exist."
                    IsError = $false
                })
            break
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
