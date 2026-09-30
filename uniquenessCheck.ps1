#################################################
# HelloID-Conn-Prov-Target-Key2Belastingen-UniquenessCheck
# PowerShell V2
#################################################

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

function Resolve-OracleError {
    [CmdletBinding()]
    param ([Parameter(Mandatory)][object]$ErrorObject)
    $result = [PSCustomObject]@{
        ScriptLineNumber = $ErrorObject.InvocationInfo.ScriptLineNumber
        Line             = $ErrorObject.InvocationInfo.Line
        ErrorDetails     = $ErrorObject.Exception.Message
        FriendlyMessage  = $ErrorObject.Exception.Message
    }
    if ($ErrorObject.Exception.InnerException) {
        $result.FriendlyMessage = $ErrorObject.Exception.InnerException.Message
    }
    Write-Output $result
}
#endregion functions

try {
    $gebrCode = $actionContext.Data.GEBRCODE
    if ([string]::IsNullOrEmpty($gebrCode)) {
        throw 'The mapped GEBRCODE is empty.'
    }

    $connectionString = "Data Source=$($actionContext.Configuration.DataSource)"
    $splatNewOracleConnection = @{
        ConnectionString = $connectionString
        Username         = $actionContext.Configuration.Username
        Password         = $actionContext.Configuration.Password
    }
    $actionMessage = 'opening Oracle connection'
    $connection = New-OracleConnection @splatNewOracleConnection

    $actionMessage = "checking uniqueness of GEBRCODE [$gebrCode]"
    $queryCheckUniqueGebrCode = "
    SELECT
        GEBRCODE
    FROM WMS_GEBRCODE
    WHERE GEBRCODE = '$($gebrCode.Replace("'", "''"))'
    "
    $splatQueryCheckUniqueGebrCode = @{
        Connection = $connection
        Query      = $queryCheckUniqueGebrCode
        NonQuery   = $false
    }
    $accounts = @(Invoke-OracleQuery @splatQueryCheckUniqueGebrCode)

    if (($accounts | Measure-Object).Count -gt 1) {
        throw "Multiple WMS_GEBRCODE records found with GEBRCODE [$gebrCode]."
    }
    elseif (($accounts | Measure-Object).Count -eq 1) {
        if (-not [string]::IsNullOrEmpty($actionContext.References.Account) -and $accounts[0].GEBRCODE -eq $actionContext.References.Account) {
            Write-Information "GEBRCODE [$gebrCode] belongs to the current account and is therefore valid."
        }
        else {
            Write-Information "GEBRCODE [$gebrCode] is not unique."
            [void]$outputContext.NonUniqueFields.Add('GEBRCODE')
        }
    }
    else {
        Write-Information "GEBRCODE [$gebrCode] is unique."
    }

    $outputContext.Success = $true
}
catch {
    $outputContext.Success = $false
    $ex = $PSItem
    if ($ex.Exception.GetBaseException().GetType().FullName -eq 'System.Data.OracleClient.OracleException') {
        $errorObj = Resolve-OracleError -ErrorObject $ex
        $errorMessage = "Error $actionMessage. Error: $($errorObj.FriendlyMessage)"
        Write-Warning "Error at Line [$($errorObj.ScriptLineNumber)]: $($errorObj.Line). Error: $($errorObj.ErrorDetails)"
    }
    else {
        $errorMessage = "Error $actionMessage. Error: $($ex.Exception.Message)"
        Write-Warning "Error at Line [$($ex.InvocationInfo.ScriptLineNumber)]: $($ex.InvocationInfo.Line). Error: $($ex.Exception.Message)"
    }
    Write-Error $errorMessage
}
finally {
    if ($connection -and $connection.State -eq 'Open') {
        $connection.Close()
    }
    if ($connection) {
        $connection.Dispose()
    }
    $outputContext.NonUniqueFields = @($outputContext.NonUniqueFields | Sort-Object -Unique)
}
