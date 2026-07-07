[CmdletBinding()]
param(
    [string]$ServerInstance = ".",
    [string]$OutputPath = "",
    [int]$QueryTimeoutSeconds = 15,
    [string]$SqlUsername = "",
    [string]$SqlPassword = "",
    [string]$SqlCredentialPath = "",
    [switch]$UseSqlAuthentication,
    [switch]$SkipIndexFragmentation,
    [switch]$OpenReport
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = "C:\SQL Reports"
}

if ([string]::IsNullOrWhiteSpace($SqlCredentialPath)) {
    $scriptRoot = if ([string]::IsNullOrWhiteSpace($PSScriptRoot)) { (Get-Location).Path } else { $PSScriptRoot }
    $SqlCredentialPath = Join-Path -Path $scriptRoot -ChildPath "sql-prod-credential.xml"
}

if ($UseSqlAuthentication) {
    if ([string]::IsNullOrWhiteSpace($SqlUsername) -or [string]::IsNullOrWhiteSpace($SqlPassword)) {
        if (-not (Test-Path -LiteralPath $SqlCredentialPath)) {
            throw "SQL credential file not found: $SqlCredentialPath. Create it once with: `$cred = Get-Credential; `$cred | Export-Clixml -Path `"$SqlCredentialPath`""
        }

        $sqlCredential = Import-Clixml -LiteralPath $SqlCredentialPath
        if ($null -eq $sqlCredential -or $null -eq $sqlCredential.GetNetworkCredential()) {
            throw "SQL credential file is invalid or cannot be decrypted by this Windows account: $SqlCredentialPath"
        }

        $networkCredential = $sqlCredential.GetNetworkCredential()
        if ([string]::IsNullOrWhiteSpace($SqlUsername)) {
            $SqlUsername = $networkCredential.UserName
        }

        $securePassword = $sqlCredential.Password
        $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($securePassword)
        try {
            $SqlPassword = [Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
        } finally {
            if ($bstr -ne [IntPtr]::Zero) {
                [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
            }
        }
    }
}

function Get-ServerCandidates {
    param([string]$RequestedServer)

    $candidates = [System.Collections.Generic.List[string]]::new()

    foreach ($value in @(
        $RequestedServer,
        ".",
        "(local)",
        "localhost",
        $env:COMPUTERNAME,
        "127.0.0.1",
        "np:\\.\pipe\sql\query"
    )) {
        if (-not [string]::IsNullOrWhiteSpace($value) -and -not $candidates.Contains($value)) {
            $candidates.Add($value)
        }
    }

    try {
        $sqlServices = Get-Service -ErrorAction Stop | Where-Object { $_.Name -like "MSSQL*" -and $_.Name -ne "MSSQLFDLauncher" }
        foreach ($service in $sqlServices) {
            if ($service.Name -eq "MSSQLSERVER") {
                $instanceName = $env:COMPUTERNAME
            } elseif ($service.Name -like "MSSQL`$*") {
                $instanceSuffix = $service.Name.Substring(6)
                $instanceName = "$($env:COMPUTERNAME)\$instanceSuffix"
            } else {
                continue
            }

            if (-not $candidates.Contains($instanceName)) {
                $candidates.Add($instanceName)
            }
        }
    } catch {
        Write-Verbose "Unable to enumerate SQL services: $($_.Exception.Message)"
    }

    return $candidates.ToArray()
}

function New-SqlConnection {
    param(
        [string]$ServerName,
        [string]$Database = "master"
    )

    if (-not [string]::IsNullOrWhiteSpace($SqlUsername) -and -not [string]::IsNullOrWhiteSpace($SqlPassword)) {
        $connectionString = "Server=$ServerName;Database=$Database;User ID=$SqlUsername;Password=$SqlPassword;Encrypt=no;TrustServerCertificate=true;Connect Timeout=3;Application Name=MSSQLDailyHealthCheck;"
    } else {
        $connectionString = "Server=$ServerName;Database=$Database;Integrated Security=true;Encrypt=no;TrustServerCertificate=true;Connect Timeout=3;Application Name=MSSQLDailyHealthCheck;"
    }

    return [System.Data.SqlClient.SqlConnection]::new($connectionString)
}

function Open-WorkingConnection {
    param([string[]]$Candidates)

    $lastError = $null
    $messages = [System.Collections.Generic.List[string]]::new()

    foreach ($candidate in $Candidates) {
        try {
            $connection = New-SqlConnection -ServerName $candidate
            $connection.Open()
            return [PSCustomObject]@{
                Connection = $connection
                Target     = $candidate
            }
        } catch {
            $lastError = $_
            $messages.Add(("[{0}] {1}" -f $candidate, $_.Exception.Message))
            Write-Verbose "Connection failed for [$candidate]: $($_.Exception.Message)"
        }
    }

    $detail = $messages -join "; "
    if ($detail -match "Cannot generate SSPI context|target principal name") {
        throw "Unable to connect to SQL Server using Windows authentication. The local SQL service is running, but authentication is failing with an SSPI/SPN error. Tried: $detail"
    }

    if ($null -ne $lastError) {
        throw "Unable to connect to any SQL Server target. Tried: $detail"
    }

    throw "Unable to connect to any SQL Server target."
}

function Invoke-QueryTable {
    param(
        [string]$ServerName,
        [string]$Query,
        [string]$Database = "master",
        [int]$TimeoutSeconds = 15
    )

    $connection = $null

    try {
        $connection = New-SqlConnection -ServerName $ServerName -Database $Database
        $connection.Open()
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $command.CommandTimeout = $TimeoutSeconds

        $adapter = [System.Data.SqlClient.SqlDataAdapter]::new($command)
        $dataSet = [System.Data.DataSet]::new()
        [void]$adapter.Fill($dataSet)

        if ($dataSet.Tables.Count -gt 0) {
            return ,$dataSet.Tables[0]
        }

        return ,[System.Data.DataTable]::new()
    } finally {
        if ($null -ne $connection -and $connection.State -ne [System.Data.ConnectionState]::Closed) {
            $connection.Close()
        }
    }
}

function Convert-DataTableToObjects {
    param([System.Data.DataTable]$Table)

    $rows = @()

    if ($null -eq $Table) {
        return $rows
    }

    foreach ($dr in $Table.Rows) {
        $item = [ordered]@{}
        foreach ($column in $Table.Columns) {
            $item[$column.ColumnName] = $dr[$column.ColumnName]
        }
        $rows += [PSCustomObject]$item
    }

    return $rows
}

function Get-StatusClass {
    param([string]$Status)

    $normalizedStatus = if ($null -eq $Status) { "" } else { $Status.ToUpperInvariant() }

    switch ($normalizedStatus) {
        "CRITICAL" { return "critical" }
        "WARNING" { return "warning" }
        default { return "healthy" }
    }
}

function New-StatusBadge {
    param([string]$Status)

    $css = Get-StatusClass -Status $Status
    return "<span class=""badge $css"">$Status</span>"
}

function Escape-Html {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return ""
    }

    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Format-CellValue {
    param([object]$Value)

    if ($null -eq $Value -or $Value -is [System.DBNull]) {
        return ""
    }

    if ($Value -is [datetime]) {
        return $Value.ToString("yyyy-MM-dd HH:mm:ss")
    }

    return [string]$Value
}

function Get-ServerIpAddress {
    param([string]$HostName)

    if ([string]::IsNullOrWhiteSpace($HostName)) {
        return "Unavailable"
    }

    try {
        $addresses = [System.Net.Dns]::GetHostAddresses($HostName) |
            Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork } |
            ForEach-Object { $_.IPAddressToString } |
            Select-Object -Unique

        if ($addresses) {
            return ($addresses -join ", ")
        }
    } catch {
    }

    return "Unavailable"
}

function New-SectionTable {
    param(
        [string]$Title,
        [string]$Subtitle,
        [object[]]$Rows,
        [string[]]$Columns,
        [string]$EmptyMessage = "No data available.",
        [switch]$FullWidth,
        [switch]$Scrollable
    )

    $panelClass = if ($FullWidth) { "panel full" } else { "panel" }
    $tableWrapClass = if ($Scrollable) { "table-wrap scroll-y" } else { "table-wrap" }
    $subtitleHtml = if ([string]::IsNullOrWhiteSpace($Subtitle)) { "" } else { "<span>$([System.Net.WebUtility]::HtmlEncode($Subtitle))</span>" }

    if (-not $Rows -or $Rows.Count -eq 0) {
        return @"
<section class="$panelClass">
    <div class="panel-head">
        <h2>$([System.Net.WebUtility]::HtmlEncode($Title))</h2>
        $subtitleHtml
    </div>
    <div class="panel-body">
        <div class="empty-state">$([System.Net.WebUtility]::HtmlEncode($EmptyMessage))</div>
    </div>
</section>
"@
    }

    $headerHtml = ($Columns | ForEach-Object { "<th>$([System.Net.WebUtility]::HtmlEncode($_))</th>" }) -join ""
    $rowHtml = foreach ($row in $Rows) {
        $rowStatus = if ($row.PSObject.Properties.Name -contains "Status") { [string]$row.Status } else { "HEALTHY" }
        $rowClass = "row-" + (Get-StatusClass -Status $rowStatus)
        $cells = foreach ($column in $Columns) {
            $value = if ($row.PSObject.Properties.Name -contains $column) { $row.$column } else { $null }
            if ($column -eq "Status") {
                "<td>$(New-StatusBadge -Status ([string]$value))</td>"
            } else {
                "<td>$([System.Net.WebUtility]::HtmlEncode((Format-CellValue -Value $value)))</td>"
            }
        }
        "<tr class=""$rowClass"">$($cells -join '')</tr>"
    }

    return @"
<section class="$panelClass">
    <div class="panel-head">
        <h2>$([System.Net.WebUtility]::HtmlEncode($Title))</h2>
        $subtitleHtml
    </div>
    <div class="panel-body">
        <div class="$tableWrapClass">
            <table>
                <thead>
                    <tr>$headerHtml</tr>
                </thead>
                <tbody>
                    $($rowHtml -join [Environment]::NewLine)
                </tbody>
            </table>
        </div>
    </div>
</section>
"@
}

function New-SummaryCard {
    param(
        [string]$Label,
        [string]$Value,
        [string]$Note,
        [string]$Status = "HEALTHY"
    )

    $css = Get-StatusClass -Status $Status
    return @"
<article class="summary-card status-$css">
    <div class="label">$([System.Net.WebUtility]::HtmlEncode($Label))</div>
    <div class="value">$([System.Net.WebUtility]::HtmlEncode($Value))</div>
    <div class="note">$([System.Net.WebUtility]::HtmlEncode($Note))</div>
</article>
"@
}

function Add-Status {
    param(
        [object[]]$Rows,
        [scriptblock]$Rule
    )

    foreach ($row in $Rows) {
        $status = & $Rule $row
        if ($row.PSObject.Properties.Name -contains "Status") {
            $row.Status = $status
        } else {
            $row | Add-Member -NotePropertyName Status -NotePropertyValue $status
        }
    }
}

function Get-OverallStatus {
    param([object[][]]$Groups)

    $statuses = @()
    foreach ($group in $Groups) {
        if ($group) {
            $statuses += $group | ForEach-Object { [string]$_.Status }
        }
    }

    if ($statuses -contains "CRITICAL") { return "CRITICAL" }
    if ($statuses -contains "WARNING") { return "WARNING" }
    return "HEALTHY"
}

function Try-Query {
    param(
        [string]$ServerName,
        [string]$Query,
        [string]$Database = "master",
        [string]$Name
    )

    try {
        $table = Invoke-QueryTable -ServerName $ServerName -Database $Database -Query $Query -TimeoutSeconds $QueryTimeoutSeconds
        return [PSCustomObject]@{
            Name  = $Name
            Table = $table
            Error = $null
        }
    } catch {
        return [PSCustomObject]@{
            Name  = $Name
            Table = $null
            Error = $_.Exception.Message
        }
    }
}

$templatePath = Join-Path $PSScriptRoot "mssql-daily-health-template.html"
if (-not (Test-Path -LiteralPath $templatePath)) {
    throw "Template file not found: $templatePath"
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null
}

$connectionInfo = Open-WorkingConnection -Candidates (Get-ServerCandidates -RequestedServer $ServerInstance)
$connectionInfo.Connection.Close()
$resolvedTarget = $connectionInfo.Target

Write-Host "Connected to SQL Server target: $resolvedTarget" -ForegroundColor Green

$serverInfoResult = Try-Query -ServerName $resolvedTarget -Name "ServerInfo" -Query @"
SELECT
    CAST(SERVERPROPERTY('ServerName') AS nvarchar(256)) AS ServerName,
    CAST(SERVERPROPERTY('MachineName') AS nvarchar(256)) AS MachineName,
    CAST(SERVERPROPERTY('ComputerNamePhysicalNetBIOS') AS nvarchar(256)) AS ActiveNode,
    CAST(SERVERPROPERTY('Edition') AS nvarchar(256)) AS Edition,
    CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(256)) AS ProductVersion,
    CAST(SERVERPROPERTY('ProductLevel') AS nvarchar(256)) AS ProductLevel,
    CAST(SERVERPROPERTY('IsClustered') AS int) AS IsClustered,
    CAST(SERVERPROPERTY('IsHadrEnabled') AS int) AS IsHadrEnabled,
    CAST(SERVERPROPERTY('EngineEdition') AS int) AS EngineEdition,
    sqlserver_start_time AS SqlStartTime
FROM sys.dm_os_sys_info;
"@

if ($null -eq $serverInfoResult.Table -or $serverInfoResult.Table.Rows.Count -eq 0) {
    throw "Unable to retrieve SQL Server metadata from $resolvedTarget."
}

$serverInfoRows = @(Convert-DataTableToObjects -Table $serverInfoResult.Table)
$serverInfo = $serverInfoRows | Select-Object -First 1

$databaseResult = Try-Query -ServerName $resolvedTarget -Name "DatabaseHealth" -Query @"
CREATE TABLE #LogShippingSecondary (
    DatabaseName sysname NOT NULL PRIMARY KEY
);

IF OBJECT_ID('msdb.dbo.log_shipping_monitor_secondary') IS NOT NULL
BEGIN
    INSERT INTO #LogShippingSecondary (DatabaseName)
    SELECT DISTINCT secondary_database
    FROM msdb.dbo.log_shipping_monitor_secondary
    WHERE secondary_database IS NOT NULL;
END;

SELECT
    d.name AS [Database],
    d.state_desc AS [State],
    d.recovery_model_desc AS [RecoveryModel],
    CAST(ISNULL(sz.SizeMB, 0) AS decimal(18,2)) AS [SizeMB],
    fullBackup.backup_finish_date AS [LastFullBackup],
    diffBackup.backup_finish_date AS [LastDiffBackup],
    logBackup.backup_finish_date AS [LastLogBackup],
    CASE
        WHEN ls.DatabaseName IS NOT NULL THEN 'Log Shipping Secondary'
        WHEN d.is_distributor = 1 THEN 'Replication Distributor'
        WHEN SERVERPROPERTY('IsHadrEnabled') = 1
             AND EXISTS (
                SELECT 1
                FROM sys.dm_hadr_database_replica_states drs
                WHERE drs.database_id = d.database_id
                  AND drs.is_local = 1
             ) THEN 'Always On / HA'
        ELSE 'Standalone / Primary'
    END AS [RoleHint]
FROM sys.databases d
LEFT JOIN (
    SELECT database_id, SUM(size) * 8.0 / 1024 AS SizeMB
    FROM sys.master_files
    GROUP BY database_id
) sz
    ON d.database_id = sz.database_id
LEFT JOIN #LogShippingSecondary ls
    ON d.name = ls.DatabaseName
OUTER APPLY (
    SELECT TOP (1) backup_finish_date
    FROM msdb.dbo.backupset b
    WHERE b.database_name = d.name
      AND b.type = 'D'
    ORDER BY backup_finish_date DESC
) fullBackup
OUTER APPLY (
    SELECT TOP (1) backup_finish_date
    FROM msdb.dbo.backupset b
    WHERE b.database_name = d.name
      AND b.type = 'I'
    ORDER BY backup_finish_date DESC
) diffBackup
OUTER APPLY (
    SELECT TOP (1) backup_finish_date
    FROM msdb.dbo.backupset b
    WHERE b.database_name = d.name
      AND b.type = 'L'
    ORDER BY backup_finish_date DESC
) logBackup
WHERE d.database_id > 4
ORDER BY d.name;
"@

$logResult = Try-Query -ServerName $resolvedTarget -Name "LogHealth" -Query @"
CREATE TABLE #LogSpace (
    [Database Name] sysname,
    [Log Size (MB)] decimal(18,2),
    [Log Space Used (%)] decimal(18,2),
    [Status] int
);

INSERT INTO #LogSpace
EXEC ('DBCC SQLPERF(LOGSPACE)');

SELECT
    [Database Name] AS [Database],
    [Log Size (MB)] AS [LogSizeMB],
    CAST(([Log Size (MB)] * [Log Space Used (%)] / 100.0) AS decimal(18,2)) AS [UsedMB],
    [Log Space Used (%)] AS [UsedPercent]
FROM #LogSpace
WHERE [Database Name] IN (SELECT name FROM sys.databases WHERE database_id > 4)
ORDER BY [Log Space Used (%)] DESC;
"@

$dbFileResult = Try-Query -ServerName $resolvedTarget -Name "DatabaseFileUsage" -Query @"
CREATE TABLE #DbFileUsage (
    [Database] sysname NOT NULL,
    [LogicalName] sysname NOT NULL,
    [SizeMB] decimal(18,2) NULL,
    [UsedMB] decimal(18,2) NULL,
    [FreeMB] decimal(18,2) NULL,
    [UsedPercent] decimal(18,2) NULL
);

DECLARE @databaseName sysname;
DECLARE @sql nvarchar(max);

DECLARE database_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.databases
WHERE database_id > 4
  AND state_desc = 'ONLINE'
  AND HAS_DBACCESS(name) = 1;

OPEN database_cursor;
FETCH NEXT FROM database_cursor INTO @databaseName;

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'USE ' + QUOTENAME(@databaseName) + N';
        INSERT INTO #DbFileUsage ([Database], [LogicalName], [SizeMB], [UsedMB], [FreeMB], [UsedPercent])
        SELECT
            DB_NAME() AS [Database],
            name AS [LogicalName],
            CAST(size * 8.0 / 1024 AS decimal(18,2)) AS [SizeMB],
            CAST(FILEPROPERTY(name, ''SpaceUsed'') * 8.0 / 1024 AS decimal(18,2)) AS [UsedMB],
            CAST((size - FILEPROPERTY(name, ''SpaceUsed'')) * 8.0 / 1024 AS decimal(18,2)) AS [FreeMB],
            CAST((FILEPROPERTY(name, ''SpaceUsed'') * 100.0) / NULLIF(size, 0) AS decimal(18,2)) AS [UsedPercent]
        FROM sys.database_files
        WHERE type = 0;';

    EXEC sys.sp_executesql @sql;
    FETCH NEXT FROM database_cursor INTO @databaseName;
END;

CLOSE database_cursor;
DEALLOCATE database_cursor;

SELECT TOP (25)
    [Database],
    [LogicalName],
    [SizeMB],
    [UsedMB],
    [FreeMB],
    [UsedPercent]
FROM #DbFileUsage
ORDER BY [UsedPercent] DESC, [SizeMB] DESC;
"@

$diskResult = Try-Query -ServerName $resolvedTarget -Name "DiskHealth" -Query @"
CREATE TABLE #DiskHealth (
    [Drive] nvarchar(260) NOT NULL,
    [TotalGB] decimal(18,2) NULL,
    [FreeGB] decimal(18,2) NULL,
    [FreePercent] decimal(18,2) NULL
);

BEGIN TRY
    INSERT INTO #DiskHealth ([Drive], [TotalGB], [FreeGB], [FreePercent])
    SELECT DISTINCT
        vs.volume_mount_point AS [Drive],
        CAST(vs.total_bytes / 1024.0 / 1024 / 1024 AS decimal(18,2)) AS [TotalGB],
        CAST(vs.available_bytes / 1024.0 / 1024 / 1024 AS decimal(18,2)) AS [FreeGB],
        CAST((vs.available_bytes * 100.0) / NULLIF(vs.total_bytes, 0) AS decimal(18,2)) AS [FreePercent]
    FROM sys.master_files mf
    CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs;
END TRY
BEGIN CATCH
END CATCH;

CREATE TABLE #FixedDrives (
    [DriveLetter] nvarchar(10) NOT NULL,
    [FreeMB] int NULL
);

BEGIN TRY
    INSERT INTO #FixedDrives ([DriveLetter], [FreeMB])
    EXEC master.dbo.xp_fixeddrives;
END TRY
BEGIN CATCH
END CATCH;

INSERT INTO #DiskHealth ([Drive], [TotalGB], [FreeGB], [FreePercent])
SELECT
    UPPER(LEFT(fd.DriveLetter, 1)) + N':\' AS [Drive],
    NULL AS [TotalGB],
    CAST(fd.FreeMB / 1024.0 AS decimal(18,2)) AS [FreeGB],
    NULL AS [FreePercent]
FROM #FixedDrives fd
WHERE NOT EXISTS (
    SELECT 1
    FROM #DiskHealth dh
    WHERE UPPER(LEFT(dh.[Drive], 2)) = UPPER(LEFT(fd.DriveLetter, 1)) + N':'
);

IF NOT EXISTS (SELECT 1 FROM #DiskHealth)
BEGIN
    INSERT INTO #DiskHealth ([Drive], [TotalGB], [FreeGB], [FreePercent])
    VALUES (N'Unavailable', NULL, NULL, NULL);
END;

SELECT
    [Drive],
    [TotalGB],
    [FreeGB],
    [FreePercent]
FROM #DiskHealth
ORDER BY [Drive];
"@

$blockingResult = Try-Query -ServerName $resolvedTarget -Name "Blocking" -Query @"
SELECT TOP (15)
    r.blocking_session_id AS [BlockingSessionID],
    r.session_id AS [SessionID],
    r.wait_type AS [WaitType],
    r.wait_time AS [WaitMs],
    r.wait_resource AS [WaitResource],
    DB_NAME(r.database_id) AS [Database],
    LEFT(REPLACE(REPLACE(t.text, CHAR(13), ' '), CHAR(10), ' '), 120) AS [Statement]
FROM sys.dm_exec_requests r
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.blocking_session_id <> 0
  AND r.session_id <> @@SPID
  AND r.session_id > 50
ORDER BY r.wait_time DESC;
"@

$jobsResult = Try-Query -ServerName $resolvedTarget -Name "FailedJobs" -Database "msdb" -Query @"
SELECT TOP (10)
    sj.name AS [JobName],
    msdb.dbo.agent_datetime(h.run_date, h.run_time) AS [RunDateTime],
    LEFT(h.message, 160) AS [Message]
FROM msdb.dbo.sysjobhistory h
INNER JOIN msdb.dbo.sysjobs sj
    ON h.job_id = sj.job_id
WHERE h.step_id = 0
  AND h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -48, GETDATE())
ORDER BY msdb.dbo.agent_datetime(h.run_date, h.run_time) DESC;
"@

$sqlErrorLogResult = Try-Query -ServerName $resolvedTarget -Name "SqlErrorLog" -Query @"
CREATE TABLE #SqlErrorLog (
    [LogDate] datetime NULL,
    [ProcessInfo] nvarchar(100) NULL,
    [Text] nvarchar(max) NULL
);

DECLARE @StartTime DATETIME;
SET @StartTime = DATEADD(HOUR, -48, GETDATE());

INSERT INTO #SqlErrorLog
EXEC master.dbo.xp_readerrorlog 0, 1, NULL, NULL, @StartTime, NULL, N'desc';

SELECT TOP (50)
    [LogDate],
    [ProcessInfo],
    LEFT(REPLACE(REPLACE([Text], CHAR(13), ' '), CHAR(10), ' '), 240) AS [Message]
FROM #SqlErrorLog
ORDER BY [LogDate] DESC;
"@

$agentResult = Try-Query -ServerName $resolvedTarget -Name "AgentStatus" -Query @"
SELECT
    servicename AS [ServiceName],
    startup_type_desc AS [StartupType],
    status_desc AS [ServiceStatus],
    last_startup_time AS [LastStartupTime]
FROM sys.dm_server_services
WHERE servicename LIKE 'SQL Server (%'
   OR servicename LIKE 'SQL Server Agent (%'
ORDER BY servicename;
"@

$haResult = Try-Query -ServerName $resolvedTarget -Name "AlwaysOn" -Query @"
IF SERVERPROPERTY('IsHadrEnabled') = 1
BEGIN
    SELECT
        DB_NAME(drs.database_id) AS [Database],
        ag.name AS [AvailabilityGroup],
        ar.replica_server_name AS [ReplicaServer],
        drs.synchronization_state_desc AS [SyncState],
        drs.synchronization_health_desc AS [SyncHealth],
        ars.role_desc AS [ReplicaRole],
        drs.database_state_desc AS [DatabaseState]
    FROM sys.dm_hadr_database_replica_states drs
    INNER JOIN sys.dm_hadr_availability_replica_states ars
        ON drs.replica_id = ars.replica_id
    INNER JOIN sys.availability_replicas ar
        ON ars.replica_id = ar.replica_id
    INNER JOIN sys.availability_groups ag
        ON ar.group_id = ag.group_id
    WHERE drs.is_local = 1
    ORDER BY DB_NAME(drs.database_id);
END
"@

$clusterResult = Try-Query -ServerName $resolvedTarget -Name "ClusterInfo" -Query @"
IF SERVERPROPERTY('IsClustered') = 1
BEGIN
    SELECT
        NodeName,
        status_description AS [ClusterNodeStatus],
        is_current_owner AS [IsCurrentOwner]
    FROM sys.dm_os_cluster_nodes
    ORDER BY is_current_owner DESC, NodeName;
END
"@

$logShippingResult = Try-Query -ServerName $resolvedTarget -Name "LogShipping" -Database "msdb" -Query @"
CREATE TABLE #LogShippingHealth (
    [Role] nvarchar(20),
    [PrimaryServer] sysname NULL,
    [SecondaryServer] sysname NULL,
    [Database] sysname NULL,
    [LastBackupDate] datetime NULL,
    [LastCopiedDate] datetime NULL,
    [LastRestoredDate] datetime NULL,
    [MinutesSinceLastBackup] int NULL,
    [MinutesSinceLastCopy] int NULL,
    [MinutesSinceLastRestore] int NULL
);

IF OBJECT_ID('msdb.dbo.log_shipping_monitor_primary') IS NOT NULL
BEGIN
    INSERT INTO #LogShippingHealth ([Role], [PrimaryServer], [Database], [LastBackupDate], [MinutesSinceLastBackup])
    SELECT
        'Primary',
        primary_server,
        primary_database,
        last_backup_date,
        DATEDIFF(MINUTE, last_backup_date, GETDATE())
    FROM msdb.dbo.log_shipping_monitor_primary;
END;

IF OBJECT_ID('msdb.dbo.log_shipping_monitor_secondary') IS NOT NULL
BEGIN
    INSERT INTO #LogShippingHealth ([Role], [PrimaryServer], [SecondaryServer], [Database], [LastCopiedDate], [LastRestoredDate], [MinutesSinceLastCopy], [MinutesSinceLastRestore])
    SELECT
        'Secondary',
        primary_server,
        secondary_server,
        secondary_database,
        last_copied_date,
        last_restored_date,
        DATEDIFF(MINUTE, last_copied_date, GETDATE()),
        DATEDIFF(MINUTE, last_restored_date, GETDATE())
    FROM msdb.dbo.log_shipping_monitor_secondary ls
    INNER JOIN master.sys.databases d
        ON d.name = ls.secondary_database
    WHERE d.state_desc IN ('ONLINE', 'RESTORING')
      AND UPPER(ls.secondary_server) IN (
          UPPER(CONVERT(sysname, @@SERVERNAME)),
          UPPER(CONVERT(sysname, SERVERPROPERTY('ServerName'))),
          UPPER(CONVERT(sysname, SERVERPROPERTY('MachineName')))
      )
      AND UPPER(ls.primary_server) NOT IN (
          UPPER(CONVERT(sysname, @@SERVERNAME)),
          UPPER(CONVERT(sysname, SERVERPROPERTY('ServerName'))),
          UPPER(CONVERT(sysname, SERVERPROPERTY('MachineName')))
      );
END;

SELECT *
FROM #LogShippingHealth
ORDER BY [Role], [Database];
"@

$replicationResult = Try-Query -ServerName $resolvedTarget -Name "ReplicationHealth" -Database "msdb" -Query @"
CREATE TABLE #ReplicationHealth (
    [Role] nvarchar(40) NULL,
    [Component] nvarchar(80) NULL,
    [Database] sysname NULL,
    [Publication] sysname NULL,
    [Subscriber] sysname NULL,
    [SubscriberDatabase] sysname NULL,
    [LatencySeconds] int NULL,
    [PendingCommands] int NULL,
    [AgentJob] sysname NULL,
    [LastRunDateTime] datetime NULL,
    [LastRunStatus] nvarchar(20) NULL,
    [Message] nvarchar(4000) NULL
);

IF DB_ID(N'distribution') IS NOT NULL
   AND OBJECT_ID(N'distribution.dbo.MSreplication_monitordata') IS NOT NULL
BEGIN
    BEGIN TRY
        DECLARE @monitorSql nvarchar(max) = N'
            INSERT INTO #ReplicationHealth ([Role], [Component], [Database], [Publication], [Subscriber], [SubscriberDatabase], [LatencySeconds], [PendingCommands], [LastRunStatus], [Message])
            SELECT
                CASE
                    WHEN md.subscriber IS NOT NULL AND md.subscriber <> @@SERVERNAME THEN N''Publisher''
                    WHEN md.subscriber = @@SERVERNAME THEN N''Subscriber''
                    ELSE N''Replication''
                END AS [Role],
                CASE md.agent_type
                    WHEN 1 THEN N''Snapshot Agent''
                    WHEN 2 THEN N''Log Reader Agent''
                    WHEN 3 THEN N''Distribution Agent''
                    WHEN 4 THEN N''Merge Agent''
                    WHEN 9 THEN N''Queue Reader Agent''
                    ELSE N''Replication Agent''
                END AS [Component],
                md.publisher_db AS [Database],
                md.publication AS [Publication],
                md.subscriber AS [Subscriber],
                md.subscriber_db AS [SubscriberDatabase],
                TRY_CONVERT(int, md.latency) AS [LatencySeconds],
                TRY_CONVERT(int, md.pendingcmdcount) AS [PendingCommands],
                CASE md.status
                    WHEN 1 THEN N''Started''
                    WHEN 2 THEN N''Succeeded''
                    WHEN 3 THEN N''In Progress''
                    WHEN 4 THEN N''Idle''
                    WHEN 5 THEN N''Retrying''
                    WHEN 6 THEN N''Failed''
                    ELSE CONVERT(nvarchar(20), md.status)
                END AS [LastRunStatus],
                LEFT(
                    CONCAT(
                        N''Latency threshold: '', COALESCE(CONVERT(nvarchar(20), md.latencythreshold), N''N/A''),
                        N''; last sync: '', COALESCE(CONVERT(nvarchar(30), md.last_distsync, 120), N''N/A'')
                    ),
                    4000
                ) AS [Message]
            FROM distribution.dbo.MSreplication_monitordata md
            WHERE md.agent_type IN (2, 3, 4)
              AND (
                  md.publisher = @@SERVERNAME
                  OR md.subscriber = @@SERVERNAME
                  OR md.publisher_db IN (SELECT name FROM master.sys.databases WHERE is_published = 1 OR is_merge_published = 1)
                  OR md.subscriber_db IN (SELECT name FROM master.sys.databases WHERE is_subscribed = 1)
              );';

        EXEC sys.sp_executesql @monitorSql;
    END TRY
    BEGIN CATCH
        INSERT INTO #ReplicationHealth ([Role], [Component], [Message])
        VALUES ('Replication', 'Monitor Data', LEFT(ERROR_MESSAGE(), 4000));
    END CATCH;
END;

DECLARE @replicationDb sysname;
DECLARE @replicationSql nvarchar(max);

DECLARE replication_database_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM master.sys.databases
WHERE database_id > 4
  AND state_desc = 'ONLINE'
  AND HAS_DBACCESS(name) = 1;

OPEN replication_database_cursor;
FETCH NEXT FROM replication_database_cursor INTO @replicationDb;

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @replicationSql = N'USE ' + QUOTENAME(@replicationDb) + N';
        IF OBJECT_ID(N''dbo.syspublications'', N''U'') IS NOT NULL
        BEGIN
            INSERT INTO #ReplicationHealth ([Role], [Component], [Database], [Publication], [LastRunStatus], [Message])
            SELECT
                N''Publisher'',
                N''Publication'',
                DB_NAME(),
                name,
                CASE status WHEN 1 THEN N''Active'' ELSE N''Inactive'' END,
                LEFT(description, 4000)
            FROM dbo.syspublications
            WHERE name IS NOT NULL
              AND NOT EXISTS (
                  SELECT 1
                  FROM #ReplicationHealth rh
                  WHERE rh.[Role] = N''Publisher''
                    AND rh.[Component] = N''Publication''
                    AND rh.[Database] = DB_NAME()
                    AND rh.[Publication] = dbo.syspublications.name
              );
        END;

        IF OBJECT_ID(N''dbo.MSreplication_subscriptions'', N''U'') IS NOT NULL
        BEGIN
            INSERT INTO #ReplicationHealth ([Role], [Component], [Database], [Publication], [Subscriber], [SubscriberDatabase], [LastRunStatus], [Message])
            SELECT
                N''Subscriber'',
                N''Subscription'',
                publisher_db,
                publication,
                @@SERVERNAME,
                DB_NAME(),
                N''Configured'',
                LEFT(CONCAT(N''Publisher: '', publisher, N''; subscription database: '', DB_NAME()), 4000)
            FROM dbo.MSreplication_subscriptions
            WHERE publication IS NOT NULL
              AND NOT EXISTS (
                  SELECT 1
                  FROM #ReplicationHealth rh
                  WHERE rh.[Role] = N''Subscriber''
                    AND rh.[Component] = N''Subscription''
                    AND rh.[Publication] = dbo.MSreplication_subscriptions.publication
                    AND rh.[SubscriberDatabase] = DB_NAME()
              );
        END;';

    BEGIN TRY
        EXEC sys.sp_executesql @replicationSql;
    END TRY
    BEGIN CATCH
    END CATCH;

    FETCH NEXT FROM replication_database_cursor INTO @replicationDb;
END;

CLOSE replication_database_cursor;
DEALLOCATE replication_database_cursor;

INSERT INTO #ReplicationHealth ([Role], [Component], [Database], [Message])
SELECT
    CASE
        WHEN is_published = 1 OR is_merge_published = 1 THEN 'Publisher'
        WHEN is_subscribed = 1 THEN 'Subscriber'
        ELSE 'Distributor'
    END AS [Role],
    'Local Replication Database' AS [Component],
    name AS [Database],
    'Local database is configured for replication; latency is reported when monitor data is available.' AS [Message]
FROM master.sys.databases
WHERE is_published = 1
   OR is_subscribed = 1
   OR is_merge_published = 1
   OR is_distributor = 1;

IF OBJECT_ID('msdb.dbo.sysjobs') IS NOT NULL
   AND OBJECT_ID('msdb.dbo.syscategories') IS NOT NULL
   AND OBJECT_ID('msdb.dbo.sysjobhistory') IS NOT NULL
BEGIN
    INSERT INTO #ReplicationHealth ([Role], [Component], [Database], [AgentJob], [LastRunDateTime], [LastRunStatus], [Message])
    SELECT
        'Distributor',
        'Cleanup Job',
        NULL,
        j.name,
        lastRun.LastRunDateTime,
        CASE lastRun.run_status
            WHEN 0 THEN 'Failed'
            WHEN 1 THEN 'Succeeded'
            WHEN 2 THEN 'Retry'
            WHEN 3 THEN 'Canceled'
            WHEN 4 THEN 'In Progress'
            ELSE 'Unknown'
        END,
        LEFT(lastRun.message, 4000)
    FROM msdb.dbo.sysjobs j
    INNER JOIN msdb.dbo.syscategories c
        ON j.category_id = c.category_id
    OUTER APPLY (
        SELECT TOP (1)
            msdb.dbo.agent_datetime(h.run_date, h.run_time) AS LastRunDateTime,
            h.run_status,
            h.message
        FROM msdb.dbo.sysjobhistory h
        WHERE h.job_id = j.job_id
          AND h.step_id = 0
        ORDER BY h.instance_id DESC
    ) lastRun
    WHERE j.name LIKE '%cleanup%'
      AND (
          c.name LIKE 'REPL-%'
          OR j.name LIKE '%distribution%'
          OR j.name LIKE '%replication%'
          OR j.name LIKE '%subscription%'
      );
END;

SELECT
    NULLIF(RTRIM([Role]), '') AS [Role],
    [Component],
    [Database],
    [Publication],
    [Subscriber],
    [SubscriberDatabase],
    [LatencySeconds],
    [PendingCommands],
    [AgentJob],
    [LastRunDateTime],
    [LastRunStatus],
    [Message]
FROM #ReplicationHealth
ORDER BY [Component], [Database], [AgentJob];
"@

$systemResourceResult = Try-Query -ServerName $resolvedTarget -Name "SystemResourceUsage" -Query @"
WITH LatestCpu AS (
    SELECT TOP (1)
        DATEADD(ms, -1 * ((si.cpu_ticks / NULLIF(si.cpu_ticks / si.ms_ticks, 0)) - rb.[timestamp]), GETDATE()) AS [CaptureTime],
        x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]', 'int') AS [SystemIdlePercent],
        x.record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]', 'int') AS [SqlCPUPercent]
    FROM sys.dm_os_ring_buffers rb
    CROSS JOIN sys.dm_os_sys_info si
    CROSS APPLY (SELECT CAST(rb.record AS xml) AS record) x
    WHERE rb.ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
      AND rb.record LIKE '%<SystemHealth>%'
    ORDER BY rb.[timestamp] DESC
)
SELECT TOP (1)
    CAST(ISNULL(100 - LatestCpu.SystemIdlePercent, 0) AS decimal(18,2)) AS [TotalCPUPercent],
    CAST(ISNULL(LatestCpu.SqlCPUPercent, 0) AS decimal(18,2)) AS [SqlCPUPercent],
    CAST(ISNULL((100 - LatestCpu.SystemIdlePercent) - LatestCpu.SqlCPUPercent, 0) AS decimal(18,2)) AS [OtherCPUPercent],
    CAST(mem.total_physical_memory_kb / 1024.0 / 1024 AS decimal(18,2)) AS [TotalMemoryGB],
    CAST(mem.available_physical_memory_kb / 1024.0 / 1024 AS decimal(18,2)) AS [AvailableMemoryGB],
    CAST(((mem.total_physical_memory_kb - mem.available_physical_memory_kb) * 100.0) / NULLIF(mem.total_physical_memory_kb, 0) AS decimal(18,2)) AS [UsedMemoryPercent],
    CAST(pm.physical_memory_in_use_kb / 1024.0 / 1024 AS decimal(18,2)) AS [SqlMemoryGB],
    CAST((pm.physical_memory_in_use_kb * 100.0) / NULLIF(mem.total_physical_memory_kb, 0) AS decimal(18,2)) AS [SqlMemoryPercent],
    CAST(
        CASE
            WHEN ((mem.total_physical_memory_kb - mem.available_physical_memory_kb - pm.physical_memory_in_use_kb) * 100.0) / NULLIF(mem.total_physical_memory_kb, 0) < 0 THEN 0
            ELSE ((mem.total_physical_memory_kb - mem.available_physical_memory_kb - pm.physical_memory_in_use_kb) * 100.0) / NULLIF(mem.total_physical_memory_kb, 0)
        END AS decimal(18,2)
    ) AS [OtherMemoryPercent],
    LatestCpu.CaptureTime
FROM sys.dm_os_sys_memory mem
CROSS JOIN sys.dm_os_process_memory pm
OUTER APPLY (SELECT TOP (1) * FROM LatestCpu) LatestCpu;
"@

if ($SkipIndexFragmentation) {
    $indexFragmentationResult = [PSCustomObject]@{
        Name  = "IndexFragmentation"
        Table = $null
        Error = $null
    }
} else {
    $indexFragmentationResult = Try-Query -ServerName $resolvedTarget -Name "IndexFragmentation" -Query @"
CREATE TABLE #IndexFragmentation (
    [Database] sysname NOT NULL,
    [SchemaName] sysname NULL,
    [TableName] sysname NULL,
    [IndexName] sysname NULL,
    [IndexType] nvarchar(60) NULL,
    [AvgFragmentationPercent] decimal(18,2) NULL,
    [PageCount] bigint NULL
);

DECLARE @databaseName sysname;
DECLARE @sql nvarchar(max);

DECLARE database_cursor CURSOR LOCAL FAST_FORWARD FOR
SELECT name
FROM sys.databases
WHERE database_id > 4
  AND state_desc = 'ONLINE'
  AND HAS_DBACCESS(name) = 1;

OPEN database_cursor;
FETCH NEXT FROM database_cursor INTO @databaseName;

WHILE @@FETCH_STATUS = 0
BEGIN
    SET @sql = N'USE ' + QUOTENAME(@databaseName) + N';
        INSERT INTO #IndexFragmentation ([Database], [SchemaName], [TableName], [IndexName], [IndexType], [AvgFragmentationPercent], [PageCount])
        SELECT
            DB_NAME() AS [Database],
            OBJECT_SCHEMA_NAME(ips.object_id) AS [SchemaName],
            OBJECT_NAME(ips.object_id) AS [TableName],
            i.name AS [IndexName],
            ips.index_type_desc AS [IndexType],
            CAST(ips.avg_fragmentation_in_percent AS decimal(18,2)) AS [AvgFragmentationPercent],
            ips.page_count AS [PageCount]
        FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, ''LIMITED'') ips
        INNER JOIN sys.indexes i
            ON ips.object_id = i.object_id
           AND ips.index_id = i.index_id
        WHERE ips.index_id > 0
          AND ips.page_count > 0
          AND i.name IS NOT NULL
          AND (
              ips.avg_fragmentation_in_percent > 30
              OR ips.page_count > 10000
          );';

    EXEC sys.sp_executesql @sql;
    FETCH NEXT FROM database_cursor INTO @databaseName;
END;

CLOSE database_cursor;
DEALLOCATE database_cursor;

SELECT
    [Database],
    [SchemaName],
    [TableName],
    [IndexName],
    [IndexType],
    [AvgFragmentationPercent],
    [PageCount]
FROM #IndexFragmentation
ORDER BY
    CASE WHEN [AvgFragmentationPercent] > 30 OR [PageCount] > 10000 THEN 0 ELSE 1 END,
    [AvgFragmentationPercent] DESC,
    [PageCount] DESC,
    [Database],
    [TableName],
    [IndexName];
"@
}

$waitsResult = Try-Query -ServerName $resolvedTarget -Name "WaitStats" -Query @"
SELECT TOP (10)
    wait_type AS [WaitType],
    CAST(wait_time_ms / 1000.0 AS decimal(18,2)) AS [WaitSeconds],
    CAST(signal_wait_time_ms / 1000.0 AS decimal(18,2)) AS [SignalSeconds],
    waiting_tasks_count AS [WaitingTasks],
    CAST(signal_wait_time_ms * 100.0 / NULLIF(wait_time_ms, 0) AS decimal(18,2)) AS [SignalWaitPercent],
    CAST(wait_time_ms * 1.0 / NULLIF(waiting_tasks_count, 0) AS decimal(18,2)) AS [AvgWaitMs]
FROM sys.dm_os_wait_stats
WHERE wait_type NOT LIKE 'SLEEP%'
  AND wait_type NOT IN ('CLR_SEMAPHORE','LAZYWRITER_SLEEP','RESOURCE_QUEUE','XE_TIMER_EVENT',
                        'XE_DISPATCHER_WAIT','FT_IFTS_SCHEDULER_IDLE_WAIT','BROKER_TASK_STOP',
                        'BROKER_TO_FLUSH','SQLTRACE_BUFFER_FLUSH','CLR_AUTO_EVENT',
                        'CLR_MANUAL_EVENT','REQUEST_FOR_DEADLOCK_SEARCH','BROKER_EVENTHANDLER',
                        'DISPATCHER_QUEUE_SEMAPHORE','BROKER_RECEIVE_WAITFOR','ONDEMAND_TASK_QUEUE',
                        'DBMIRROR_EVENTS_QUEUE','DBMIRRORING_CMD','DIRTY_PAGE_POLL',
                        'HADR_FILESTREAM_IOMGR_IOCOMPLETION','SP_SERVER_DIAGNOSTICS_SLEEP')
ORDER BY wait_time_ms DESC;
"@

$databases = @(Convert-DataTableToObjects -Table $databaseResult.Table)
$logs = @(Convert-DataTableToObjects -Table $logResult.Table)
$dbFiles = @(Convert-DataTableToObjects -Table $dbFileResult.Table)
$disks = @(Convert-DataTableToObjects -Table $diskResult.Table)
$blocking = @(Convert-DataTableToObjects -Table $blockingResult.Table)
$jobs = @(Convert-DataTableToObjects -Table $jobsResult.Table)
$sqlErrorLogRows = @(Convert-DataTableToObjects -Table $sqlErrorLogResult.Table)
$agents = @(Convert-DataTableToObjects -Table $agentResult.Table)
$haRows = @(Convert-DataTableToObjects -Table $haResult.Table)
$clusterRows = @(Convert-DataTableToObjects -Table $clusterResult.Table)
$logShippingRows = @(Convert-DataTableToObjects -Table $logShippingResult.Table)
$replicationRows = @(Convert-DataTableToObjects -Table $replicationResult.Table)
$systemResourceRows = @(Convert-DataTableToObjects -Table $systemResourceResult.Table)
$indexFragmentationRows = @(Convert-DataTableToObjects -Table $indexFragmentationResult.Table)
$waits = @(Convert-DataTableToObjects -Table $waitsResult.Table)

Add-Status -Rows $databases -Rule {
    param($row)
    if ($row.RoleHint -eq "Log Shipping Secondary" -and $row.State -eq "RESTORING") { return "HEALTHY" }
    if ($row.State -ne "ONLINE") { return "CRITICAL" }
    if ($row.RoleHint -eq "Log Shipping Secondary" -or $row.RoleHint -eq "Replication Distributor") { return "HEALTHY" }

    if ($null -eq $row.LastFullBackup -or $row.LastFullBackup -is [System.DBNull]) { return "WARNING" }
    $lastFullAgeHours = ((Get-Date) - [datetime]$row.LastFullBackup).TotalHours
    if ($lastFullAgeHours -gt 192) { return "WARNING" }

    if ($lastFullAgeHours -gt 30) {
        if ($null -eq $row.LastDiffBackup -or $row.LastDiffBackup -is [System.DBNull]) { return "WARNING" }
        if (((Get-Date) - [datetime]$row.LastDiffBackup).TotalHours -gt 30) { return "WARNING" }
    }

    if ($row.RecoveryModel -eq "FULL") {
        if ($null -eq $row.LastLogBackup -or $row.LastLogBackup -is [System.DBNull]) { return "WARNING" }
        if (((Get-Date) - [datetime]$row.LastLogBackup).TotalMinutes -gt 120) { return "WARNING" }
    }
    return "HEALTHY"
}

Add-Status -Rows $logs -Rule {
    param($row)
    $used = [decimal]$row.UsedPercent
    if ($used -ge 90) { return "CRITICAL" }
    if ($used -ge 80) { return "WARNING" }
    return "HEALTHY"
}

Add-Status -Rows $dbFiles -Rule {
    param($row)
    return "HEALTHY"
}

Add-Status -Rows $disks -Rule {
    param($row)
    if ($row.Drive -eq "Unavailable") { return "WARNING" }

    if ($null -ne $row.FreePercent -and $row.FreePercent -isnot [System.DBNull]) {
        $free = [decimal]$row.FreePercent
        if ($free -lt 10) { return "CRITICAL" }
        if ($free -lt 20) { return "WARNING" }
        return "HEALTHY"
    }

    $freeGb = if ($null -eq $row.FreeGB -or $row.FreeGB -is [System.DBNull]) { $null } else { [decimal]$row.FreeGB }
    if ($null -ne $freeGb) {
        if ($freeGb -lt 5) { return "CRITICAL" }
        if ($freeGb -lt 10) { return "WARNING" }
    }
    return "HEALTHY"
}

Add-Status -Rows $blocking -Rule {
    param($row)
    $wait = [int64]$row.WaitMs
    if ($wait -ge 30000) { return "CRITICAL" }
    if ($wait -ge 10000) { return "WARNING" }
    return "HEALTHY"
}

Add-Status -Rows $jobs -Rule { param($row) "CRITICAL" }

Add-Status -Rows $sqlErrorLogRows -Rule { param($row) "HEALTHY" }

Add-Status -Rows $agents -Rule {
    param($row)
    if ($row.ServiceName -like "SQL Server Agent*" -and $row.ServiceStatus -ne "Running") { return "WARNING" }
    if ($row.ServiceStatus -ne "Running") { return "WARNING" }
    return "HEALTHY"
}

Add-Status -Rows $haRows -Rule {
    param($row)
    if ($row.SyncHealth -ne "HEALTHY") { return "CRITICAL" }
    return "HEALTHY"
}

Add-Status -Rows $clusterRows -Rule {
    param($row)
    if ($row.ClusterNodeStatus -ne "Up") { return "WARNING" }
    return "HEALTHY"
}

Add-Status -Rows $logShippingRows -Rule {
    param($row)
    if ($row.Role -eq "Primary") {
        if ($null -eq $row.LastBackupDate -or $row.LastBackupDate -is [System.DBNull]) { return "CRITICAL" }
        if ([int]$row.MinutesSinceLastBackup -gt 120) { return "CRITICAL" }
        if ([int]$row.MinutesSinceLastBackup -gt 60) { return "WARNING" }
        return "HEALTHY"
    }

    if ($null -eq $row.LastRestoredDate -or $row.LastRestoredDate -is [System.DBNull]) { return "CRITICAL" }
    if ([int]$row.MinutesSinceLastRestore -gt 120) { return "CRITICAL" }
    if ([int]$row.MinutesSinceLastRestore -gt 60) { return "WARNING" }
    return "HEALTHY"
}

Add-Status -Rows $replicationRows -Rule {
    param($row)
    if ($row.LastRunStatus -eq "Failed") { return "CRITICAL" }
    if ($row.LastRunStatus -eq "Retry" -or $row.LastRunStatus -eq "Retrying" -or $row.LastRunStatus -eq "Canceled" -or $row.LastRunStatus -eq "Unknown") { return "WARNING" }
    if ($null -ne $row.LatencySeconds -and $row.LatencySeconds -isnot [System.DBNull]) {
        $latency = [int]$row.LatencySeconds
        if ($latency -gt 900) { return "CRITICAL" }
        if ($latency -gt 300) { return "WARNING" }
    }
    return "HEALTHY"
}

Add-Status -Rows $systemResourceRows -Rule {
    param($row)
    $cpu = if ($null -eq $row.TotalCPUPercent -or $row.TotalCPUPercent -is [System.DBNull]) { 0 } else { [decimal]$row.TotalCPUPercent }
    $memory = if ($null -eq $row.UsedMemoryPercent -or $row.UsedMemoryPercent -is [System.DBNull]) { 0 } else { [decimal]$row.UsedMemoryPercent }
    if ($cpu -ge 95 -or $memory -ge 95) { return "CRITICAL" }
    if ($cpu -ge 80 -or $memory -ge 90) { return "WARNING" }
    return "HEALTHY"
}

Add-Status -Rows $indexFragmentationRows -Rule {
    param($row)
    return "HEALTHY"
}

Add-Status -Rows $waits -Rule {
    param($row)
    $signalWaitPercent = if ($null -eq $row.SignalWaitPercent -or $row.SignalWaitPercent -is [System.DBNull]) { 0 } else { [decimal]$row.SignalWaitPercent }
    $avgWaitMs = if ($null -eq $row.AvgWaitMs -or $row.AvgWaitMs -is [System.DBNull]) { 0 } else { [decimal]$row.AvgWaitMs }
    if ($signalWaitPercent -gt 25 -or $avgWaitMs -gt 500) { return "WARNING" }
    return "HEALTHY"
}

$databaseCount = $databases.Count
$criticalCount = @($databases + $logs + $dbFiles + $disks + $blocking + $jobs + $sqlErrorLogRows + $agents + $haRows + $clusterRows + $logShippingRows + $replicationRows + $systemResourceRows + $indexFragmentationRows | Where-Object { $_.Status -eq "CRITICAL" }).Count
$warningCount = @($databases + $logs + $dbFiles + $disks + $blocking + $jobs + $sqlErrorLogRows + $agents + $haRows + $clusterRows + $logShippingRows + $replicationRows + $systemResourceRows + $indexFragmentationRows | Where-Object { $_.Status -eq "WARNING" }).Count
$overallStatus = Get-OverallStatus -Groups @($databases, $logs, $dbFiles, $disks, $blocking, $jobs, $sqlErrorLogRows, $agents, $haRows, $clusterRows, $logShippingRows, $replicationRows, $systemResourceRows, $indexFragmentationRows)

$maxLog = if ($logs.Count -gt 0) { ($logs | Sort-Object UsedPercent -Descending | Select-Object -First 1) } else { $null }
$lowestDisk = if ($disks.Count -gt 0) {
    $percentDisk = $disks | Where-Object { $null -ne $_.FreePercent -and $_.FreePercent -isnot [System.DBNull] } | Sort-Object FreePercent | Select-Object -First 1
    if ($percentDisk) { $percentDisk } else { $disks | Sort-Object FreeGB | Select-Object -First 1 }
} else { $null }
$lastSqlStart = if ($serverInfo.SqlStartTime) { [datetime]$serverInfo.SqlStartTime } else { $null }
$uptimeDays = if ($lastSqlStart) { [math]::Round(((Get-Date) - $lastSqlStart).TotalDays, 1) } else { 0 }
$sqlUptimeNote = if ($lastSqlStart) { "Started {0}" -f $lastSqlStart.ToString("yyyy-MM-dd HH:mm") } else { "Unknown" }
$maxLogValue = if ($maxLog) { "{0}%" -f ([math]::Round([decimal]$maxLog.UsedPercent, 2)) } else { "N/A" }
$maxLogNote = if ($maxLog) { [string]$maxLog.Database } else { "No user database log data" }
$maxLogStatus = if ($maxLog) { [string]$maxLog.Status } else { "HEALTHY" }
$lowestDiskValue = if ($lowestDisk -and $null -ne $lowestDisk.FreePercent -and $lowestDisk.FreePercent -isnot [System.DBNull]) {
    "{0}%" -f ([math]::Round([decimal]$lowestDisk.FreePercent, 2))
} elseif ($lowestDisk -and $null -ne $lowestDisk.FreeGB -and $lowestDisk.FreeGB -isnot [System.DBNull]) {
    "{0} GB free" -f ([math]::Round([decimal]$lowestDisk.FreeGB, 2))
} else { "N/A" }
$lowestDiskNote = if ($lowestDisk) { [string]$lowestDisk.Drive } else { "Disk data unavailable" }
$lowestDiskStatus = if ($lowestDisk) { [string]$lowestDisk.Status } else { "HEALTHY" }
$failedJobsStatus = if ($jobs.Count -gt 0) { "CRITICAL" } else { "HEALTHY" }
$haClusterValue = if ([int]$serverInfo.IsClustered -eq 1) { "WSFC" } elseif ([int]$serverInfo.IsHadrEnabled -eq 1) { "Always On" } else { "Standalone" }
$logShippingCriticalCount = @($logShippingRows | Where-Object { $_.Status -eq "CRITICAL" }).Count
$logShippingWarningCount = @($logShippingRows | Where-Object { $_.Status -eq "WARNING" }).Count
$logShippingStatus = if ($logShippingCriticalCount -gt 0) { "CRITICAL" } elseif ($logShippingWarningCount -gt 0) { "WARNING" } else { "HEALTHY" }
$replicationCriticalCount = @($replicationRows | Where-Object { $_.Status -eq "CRITICAL" }).Count
$replicationWarningCount = @($replicationRows | Where-Object { $_.Status -eq "WARNING" }).Count
$replicationStatus = if ($replicationCriticalCount -gt 0) { "CRITICAL" } elseif ($replicationWarningCount -gt 0) { "WARNING" } else { "HEALTHY" }
$systemResource = if ($systemResourceRows.Count -gt 0) { $systemResourceRows | Select-Object -First 1 } else { $null }
$cpuValue = if ($systemResource) { "{0}%" -f ([math]::Round([decimal]$systemResource.TotalCPUPercent, 2)) } else { "N/A" }
$cpuNote = if ($systemResource) { "SQL {0}%, other {1}%" -f ([math]::Round([decimal]$systemResource.SqlCPUPercent, 2)), ([math]::Round([decimal]$systemResource.OtherCPUPercent, 2)) } else { "CPU data unavailable" }
$ramValue = if ($systemResource) { "{0}%" -f ([math]::Round([decimal]$systemResource.UsedMemoryPercent, 2)) } else { "N/A" }
$ramNote = if ($systemResource) { "SQL {0}%, other {1}%" -f ([math]::Round([decimal]$systemResource.SqlMemoryPercent, 2)), ([math]::Round([decimal]$systemResource.OtherMemoryPercent, 2)) } else { "Memory data unavailable" }
$systemResourceStatus = if ($systemResource) { [string]$systemResource.Status } else { "HEALTHY" }
$serverIpAddress = Get-ServerIpAddress -HostName ([string]$serverInfo.MachineName)

$summaryCards = @(
    New-SummaryCard -Label "Overall Health" -Value $overallStatus -Note "$criticalCount critical, $warningCount warning findings" -Status $overallStatus
    New-SummaryCard -Label "SQL Uptime" -Value ("{0} days" -f $uptimeDays) -Note $sqlUptimeNote -Status "HEALTHY"
    New-SummaryCard -Label "Database Count" -Value ([string]$databaseCount) -Note "User databases included in this daily check" -Status "HEALTHY"
    New-SummaryCard -Label "Highest Log Use" -Value $maxLogValue -Note $maxLogNote -Status $maxLogStatus
    New-SummaryCard -Label "Lowest Disk Free" -Value $lowestDiskValue -Note $lowestDiskNote -Status $lowestDiskStatus
    New-SummaryCard -Label "Failed Jobs" -Value ([string]$jobs.Count) -Note "Recent SQL Agent job failures" -Status $failedJobsStatus
    New-SummaryCard -Label "HA / Cluster" -Value $haClusterValue -Note ("Active node: {0}" -f $serverInfo.ActiveNode) -Status "HEALTHY"
    New-SummaryCard -Label "Log Shipping" -Value ([string]$logShippingRows.Count) -Note "Primary/secondary monitor rows" -Status $logShippingStatus
) -join [Environment]::NewLine

$systemResourceCss = Get-StatusClass -Status $systemResourceStatus
$resourceStrip = @(
    "<div class=""resource-item status-$systemResourceCss""><div class=""resource-label"">CPU Usage</div><div class=""resource-value"">$([string](Escape-Html -Value $cpuValue))</div><div class=""resource-note"">$([string](Escape-Html -Value $cpuNote))</div></div>"
    "<div class=""resource-item status-$systemResourceCss""><div class=""resource-label"">RAM Usage</div><div class=""resource-value"">$([string](Escape-Html -Value $ramValue))</div><div class=""resource-note"">$([string](Escape-Html -Value $ramNote))</div></div>"
) -join [Environment]::NewLine

$indexFragmentationPanel = if ($SkipIndexFragmentation) {
    ""
} else {
    New-SectionTable -Title "Index Fragmentation" -Subtitle "Advisory only: fragmentation above 30% or page count above 10000 does not affect overall health" -Rows $indexFragmentationRows -Columns @("Database", "SchemaName", "TableName", "IndexName", "IndexType", "AvgFragmentationPercent", "PageCount") -FullWidth -Scrollable
}

$panels = @(
    New-SectionTable -Title "Database Health" -Subtitle "State, recovery model, size, role, and last full/differential/log backups" -Rows $databases -Columns @("Database", "State", "RecoveryModel", "SizeMB", "RoleHint", "LastFullBackup", "LastDiffBackup", "LastLogBackup", "Status") -FullWidth
    New-SectionTable -Title "Database File Usage" -Subtitle "Advisory only: high data-file allocation is expected for growing production databases and does not affect overall health" -Rows $dbFiles -Columns @("Database", "LogicalName", "SizeMB", "UsedMB", "FreeMB", "UsedPercent", "Status")
    New-SectionTable -Title "Transaction Log Usage" -Subtitle "Top user databases by log utilization" -Rows $logs -Columns @("Database", "LogSizeMB", "UsedMB", "UsedPercent", "Status")
    New-SectionTable -Title "Storage Health" -Subtitle "Volume free space for SQL data/log locations" -Rows $disks -Columns @("Drive", "TotalGB", "FreeGB", "FreePercent", "Status")
    New-SectionTable -Title "Blocking Sessions" -Subtitle "Only active user requests with blocking_session_id <> 0 are shown" -Rows $blocking -Columns @("BlockingSessionID", "SessionID", "WaitType", "WaitMs", "WaitResource", "Database", "Statement", "Status") -EmptyMessage "No active blocking sessions detected."
    New-SectionTable -Title "SQL Services" -Subtitle "Engine and Agent service visibility from SQL Server" -Rows $agents -Columns @("ServiceName", "StartupType", "ServiceStatus", "LastStartupTime", "Status")
    New-SectionTable -Title "Failed SQL Agent Jobs" -Subtitle "Failed SQL Agent job outcomes from the last 48 hours" -Rows $jobs -Columns @("JobName", "RunDateTime", "Message", "Status") -EmptyMessage "No SQL Agent job failures found in the last 48 hours." -Scrollable
    New-SectionTable -Title "Always On Health" -Subtitle "Rendered only when HADR is enabled and replica state is visible" -Rows $haRows -Columns @("Database", "AvailabilityGroup", "ReplicaServer", "SyncState", "SyncHealth", "ReplicaRole", "DatabaseState", "Status") -EmptyMessage "Always On is not enabled or no local replica state was returned." -FullWidth
    New-SectionTable -Title "Log Shipping Health" -Subtitle "Primary backup and secondary copy/restore monitor status" -Rows $logShippingRows -Columns @("Role", "PrimaryServer", "SecondaryServer", "Database", "LastBackupDate", "LastCopiedDate", "LastRestoredDate", "MinutesSinceLastBackup", "MinutesSinceLastCopy", "MinutesSinceLastRestore", "Status") -EmptyMessage "No log shipping monitor rows were found on this instance." -FullWidth
    New-SectionTable -Title "Replication Health" -Subtitle "Replication health, latency, pending commands, and cleanup job status" -Rows $replicationRows -Columns @("Role", "Component", "Database", "Publication", "Subscriber", "SubscriberDatabase", "LatencySeconds", "PendingCommands", "AgentJob", "LastRunDateTime", "LastRunStatus", "Message", "Status") -EmptyMessage "No publisher/subscriber replication monitor rows or cleanup jobs were found on this instance." -FullWidth
    $indexFragmentationPanel
    New-SectionTable -Title "SQL Error Log" -Subtitle "Newest 50 current error log rows from the last 48 hours" -Rows $sqlErrorLogRows -Columns @("LogDate", "ProcessInfo", "Message") -EmptyMessage "No SQL error log rows found in the last 48 hours." -FullWidth -Scrollable
    New-SectionTable -Title "Top Wait Statistics" -Subtitle "Advisory only: wait rows do not affect overall health" -Rows $waits -Columns @("WaitType", "WaitSeconds", "SignalSeconds", "WaitingTasks", "SignalWaitPercent", "AvgWaitMs") -FullWidth -Scrollable
) -join [Environment]::NewLine

$overallCss = Get-StatusClass -Status $overallStatus
$overallPill = "<div class=""pill $overallCss"">Overall Status: $overallStatus</div>"

$htmlTemplate = Get-Content -Raw -LiteralPath $templatePath
$htmlContent = $htmlTemplate.Replace("__OVERALL_PILL__", $overallPill)
$htmlContent = $htmlContent.Replace("__SERVER_NAME__", [string](Escape-Html -Value $serverInfo.ServerName))
$htmlContent = $htmlContent.Replace("__SERVER_IP__", [string](Escape-Html -Value $serverIpAddress))
$htmlContent = $htmlContent.Replace("__SERVER_TARGET__", [string](Escape-Html -Value $resolvedTarget))
$htmlContent = $htmlContent.Replace("__GENERATED_AT__", (Get-Date -Format "yyyy-MM-dd HH:mm:ss"))
$htmlContent = $htmlContent.Replace("__EDITION__", [string](Escape-Html -Value $serverInfo.Edition))
$htmlContent = $htmlContent.Replace("__VERSION__", [string](Escape-Html -Value ("{0} ({1})" -f $serverInfo.ProductVersion, $serverInfo.ProductLevel)))
$htmlContent = $htmlContent.Replace("__DATABASE_COUNT__", [string]$databaseCount)
$htmlContent = $htmlContent.Replace("__SUMMARY_CARDS__", $summaryCards)
$htmlContent = $htmlContent.Replace("__RESOURCE_STRIP__", $resourceStrip)
$htmlContent = $htmlContent.Replace("__SECTION_PANELS__", $panels)

$warnings = @(@(
    $databaseResult,
    $logResult,
    $dbFileResult,
    $diskResult,
    $blockingResult,
    $jobsResult,
    $sqlErrorLogResult,
    $agentResult,
    $haResult,
    $clusterResult,
    $logShippingResult,
    $replicationResult,
    $systemResourceResult,
    $indexFragmentationResult,
    $waitsResult
 ) | Where-Object { $_.Error })

if ($warnings.Count -gt 0) {
    $warningLines = ($warnings | ForEach-Object {
        "<div class=""warning-state"">$([System.Net.WebUtility]::HtmlEncode($_.Name)): $([System.Net.WebUtility]::HtmlEncode($_.Error))</div>"
    }) -join [Environment]::NewLine
    $htmlContent = $htmlContent.Replace("</main>", "$warningLines`r`n    </main>")
}

$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$reportPath = Join-Path $OutputPath "MSSQL_Daily_Health_$timestamp.html"
[System.IO.File]::WriteAllText($reportPath, $htmlContent, [System.Text.Encoding]::UTF8)

Write-Host "HTML report created: $reportPath" -ForegroundColor Cyan

if ($OpenReport) {
    Start-Process -FilePath $reportPath
}

[PSCustomObject]@{
    ReportPath      = $reportPath
    ServerTarget    = $resolvedTarget
    ServerName      = [string]$serverInfo.ServerName
    OverallStatus   = $overallStatus
    CriticalCount   = $criticalCount
    WarningCount    = $warningCount
    DatabaseCount   = $databaseCount
}
