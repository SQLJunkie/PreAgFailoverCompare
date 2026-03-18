# Compare-SqlInstanceConfigs.ps1
# SQL Server instance comparison — self-contained script
param(
    [Parameter(Mandatory=$true)]
    [string]$SourceInstance,

    [Parameter(Mandatory=$true)]
    [string]$TargetInstance
)

# Ensure dbatools is loaded at exactly the required version.
# -Force reloads if a different version is already in the session.
$requiredDbatoolsVersion = [version]'2.7.25'
$loadedDbatools = Get-Module -Name dbatools
if (-not $loadedDbatools -or $loadedDbatools.Version -ne $requiredDbatoolsVersion) {
    Import-Module dbatools -RequiredVersion $requiredDbatoolsVersion -Force -ErrorAction Stop
}

function Add-Row {
    param(
        [string]$Scope,        # Server | Database
        [string]$Category,     # sp_configure | StartupParameters | TraceFlag | Login | DbOptions | etc.
        [string]$Item,         # Name of the thing being compared
        [string]$Property,     # Property name
        [object]$SourceValue,
        [object]$TargetValue,
        [string]$DatabaseName  # Optional, for DB-scoped rows
    )

    $match = $null
    if ($null -eq $SourceValue -and $null -eq $TargetValue) {
        $match = $true
    }
    elseif ($SourceValue -eq $TargetValue) {
        $match = $true
    }
    else {
        $match = $false
    }

    $script:rows += [pscustomobject]@{
        Scope        = $Scope
        Category     = $Category
        Item         = $Item
        Property     = $Property
        Source       = $SourceValue
        Target       = $TargetValue
        Match        = $match
        DatabaseName = $DatabaseName
    }
}

function Get-AgDbNamesSafe {
    param(
        [Microsoft.SqlServer.Management.Smo.Server]$Server
    )

    # Query sys.databases directly — replica_id IS NOT NULL is the authoritative
    # indicator that a database belongs to an AG on any replica (primary or secondary).
    # SMO AvailabilityGroups navigation is unreliable on secondaries.
    try {
        $result = $Server.Databases['master'].ExecuteWithResults(
            'SELECT name FROM sys.databases WHERE replica_id IS NOT NULL'
        )
        return @($result.Tables[0].Rows | ForEach-Object { $_.name })
    }
    catch {
        return @()
    }
}

function Compare-TempDbFileSet {
    param(
        [object[]]$SrcFiles,
        [object[]]$TgtFiles,
        [string]$LabelPrefix
    )
    $maxCount = [Math]::Max($SrcFiles.Count, $TgtFiles.Count)
    for ($i = 0; $i -lt $maxCount; $i++) {
        $sf    = if ($i -lt $SrcFiles.Count) { $SrcFiles[$i] } else { $null }
        $tf    = if ($i -lt $TgtFiles.Count) { $TgtFiles[$i] } else { $null }
        $label = "$LabelPrefix $($i + 1)"
        Add-Row 'Server' 'TempDbConfig' $label 'LogicalName' ($sf.LogicalName) ($tf.LogicalName)
        Add-Row 'Server' 'TempDbConfig' $label 'SizeMB'      ($sf.SizeMB)      ($tf.SizeMB)
        Add-Row 'Server' 'TempDbConfig' $label 'AutoGrowth'  ($sf.AutoGrowth)  ($tf.AutoGrowth)
    }
}

function Compare-ServerLevel {
    param(
        [Microsoft.SqlServer.Management.Smo.Server]$Source,
        [Microsoft.SqlServer.Management.Smo.Server]$Target
    )

    $Source.Refresh()
    $Target.Refresh()

    #
    # Server properties
    #
    $props = @(
        'ProductLevel',
        'ProductVersion',
        'Edition',
        'EngineEdition',
        'IsClustered',
        'IsHadrEnabled',
        'IsPolyBaseInstalled',
        'IsXTPSupported',
        'IsColumnStoreIndexSupported',
        'IsIntegratedSecurityOnly'
    )

    foreach ($p in $props) {
        $srcVal = $Source.$p
        $tgtVal = $Target.$p
        Add-Row 'Server' 'ServerProperty' $p 'Value' $srcVal $tgtVal
    }

    #
    # sp_configure (one-way: source -> target only)
    #
    $srcCfg = Get-DbaSpConfigure -SqlInstance $Source
    $tgtCfg = Get-DbaSpConfigure -SqlInstance $Target

    # Internal read-only version metadata — always differ between SQL Server versions, not user-configurable
    $skipSpConfigItems = @('VersionHighPartOfSqlServer', 'VersionLowPartOfSqlServer')

    foreach ($c in $srcCfg.Name) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        if ($c -in $skipSpConfigItems)        { continue }

        $s = $srcCfg | Where-Object Name -eq $c
        $t = $tgtCfg | Where-Object Name -eq $c

        if (-not $t) {
            Add-Row 'Server' 'sp_configure' $c 'ExistsOnTarget' $true $false
            continue
        }

        Add-Row 'Server' 'sp_configure' $c 'ConfiguredValue' $s.ConfiguredValue $t.ConfiguredValue
        Add-Row 'Server' 'sp_configure' $c 'RunningValue'   $s.RunningValue    $t.RunningValue
    }

    #
    # Startup parameters + trace flags
    # Uses Get-DbaStartupParameter (WMI-based) instead of Server.StartupParameters
    # (SMO property reads the remote registry directly, which is often inaccessible)
    #
    $srcStartup = $null
    $tgtStartup = $null

    try {
        $srcStartup = Get-DbaStartupParameter -SqlInstance $Source -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve startup parameters for source '$($Source.Name)': $($_.Exception.Message)"
    }

    try {
        $tgtStartup = Get-DbaStartupParameter -SqlInstance $Target -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve startup parameters for target '$($Target.Name)': $($_.Exception.Message)"
    }

    if ($srcStartup -or $tgtStartup) {

        # Named startup parameter properties
        $startupProps = @(
            'MasterData',
            'MasterLog',
            'ErrorLog',
            'SingleUser',
            'MinimalStart',
            'CommandPromptStart',
            'NoLoggingToWinEvents',
            'WithoutNetworkSupport'
        )

        foreach ($p in $startupProps) {
            $srcVal = if ($srcStartup) { $srcStartup.$p } else { $null }
            $tgtVal = if ($tgtStartup) { $tgtStartup.$p } else { $null }
            Add-Row 'Server' 'StartupParameters' 'StartupParameters' $p $srcVal $tgtVal
        }

        # Trace flags — TraceFlags property is a comma/semicolon-delimited string of flag numbers
        $srcFlags = @(($srcStartup.TraceFlags -split '[,;]') | Where-Object { $_ -match '^\d+$' })
        $tgtFlags = @(($tgtStartup.TraceFlags -split '[,;]') | Where-Object { $_ -match '^\d+$' })

        $allFlags = ($srcFlags + $tgtFlags) | Sort-Object -Unique
        foreach ($f in $allFlags) {
            $srcHas = $srcFlags -contains $f
            $tgtHas = $tgtFlags -contains $f
            Add-Row 'Server' 'TraceFlag' $f 'EnabledAtStartup' $srcHas $tgtHas
        }

        Add-Row 'Server' 'TraceFlagSummary' 'TraceFlags' 'StartupFlags' ($srcFlags -join ',') ($tgtFlags -join ',')
    }
    else {
        Add-Row 'Server' 'StartupParameters' 'StartupParameters' 'Parameters' $null $null
    }

    #
    # Windows service properties — IFI, LPIM, service accounts, start modes
    #

    # Services first — we need the Engine service account name to filter privilege results correctly.
    # Get-DbaPrivilege returns one row per account that holds any of the checked privileges on the
    # host; without filtering, $srcPriv would be an array and .InstantFileInitialization would return
    # a space-joined list of all values instead of a single bool.
    $srcSvcs = $null
    $tgtSvcs = $null

    try {
        $srcSvcs = Get-DbaService -SqlInstance $Source -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve services for source '$($Source.Name)': $($_.Exception.Message)"
    }

    try {
        $tgtSvcs = Get-DbaService -SqlInstance $Target -ErrorAction Stop
    }
    catch {
        Write-Warning "Could not retrieve services for target '$($Target.Name)': $($_.Exception.Message)"
    }

    $srcEngineAcct = ($srcSvcs | Where-Object { $_.ServiceType -eq 'Engine' } | Select-Object -First 1).StartName
    $tgtEngineAcct = ($tgtSvcs | Where-Object { $_.ServiceType -eq 'Engine' } | Select-Object -First 1).StartName

    # Get-DbaPrivilege is a Windows-level query — requires ComputerName, not SqlInstance.
    # Filter to the Engine service account row so we get a single object, not an array.
    $srcPriv = $null
    $tgtPriv = $null

    try {
        $srcAllPriv = Get-DbaPrivilege -ComputerName $Source.ComputerName -ErrorAction Stop
        $srcPriv = if ($srcEngineAcct) {
            $srcAllPriv | Where-Object { $_.User -eq $srcEngineAcct } | Select-Object -First 1
        } else {
            $srcAllPriv | Select-Object -First 1
        }
    }
    catch {
        Write-Warning "Could not retrieve privileges for source '$($Source.ComputerName)': $($_.Exception.Message)"
    }

    try {
        $tgtAllPriv = Get-DbaPrivilege -ComputerName $Target.ComputerName -ErrorAction Stop
        $tgtPriv = if ($tgtEngineAcct) {
            $tgtAllPriv | Where-Object { $_.User -eq $tgtEngineAcct } | Select-Object -First 1
        } else {
            $tgtAllPriv | Select-Object -First 1
        }
    }
    catch {
        Write-Warning "Could not retrieve privileges for target '$($Target.ComputerName)': $($_.Exception.Message)"
    }

    $srcIFI  = if ($srcPriv) { $srcPriv.InstantFileInitialization } else { $null }
    $tgtIFI  = if ($tgtPriv) { $tgtPriv.InstantFileInitialization } else { $null }
    $srcLPIM = if ($srcPriv) { $srcPriv.LockPagesInMemory }         else { $null }
    $tgtLPIM = if ($tgtPriv) { $tgtPriv.LockPagesInMemory }         else { $null }

    Add-Row 'Server' 'ServiceProperty' 'SQL Server Engine' 'Instant File Initialization' $srcIFI $tgtIFI
    Add-Row 'Server' 'ServiceProperty' 'SQL Server Engine' 'Lock Pages in Memory'        $srcLPIM $tgtLPIM

    # Engine and Agent service account, start mode, and state
    $svcTypeMap = [ordered]@{
        'Engine' = 'SQL Server Engine'
        'Agent'  = 'SQL Server Agent'
    }

    foreach ($svcType in $svcTypeMap.Keys) {
        $label  = $svcTypeMap[$svcType]
        $srcSvc = $srcSvcs | Where-Object { $_.ServiceType -eq $svcType } | Select-Object -First 1
        $tgtSvc = $tgtSvcs | Where-Object { $_.ServiceType -eq $svcType } | Select-Object -First 1

        if (-not ($srcSvc -or $tgtSvc)) { continue }

        $srcAcct  = if ($srcSvc) { $srcSvc.StartName } else { $null }
        $tgtAcct  = if ($tgtSvc) { $tgtSvc.StartName } else { $null }
        $srcMode  = if ($srcSvc) { $srcSvc.StartMode } else { $null }
        $tgtMode  = if ($tgtSvc) { $tgtSvc.StartMode } else { $null }
        $srcState = if ($srcSvc) { $srcSvc.State     } else { $null }
        $tgtState = if ($tgtSvc) { $tgtSvc.State     } else { $null }

        Add-Row 'Server' 'ServiceProperty' $label 'ServiceAccount' $srcAcct  $tgtAcct
        Add-Row 'Server' 'ServiceProperty' $label 'StartMode'      $srcMode  $tgtMode
        Add-Row 'Server' 'ServiceProperty' $label 'State'          $srcState $tgtState
    }

    #
    # Linked servers
    #
    # LinkedServers is a lazily-loaded SMO collection — Refresh() forces it to populate.
    $Source.LinkedServers.Refresh()
    $Target.LinkedServers.Refresh()
    $srcLinkedServers = @($Source.LinkedServers)
    $tgtLinkedServers = @($Target.LinkedServers)

    # sys.servers has reliable bit columns for two options that SMO does not expose correctly:
    #   is_data_access_enabled                        → SSMS "Data Access"
    #   is_remote_proc_transaction_promotion_enabled  → SSMS "Enable Promotion of Distributed Transactions"
    $lsOptionsQuery = @"
SELECT name,
       CAST(is_data_access_enabled                       AS bit) AS DataAccess,
       CAST(is_remote_proc_transaction_promotion_enabled AS bit) AS DistribXactProcEnabled
FROM sys.servers
WHERE is_linked = 1
"@
    $srcLsOptions = Invoke-DbaQuery -SqlInstance $Source -Database master -Query $lsOptionsQuery -As PSObject
    $tgtLsOptions = Invoke-DbaQuery -SqlInstance $Target -Database master -Query $lsOptionsQuery -As PSObject

    $allLinkedServerNames = (@($srcLinkedServers.Name) + @($tgtLinkedServers.Name)) | Sort-Object -Unique

    foreach ($lsName in $allLinkedServerNames) {
        $srcLs     = $srcLinkedServers | Where-Object { $_.Name -eq $lsName }
        $tgtLs     = $tgtLinkedServers | Where-Object { $_.Name -eq $lsName }
        $srcSsOpts = $srcLsOptions     | Where-Object { $_.name -eq $lsName }
        $tgtSsOpts = $tgtLsOptions     | Where-Object { $_.name -eq $lsName }

        Add-Row 'Server' 'LinkedServer' $lsName 'Exists' (!!$srcLs) (!!$tgtLs)

        if ($srcLs -and $tgtLs) {
            # Connection & identity
            Add-Row 'Server' 'LinkedServer' $lsName 'DataSource'             $srcLs.DataSource             $tgtLs.DataSource
            Add-Row 'Server' 'LinkedServer' $lsName 'ProviderName'           $srcLs.ProviderName           $tgtLs.ProviderName
            Add-Row 'Server' 'LinkedServer' $lsName 'ProductName'            $srcLs.ProductName            $tgtLs.ProductName
            Add-Row 'Server' 'LinkedServer' $lsName 'Catalog'                $srcLs.Catalog                $tgtLs.Catalog

            # Data access & RPC (DataAccess and DistribXactProcEnabled from sys.servers — SMO doesn't expose these correctly)
            Add-Row 'Server' 'LinkedServer' $lsName 'DataAccess'             ([bool]$srcSsOpts.DataAccess)             ([bool]$tgtSsOpts.DataAccess)
            Add-Row 'Server' 'LinkedServer' $lsName 'Rpc'                    $srcLs.Rpc                                $tgtLs.Rpc
            Add-Row 'Server' 'LinkedServer' $lsName 'RpcOut'                 $srcLs.RpcOut                             $tgtLs.RpcOut
            Add-Row 'Server' 'LinkedServer' $lsName 'DistribXactProcEnabled' ([bool]$srcSsOpts.DistribXactProcEnabled) ([bool]$tgtSsOpts.DistribXactProcEnabled)

            # Replication roles
            Add-Row 'Server' 'LinkedServer' $lsName 'Distributor'            $srcLs.Distributor            $tgtLs.Distributor
            Add-Row 'Server' 'LinkedServer' $lsName 'Publisher'              $srcLs.Publisher              $tgtLs.Publisher
            Add-Row 'Server' 'LinkedServer' $lsName 'Subscriber'             $srcLs.Subscriber             $tgtLs.Subscriber

            # Collation
            Add-Row 'Server' 'LinkedServer' $lsName 'UseRemoteCollation'     $srcLs.UseRemoteCollation     $tgtLs.UseRemoteCollation
            Add-Row 'Server' 'LinkedServer' $lsName 'CollationCompatible'    $srcLs.CollationCompatible    $tgtLs.CollationCompatible
            Add-Row 'Server' 'LinkedServer' $lsName 'CollationName'          $srcLs.CollationName          $tgtLs.CollationName

            # Timeouts
            Add-Row 'Server' 'LinkedServer' $lsName 'ConnectTimeout'         $srcLs.ConnectTimeout         $tgtLs.ConnectTimeout
            Add-Row 'Server' 'LinkedServer' $lsName 'QueryTimeout'           $srcLs.QueryTimeout           $tgtLs.QueryTimeout

            # Provider options
            Add-Row 'Server' 'LinkedServer' $lsName 'LazySchemaValidation'   $srcLs.LazySchemaValidation   $tgtLs.LazySchemaValidation

            # Login security mappings
            $srcLsLogins = @($srcLs.LinkedServerLogins)
            $tgtLsLogins = @($tgtLs.LinkedServerLogins)
            $allLsLoginNames = (@($srcLsLogins.Name) + @($tgtLsLogins.Name)) | Sort-Object -Unique

            foreach ($loginName in $allLsLoginNames) {
                $sl = $srcLsLogins | Where-Object { $_.Name -eq $loginName }
                $tl = $tgtLsLogins | Where-Object { $_.Name -eq $loginName }
                $loginLabel = "$lsName / Login: $loginName"

                Add-Row 'Server' 'LinkedServer' $loginLabel 'Mapped'      (!!$sl)         (!!$tl)
                if ($sl -and $tl) {
                    Add-Row 'Server' 'LinkedServer' $loginLabel 'RemoteUser'  $sl.RemoteUser  $tl.RemoteUser
                    Add-Row 'Server' 'LinkedServer' $loginLabel 'Impersonate' $sl.Impersonate $tl.Impersonate
                }
            }
        }
    }

    #
    # TempDB file configuration
    #
    $tempDbQuery = @"
SELECT file_id,
       type_desc,
       name AS LogicalName,
       CAST(size * 8.0 / 1024 AS decimal(10,0)) AS SizeMB,
       CASE is_percent_growth
           WHEN 1 THEN CAST(growth AS varchar) + '%'
           ELSE        CAST(CAST(growth * 8.0 / 1024 AS decimal(10,0)) AS varchar) + ' MB'
       END AS AutoGrowth
FROM sys.master_files
WHERE database_id = DB_ID('tempdb')
ORDER BY type, file_id
"@

    $srcTempDb = Invoke-DbaQuery -SqlInstance $Source -Database master -Query $tempDbQuery -As PSObject
    $tgtTempDb = Invoke-DbaQuery -SqlInstance $Target -Database master -Query $tempDbQuery -As PSObject

    # Summary counts
    # Filter once and derive summary counts from the same arrays used for per-file comparison.
    # Ordinal position (not file_id) is used since file_id is instance-local and can diverge
    # if files were ever removed and re-added.
    $srcDataFiles = @($srcTempDb | Where-Object { $_.type_desc -eq 'ROWS' } | Sort-Object file_id)
    $tgtDataFiles = @($tgtTempDb | Where-Object { $_.type_desc -eq 'ROWS' } | Sort-Object file_id)
    $srcLogFiles  = @($srcTempDb | Where-Object { $_.type_desc -eq 'LOG'  } | Sort-Object file_id)
    $tgtLogFiles  = @($tgtTempDb | Where-Object { $_.type_desc -eq 'LOG'  } | Sort-Object file_id)

    Add-Row 'Server' 'TempDbConfig' 'TempDB Summary' 'DataFileCount' $srcDataFiles.Count $tgtDataFiles.Count
    Add-Row 'Server' 'TempDbConfig' 'TempDB Summary' 'LogFileCount'  $srcLogFiles.Count  $tgtLogFiles.Count

    Compare-TempDbFileSet -SrcFiles $srcDataFiles -TgtFiles $tgtDataFiles -LabelPrefix 'Data File'
    Compare-TempDbFileSet -SrcFiles $srcLogFiles  -TgtFiles $tgtLogFiles  -LabelPrefix 'Log File'

    #
    # SQL Server Error Log configuration
    # Read directly from the registry via xp_instance_regread — more reliable than SMO/dbatools
    # for these values. NumErrorLogs defaults to 6 if never explicitly set; ErrorLogSizeKb
    # defaults to 0 (unlimited) if never explicitly set.
    #
    try {
        $errLogQuery = @"
DECLARE @numLogs   int
DECLARE @logSizeKb int
EXEC xp_instance_regread N'HKEY_LOCAL_MACHINE',
     N'Software\Microsoft\MSSQLServer\MSSQLServer',
     N'NumErrorLogs', @numLogs OUTPUT
EXEC xp_instance_regread N'HKEY_LOCAL_MACHINE',
     N'Software\Microsoft\MSSQLServer\MSSQLServer',
     N'ErrorLogSizeKb', @logSizeKb OUTPUT
SELECT ISNULL(@numLogs,   6) AS NumErrorLogs,
       ISNULL(@logSizeKb, 0) AS ErrorLogSizeKb
"@
        $srcErrLog = Invoke-DbaQuery -SqlInstance $Source -Database master -Query $errLogQuery -As PSObject
        $tgtErrLog = Invoke-DbaQuery -SqlInstance $Target -Database master -Query $errLogQuery -As PSObject

        Add-Row 'Server' 'ErrorLogConfig' 'Error Log' 'NumberOfLogFiles' $srcErrLog.NumErrorLogs   $tgtErrLog.NumErrorLogs
        Add-Row 'Server' 'ErrorLogConfig' 'Error Log' 'LogSizeKb'        $srcErrLog.ErrorLogSizeKb $tgtErrLog.ErrorLogSizeKb
    }
    catch {
        Write-Warning "Could not retrieve error log configuration: $_"
    }

    #
    # SQL Agent job history configuration
    #
    try {
        $srcAgent = Get-DbaAgentServer -SqlInstance $Source -ErrorAction Stop
        $tgtAgent = Get-DbaAgentServer -SqlInstance $Target -ErrorAction Stop

        Add-Row 'Server' 'AgentHistory' 'Agent History' 'MaximumHistoryRows'    $srcAgent.MaximumHistoryRows    $tgtAgent.MaximumHistoryRows
        Add-Row 'Server' 'AgentHistory' 'Agent History' 'MaximumJobHistoryRows' $srcAgent.MaximumJobHistoryRows $tgtAgent.MaximumJobHistoryRows
    }
    catch {
        Write-Warning "Could not retrieve SQL Agent configuration: $_"
    }
}

function Compare-Logins {
    param(
        [Microsoft.SqlServer.Management.Smo.Server]$Source,
        [Microsoft.SqlServer.Management.Smo.Server]$Target
    )

    $Source.Refresh()
    $Target.Refresh()

    # Exclude ## logins — these are internal SQL Server system certificates (e.g., ##MS_PolicyEventProcessingLogin##)
    # that are expected to have different SIDs on every instance and are not relevant to failover readiness.
    $srcLogins = Get-DbaLogin -SqlInstance $Source | Where-Object { $_.Name -notlike '##*' }
    $tgtLogins = Get-DbaLogin -SqlInstance $Target | Where-Object { $_.Name -notlike '##*' }

    foreach ($l in $srcLogins.Name) {
        $s = $srcLogins | Where-Object Name -eq $l
        $t = $tgtLogins | Where-Object Name -eq $l

        Add-Row 'Server' 'Login' $l 'Exists' $true (!!$t)

        if ($t) {
            Add-Row 'Server' 'Login' $l 'SID' ([System.BitConverter]::ToString($s.SID)) ([System.BitConverter]::ToString($t.SID))
            Add-Row 'Server' 'Login' $l 'IsDisabled' $s.IsDisabled $t.IsDisabled
            Add-Row 'Server' 'Login' $l 'DefaultDatabase' $s.DefaultDatabase $t.DefaultDatabase
        }
    }
}

function Compare-NonAgDatabases {
    param(
        [Microsoft.SqlServer.Management.Smo.Server]$Source,
        [Microsoft.SqlServer.Management.Smo.Server]$Target
    )

    $Source.Refresh()
    $Target.Refresh()

    $srcAgDbs = Get-AgDbNamesSafe -Server $Source
    $tgtAgDbs = Get-AgDbNamesSafe -Server $Target

    # Track AG database existence — source vs target, for inventory reporting
    $allAgDbNames = (@($srcAgDbs) + @($tgtAgDbs)) | Sort-Object -Unique
    foreach ($agDbName in $allAgDbNames) {
        $srcHas = $srcAgDbs -contains $agDbName
        $tgtHas = $tgtAgDbs -contains $agDbName
        Add-Row 'Database' 'AgDatabase' $agDbName 'ExistsOnTarget' $srcHas $tgtHas $agDbName
    }

    # Exclude any DB that is in the AG on EITHER instance — not just its own side
    $srcDbNames = $Source.Databases.Name | Where-Object { $_ -notin $allAgDbNames }
    $tgtDbNames = $Target.Databases.Name | Where-Object { $_ -notin $allAgDbNames }

    $allDbNames = ($srcDbNames + $tgtDbNames) | Sort-Object -Unique

    # Batch-query sys.databases once for all comparable DB properties.
    # Using -As PSObject so property access via variable ($row.$col) works reliably.
    # Results are stored in hashtables keyed by DB name for efficient per-database lookup.
    # snapshot_isolation_state: 0=Disabled, 1=Enabled, 2=PendingOff, 3=PendingOn
    # Values 1 and 3 mean snapshot isolation is effectively allowed.
    $sysDbQuery = @"
SELECT
    name,
    SUSER_SNAME(owner_sid) AS owner_name,
    compatibility_level,
    collation_name,
    user_access,
    user_access_desc,
    CAST(is_auto_close_on                          AS bit) AS is_auto_close_on,
    CAST(is_auto_shrink_on                         AS bit) AS is_auto_shrink_on,
    state,
    state_desc,
    CAST(is_in_standby                             AS bit) AS is_in_standby,
    CAST(is_supplemental_logging_enabled           AS bit) AS is_supplemental_logging_enabled,
    CAST(is_read_committed_snapshot_on             AS bit) AS is_read_committed_snapshot_on,
    recovery_model,
    recovery_model_desc,
    page_verify_option,
    page_verify_option_desc,
    CAST(is_auto_create_stats_on                   AS bit) AS is_auto_create_stats_on,
    CAST(is_auto_create_stats_incremental_on       AS bit) AS is_auto_create_stats_incremental_on,
    CAST(is_auto_update_stats_on                   AS bit) AS is_auto_update_stats_on,
    CAST(is_auto_update_stats_async_on             AS bit) AS is_auto_update_stats_async_on,
    CAST(is_ansi_null_default_on                   AS bit) AS is_ansi_null_default_on,
    CAST(is_ansi_nulls_on                          AS bit) AS is_ansi_nulls_on,
    CAST(is_ansi_padding_on                        AS bit) AS is_ansi_padding_on,
    CAST(is_ansi_warnings_on                       AS bit) AS is_ansi_warnings_on,
    CAST(is_arithabort_on                          AS bit) AS is_arithabort_on,
    CAST(is_concat_null_yields_null_on             AS bit) AS is_concat_null_yields_null_on,
    CAST(is_numeric_roundabort_on                  AS bit) AS is_numeric_roundabort_on,
    CAST(is_quoted_identifier_on                   AS bit) AS is_quoted_identifier_on,
    CAST(is_recursive_triggers_on                  AS bit) AS is_recursive_triggers_on,
    CAST(is_cursor_close_on_commit_on              AS bit) AS is_cursor_close_on_commit_on,
    CAST(is_local_cursor_default                   AS bit) AS is_local_cursor_default,
    CAST(is_fulltext_enabled                       AS bit) AS is_fulltext_enabled,
    CAST(is_trustworthy_on                         AS bit) AS is_trustworthy_on,
    CAST(is_db_chaining_on                         AS bit) AS is_db_chaining_on,
    CAST(is_parameterization_forced                AS bit) AS is_parameterization_forced,
    CAST(is_master_key_encrypted_by_server         AS bit) AS is_master_key_encrypted_by_server,
    CAST(is_query_store_on                         AS bit) AS is_query_store_on,
    CAST(is_published                              AS bit) AS is_published,
    CAST(is_subscribed                             AS bit) AS is_subscribed,
    CAST(is_merge_published                        AS bit) AS is_merge_published,
    CAST(is_distributor                            AS bit) AS is_distributor,
    CAST(is_sync_with_backup                       AS bit) AS is_sync_with_backup,
    CAST(is_broker_enabled                         AS bit) AS is_broker_enabled,
    CAST(is_date_correlation_on                    AS bit) AS is_date_correlation_on,
    CAST(is_cdc_enabled                            AS bit) AS is_cdc_enabled,
    CAST(is_encrypted                              AS bit) AS is_encrypted,
    CAST(is_honor_broker_priority_on               AS bit) AS is_honor_broker_priority_on,
    -- snapshot_isolation_state: 0=Disabled, 1=Enabled, 2=PendingOff, 3=PendingOn
    -- Values 1 and 3 mean snapshot isolation is effectively allowed.
    CAST(CASE WHEN snapshot_isolation_state IN (1,3) THEN 1 ELSE 0 END AS bit) AS AllowSnapshotIsolation,
    default_language_lcid,
    default_language_name,
    default_fulltext_language_name,
    CAST(is_nested_triggers_on                     AS bit) AS is_nested_triggers_on,
    CAST(is_transform_noise_words_on               AS bit) AS is_transform_noise_words_on,
    two_digit_year_cutoff,
    containment,
    containment_desc,
    target_recovery_time_in_seconds,
    delayed_durability,
    delayed_durability_desc,
    CAST(is_memory_optimized_elevate_to_snapshot_on AS bit) AS is_memory_optimized_elevate_to_snapshot_on,
    CAST(is_federation_member                      AS bit) AS is_federation_member,
    CAST(is_remote_data_archive_enabled            AS bit) AS is_remote_data_archive_enabled,
    CAST(is_mixed_page_allocation_on               AS bit) AS is_mixed_page_allocation_on,
    CAST(is_temporal_history_retention_enabled     AS bit) AS is_temporal_history_retention_enabled
FROM sys.databases
"@
    $srcSysDbHash = @{}
    Invoke-DbaQuery -SqlInstance $Source -Database master -Query $sysDbQuery -As PSObject |
        ForEach-Object { $srcSysDbHash[$_.name] = $_ }

    $tgtSysDbHash = @{}
    Invoke-DbaQuery -SqlInstance $Target -Database master -Query $sysDbQuery -As PSObject |
        ForEach-Object { $tgtSysDbHash[$_.name] = $_ }

    # Columns to compare from sys.databases, in display order.
    # These are the result column names from $sysDbQuery above.
    $sysDbCols = @(
        'owner_name', 'compatibility_level', 'collation_name',
        'user_access', 'user_access_desc',
        'is_auto_close_on', 'is_auto_shrink_on',
        'state', 'state_desc', 'is_in_standby', 'is_supplemental_logging_enabled',
        'is_read_committed_snapshot_on',
        'recovery_model', 'recovery_model_desc',
        'page_verify_option', 'page_verify_option_desc',
        'is_auto_create_stats_on', 'is_auto_create_stats_incremental_on',
        'is_auto_update_stats_on', 'is_auto_update_stats_async_on',
        'is_ansi_null_default_on', 'is_ansi_nulls_on', 'is_ansi_padding_on',
        'is_ansi_warnings_on', 'is_arithabort_on', 'is_concat_null_yields_null_on',
        'is_numeric_roundabort_on', 'is_quoted_identifier_on',
        'is_recursive_triggers_on', 'is_cursor_close_on_commit_on', 'is_local_cursor_default',
        'is_fulltext_enabled', 'is_trustworthy_on', 'is_db_chaining_on',
        'is_parameterization_forced', 'is_master_key_encrypted_by_server', 'is_query_store_on',
        'is_published', 'is_subscribed', 'is_merge_published', 'is_distributor',
        'is_sync_with_backup', 'is_broker_enabled', 'is_date_correlation_on',
        'is_cdc_enabled', 'is_encrypted', 'is_honor_broker_priority_on',
        'AllowSnapshotIsolation',
        'default_language_lcid', 'default_language_name', 'default_fulltext_language_name',
        'is_nested_triggers_on', 'is_transform_noise_words_on', 'two_digit_year_cutoff',
        'containment', 'containment_desc', 'target_recovery_time_in_seconds',
        'delayed_durability', 'delayed_durability_desc',
        'is_memory_optimized_elevate_to_snapshot_on',
        'is_federation_member', 'is_remote_data_archive_enabled',
        'is_mixed_page_allocation_on', 'is_temporal_history_retention_enabled'
    )

    foreach ($dbName in $allDbNames) {
        $sdb = $Source.Databases[$dbName]
        $tdb = $Target.Databases[$dbName]

        $srcExists = $null -ne $sdb
        $tgtExists = $null -ne $tdb

        Add-Row 'Database' 'Database' $dbName 'ExistsOnTarget' $srcExists $tgtExists $dbName

        if (-not ($srcExists -and $tgtExists)) {
            continue
        }

        $sdb.Refresh()
        $tdb.Refresh()

        # ReadOnly is not in sys.databases — read from SMO directly
        Add-Row 'Database' 'DbProperty' $dbName 'ReadOnly' $sdb.ReadOnly $tdb.ReadOnly $dbName

        # All other DB properties come from sys.databases (see batch query above)
        $srcSysRow = $srcSysDbHash[$dbName]
        $tgtSysRow = $tgtSysDbHash[$dbName]

        foreach ($col in $sysDbCols) {
            $srcVal = if ($srcSysRow) { $srcSysRow.$col } else { $null }
            $tgtVal = if ($tgtSysRow) { $tgtSysRow.$col } else { $null }
            Add-Row 'Database' 'DbProperty' $dbName $col $srcVal $tgtVal $dbName
        }

        #
        # Users, roles, permissions (kept simple and structural)
        #

        # Guard: skip security enumeration for databases that are not in a normal, online, read-write state.
        # SMO throws a T-SQL execution exception when trying to enumerate Users/Roles/Permissions on
        # databases that are offline, restoring, in emergency mode, or otherwise inaccessible.
        $skipSecurity = ($sdb.Status -ne [Microsoft.SqlServer.Management.Smo.DatabaseStatus]::Normal) -or
                        ($tdb.Status -ne [Microsoft.SqlServer.Management.Smo.DatabaseStatus]::Normal)

        if ($skipSecurity) {
            Write-Warning "Skipping Users/Roles/Permissions for '$dbName' — database is not in Normal status (Source: $($sdb.Status), Target: $($tdb.Status))"
            Add-Row 'Database' 'User' "$dbName / (skipped)" 'ExistsOnTarget' $null $null $dbName
        }
        else {
            # Users
            $srcUsers = @()
            $tgtUsers = @()

            try   { $srcUsers = @($sdb.Users | Where-Object { -not $_.IsSystemObject }) }
            catch { Write-Warning "Could not enumerate users for '$dbName' on source: $($_.Exception.Message)" }

            try   { $tgtUsers = @($tdb.Users | Where-Object { -not $_.IsSystemObject }) }
            catch { Write-Warning "Could not enumerate users for '$dbName' on target: $($_.Exception.Message)" }

            $srcUserHash = @{}; $srcUsers | ForEach-Object { $srcUserHash[$_.Name] = $true }
            $tgtUserHash = @{}; $tgtUsers | ForEach-Object { $tgtUserHash[$_.Name] = $true }
            $allUsers = (@($srcUserHash.Keys) + @($tgtUserHash.Keys)) | Sort-Object -Unique
            foreach ($u in $allUsers) {
                Add-Row 'Database' 'User' "$dbName / $u" 'ExistsOnTarget' $srcUserHash.ContainsKey($u) $tgtUserHash.ContainsKey($u) $dbName
            }

            # Roles
            $srcRoles = @()
            $tgtRoles = @()

            try   { $srcRoles = @($sdb.Roles | Where-Object { -not $_.IsFixedRole }) }
            catch { Write-Warning "Could not enumerate roles for '$dbName' on source: $($_.Exception.Message)" }

            try   { $tgtRoles = @($tdb.Roles | Where-Object { -not $_.IsFixedRole }) }
            catch { Write-Warning "Could not enumerate roles for '$dbName' on target: $($_.Exception.Message)" }

            $srcRoleHash = @{}; $srcRoles | ForEach-Object { $srcRoleHash[$_.Name] = $true }
            $tgtRoleHash = @{}; $tgtRoles | ForEach-Object { $tgtRoleHash[$_.Name] = $true }
            $allRoles = (@($srcRoleHash.Keys) + @($tgtRoleHash.Keys)) | Sort-Object -Unique
            foreach ($r in $allRoles) {
                Add-Row 'Database' 'Role' "$dbName / $r" 'ExistsOnTarget' $srcRoleHash.ContainsKey($r) $tgtRoleHash.ContainsKey($r) $dbName
            }

            # Permissions (high-level)
            $srcPerms   = @()
            $tgtPerms   = @()
            $srcPermKey = @()
            $tgtPermKey = @()

            try {
                $srcPerms   = $sdb.EnumDatabasePermissions()
                $srcPermKey = $srcPerms | ForEach-Object { "$($_.Grantee)/$($_.PermissionType)/$($_.PermissionState)" }
            }
            catch { Write-Warning "Could not enumerate permissions for '$dbName' on source: $($_.Exception.Message)" }

            try {
                $tgtPerms   = $tdb.EnumDatabasePermissions()
                $tgtPermKey = $tgtPerms | ForEach-Object { "$($_.Grantee)/$($_.PermissionType)/$($_.PermissionState)" }
            }
            catch { Write-Warning "Could not enumerate permissions for '$dbName' on target: $($_.Exception.Message)" }

            $srcPermSet = [System.Collections.Generic.HashSet[string]]::new([string[]]@($srcPermKey))
            $tgtPermSet = [System.Collections.Generic.HashSet[string]]::new([string[]]@($tgtPermKey))
            $allPerms = (@($srcPermKey) + @($tgtPermKey)) | Sort-Object -Unique
            foreach ($p in $allPerms) {
                Add-Row 'Database' 'Permission' "$dbName / $p" 'ExistsOnTarget' $srcPermSet.Contains($p) $tgtPermSet.Contains($p) $dbName
            }
        }
    }
}

function Write-HtmlReport {
    param(
        [string]$SourceInstance,
        [string]$TargetInstance
    )

    $html = @"
<html>
<head>
<title>SQL Server Instance Comparison Report</title>
<style>
body { font-family: Arial, sans-serif; font-size: 13px; }
table { border-collapse: collapse; margin-bottom: 10px; width: 100%; }
th, td { border: 1px solid #ccc; padding: 4px 6px; }
th { background-color: #eee; text-align: left; }
tr.mismatch { background-color: #ffe0e0; }
tr.match { background-color: #e8ffe8; }
h1 { font-size: 28px; margin-bottom: 8px; }
h2 { font-size: 22px; margin-top: 30px; border-bottom: 2px solid #999; padding-bottom: 4px; }
details { margin: 14px 0; }
details[open] summary { margin-bottom: 6px; }
summary {
    cursor: pointer;
    font-size: 17px;
    font-weight: bold;
    color: #333;
    padding: 4px 2px;
    user-select: none;
    list-style: none;
}
summary::-webkit-details-marker { display: none; }
summary::before { content: '\25B6\00A0'; font-size: 11px; color: #666; }
details[open] summary::before { content: '\25BC\00A0'; font-size: 11px; color: #666; }
</style>
</head>
<body>
<h1>SQL Server Instance Comparison Report</h1>
<p style="font-size: 16px; line-height: 1.8;">
<b>Source:</b> $SourceInstance<br>
<b>Target:</b> $TargetInstance<br>
<b>Generated:</b> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
</p>
"@

    #
    # Server-level sections — one table per logical category group
    #
    $serverRows = $script:rows | Where-Object Scope -eq 'Server'

    if ($serverRows) {
        $html += "<h2>Server-Level Comparison</h2>`n"

        # Ordered map: section heading → one or more Category values to include
        $serverSections = [ordered]@{
            'Server Properties'          = @('ServerProperty')
            'Service Properties'         = @('ServiceProperty')
            'TempDB Configuration'       = @('TempDbConfig')
            'Error Log Configuration'    = @('ErrorLogConfig')
            'SQL Agent History'          = @('AgentHistory')
            'sp_configure'               = @('sp_configure')
            'Startup Parameters'         = @('StartupParameters')
            'Trace Flags'                = @('TraceFlag', 'TraceFlagSummary')
            'Linked Servers'             = @('LinkedServer')
            'Logins'                     = @('Login')
        }

        foreach ($sectionTitle in $serverSections.Keys) {
            $cats       = $serverSections[$sectionTitle]
            $sectionRows = $serverRows | Where-Object { $_.Category -in $cats }
            if (-not $sectionRows) { continue }

            $html += "<details open>`n"
            $html += "<summary>$sectionTitle</summary>`n"
            $html += "<table>`n"
            $html += "<tr><th>Item</th><th>Property</th><th>Source ($SourceInstance)</th><th>Target ($TargetInstance)</th><th>Match</th></tr>`n"

            foreach ($row in $sectionRows) {
                $class = if ($row.Match) { 'match' } else { 'mismatch' }
                $html += "<tr class='$class'><td>$($row.Item)</td><td>$($row.Property)</td><td>$($row.Source)</td><td>$($row.Target)</td><td>$($row.Match)</td></tr>`n"
            }

            $html += "</table>`n"
            $html += "</details>`n"
        }
    }

    #
    # Database-level sections — one table per database
    #
    $dbRows = $script:rows | Where-Object Scope -eq 'Database'

    if ($dbRows) {
        $html += "<h2>Database-Level Comparison</h2>`n"

        # Source database inventory — all source databases (AG and non-AG) with True/False for presence on target
        $srcDbRows = $dbRows | Where-Object { $_.Category -in @('Database', 'AgDatabase') -and [bool]$_.Source }
        $html += "<details open>`n"
        $html += "<summary>Source Database Inventory</summary>`n"
        $html += "<table>`n"
        $html += "<tr><th>Database</th><th>Type</th><th>Present on Target ($TargetInstance)</th></tr>`n"
        foreach ($row in ($srcDbRows | Sort-Object DatabaseName)) {
            $class = if ($row.Match) { 'match' } else { 'mismatch' }
            $type  = if ($row.Category -eq 'AgDatabase') { 'AG' } else { 'Non-AG' }
            $html += "<tr class='$class'><td>$($row.DatabaseName)</td><td>$type</td><td>$($row.Target)</td></tr>`n"
        }
        $html += "</table>`n"
        $html += "</details>`n"

        # Per-database detail sections — exclude AgDatabase rows (those are inventory-only)
        $databases = ($dbRows | Where-Object { $_.Category -ne 'AgDatabase' } |
            Select-Object -ExpandProperty DatabaseName | Sort-Object -Unique)

        foreach ($db in $databases) {
            $thisDbRows = $dbRows | Where-Object { $_.DatabaseName -eq $db }

            $html += "<details open>`n"
            $html += "<summary>$db</summary>`n"
            $html += "<table>`n"
            $html += "<tr><th>Category</th><th>Item</th><th>Property</th><th>Source ($SourceInstance)</th><th>Target ($TargetInstance)</th><th>Match</th></tr>`n"

            foreach ($row in $thisDbRows) {
                $class = if ($row.Match) { 'match' } else { 'mismatch' }
                $html += "<tr class='$class'><td>$($row.Category)</td><td>$($row.Item)</td><td>$($row.Property)</td><td>$($row.Source)</td><td>$($row.Target)</td><td>$($row.Match)</td></tr>`n"
            }

            $html += "</table>`n"
            $html += "</details>`n"
        }
    }

    # Mismatch Summary — all rows where Match is false, excluding CompatibilityLevel
    $excludeProperties = @('CompatibilityLevel', 'ServiceAccount', 'MasterData', 'MasterLog', 'ErrorLog')
    $excludeItems      = @('VersionHighPartOfSqlServer', 'VersionLowPartOfSqlServer')

    $mismatchRows = $script:rows | Where-Object {
        $_.Match -eq $false -and
        $_.Property -notin $excludeProperties -and
        $_.Item     -notin $excludeItems
    }

    $html += "<h2>Mismatch Summary</h2>`n"

    if ($mismatchRows) {
        $html += "<details open>`n"
        $html += "<summary>All Mismatches ($($mismatchRows.Count))</summary>`n"
        $html += "<table>`n"
        $html += "<tr><th>Scope</th><th>Category</th><th>Item</th><th>Property</th><th>Source ($SourceInstance)</th><th>Target ($TargetInstance)</th></tr>`n"
        foreach ($row in $mismatchRows) {
            $html += "<tr class='mismatch'><td>$($row.Scope)</td><td>$($row.Category)</td><td>$($row.Item)</td><td>$($row.Property)</td><td>$($row.Source)</td><td>$($row.Target)</td></tr>`n"
        }
        $html += "</table>`n"
        $html += "</details>`n"
    } else {
        $html += "<p style='color: green; font-size: 15px;'>No mismatches found.</p>`n"
    }

    $html += "</body></html>"

    $timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
    $reportName = "SqlInstanceComparison_${SourceInstance}_vs_${TargetInstance}_$timestamp.html"
    $reportName = $reportName -replace "[:\\\/]", "_"

    $reportPath = Join-Path $PSScriptRoot $reportName

    $html | Out-File -FilePath $reportPath -Encoding UTF8

    Write-Host "HTML report written to: $reportPath" -ForegroundColor Green

    return $reportPath
}

#
# Entry point
#
Write-Host "Connecting to SQL instances..." -ForegroundColor Cyan

try {
    $src = Connect-DbaInstance -SqlInstance $SourceInstance -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to source instance '$SourceInstance': $_"
    exit 1
}

try {
    $tgt = Connect-DbaInstance -SqlInstance $TargetInstance -TrustServerCertificate -ErrorAction Stop
}
catch {
    Write-Error "Failed to connect to target instance '$TargetInstance': $_"
    exit 1
}

$script:rows = @()

Compare-ServerLevel    -Source $src -Target $tgt
Compare-Logins         -Source $src -Target $tgt
Compare-NonAgDatabases -Source $src -Target $tgt

$myReport = Write-HtmlReport -SourceInstance $SourceInstance -TargetInstance $TargetInstance

Write-Host "Comparison complete. HTML report generated." -ForegroundColor Green

Start-Process $myReport
