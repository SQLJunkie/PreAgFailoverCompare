# SqlCompareFunctions.ps1
# Core comparison functions for SQL Server instance configuration

# Global collection for rows
$script:rows = @()

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

    $agDbNames = @()

    try {
        $Server.Refresh()
        $availabilityGroups = $Server.AvailabilityGroups
        if ($availabilityGroups) {
            foreach ($ag in $availabilityGroups) {
                foreach ($db in $ag.AvailabilityDatabases) {
                    if ($db.Name -and -not $agDbNames.Contains($db.Name)) {
                        $agDbNames += $db.Name
                    }
                }
            }
        }
    }
    catch {
        # If AG metadata isn't accessible, just return empty
    }

    return $agDbNames
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
    $srcDataCount = @($srcTempDb | Where-Object { $_.type_desc -eq 'ROWS' }).Count
    $tgtDataCount = @($tgtTempDb | Where-Object { $_.type_desc -eq 'ROWS' }).Count
    $srcLogCount  = @($srcTempDb | Where-Object { $_.type_desc -eq 'LOG'  }).Count
    $tgtLogCount  = @($tgtTempDb | Where-Object { $_.type_desc -eq 'LOG'  }).Count

    Add-Row 'Server' 'TempDbConfig' 'TempDB Summary' 'DataFileCount' $srcDataCount $tgtDataCount
    Add-Row 'Server' 'TempDbConfig' 'TempDB Summary' 'LogFileCount'  $srcLogCount  $tgtLogCount

    # Per data file: size and autogrowth (keyed by file_id for cross-instance alignment)
    $srcDataFiles = @($srcTempDb | Where-Object { $_.type_desc -eq 'ROWS' })
    $tgtDataFiles = @($tgtTempDb | Where-Object { $_.type_desc -eq 'ROWS' })

    $allDataFileIds = (@($srcDataFiles.file_id) + @($tgtDataFiles.file_id)) | Sort-Object -Unique
    foreach ($fid in $allDataFileIds) {
        $sf    = $srcDataFiles | Where-Object { $_.file_id -eq $fid }
        $tf    = $tgtDataFiles | Where-Object { $_.file_id -eq $fid }
        $label = "Data File $fid"
        Add-Row 'Server' 'TempDbConfig' $label 'SizeMB'     ($sf.SizeMB)     ($tf.SizeMB)
        Add-Row 'Server' 'TempDbConfig' $label 'AutoGrowth' ($sf.AutoGrowth) ($tf.AutoGrowth)
    }

    # Log file(s)
    $srcLogFiles = @($srcTempDb | Where-Object { $_.type_desc -eq 'LOG' })
    $tgtLogFiles = @($tgtTempDb | Where-Object { $_.type_desc -eq 'LOG' })

    $allLogFileIds = (@($srcLogFiles.file_id) + @($tgtLogFiles.file_id)) | Sort-Object -Unique
    foreach ($fid in $allLogFileIds) {
        $sf    = $srcLogFiles | Where-Object { $_.file_id -eq $fid }
        $tf    = $tgtLogFiles | Where-Object { $_.file_id -eq $fid }
        $label = "Log File $fid"
        Add-Row 'Server' 'TempDbConfig' $label 'SizeMB'     ($sf.SizeMB)     ($tf.SizeMB)
        Add-Row 'Server' 'TempDbConfig' $label 'AutoGrowth' ($sf.AutoGrowth) ($tf.AutoGrowth)
    }

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

    $srcDbNames = $Source.Databases.Name | Where-Object { $_ -notin $srcAgDbs }
    $tgtDbNames = $Target.Databases.Name | Where-Object { $_ -notin $tgtAgDbs }

    $allDbNames = ($srcDbNames + $tgtDbNames) | Sort-Object -Unique

    # Batch-query sys.databases once for properties not exposed as simple booleans in SMO.
    # AllowSnapshotIsolation → SnapshotIsolationState enum on Database (not a bool)
    # ReadCommittedSnapshot  → no direct SMO bool property in this version
    # ConcatNullYieldsNull   → DatabaseOptions.ConcatenateNullYieldsNull inconsistently accessible
    # Using -As PSObject so property access via variable ($row.$prop) works reliably.
    # Results are stored in hashtables keyed by DB name for efficient per-database lookup.
    # is_allow_snapshot_isolation_on does not exist in sys.databases.
    # snapshot_isolation_state: 0=Disabled, 1=Enabled, 2=PendingOff, 3=PendingOn
    # Values 1 and 3 mean snapshot isolation is effectively allowed.
    $sysDbQuery = @"
SELECT name,
       CAST(CASE WHEN snapshot_isolation_state IN (1, 3) THEN 1 ELSE 0 END AS bit) AS AllowSnapshotIsolation,
       CAST(is_read_committed_snapshot_on AS bit) AS ReadCommittedSnapshot,
       CAST(is_concat_null_yields_null_on AS bit) AS ConcatNullYieldsNull
FROM sys.databases
"@
    $srcSysDbHash = @{}
    Invoke-DbaQuery -SqlInstance $Source -Database master -Query $sysDbQuery -As PSObject |
        ForEach-Object { $srcSysDbHash[$_.name] = $_ }

    $tgtSysDbHash = @{}
    Invoke-DbaQuery -SqlInstance $Target -Database master -Query $sysDbQuery -As PSObject |
        ForEach-Object { $tgtSysDbHash[$_.name] = $_ }

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

        # Basic DB properties
        $dbProps = @(
            'Owner',
            'Collation',
            'RecoveryModel',
            'CompatibilityLevel',
            'ContainmentType',
            'PageVerify',
            'AutoClose',
            'AutoShrink',
            'ReadOnly'
        )

        foreach ($p in $dbProps) {
            $srcVal = $sdb.$p
            $tgtVal = $tdb.$p
            Add-Row 'Database' 'DbProperty' $dbName $p $srcVal $tgtVal $dbName
        }

        # Database options
        $srcOpt = $sdb.DatabaseOptions
        $tgtOpt = $tdb.DatabaseOptions

        if ($srcOpt -and $tgtOpt) {
            # Properties confirmed accessible via DatabaseOptions
            $dbOptMap = @{
                "AnsiNullDefault"           = "ANSI Null Default"
                "AnsiNullsEnabled"          = "ANSI Nulls Enabled"
                "AnsiPaddingEnabled"        = "ANSI Padding Enabled"
                "AnsiWarningsEnabled"       = "ANSI Warnings Enabled"
                "ArithmeticAbortEnabled"    = "Arithmetic Abort Enabled"
                "NumericRoundAbortEnabled"  = "Numeric Round Abort"
                "QuotedIdentifiersEnabled"  = "Quoted Identifier Enabled"
                "RecursiveTriggersEnabled"  = "Recursive Triggers Enabled"
                "Trustworthy"               = "Trustworthy"
                "AutoUpdateStatisticsAsync" = "Auto Update Stats Async"
            }

            foreach ($prop in $dbOptMap.Keys) {
                $label = $dbOptMap[$prop]
                Add-Row 'Database' 'DbOptions' "$dbName / $label" $prop $srcOpt.$prop $tgtOpt.$prop $dbName
            }

            # Properties read directly from Database object (not DatabaseOptions)
            $dbDirectMap = @{
                "AutoUpdateStatisticsEnabled" = "Auto Update Statistics"
                "AutoCreateStatisticsEnabled" = "Auto Create Statistics"
            }

            foreach ($prop in $dbDirectMap.Keys) {
                $label = $dbDirectMap[$prop]
                Add-Row 'Database' 'DbOptions' "$dbName / $label" $prop $sdb.$prop $tdb.$prop $dbName
            }

            # Properties sourced from sys.databases (not accessible as simple booleans via SMO)
            $srcSysRow = $srcSysDbHash[$dbName]
            $tgtSysRow = $tgtSysDbHash[$dbName]

            $sysDbOptMap = @{
                "AllowSnapshotIsolation" = "Allow Snapshot Isolation"
                "ReadCommittedSnapshot"  = "Read Committed Snapshot"
                "ConcatNullYieldsNull"   = "Concat Null Yields Null"
            }

            foreach ($prop in $sysDbOptMap.Keys) {
                $label  = $sysDbOptMap[$prop]
                $srcVal = if ($srcSysRow) { [bool]$srcSysRow.$prop } else { $null }
                $tgtVal = if ($tgtSysRow) { [bool]$tgtSysRow.$prop } else { $null }
                Add-Row 'Database' 'DbOptions' "$dbName / $label" $prop $srcVal $tgtVal $dbName
            }
        }

        #
        # Users, roles, permissions (kept simple and structural)
        #

        # Users
        $srcUsers = $sdb.Users | Where-Object { -not $_.IsSystemObject }
        $tgtUsers = $tdb.Users | Where-Object { -not $_.IsSystemObject }

        $allUsers = (@($srcUsers.Name) + @($tgtUsers.Name)) | Sort-Object -Unique
        foreach ($u in $allUsers) {
            $su = $srcUsers | Where-Object Name -eq $u
            $tu = $tgtUsers | Where-Object Name -eq $u

            Add-Row 'Database' 'User' "$dbName / $u" 'ExistsOnTarget' (!!$su) (!!$tu) $dbName
        }

        # Roles
        $srcRoles = $sdb.Roles | Where-Object { -not $_.IsFixedRole }
        $tgtRoles = $tdb.Roles | Where-Object { -not $_.IsFixedRole }

        $allRoles = (@($srcRoles.Name) + @($tgtRoles.Name)) | Sort-Object -Unique
        foreach ($r in $allRoles) {
            $sr = $srcRoles | Where-Object Name -eq $r
            $tr = $tgtRoles | Where-Object Name -eq $r

            Add-Row 'Database' 'Role' "$dbName / $r" 'ExistsOnTarget' (!!$sr) (!!$tr) $dbName
        }

        # Permissions (high-level)
        $srcPerms = $sdb.EnumDatabasePermissions()
        $tgtPerms = $tdb.EnumDatabasePermissions()

        $srcPermKey = $srcPerms | ForEach-Object { "$($_.Grantee)/$($_.PermissionType)/$($_.PermissionState)" }
        $tgtPermKey = $tgtPerms | ForEach-Object { "$($_.Grantee)/$($_.PermissionType)/$($_.PermissionState)" }

        $allPerms = (@($srcPermKey) + @($tgtPermKey)) | Sort-Object -Unique
        foreach ($p in $allPerms) {
            $srcHas = $srcPermKey -contains $p
            $tgtHas = $tgtPermKey -contains $p
            Add-Row 'Database' 'Permission' "$dbName / $p" 'ExistsOnTarget' $srcHas $tgtHas $dbName
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

function Invoke-SqlInstanceComparison {
    param(
        [string]$SourceInstance,
        [string]$TargetInstance
    )

    $script:rows = @()

    $src = Connect-DbaInstance -SqlInstance $SourceInstance -TrustServerCertificate -ErrorAction Stop
    $tgt = Connect-DbaInstance -SqlInstance $TargetInstance -TrustServerCertificate -ErrorAction Stop

    Compare-ServerLevel -Source $src -Target $tgt
    Compare-Logins      -Source $src -Target $tgt
    Compare-NonAgDatabases -Source $src -Target $tgt

    $reportPath = Write-HtmlReport -SourceInstance $SourceInstance -TargetInstance $TargetInstance
    return $reportPath
}

