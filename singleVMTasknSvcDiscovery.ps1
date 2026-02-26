#requires -Version 5.1
param(
  [string] $OutputRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Log { param([string]$m) Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) }

function Ensure-ImportExcel {
  try {
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
      Log "ImportExcel not found. Attempting to install (CurrentUser)..."
      Install-Module ImportExcel -Scope CurrentUser -Force -Confirm:$false -ErrorAction Stop
    }

    Import-Module ImportExcel -ErrorAction Stop | Out-Null
    Log "ImportExcel module loaded successfully."
    return $true
  }
  catch {
    Log ("WARNING: ImportExcel unavailable. Falling back to CSV. Reason: {0}" -f $_.Exception.Message)
    return $false
  }
}

function Get-InstalledApps {
  $apps = @()
  $paths = @(
    "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
  )
  foreach ($p in $paths) {
    if (Test-Path $p) {
      Get-ItemProperty $p -ErrorAction SilentlyContinue | ForEach-Object {
        $dnProp = $_.PSObject.Properties['DisplayName']
        if ($dnProp -and $dnProp.Value -and $dnProp.Value.ToString().Trim().Length -gt 0) {
          $verProp = $_.PSObject.Properties['DisplayVersion']
          $pubProp = $_.PSObject.Properties['Publisher']
          $idProp  = $_.PSObject.Properties['InstallDate']
          $apps += [pscustomobject]@{
            DisplayName    = $dnProp.Value
            DisplayVersion = if ($verProp) { $verProp.Value } else { $null }
            Publisher      = if ($pubProp) { $pubProp.Value } else { $null }
            InstallDate    = if ($idProp)  { $idProp.Value } else { $null }
          }
        }
      }
    }
  }
  return ($apps | Sort-Object DisplayName -Unique)
}

function Get-ObservedListeningPorts {
  # Observed = what the OS reports as actively listening (TCP Listen + UDP endpoints)
  $svcMap = @{}
  try {
    Get-CimInstance Win32_Service -ErrorAction SilentlyContinue | ForEach-Object {
      $pidProp  = $_.PSObject.Properties['ProcessId']
      $nameProp = $_.PSObject.Properties['Name']
      if ($pidProp -and $nameProp -and $pidProp.Value -and $pidProp.Value -ne 0) {
        $svcMap[$pidProp.Value] = $nameProp.Value
      }
    }
  } catch { }

  $ports = @()

  try {
    Get-NetTCPConnection -State Listen -ErrorAction Stop | ForEach-Object {
      $ownerPid = $null
      $ownProp = $_.PSObject.Properties['OwningProcess']
      if ($ownProp) { $ownerPid = $ownProp.Value }

      $procName = $null
      if ($ownerPid -ne $null) {
        try { $procName = (Get-Process -Id $ownerPid -ErrorAction Stop).ProcessName } catch { }
      }

      $svcName = $null
      if ($ownerPid -ne $null -and $svcMap.ContainsKey($ownerPid)) { $svcName = $svcMap[$ownerPid] }

      $ports += [pscustomobject]@{
        Protocol      = "TCP"
        LocalAddress  = $_.LocalAddress
        LocalPort     = $_.LocalPort
        OwningProcess = $ownerPid
        ProcessName   = $procName
        Service       = $svcName
      }
    }
  } catch { }

  try {
    Get-NetUDPEndpoint -ErrorAction Stop | ForEach-Object {
      $ownerPid = $null
      $ownProp = $_.PSObject.Properties['OwningProcess']
      if ($ownProp) { $ownerPid = $ownProp.Value }

      $procName = $null
      if ($ownerPid -ne $null) {
        try { $procName = (Get-Process -Id $ownerPid -ErrorAction Stop).ProcessName } catch { }
      }

      $svcName = $null
      if ($ownerPid -ne $null -and $svcMap.ContainsKey($ownerPid)) { $svcName = $svcMap[$ownerPid] }

      $ports += [pscustomobject]@{
        Protocol      = "UDP"
        LocalAddress  = $_.LocalAddress
        LocalPort     = $_.LocalPort
        OwningProcess = $ownerPid
        ProcessName   = $procName
        Service       = $svcName
      }
    }
  } catch { }

  return ($ports | Sort-Object Protocol, LocalPort, ProcessName)
}

function Get-IISInfo {
  $iisPresent = $false
  try { if (Get-Service -Name W3SVC -ErrorAction SilentlyContinue) { $iisPresent = $true } } catch { }

  $sites    = @()
  $pools    = @()
  $apps     = @()
  $bindings = @()

  if (-not $iisPresent) {
    return [pscustomobject]@{ Present=$false; Sites=@(); AppPools=@(); Apps=@(); Bindings=@() }
  }

  if (-not (Get-Module -ListAvailable -Name WebAdministration)) {
    return [pscustomobject]@{ Present=$true; Sites=@(); AppPools=@(); Apps=@(); Bindings=@() }
  }

  try {
    Import-Module WebAdministration -ErrorAction Stop | Out-Null

    try { $sites = @(Get-Website | Select-Object Name, State, PhysicalPath, ApplicationPool, ID) } catch { $sites=@() }

    try {
      $pools = @(Get-ChildItem IIS:\AppPools | ForEach-Object {
        [pscustomobject]@{
          Name         = $_.Name
          State        = $_.State
          Runtime      = $_.managedRuntimeVersion
          PipelineMode = $_.managedPipelineMode
          IdentityType = $_.processModel.identityType
        }
      })
    } catch { $pools=@() }

    try {
      $apps = @()
      foreach ($s in $sites) {
        try {
          @(Get-WebApplication -Site $s.Name) | ForEach-Object {
            $apps += [pscustomobject]@{
              Site            = $s.Name
              Path            = $_.Path
              PhysicalPath    = $_.PhysicalPath
              ApplicationPool = $_.ApplicationPool
            }
          }
        } catch { }
      }
      $apps = @($apps)
    } catch { $apps=@() }

    try {
      $bindings = @()
      foreach ($s in @(Get-Website)) {
        foreach ($b in $s.Bindings.Collection) {
          $bi = $b.bindingInformation
          $ipPart = $null; $portPart = $null; $hostPart = $null
          if ($bi) {
            $parts = $bi.Split(":")
            if ($parts.Count -ge 1) { $ipPart = $parts[0] }
            if ($parts.Count -ge 2) { $portPart = $parts[1] }
            if ($parts.Count -ge 3) { $hostPart = $parts[2] }
          }
          $bindings += [pscustomobject]@{
            Site        = $s.Name
            Protocol    = $b.protocol
            IP          = $ipPart
            Port        = $portPart
            HostHeader  = $hostPart
            BindingInfo = $bi
          }
        }
      }
      $bindings = @($bindings)
    } catch { $bindings=@() }

  } catch {
    return [pscustomobject]@{ Present=$true; Sites=@(); AppPools=@(); Apps=@(); Bindings=@() }
  }

  return [pscustomobject]@{
    Present  = $true
    Sites    = @($sites)
    AppPools = @($pools)
    Apps     = @($apps)
    Bindings = @($bindings)
  }
}

# ---------------- Collect data ----------------
Log "Collecting summary..."
$cs = Get-CimInstance Win32_ComputerSystem
$os = Get-CimInstance Win32_OperatingSystem

$diskTotalGB = 0
$diskFreeGB  = 0
Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
  if ($_.Size)      { $diskTotalGB += ($_.Size/1GB) }
  if ($_.FreeSpace) { $diskFreeGB  += ($_.FreeSpace/1GB) }
}

$primaryIP = $null
try {
  $cfg = Get-NetIPConfiguration -ErrorAction Stop |
    Where-Object { $_.IPv4Address -and $_.NetAdapter -and $_.NetAdapter.Status -eq "Up" } |
    Select-Object -First 1
  if ($cfg -and $cfg.IPv4Address -and $cfg.IPv4Address[0] -and $cfg.IPv4Address[0].IPAddress) {
    $primaryIP = $cfg.IPv4Address[0].IPAddress
  }
} catch { }

$cpuPct = $null
$availMemMB = $null
try { $cpuPct = [math]::Round((Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue, 1) } catch { }
try { $availMemMB = [math]::Round((Get-Counter '\Memory\Available MBytes').CounterSamples.CookedValue, 0) } catch { }

Log "Collecting observed listening ports..."
$observedPorts = @(Get-ObservedListeningPorts)

Log "Collecting installed apps..."
$apps = @(Get-InstalledApps)

Log "Checking IIS..."
$iisInfo = Get-IISInfo

$dashboard = @([pscustomobject]@{
  ComputerName = $env:COMPUTERNAME
  Domain       = $cs.Domain
  OS           = $os.Caption
  OSVersion    = $os.Version
  CPUCount     = $cs.NumberOfLogicalProcessors
  RAM_GB       = [math]::Round($cs.TotalPhysicalMemory/1GB, 1)
  DiskTotal_GB = [math]::Round($diskTotalGB, 1)
  DiskFree_GB  = [math]::Round($diskFreeGB, 1)
  PrimaryIPv4  = $primaryIP
  CPU_Pct      = $cpuPct
  AvailMem_MB  = $availMemMB
  IIS          = [bool]$iisInfo.Present
  CollectedAt  = (Get-Date).ToString("s")
})

# ---------------- Export (Excel if possible, else CSV) ----------------
$stamp  = Get-Date -Format "yyyyMMdd-HHmmss"
$outDir = Join-Path $OutputRoot ("SingleVM-MinDiscovery-{0}" -f $stamp)
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

if (Ensure-ImportExcel) {
  Log "ImportExcel found. Writing Excel..."
  Import-Module ImportExcel -ErrorAction Stop | Out-Null

  $xlsx = Join-Path $outDir "SingleVM-Minimal-Discovery.xlsx"
  if (Test-Path $xlsx) { Remove-Item $xlsx -Force }

  $dashboard     | Export-Excel -Path $xlsx -WorksheetName "00-Dashboard"              -AutoSize -FreezeTopRow -BoldTopRow
  $observedPorts | Export-Excel -Path $xlsx -WorksheetName "Observed_Listening_Ports" -AutoSize -FreezeTopRow -BoldTopRow -Append
  $apps          | Export-Excel -Path $xlsx -WorksheetName "Installed_Apps"           -AutoSize -FreezeTopRow -BoldTopRow -Append

  if ($iisInfo.Present -eq $true) {
    if (@($iisInfo.Sites).Count    -gt 0) { @($iisInfo.Sites)    | Export-Excel -Path $xlsx -WorksheetName "IIS_Sites"    -AutoSize -FreezeTopRow -BoldTopRow -Append }
    if (@($iisInfo.AppPools).Count -gt 0) { @($iisInfo.AppPools) | Export-Excel -Path $xlsx -WorksheetName "IIS_AppPools" -AutoSize -FreezeTopRow -BoldTopRow -Append }
    if (@($iisInfo.Apps).Count     -gt 0) { @($iisInfo.Apps)     | Export-Excel -Path $xlsx -WorksheetName "IIS_Apps"     -AutoSize -FreezeTopRow -BoldTopRow -Append }
    if (@($iisInfo.Bindings).Count -gt 0) { @($iisInfo.Bindings) | Export-Excel -Path $xlsx -WorksheetName "IIS_Bindings" -AutoSize -FreezeTopRow -BoldTopRow -Append }
  }

  Log "DONE."
  Log ("Output folder: {0}" -f $outDir)
  Log ("Excel file:    {0}" -f $xlsx)

} else {
  Log "ImportExcel not available. Writing CSV (single file)..."

  $csv = Join-Path $outDir "SingleVM-Minimal-Discovery.csv"
  if (Test-Path $csv) { Remove-Item $csv -Force }

  $rows = New-Object System.Collections.Generic.List[object]

  # Dashboard as KV rows (CSV-friendly)
  $dashObj = $dashboard | Select-Object -First 1
  foreach ($p in $dashObj.PSObject.Properties) {
    $rows.Add([pscustomobject]@{
      Section = "Dashboard"
      Name    = $p.Name
      Value   = [string]$p.Value
      Col1    = $null
      Col2    = $null
      Col3    = $null
      Col4    = $null
    }) | Out-Null
  }

  foreach ($p in $observedPorts) {
    $rows.Add([pscustomobject]@{
      Section = "Observed_Listening_Ports"
      Name    = $p.Protocol
      Value   = $p.LocalPort
      Col1    = $p.LocalAddress
      Col2    = $p.ProcessName
      Col3    = $p.OwningProcess
      Col4    = $p.Service
    }) | Out-Null
  }

  foreach ($a in $apps) {
    $rows.Add([pscustomobject]@{
      Section = "Installed_Apps"
      Name    = $a.DisplayName
      Value   = $a.DisplayVersion
      Col1    = $a.Publisher
      Col2    = $a.InstallDate
      Col3    = $null
      Col4    = $null
    }) | Out-Null
  }

  if ($iisInfo.Present -eq $true) {
    foreach ($s in @($iisInfo.Sites)) {
      $rows.Add([pscustomobject]@{
        Section = "IIS_Sites"
        Name    = $s.Name
        Value   = $s.State
        Col1    = $s.PhysicalPath
        Col2    = $s.ApplicationPool
        Col3    = $s.ID
        Col4    = $null
      }) | Out-Null
    }

    foreach ($pp in @($iisInfo.AppPools)) {
      $rows.Add([pscustomobject]@{
        Section = "IIS_AppPools"
        Name    = $pp.Name
        Value   = $pp.State
        Col1    = $pp.Runtime
        Col2    = $pp.PipelineMode
        Col3    = $pp.IdentityType
        Col4    = $null
      }) | Out-Null
    }

    foreach ($wa in @($iisInfo.Apps)) {
      $rows.Add([pscustomobject]@{
        Section = "IIS_Apps"
        Name    = $wa.Site
        Value   = $wa.Path
        Col1    = $wa.PhysicalPath
        Col2    = $wa.ApplicationPool
        Col3    = $null
        Col4    = $null
      }) | Out-Null
    }

    foreach ($b in @($iisInfo.Bindings)) {
      $rows.Add([pscustomobject]@{
        Section = "IIS_Bindings"
        Name    = $b.Site
        Value   = $b.Protocol
        Col1    = $b.IP
        Col2    = $b.Port
        Col3    = $b.HostHeader
        Col4    = $b.BindingInfo
      }) | Out-Null
    }
  }

  $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

  Log "DONE."
  Log ("Output folder: {0}" -f $outDir)
  Log ("CSV file:      {0}" -f $csv)
}

# ==================================================================================================
# ADD-ON (NEW FUNCTIONS ADDED): Scheduled Tasks + Windows Services inventory
#   - Add-only. Does not modify original flow.
#   - Excel: appends worksheets Scheduled_Tasks, Windows_Services
#   - CSV: creates Scheduled_Tasks.csv, Windows_Services.csv in $outDir
# ==================================================================================================

function Get-ScheduledTasksInventory {
  $scheduledTasks = @()

  $hasScheduledTasksModule = $false
  try { $hasScheduledTasksModule = [bool](Get-Module -ListAvailable -Name ScheduledTasks) } catch { $hasScheduledTasksModule = $false }

  if (-not $hasScheduledTasksModule) {
    return @([pscustomobject]@{ Note = "ScheduledTasks module not available; scheduled task enumeration skipped." })
  }

  try {
    Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
    $allTasks = @(Get-ScheduledTask -ErrorAction Stop)

    foreach ($t in $allTasks) {
      $info = $null
      try { $info = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop } catch { $info = $null }

      $actionText = $null
      try {
        $actionText = ($t.Actions | ForEach-Object {
          $exe = $_.Execute
          $arg = $_.Arguments
          $wd  = $_.WorkingDirectory
          ("Execute={0}; Arguments={1}; WorkingDir={2}" -f $exe, $arg, $wd)
        }) -join " | "
      } catch { $actionText = $null }

      $triggerText = $null
      try {
        $triggerText = ($t.Triggers | ForEach-Object {
          $start   = $_.StartBoundary
          $end     = $_.EndBoundary
          $enabled = $_.Enabled
          $type    = $_.CimClass.CimClassName
          ("Type={0}; Start={1}; End={2}; Enabled={3}" -f $type, $start, $end, $enabled)
        }) -join " | "
      } catch { $triggerText = $null }

      $principalUser = $null
      $principalType = $null
      $runLevel      = $null
      try {
        $principalUser = $t.Principal.UserId
        $principalType = $t.Principal.LogonType
        $runLevel      = $t.Principal.RunLevel
      } catch { }

      $scheduledTasks += [pscustomobject]@{
        TaskName           = $t.TaskName
        TaskPath           = $t.TaskPath
        State              = $t.State
        Author             = $t.Author
        Description        = $t.Description
        PrincipalUserId    = $principalUser
        PrincipalLogon     = $principalType
        RunLevel           = $runLevel
        Actions            = $actionText
        Triggers           = $triggerText
        LastRunTime        = if ($info) { $info.LastRunTime } else { $null }
        LastTaskResult     = if ($info) { $info.LastTaskResult } else { $null }
        NextRunTime        = if ($info) { $info.NextRunTime } else { $null }
        NumberOfMissedRuns = if ($info) { $info.NumberOfMissedRuns } else { $null }
      }
    }

    return $scheduledTasks
  }
  catch {
    return @([pscustomobject]@{ Note = ("Scheduled task enumeration failed: {0}" -f $_.Exception.Message) })
  }
}

function Get-WindowsServicesInventory {
  try {
    return @(Get-CimInstance Win32_Service -ErrorAction Stop | ForEach-Object {
      [pscustomobject]@{
        Name        = $_.Name
        DisplayName = $_.DisplayName
        State       = $_.State
        Status      = $_.Status
        StartMode   = $_.StartMode
        StartName   = $_.StartName
        ProcessId   = $_.ProcessId
        PathName    = $_.PathName
        ServiceType = $_.ServiceType
        Description = $_.Description
      }
    } | Sort-Object DisplayName)
  }
  catch {
    return @([pscustomobject]@{ Note = ("Windows services enumeration failed: {0}" -f $_.Exception.Message) })
  }
}

function Export-AddOnInventory {
  param(
    [object[]]$ScheduledTasks,
    [object[]]$WindowsServices
  )

  $excelExists = $false
  try { if ($xlsx -and (Test-Path $xlsx)) { $excelExists = $true } } catch { $excelExists = $false }

  if ($excelExists -and (Ensure-ImportExcel)) {
    Log "Appending Scheduled Tasks & Windows Services into existing Excel..."
    try {
      Import-Module ImportExcel -ErrorAction Stop | Out-Null
      $ScheduledTasks  | Export-Excel -Path $xlsx -WorksheetName "Scheduled_Tasks"  -AutoSize -FreezeTopRow -BoldTopRow -Append
      $WindowsServices | Export-Excel -Path $xlsx -WorksheetName "Windows_Services" -AutoSize -FreezeTopRow -BoldTopRow -Append
      Log "Excel updated with Scheduled Tasks & Windows Services."
    } catch {
      Log ("WARNING: Failed to append to Excel. Reason: {0}" -f $_.Exception.Message)
    }
  }
  else {
    try {
      if (-not $outDir) { $outDir = $OutputRoot }

      $taskCsv = Join-Path $outDir "Scheduled_Tasks.csv"
      $svcCsv  = Join-Path $outDir "Windows_Services.csv"

      $ScheduledTasks  | Export-Csv -Path $taskCsv -NoTypeInformation -Encoding UTF8
      $WindowsServices | Export-Csv -Path $svcCsv  -NoTypeInformation -Encoding UTF8

      Log ("Additional CSV created: {0}" -f $taskCsv)
      Log ("Additional CSV created: {0}" -f $svcCsv)
    } catch {
      Log ("WARNING: Failed to write Scheduled Tasks/Services CSV. Reason: {0}" -f $_.Exception.Message)
    }
  }
}

try {
  Log "Collecting Scheduled Tasks (Task Scheduler)..."
  $scheduledTasksInventory = @(Get-ScheduledTasksInventory)

  Log "Collecting Windows Services..."
  $windowsServicesInventory = @(Get-WindowsServicesInventory)

  Export-AddOnInventory -ScheduledTasks $scheduledTasksInventory -WindowsServices $windowsServicesInventory
}
catch {
  Log ("WARNING: Scheduler/Service inventory add-on failed. Reason: {0}" -f $_.Exception.Message)
}
# ==================================================================================================