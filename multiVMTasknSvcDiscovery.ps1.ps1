#requires -Version 5.1
param(
  [string] $ServerListPath = ".\servers.txt",
  [string] $OutputRoot = ".",
  [string] $Username = "",
  [string] $Password = ""  #If interactive, leave blank and you will be prompted.
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Log { param([string]$m) Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) }
function Release-ComObjectSafe { param($obj) try { if ($null -ne $obj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } } catch { } }
function ToArray { param($x) @($x) }

if (-not (Test-Path $ServerListPath)) { throw "servers.txt not found: $ServerListPath" }
$servers = Get-Content $ServerListPath | ForEach-Object { $_.Trim() } | Where-Object { $_ -and -not $_.StartsWith("#") } | Select-Object -Unique
if (@($servers).Count -eq 0) { throw "servers.txt is empty." }

# Build credential
$cred = $null
if ($Username -and $Password) {
  $sec = ConvertTo-SecureString $Password -AsPlainText -Force
  $cred = New-Object System.Management.Automation.PSCredential($Username, $sec)
} else {
  $cred = Get-Credential -Message "Enter credentials that work on all target servers"
}

# ---------- Remote collectors (runs on each target) ----------
$remoteSb = {
  Set-StrictMode -Version Latest
  $ErrorActionPreference = "Stop"

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
    $apps | Sort-Object DisplayName -Unique
  }

  function Get-ObservedListeningPorts {
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

    $ports | Sort-Object Protocol, LocalPort, ProcessName
  }

  function Get-IISInfo {
    $iisPresent = $false
    try { if (Get-Service -Name W3SVC -ErrorAction SilentlyContinue) { $iisPresent = $true } } catch { }

    if (-not $iisPresent) { return [pscustomobject]@{ Present=$false; Sites=@(); AppPools=@(); Apps=@(); Bindings=@() } }
    if (-not (Get-Module -ListAvailable -Name WebAdministration)) { return [pscustomobject]@{ Present=$true; Sites=@(); AppPools=@(); Apps=@(); Bindings=@() } }

    $sites=@(); $pools=@(); $apps=@(); $bindings=@()
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
            $ipPart=$null; $portPart=$null; $hostPart=$null
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
    } catch { }

    [pscustomobject]@{ Present=$true; Sites=@($sites); AppPools=@($pools); Apps=@($apps); Bindings=@($bindings) }
  }

  # ===================== NEW LOGIC: Task Scheduler =====================
  function Get-ScheduledTasksInfo {
    $tasksOut = @()

    # Prefer ScheduledTasks module (rich objects)
    if (Get-Module -ListAvailable -Name ScheduledTasks) {
      try {
        Import-Module ScheduledTasks -ErrorAction Stop | Out-Null
        $tasks = @(Get-ScheduledTask -ErrorAction Stop)

        foreach ($t in $tasks) {
          $info = $null
          try { $info = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction Stop } catch { $info = $null }

          $actions = $null
          try {
            $actions = ($t.Actions | ForEach-Object {
              ("Execute={0}; Arguments={1}; WorkingDir={2}" -f $_.Execute, $_.Arguments, $_.WorkingDirectory)
            }) -join " | "
          } catch { $actions = $null }

          $triggers = $null
          try {
            $triggers = ($t.Triggers | ForEach-Object {
              $type = $_.CimClass.CimClassName
              ("Type={0}; Start={1}; End={2}; Enabled={3}" -f $type, $_.StartBoundary, $_.EndBoundary, $_.Enabled)
            }) -join " | "
          } catch { $triggers = $null }

          $principalUser = $null; $principalLogon = $null; $runLevel = $null
          try {
            $principalUser  = $t.Principal.UserId
            $principalLogon = $t.Principal.LogonType
            $runLevel       = $t.Principal.RunLevel
          } catch { }

          $tasksOut += [pscustomobject]@{
            TaskName           = $t.TaskName
            TaskPath           = $t.TaskPath
            State              = $t.State
            Author             = $t.Author
            Description        = $t.Description
            PrincipalUserId    = $principalUser
            PrincipalLogon     = $principalLogon
            RunLevel           = $runLevel
            Actions            = $actions
            Triggers           = $triggers
            LastRunTime        = if ($info) { $info.LastRunTime } else { $null }
            LastTaskResult     = if ($info) { $info.LastTaskResult } else { $null }
            NextRunTime        = if ($info) { $info.NextRunTime } else { $null }
            NumberOfMissedRuns = if ($info) { $info.NumberOfMissedRuns } else { $null }
          }
        }

        return $tasksOut
      } catch {
        return @([pscustomobject]@{ Note = ("ScheduledTasks module enumeration failed: {0}" -f $_.Exception.Message) })
      }
    }

    # Fallback: schtasks.exe (works even without module)
    try {
      $raw = & schtasks.exe /Query /V /FO CSV 2>$null
      if (-not $raw) { return @([pscustomobject]@{ Note = "schtasks.exe returned no data." }) }

      $csv = $raw | ConvertFrom-Csv
      return @($csv | ForEach-Object {
        [pscustomobject]@{
          TaskName       = $_."TaskName"
          Status         = $_."Status"
          NextRunTime    = $_."Next Run Time"
          LastRunTime    = $_."Last Run Time"
          LastTaskResult = $_."Last Result"
          Author         = $_."Author"
          RunAsUser      = $_."Run As User"
          ScheduleType   = $_."Schedule Type"
          StartIn        = $_."Start In"
          TaskToRun      = $_."Task To Run"
          Comment        = $_."Comment"
        }
      })
    } catch {
      return @([pscustomobject]@{ Note = ("schtasks.exe fallback failed: {0}" -f $_.Exception.Message) })
    }
  }

  # ===================== NEW LOGIC: Windows Services =====================
  function Get-WindowsServicesInfo {
    try {
      @(Get-CimInstance Win32_Service -ErrorAction Stop | ForEach-Object {
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
    } catch {
      @([pscustomobject]@{ Note = ("Service enumeration failed: {0}" -f $_.Exception.Message) })
    }
  }

  # Summary
  $cs = Get-CimInstance Win32_ComputerSystem
  $os = Get-CimInstance Win32_OperatingSystem

  $diskTotalGB = 0; $diskFreeGB = 0
  Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3" | ForEach-Object {
    if ($_.Size)      { $diskTotalGB += ($_.Size/1GB) }
    if ($_.FreeSpace) { $diskFreeGB  += ($_.FreeSpace/1GB) }
  }

  $primaryIP = $null
  try {
    $cfg = Get-NetIPConfiguration -ErrorAction Stop |
      Where-Object { $_.IPv4Address -and $_.NetAdapter -and $_.NetAdapter.Status -eq "Up" } |
      Select-Object -First 1
    if ($cfg -and $cfg.IPv4Address -and $cfg.IPv4Address[0] -and $cfg.IPv4Address[0].IPAddress) { $primaryIP = $cfg.IPv4Address[0].IPAddress }
  } catch { }

  $cpuPct = $null; $availMemMB = $null
  try { $cpuPct = [math]::Round((Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples.CookedValue, 1) } catch { }
  try { $availMemMB = [math]::Round((Get-Counter '\Memory\Available MBytes').CounterSamples.CookedValue, 0) } catch { }

  $ports = @(Get-ObservedListeningPorts)
  $apps  = @(Get-InstalledApps)
  $iis   = Get-IISInfo

  # NEW: fetch sched tasks + services on the remote server
  $schedTasks = @(Get-ScheduledTasksInfo)
  $services   = @(Get-WindowsServicesInfo)

  [pscustomobject]@{
    Summary = [pscustomobject]@{
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
      IIS          = [bool]$iis.Present
      CollectedAt  = (Get-Date).ToString("s")
    }
    ObservedPorts   = $ports
    Apps            = $apps
    IIS             = $iis
    ScheduledTasks  = $schedTasks
    Services        = $services
  }
}

# ---------- Excel COM writer ----------
function Write-Worksheet {
  param($workbook, [string]$name, $data)

  $arr = @($data)
  $ws = $workbook.Worksheets.Add()
  $ws.Name = $name

  if ($arr.Count -eq 0) { $ws.Cells.Item(1,1).Value2 = "No data"; return }

  $props = @($arr[0].PSObject.Properties | ForEach-Object { $_.Name })
  for ($c=0; $c -lt $props.Count; $c++) {
    $ws.Cells.Item(1, $c+1).Value2 = $props[$c]
    $ws.Cells.Item(1, $c+1).Font.Bold = $true
  }

  for ($r=0; $r -lt $arr.Count; $r++) {
    for ($c=0; $c -lt $props.Count; $c++) {
      $val = $arr[$r].PSObject.Properties[$props[$c]].Value
      $ws.Cells.Item($r+2, $c+1).Value2 = if ($null -ne $val) { [string]$val } else { "" }
    }
  }

  $ws.Columns.AutoFit() | Out-Null
  $ws.Application.ActiveWindow.SplitRow = 1
  $ws.Application.ActiveWindow.FreezePanes = $true
}

function Try-WriteExcelCom {
  param(
    [string]$xlsxPath,
    $dash, $ports, $apps,
    $iisSites, $iisPools, $iisApps, $iisBindings,
    $schedTasks, $services
  )

  $excel = $null; $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -gt 0) { $wb.Worksheets.Item(1).Delete() }

    Write-Worksheet $wb "00-Dashboard"              $dash
    Write-Worksheet $wb "Observed_Listening_Ports"  $ports
    Write-Worksheet $wb "Installed_Apps"            $apps
    if (@($iisSites).Count    -gt 0) { Write-Worksheet $wb "IIS_Sites"    $iisSites }
    if (@($iisPools).Count    -gt 0) { Write-Worksheet $wb "IIS_AppPools" $iisPools }
    if (@($iisApps).Count     -gt 0) { Write-Worksheet $wb "IIS_Apps"     $iisApps }
    if (@($iisBindings).Count -gt 0) { Write-Worksheet $wb "IIS_Bindings" $iisBindings }

    # NEW sheets
    if (@($schedTasks).Count -gt 0) { Write-Worksheet $wb "Scheduled_Tasks"  $schedTasks }
    if (@($services).Count   -gt 0) { Write-Worksheet $wb "Windows_Services" $services }

    $wb.SaveAs($xlsxPath)
    $wb.Close($true)
    $excel.Quit()
    return $true
  } catch {
    return $false
  } finally {
    Release-ComObjectSafe $wb
    Release-ComObjectSafe $excel
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  }
}

# ---------- Run across servers ----------
$stamp  = Get-Date -Format "yyyyMMdd-HHmmss"
$outDir = Join-Path $OutputRoot ("MultiServer-Discovery-{0}" -f $stamp)
New-Item -ItemType Directory -Path $outDir -Force | Out-Null
$xlsx = Join-Path $outDir "Discovery-AllServers.xlsx"
$csv  = Join-Path $outDir "Discovery-AllServers.csv"

$dashAll = @()
$portsAll = @()
$appsAll = @()
$iisSitesAll = @()
$iisPoolsAll = @()
$iisAppsAll  = @()
$iisBindAll  = @()
$scheduledTasksAll = @()
$servicesAll = @()
$failAll = @()

foreach ($s in $servers) {
  Log ("Discovering: {0}" -f $s)
  try {
    $res = Invoke-Command -ComputerName $s -Credential $cred -ScriptBlock $remoteSb -ErrorAction Stop

    $sum = $res.Summary
    $dashAll += [pscustomobject]@{
      Server       = $s
      ComputerName = $sum.ComputerName
      Domain       = $sum.Domain
      OS           = $sum.OS
      OSVersion    = $sum.OSVersion
      CPUCount     = $sum.CPUCount
      RAM_GB       = $sum.RAM_GB
      DiskTotal_GB = $sum.DiskTotal_GB
      DiskFree_GB  = $sum.DiskFree_GB
      PrimaryIPv4  = $sum.PrimaryIPv4
      CPU_Pct      = $sum.CPU_Pct
      AvailMem_MB  = $sum.AvailMem_MB
      IIS          = $sum.IIS
      CollectedAt  = $sum.CollectedAt
    }

    foreach ($p in @($res.ObservedPorts)) {
      $portsAll += [pscustomobject]@{
        Server        = $s
        Protocol      = $p.Protocol
        LocalAddress  = $p.LocalAddress
        LocalPort     = $p.LocalPort
        ProcessName   = $p.ProcessName
        OwningProcess = $p.OwningProcess
        Service       = $p.Service
      }
    }

    foreach ($a in @($res.Apps)) {
      $appsAll += [pscustomobject]@{
        Server         = $s
        DisplayName    = $a.DisplayName
        DisplayVersion = $a.DisplayVersion
        Publisher      = $a.Publisher
        InstallDate    = $a.InstallDate
      }
    }

    if ($res.IIS -and $res.IIS.Present -eq $true) {
      foreach ($x in @($res.IIS.Sites))    { $iisSitesAll += ($x | Select-Object @{n="Server";e={$s}}, *) }
      foreach ($x in @($res.IIS.AppPools)) { $iisPoolsAll += ($x | Select-Object @{n="Server";e={$s}}, *) }
      foreach ($x in @($res.IIS.Apps))     { $iisAppsAll  += ($x | Select-Object @{n="Server";e={$s}}, *) }
      foreach ($x in @($res.IIS.Bindings)) { $iisBindAll  += ($x | Select-Object @{n="Server";e={$s}}, *) }
    }

    # NEW: aggregate scheduled tasks + services per server
    foreach ($t in @($res.ScheduledTasks)) { $scheduledTasksAll += ($t | Select-Object @{n="Server";e={$s}}, *) }
    foreach ($sv in @($res.Services))      { $servicesAll       += ($sv | Select-Object @{n="Server";e={$s}}, *) }

  } catch {
    $failAll += [pscustomobject]@{ Server=$s; Error=$_.Exception.Message }
  }
}

# ---------- Write output ----------
Log "Writing Excel (COM). If COM fails, writing single CSV..."
$excelOk = Try-WriteExcelCom -xlsxPath $xlsx `
  -dash $dashAll -ports $portsAll -apps $appsAll `
  -iisSites $iisSitesAll -iisPools $iisPoolsAll -iisApps $iisAppsAll -iisBindings $iisBindAll `
  -schedTasks $scheduledTasksAll -services $servicesAll

if ($excelOk) {
  if (@($failAll).Count -gt 0) {
    # Add failures in CSV next to xlsx for convenience
    $failAll | Export-Csv -Path (Join-Path $outDir "Failures.csv") -NoTypeInformation -Encoding UTF8
  }
  Log "DONE."
  Log ("Output folder: {0}" -f $outDir)
  Log ("Excel file:    {0}" -f $xlsx)
} else {
  # Single CSV with sections
  $rows = New-Object System.Collections.Generic.List[object]

  foreach ($d in @($dashAll)) {
    foreach ($pp in $d.PSObject.Properties) {
      $rows.Add([pscustomobject]@{ Section="Dashboard"; Server=$d.Server; Name=$pp.Name; Value=[string]$pp.Value; Col1=$null; Col2=$null; Col3=$null; Col4=$null }) | Out-Null
    }
  }
  foreach ($p in @($portsAll)) { $rows.Add([pscustomobject]@{ Section="Observed_Listening_Ports"; Server=$p.Server; Name=("{0}/{1}" -f $p.Protocol,$p.LocalPort); Value=$p.LocalAddress; Col1=$p.ProcessName; Col2=$p.Service; Col3=$p.OwningProcess; Col4=$null }) | Out-Null }
  foreach ($a in @($appsAll))  { $rows.Add([pscustomobject]@{ Section="Installed_Apps"; Server=$a.Server; Name=$a.DisplayName; Value=$a.DisplayVersion; Col1=$a.Publisher; Col2=$a.InstallDate; Col3=$null; Col4=$null }) | Out-Null }

  foreach ($x in @($iisSitesAll)) { $rows.Add([pscustomobject]@{ Section="IIS_Sites"; Server=$x.Server; Name=$x.Name; Value=$x.State; Col1=$x.PhysicalPath; Col2=$x.ApplicationPool; Col3=$x.ID; Col4=$null }) | Out-Null }
  foreach ($x in @($iisPoolsAll)) { $rows.Add([pscustomobject]@{ Section="IIS_AppPools"; Server=$x.Server; Name=$x.Name; Value=$x.State; Col1=$x.Runtime; Col2=$x.PipelineMode; Col3=$x.IdentityType; Col4=$null }) | Out-Null }
  foreach ($x in @($iisAppsAll))  { $rows.Add([pscustomobject]@{ Section="IIS_Apps"; Server=$x.Server; Name=$x.Site; Value=$x.Path; Col1=$x.PhysicalPath; Col2=$x.ApplicationPool; Col3=$null; Col4=$null }) | Out-Null }
  foreach ($x in @($iisBindAll))  { $rows.Add([pscustomobject]@{ Section="IIS_Bindings"; Server=$x.Server; Name=$x.Site; Value=$x.Protocol; Col1=$x.IP; Col2=$x.Port; Col3=$x.HostHeader; Col4=$x.BindingInfo }) | Out-Null }

  # NEW: Scheduled Tasks section
  foreach ($t in @($scheduledTasksAll)) {
    $taskNameProp = $t.PSObject.Properties['TaskName']
    $stateProp    = $t.PSObject.Properties['State']
    $pathProp     = $t.PSObject.Properties['TaskPath']
    $userProp     = $t.PSObject.Properties['PrincipalUserId']
    $nextProp     = $t.PSObject.Properties['NextRunTime']
    $actProp      = $t.PSObject.Properties['Actions']

    $rows.Add([pscustomobject]@{
      Section="Scheduled_Tasks"; Server=$t.Server;
      Name = if ($taskNameProp) { [string]$taskNameProp.Value } else { "" };
      Value= if ($stateProp)    { [string]$stateProp.Value }    else { "" };
      Col1 = if ($pathProp)     { [string]$pathProp.Value }     else { "" };
      Col2 = if ($userProp)     { [string]$userProp.Value }     else { "" };
      Col3 = if ($nextProp)     { [string]$nextProp.Value }     else { "" };
      Col4 = if ($actProp)      { [string]$actProp.Value }      else { "" };
    }) | Out-Null
  }

  # NEW: Windows Services section
  foreach ($sv in @($servicesAll)) {
    $rows.Add([pscustomobject]@{
      Section="Windows_Services"; Server=$sv.Server;
      Name=$sv.DisplayName; Value=$sv.State;
      Col1=$sv.Name; Col2=$sv.StartMode; Col3=$sv.StartName; Col4=$sv.PathName
    }) | Out-Null
  }

  foreach ($f in @($failAll)) {
    $rows.Add([pscustomobject]@{ Section="Failures"; Server=$f.Server; Name="Error"; Value=$f.Error; Col1=$null; Col2=$null; Col3=$null; Col4=$null }) | Out-Null
  }

  $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

  Log "DONE."
  Log ("Output folder: {0}" -f $outDir)
  Log ("CSV file:      {0}" -f $csv)
}
