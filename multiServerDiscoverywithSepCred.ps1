#requires -Version 5.1
param(
  [string] $ServerCsvPath = ".\servers.csv",
  [string] $OutputRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Log { param([string]$m) Write-Host ("[{0}] {1}" -f (Get-Date -Format "HH:mm:ss"), $m) }
function Release-ComObjectSafe { param($obj) try { if ($null -ne $obj) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } } catch { } }

if (-not (Test-Path $ServerCsvPath)) { throw "servers.csv not found: $ServerCsvPath" }
$targets = Import-Csv $ServerCsvPath
if (@($targets).Count -eq 0) { throw "servers.csv is empty." }

# --- Remote payload (runs on each server) ---
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
    ObservedPorts = $ports
    Apps         = $apps
    IIS          = $iis
  }
}

# --- Excel COM writer (offline) ---
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
  param([string]$xlsxPath, $dash, $ports, $apps, $iisSites, $iisPools, $iisApps, $iisBindings, $failures)

  $excel = $null; $wb = $null
  try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $wb = $excel.Workbooks.Add()
    while ($wb.Worksheets.Count -gt 0) { $wb.Worksheets.Item(1).Delete() }

    Write-Worksheet $wb "00-Dashboard"             $dash
    Write-Worksheet $wb "Observed_Listening_Ports" $ports
    Write-Worksheet $wb "Installed_Apps"           $apps
    if (@($iisSites).Count    -gt 0) { Write-Worksheet $wb "IIS_Sites"    $iisSites }
    if (@($iisPools).Count    -gt 0) { Write-Worksheet $wb "IIS_AppPools" $iisPools }
    if (@($iisApps).Count     -gt 0) { Write-Worksheet $wb "IIS_Apps"     $iisApps }
    if (@($iisBindings).Count -gt 0) { Write-Worksheet $wb "IIS_Bindings" $iisBindings }
    if (@($failures).Count    -gt 0) { Write-Worksheet $wb "Failures"     $failures }

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

# --- Run across servers ---
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
$failAll = @()

foreach ($t in $targets) {
  $server = $t.Server
  $user   = $t.Username
  $pass   = $t.Password

  if (-not $server -or -not $user -or -not $pass) {
    $failAll += [pscustomobject]@{ Server=$server; Error="Missing Server/Username/Password in servers.csv row." }
    continue
  }

  Log ("Discovering: {0}" -f $server)

  try {
    $sec  = ConvertTo-SecureString $pass -AsPlainText -Force
    $cred = New-Object System.Management.Automation.PSCredential($user, $sec)

    $res = Invoke-Command -ComputerName $server -Credential $cred -ScriptBlock $remoteSb -ErrorAction Stop

    $sum = $res.Summary
    $dashAll += [pscustomobject]@{
      Server       = $server
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
        Server        = $server
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
        Server         = $server
        DisplayName    = $a.DisplayName
        DisplayVersion = $a.DisplayVersion
        Publisher      = $a.Publisher
        InstallDate    = $a.InstallDate
      }
    }

    if ($res.IIS -and $res.IIS.Present -eq $true) {
      foreach ($x in @($res.IIS.Sites))    { $iisSitesAll += ($x | Select-Object @{n="Server";e={$server}}, *) }
      foreach ($x in @($res.IIS.AppPools)) { $iisPoolsAll += ($x | Select-Object @{n="Server";e={$server}}, *) }
      foreach ($x in @($res.IIS.Apps))     { $iisAppsAll  += ($x | Select-Object @{n="Server";e={$server}}, *) }
      foreach ($x in @($res.IIS.Bindings)) { $iisBindAll  += ($x | Select-Object @{n="Server";e={$server}}, *) }
    }
  }
  catch {
    $failAll += [pscustomobject]@{ Server=$server; Error=$_.Exception.Message }
  }
}

Log "Writing Excel (COM). If COM fails, writing CSV..."
$excelOk = Try-WriteExcelCom -xlsxPath $xlsx `
  -dash $dashAll -ports $portsAll -apps $appsAll `
  -iisSites $iisSitesAll -iisPools $iisPoolsAll -iisApps $iisAppsAll -iisBindings $iisBindAll `
  -failures $failAll

if ($excelOk) {
  Log "DONE."
  Log ("Output folder: {0}" -f $outDir)
  Log ("Excel file:    {0}" -f $xlsx)
} else {
  Log "Excel COM failed. Falling back to single CSV..."

  $rows = New-Object System.Collections.Generic.List[object]

  foreach ($d in @($dashAll)) {
    foreach ($pp in $d.PSObject.Properties) {
      $rows.Add([pscustomobject]@{ Section="Dashboard"; Server=$d.Server; Name=$pp.Name; Value=[string]$pp.Value; Col1=$null; Col2=$null; Col3=$null; Col4=$null }) | Out-Null
    }
  }

  foreach ($p in @($portsAll)) {
    $rows.Add([pscustomobject]@{
      Section="Observed_Listening_Ports"; Server=$p.Server; Name=("{0}/{1}" -f $p.Protocol,$p.LocalPort); Value=$p.LocalAddress;
      Col1=$p.ProcessName; Col2=$p.Service; Col3=$p.OwningProcess; Col4=$null
    }) | Out-Null
  }

  foreach ($a in @($appsAll)) {
    $rows.Add([pscustomobject]@{
      Section="Installed_Apps"; Server=$a.Server; Name=$a.DisplayName; Value=$a.DisplayVersion;
      Col1=$a.Publisher; Col2=$a.InstallDate; Col3=$null; Col4=$null
    }) | Out-Null
  }

  foreach ($x in @($iisSitesAll)) { $rows.Add([pscustomobject]@{ Section="IIS_Sites"; Server=$x.Server; Name=$x.Name; Value=$x.State; Col1=$x.PhysicalPath; Col2=$x.ApplicationPool; Col3=$x.ID; Col4=$null }) | Out-Null }
  foreach ($x in @($iisPoolsAll)) { $rows.Add([pscustomobject]@{ Section="IIS_AppPools"; Server=$x.Server; Name=$x.Name; Value=$x.State; Col1=$x.Runtime; Col2=$x.PipelineMode; Col3=$x.IdentityType; Col4=$null }) | Out-Null }
  foreach ($x in @($iisAppsAll))  { $rows.Add([pscustomobject]@{ Section="IIS_Apps"; Server=$x.Server; Name=$x.Site; Value=$x.Path; Col1=$x.PhysicalPath; Col2=$x.ApplicationPool; Col3=$null; Col4=$null }) | Out-Null }
  foreach ($x in @($iisBindAll))  { $rows.Add([pscustomobject]@{ Section="IIS_Bindings"; Server=$x.Server; Name=$x.Site; Value=$x.Protocol; Col1=$x.IP; Col2=$x.Port; Col3=$x.HostHeader; Col4=$x.BindingInfo }) | Out-Null }

  foreach ($f in @($failAll))     { $rows.Add([pscustomobject]@{ Section="Failures"; Server=$f.Server; Name="Error"; Value=$f.Error; Col1=$null; Col2=$null; Col3=$null; Col4=$null }) | Out-Null }

  $rows | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

  Log "DONE."
  Log ("Output folder: {0}" -f $outDir)
  Log ("CSV file:      {0}" -f $csv)
}
