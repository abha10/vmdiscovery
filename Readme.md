# VM Discovery Script (Read-Only)

## What it does
Collects basic discovery information from a single Windows VM to help understand active services and dependencies.

The script is read-only and does not make any system changes.

---

## What it collects
- System summary  
  OS, CPU, memory, disks, primary IP
- Observed listening ports  
  TCP/UDP ports currently in use
- Installed software
- IIS details (if present)  
  Sites, application pools, apps, bindings

---

## What it does NOT do
- No configuration changes  
- No service restarts  
- No firewall or registry changes  
- No data sent outside the VM  

---

## Output
- Attempts to create an Excel (.xlsx) report  
  (`ImportExcel` module is installed automatically if required)
- Falls back to CSV automatically if Excel is not available
- Output is saved locally in a timestamped folder

---

## Prerequisites
- Windows VM  
- PowerShell 5.1  
- Local administrator access (recommended)

---

## How to run
```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\vmdiscovery\script.ps1
```In powershell
> Unblock-File .\vmdiscovery\script.ps1
> .\vmdiscovery\script.ps1

