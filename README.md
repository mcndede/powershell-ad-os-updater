<h1>PowerShell AD OS Updater</h1>
A two-script PowerShell automation tool that queries Active Directory across multiple domains and bulk-updates an Excel spreadsheet with the current Operating System for each device — no manual lookups required.

Built to solve a real-world Desktop Support problem: updating large device lists with accurate OS data without touching each machine individually.

<h2>Environments and Technologies Used</h2>

- Windows PowerShell 5.1
- Active Directory (RSAT / ActiveDirectory module)
- Excel COM Automation (no third-party modules required)
- Microsoft Active Directory (multi-domain environment)

<h2>Operating Systems Used</h2>

- Windows 10 / Windows 11 (domain-joined endpoints)
- Windows Server (domain controllers — corp.company.com / net.company.com)

<h2>Problem This Solves</h2>

In a large enterprise environment, keeping track of which devices are running Windows 10 vs Windows 11 across multiple AD domains is tedious. Doing it manually, searching AD one device at a time and updating a spreadsheet, is slow and error-prone.

This tool automates the entire process: give it a spreadsheet of device names, and it queries AD across both domains, then writes the OS directly back into the file.

<h2>Scripts Overview</h2>

- **1_Install_Modules.ps1** — Prerequisite checker. Verifies that the ActiveDirectory (RSAT) module is installed and that Excel COM automation is available before running the main script.
- **2_Update_DeviceOS.ps1** — The main script. Opens `device.xlsx` from the Desktop, loops through every device name, queries both AD domains for the OperatingSystem attribute, and writes "Windows 10" or "Windows 11" back into the spreadsheet automatically.

<h2>How to Use</h2>

**Step 1 — Run the prerequisite check**
```powershell
# Check ActiveDirectory module (RSAT)
if (Get-Module -ListAvailable -Name ActiveDirectory) {
    Write-Host "ActiveDirectory module: OK" -ForegroundColor Green
} else {
    Write-Host "WARNING: ActiveDirectory module not found." -ForegroundColor Red
    Write-Host "Ask your IT admin to install RSAT on this machine." -ForegroundColor Yellow
}

# Check Excel is installed (needed for COM)
try {
    $excel = New-Object -ComObject Excel.Application -ErrorAction Stop
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Excel COM automation: OK" -ForegroundColor Green
} catch {
    Write-Host "WARNING: Excel not found or COM automation blocked." -ForegroundColor Red
}

Write-Host "`nSetup check complete! You can now run 2_Update_DeviceOS.ps1 anytime." -ForegroundColor Green
Read-Host "Press Enter to close"
```
This confirms that RSAT and Excel COM are available on your machine. If either is missing, it will tell you what to fix before proceeding.

**Step 2 — Prepare your spreadsheet**

Place `device.xlsx` on your Desktop. The sheet must have a column called **Device Name** and a column starting with **Window** (e.g. "Windows Version"). Make sure the file is closed in Excel before running the script.

**Step 3 — Run the main script**
```powershell

Import-Module ActiveDirectory -ErrorAction Stop

$domains = @("corp.company.com", "net.company.com")

$filePath = Join-Path ([Environment]::GetFolderPath('Desktop')) "device.xlsx"
if (-not (Test-Path $filePath)) {
    Write-Host "ERROR: device.xlsx not found on your Desktop." -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit
}

# Open Excel via COM
Write-Host "Opening $filePath via Excel..." -ForegroundColor Cyan
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($filePath)
    $sheet = $workbook.Sheets.Item(1)
} catch {
    Write-Host "ERROR: Could not open Excel file. Is it already open?" -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit
}

# Find header row - locate 'Device Name' and 'Window' columns
$headerRow = 1
$deviceNameCol = $null
$windowCol     = $null

$lastCol = $sheet.UsedRange.Columns.Count
for ($col = 1; $col -le $lastCol; $col++) {
    $header = $sheet.Cells.Item($headerRow, $col).Text.Trim()
    if ($header -eq "Device Name") { $deviceNameCol = $col }
    if ($header -like "Window*")   { $windowCol     = $col }
}

if (-not $deviceNameCol -or -not $windowCol) {
    Write-Host "ERROR: Could not find 'Device Name' or 'Window' column in sheet." -ForegroundColor Red
    $workbook.Close($false)
    $excel.Quit()
    Read-Host "Press Enter to close"
    exit
}

Write-Host "Columns found - Device Name: $deviceNameCol | Window: $windowCol" -ForegroundColor DarkGray
Write-Host "`nSearching devices...`n" -ForegroundColor Yellow

$updated  = 0
$notFound = 0
$lastRow  = $sheet.UsedRange.Rows.Count

for ($row = 2; $row -le $lastRow; $row++) {
    $deviceName = $sheet.Cells.Item($row, $deviceNameCol).Text.Trim()
    if (-not $deviceName) { continue }

    Write-Host "Searching: $deviceName" -ForegroundColor Magenta
    $found = $false

    foreach ($domain in $domains) {
        try {
            $computer = Get-ADComputer -Identity $deviceName `
                            -Server $domain `
                            -Properties OperatingSystem `
                            -ErrorAction Stop

            if ($computer) {
                $os = $computer.OperatingSystem

                if ($os -like "*Windows 11*") {
                    $sheet.Cells.Item($row, $windowCol) = "Windows 11"
                    Write-Host "  -> Found in $domain : Windows 11" -ForegroundColor Green
                } elseif ($os -like "*Windows 10*") {
                    $sheet.Cells.Item($row, $windowCol) = "Windows 10"
                    Write-Host "  -> Found in $domain : Windows 10" -ForegroundColor Green
                } else {
                    $sheet.Cells.Item($row, $windowCol) = $os
                    Write-Host "  -> Found in $domain : $os" -ForegroundColor Yellow
                }

                $found = $true
                $updated++
                break
            }
        }
        catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
            Write-Host "  -> Not found in $domain, trying next..." -ForegroundColor DarkGray
        }
        catch {
            Write-Warning "  -> Error searching $domain : $($_.Exception.Message)"
        }
    }

    if (-not $found) {
        Write-Warning "  -> '$deviceName' not found in any domain."
        $notFound++
    }
}

# Save and close
Write-Host "`nSaving Excel file..." -ForegroundColor Cyan
try {
    $workbook.Save()
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    Write-Host "Saved successfully!" -ForegroundColor Green
} catch {
    Write-Host "ERROR: Could not save file." -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit
}

Write-Host "`n============================" -ForegroundColor Cyan
Write-Host "  Updated  : $updated devices" -ForegroundColor Green
Write-Host "  Not found: $notFound devices" -ForegroundColor Yellow
Write-Host "============================`n" -ForegroundColor Cyan
Read-Host "Press Enter to close"
```
The script will:
- Open the Excel file via COM automation
- Search each device name against both AD domains
- Write the OS result ("Windows 10", "Windows 11", or the raw OS string) into the Window column
- Save and close the file automatically
- Print a summary showing how many devices were updated vs. not found

<h2>Script Walkthrough</h2>

<p>
<img src="https://us-east.storage.cloudconvert.com/tasks/6ca492b5-e3a2-4f9b-8371-5069cededee5/Searching.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Content-Sha256=UNSIGNED-PAYLOAD&X-Amz-Credential=cloudconvert-production%2F20260528%2Fva%2Fs3%2Faws4_request&X-Amz-Date=20260528T234717Z&X-Amz-Expires=86400&X-Amz-Signature=a45d39085d02ee05c5da2c602c95f098829712f139522fffa2924cce7a6a90f0&X-Amz-SignedHeaders=host&response-content-disposition=inline%3B%20filename%3D%22Searching.png%22&response-content-type=image%2Fpng&x-id=GetObject" height="80%" width="80%" alt="Script running in PowerShell terminal"/>
</p>
<p>
When the script runs, it searches each device name against the first domain (corp.company.com) and falls back to the second (net.company.com) if not found. Each result is printed to the console in real time so you can follow along as it processes the list.
</p>
<br />

<p>
<img src="https://us-east.storage.cloudconvert.com/tasks/80ca12a8-2e1c-472a-9cf7-3cca1e467afe/Excel%20Sheet.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Content-Sha256=UNSIGNED-PAYLOAD&X-Amz-Credential=cloudconvert-production%2F20260529%2Fva%2Fs3%2Faws4_request&X-Amz-Date=20260529T030127Z&X-Amz-Expires=86400&X-Amz-Signature=6df4ac9c685a1ca7af2bfaf5626043557a074f99f9a440d13680523863309e27&X-Amz-SignedHeaders=host&response-content-disposition=inline%3B%20filename%3D%22Excel%20Sheet.png%22&response-content-type=image%2Fpng&x-id=GetObject" height="80%" width="80%" alt="Excel file with OS column filled in"/>
</p>
<p>
Once complete, the Excel file is saved with the OS column filled in for every device that was found in AD. Devices not found in either domain are flagged in the console output with a warning so nothing gets silently skipped.
</p>
<br />

<p>
<img src="https://us-east.storage.cloudconvert.com/tasks/2354629c-d87e-491e-aa10-9d9d1c56a806/Summary.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Content-Sha256=UNSIGNED-PAYLOAD&X-Amz-Credential=cloudconvert-production%2F20260529%2Fva%2Fs3%2Faws4_request&X-Amz-Date=20260529T030234Z&X-Amz-Expires=86400&X-Amz-Signature=38acc303417887673dcf519ae536f1f9cf00af1fc71ecebfac8c86b25cd08ab8&X-Amz-SignedHeaders=host&response-content-disposition=inline%3B%20filename%3D%22Summary.png%22&response-content-type=image%2Fpng&x-id=GetObject" height="80%" width="80%" alt="Script summary showing updated and not found counts"/>
</p>
<p>
At the end of each run, the script prints a summary showing the total number of devices updated and the number not found across either domain.
</p>
<br />

<h2>Errors & Troubleshooting</h2>

This script went through a few iterations before working correctly. Below are the real errors encountered during development and how each one was resolved.

<p>
<img src="https://us-east.storage.cloudconvert.com/tasks/0633fc27-a76d-466b-8fe2-b247326d8998/Error%20Message.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Content-Sha256=UNSIGNED-PAYLOAD&X-Amz-Credential=cloudconvert-production%2F20260529%2Fva%2Fs3%2Faws4_request&X-Amz-Date=20260529T045947Z&X-Amz-Expires=86400&X-Amz-Signature=ce374131eee25307b9a33e3ba892459a34d752e0c9702f7962b8e4004d2cb9fb&X-Amz-SignedHeaders=host&response-content-disposition=inline%3B%20filename%3D%22Error%20Message.png%22&response-content-type=image%2Fpng&x-id=GetObject" height="80%" width="80%" alt="Error screenshot"/>
</p>
<p>
<strong>Cause:</strong> The script failed at launch because the ImportExcel PowerShell module was not installed on the target machine. Since the script depends on this module to read/write Excel data, it cannot continue without it. <br/>
<strong>Fix:</strong> The script was rewritten to use Excel COM Automation instead, which is built into Windows and requires no module installation.
</p>
<br />

<p>
<img src="[[INSERT SCREENSHOT — error 2]](https://us-east.storage.cloudconvert.com/tasks/71878be7-86f4-44eb-8f01-dbdb6bae99bb/Error%202%20from%20module%201.png?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Content-Sha256=UNSIGNED-PAYLOAD&X-Amz-Credential=cloudconvert-production%2F20260529%2Fva%2Fs3%2Faws4_request&X-Amz-Date=20260529T051312Z&X-Amz-Expires=86400&X-Amz-Signature=e5f2b2b908d48c2e983777cfdd97fb74724b1bafb46ede3d2718036e3cd8abad&X-Amz-SignedHeaders=host&response-content-disposition=inline%3B%20filename%3D%22Error%202%20from%20module%201.png%22&response-content-type=image%2Fpng&x-id=GetObject)" height="80%" width="80%" alt="Error screenshot"/>
</p>
<p>
<strong>Cause:</strong> The script crashed when it hit the "press any key to continue" pause at line 19. The method it used to wait for a keypress is not supported in every PowerShell environment.<br/>
<strong>Fix:</strong> The original script used $Host.UI.RawUI.ReadKey to pause at the end, which crashed in certain PowerShell environments. The updated script replaces it with Read-Host, which works universally.
</p>
<br />

<h2>Key Takeaways</h2>

- Excel COM automation is a powerful way to read/write .xlsx files in PowerShell without installing any extra modules — useful in locked-down enterprise environments
- Querying AD across multiple domains requires specifying the `-Server` parameter on `Get-ADComputer` for each domain separately
- Catching `ADIdentityNotFoundException` specifically (rather than a generic catch) allows the script to gracefully try the next domain instead of stopping on every miss
