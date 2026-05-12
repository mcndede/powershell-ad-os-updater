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

In a large enterprise environment, keeping track of which devices are running Windows 10 vs Windows 11 across multiple AD domains is tedious. Doing it manually — searching AD one device at a time and updating a spreadsheet — is slow and error-prone.

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
.\2_Update_DeviceOS.ps1
```
The script will:
- Open the Excel file via COM automation
- Search each device name against both AD domains
- Write the OS result ("Windows 10", "Windows 11", or the raw OS string) into the Window column
- Save and close the file automatically
- Print a summary showing how many devices were updated vs. not found

<h2>Script Walkthrough</h2>

<p>
<img src="[INSERT SCREENSHOT — script output in terminal, devices being searched]" height="80%" width="80%" alt="Script running in PowerShell terminal"/>
</p>
<p>
When the script runs, it searches each device name against the first domain (corp.company.com) and falls back to the second (net.company.com) if not found. Each result is printed to the console in real time so you can follow along as it processes the list.
</p>
<br />

<p>
<img src="[INSERT SCREENSHOT — Excel file before and after, showing the Window column populated]" height="80%" width="80%" alt="Excel file with OS column filled in"/>
</p>
<p>
Once complete, the Excel file is saved with the OS column filled in for every device that was found in AD. Devices not found in either domain are flagged in the console output with a warning so nothing gets silently skipped.
</p>
<br />

<p>
<img src="[INSERT SCREENSHOT — summary output at the end of the script run]" height="80%" width="80%" alt="Script summary showing updated and not found counts"/>
</p>
<p>
At the end of each run, the script prints a summary showing the total number of devices updated and the number not found across either domain.
</p>
<br />

<h2>Errors & Troubleshooting</h2>

This script went through several iterations before working correctly. Below are the real errors encountered during development and how each one was resolved.

<p>
<img src="[INSERT SCREENSHOT — error 1]" height="80%" width="80%" alt="Error screenshot"/>
</p>
<p>
<strong>Error:</strong> [describe the error here — paste from the troubleshooting chat]<br/>
<strong>Cause:</strong> [what was causing it]<br/>
<strong>Fix:</strong> [what change resolved it]
</p>
<br />

<p>
<img src="[INSERT SCREENSHOT — error 2]" height="80%" width="80%" alt="Error screenshot"/>
</p>
<p>
<strong>Error:</strong> [describe the error here]<br/>
<strong>Cause:</strong> [what was causing it]<br/>
<strong>Fix:</strong> [what change resolved it]
</p>
<br />

<h2>Key Takeaways</h2>

- Excel COM automation is a powerful way to read/write .xlsx files in PowerShell without installing any extra modules — useful in locked-down enterprise environments
- Querying AD across multiple domains requires specifying the `-Server` parameter on `Get-ADComputer` for each domain separately
- Catching `ADIdentityNotFoundException` specifically (rather than a generic catch) allows the script to gracefully try the next domain instead of stopping on every miss
