# ============================================================
# 2_Update_DeviceOS.ps1
# Uses Excel COM automation - no ImportExcel module needed.
# Make sure device.xlsx is on your Desktop and CLOSED in Excel.
# ============================================================

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
