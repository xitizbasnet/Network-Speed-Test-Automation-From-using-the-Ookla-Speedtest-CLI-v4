# =====================================================================
# Network Speed Test Automation
# Author      : xitiz
# Description : Runs Ookla Speedtest CLI and exports the results
#               directly to the current user's Desktop in Excel format.
# =====================================================================

# ----------------------------
# Import Required Module
# ----------------------------
Import-Module ImportExcel -ErrorAction Stop

# ----------------------------
# Configuration
# ----------------------------

# Update this path if required
$SpeedtestCLI = "C:\Tools\speedtest\speedtest.exe"

# Verify Speedtest CLI exists
if (!(Test-Path $SpeedtestCLI)) {
    Write-Host ""
    Write-Host "ERROR: Speedtest CLI not found."
    Write-Host "Expected Path: $SpeedtestCLI"
    Read-Host "Press Enter to exit"
    exit
}

# Desktop Path
$Desktop = [Environment]::GetFolderPath("Desktop")

# ----------------------------
# Run Speed Test
# ----------------------------

Write-Host ""
Write-Host "==============================================="
Write-Host "Running Internet Speed Test..."
Write-Host "Please wait..."
Write-Host "==============================================="
Write-Host ""

try {

    $Result = & $SpeedtestCLI `
        --accept-license `
        --accept-gdpr `
        --format=json

    if ([string]::IsNullOrWhiteSpace($Result)) {
        throw "No response received from Speedtest CLI."
    }

    $Speed = $Result | ConvertFrom-Json

}
catch {

    Write-Host ""
    Write-Host "Failed to execute Speed Test."
    Write-Host $_.Exception.Message
    Read-Host "Press Enter to exit"
    exit

}

# ----------------------------
# Get IPv4 Address
# ----------------------------

$IPv4 = (Get-NetIPAddress -AddressFamily IPv4 |
Where-Object {
    $_.IPAddress -notlike "169.254*" -and
    $_.IPAddress -ne "127.0.0.1"
} |
Select-Object -First 1 -ExpandProperty IPAddress)

# ----------------------------
# Determine WiFi SSID
# ----------------------------

if ($IPv4 -like '192.168.10.*') {
    $WiFiSSID = 'Corporate_Office'
}
elseif ($IPv4 -like '192.168.11.*') {
    $WiFiSSID = 'Corporate_Executive'
}
elseif ($IPv4 -like '192.168.12.*') {
    $WiFiSSID = 'Corporate_Guest'
}
elseif ($IPv4 -like '192.168.13.*') {
    $WiFiSSID = 'Corporate_WIFI'
}
else {
    $WiFiSSID = 'Unknown'
}

# ----------------------------
# Create Output File
# ----------------------------

$Date = Get-Date -Format "yyyy-MM-dd_HH_mm_ss"

$SafeSSID = $WiFiSSID -replace '\s+', '_'

$OutputFile = Join-Path $Desktop "Network_Speed_Test_${SafeSSID}_$Date.xlsx"

# ----------------------------
# Create Report (Vertical)
# ----------------------------

$Report = @(
    [PSCustomObject]@{ Parameter="Date"; Value=(Get-Date -Format "yyyy-MM-dd") }
    [PSCustomObject]@{ Parameter="Time"; Value=(Get-Date -Format "HH:mm:ss") }
    [PSCustomObject]@{ Parameter="WiFi SSID"; Value=$WiFiSSID }
    [PSCustomObject]@{ Parameter="Internal IP"; Value=$Speed.interface.internalIp }
    [PSCustomObject]@{ Parameter="External IP"; Value=$Speed.interface.externalIp }
    [PSCustomObject]@{ Parameter="Download Speed"; Value="$([math]::Round(($Speed.download.bandwidth*8)/1MB,2)) Mbps" }
    [PSCustomObject]@{ Parameter="Upload Speed"; Value="$([math]::Round(($Speed.upload.bandwidth*8)/1MB,2)) Mbps" }
    [PSCustomObject]@{ Parameter="Ping"; Value="$($Speed.ping.latency) ms" }
    [PSCustomObject]@{ Parameter="Jitter"; Value="$($Speed.ping.jitter) ms" }
    [PSCustomObject]@{ Parameter="Packet Loss"; Value="$($Speed.packetLoss) %" }
    [PSCustomObject]@{ Parameter="ISP"; Value=$Speed.isp }
    [PSCustomObject]@{ Parameter="Speed Test Server"; Value=$Speed.server.name }
    [PSCustomObject]@{ Parameter="Server Location"; Value=$Speed.server.location }
    [PSCustomObject]@{ Parameter="Country"; Value=$Speed.server.country }
    [PSCustomObject]@{ Parameter="Result URL"; Value=$Speed.result.url }
)

# ----------------------------
# Export to Excel
# ----------------------------

$Report | Export-Excel `
    -Path $OutputFile `
    -WorksheetName "Speed Test" `
    -TableName "SpeedTest" `
    -TableStyle Medium2 `
    -AutoSize `
    -AutoFilter `
    -FreezeTopRow `
    -BoldTopRow `
    -ClearSheet

# ----------------------------
# Format Worksheet
# ----------------------------

$Excel = Open-ExcelPackage -Path $OutputFile

$Worksheet = $Excel.Workbook.Worksheets["Speed Test"]

$Worksheet.View.ShowGridLines = $false

$Range = $Worksheet.Dimension.Address

$Worksheet.Cells[$Range].Style.Border.Top.Style = "Thin"
$Worksheet.Cells[$Range].Style.Border.Bottom.Style = "Thin"
$Worksheet.Cells[$Range].Style.Border.Left.Style = "Thin"
$Worksheet.Cells[$Range].Style.Border.Right.Style = "Thin"

Close-ExcelPackage $Excel

# ----------------------------
# Finished
# ----------------------------

Write-Host ""
Write-Host "==============================================="
Write-Host "Network Speed Test Completed Successfully"
Write-Host "Report Saved To:"
Write-Host $OutputFile
Write-Host "==============================================="

Invoke-Item $OutputFile

Read-Host "Press Enter to Exit"