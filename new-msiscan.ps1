param(
    [Parameter(Mandatory = $true)]
    [string]$MsiPath,

    [string]$OutputDir = ".\ScanReport"
)

# ===================== VALIDATION =====================
if (!(Test-Path $MsiPath)) { Write-Error "MSI file not found"; exit 1 }
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$msiName  = [System.IO.Path]::GetFileNameWithoutExtension($MsiPath)
$jsonPath = Join-Path $OutputDir "$msiName.json"
$mdPath   = Join-Path $OutputDir "$msiName.md"

# ===================== MSI OPEN =====================
$installer = New-Object -ComObject WindowsInstaller.Installer
$db = $installer.OpenDatabase($MsiPath, 0)

# ===================== FUNCTIONS =====================
function Get-MsiTable {
    param ($Table, $Columns)
    try {
        $view = $db.OpenView("SELECT * FROM $Table")
        $view.Execute()
        $rows = @()
        while ($r = $view.Fetch()) {
            $o = @{}
            for ($i=0; $i -lt $Columns.Count; $i++) {
                $o[$Columns[$i]] = $r.StringData($i+1)
            }
            $rows += [pscustomobject]$o
        }
        $view.Close()
        return $rows
    } catch { return @() }
}

function RegRoot {
    param($r)
    switch ($r) {
        "0" { "HKCR" }
        "1" { "HKCU" }
        "2" { "HKLM" }
        "3" { "HKU" }
        default { "Unknown" }
    }
}

function Remove-EmptyRows {
    param($data)
    if (!$data) { return @() }
    $data | Where-Object {
        $_.PSObject.Properties.Value |
        Where-Object { $_ -and $_.ToString().Trim() -ne "" }
    }
}

function Write-MarkdownTable {
    param($Title, $Data)

    $out = @("## $Title","")

    if (!$Data -or $Data.Count -eq 0) {
        $out += "_No entries found_"
        $out += ""
        return $out
    }

    $cols = $Data[0].PSObject.Properties.Name
    $out += "| " + ($cols -join " | ") + " |"
    $out += "| " + (($cols | ForEach-Object { "---" }) -join " | ") + " |"

    foreach ($row in $Data) {
        $vals = foreach ($c in $cols) {
            if ($null -eq $row.$c) { "" } else { [string]$row.$c }
        }
        $out += "| " + ($vals -join " | ") + " |"
    }

    $out += ""
    return $out
}

# ===================== RAW TABLES =====================
$serviceTable = Remove-EmptyRows (
    Get-MsiTable "ServiceInstall" @(
        "ServiceName","DisplayName","StartType","ServiceType"
    )
)

$addins = Remove-EmptyRows (
    Get-MsiTable "Class" @(
        "CLSID","ProgId","Context","Description"
    )
)

$registry = Remove-EmptyRows (
    Get-MsiTable "Registry" @(
        "Registry","Root","Key","Name","Value","Component"
    ) | ForEach-Object {
        [pscustomobject]@{
            Root  = RegRoot $_.Root
            Key   = $_.Key
            Name  = if ($_.Name) { $_.Name } else { "(Default)" }
            Value = $_.Value
        }
    }
)

# ===================== SERVICES (TABLE + REGISTRY) =====================
$servicesFromRegistry = $registry | Where-Object {
    $_.Root -eq "HKLM" -and
    $_.Key -match "^SYSTEM\\CurrentControlSet\\Services\\[^\\]+$"
} | Group-Object Key | ForEach-Object {
    $svc = $_.Group
    [pscustomobject]@{
        ServiceName = ($_.Name -split "\\")[-1]
        StartType   = ($svc | Where-Object Name -eq "Start").Value
        Type        = ($svc | Where-Object Name -eq "Type").Value
        Source      = "Registry"
    }
}

$servicesFromTable = $serviceTable | ForEach-Object {
    $_ | Add-Member Source "ServiceInstall" -Force -PassThru
}

$services = @($servicesFromTable) + @($servicesFromRegistry)

# ---- ADD RISK LABEL (FIXED) ----
$services = $services | ForEach-Object {
    $risk =
        if ($_.StartType -eq "2") { "Auto-start" }
        elseif ($_.StartType -eq "4") { "Disabled" }
        else { "Manual" }

    $_ | Add-Member -NotePropertyName Risk `
                    -NotePropertyValue $risk `
                    -Force -PassThru
}

# ===================== DRIVERS =====================
$drivers = $services | Where-Object {
    $_.Type -match "1|2"
}

# ===================== REPORT OBJECT =====================
$report = [ordered]@{
    MSI        = $msiName
    ScannedOn = Get-Date
    Services  = $services
    Drivers   = $drivers
    AddIns    = $addins
    Registry  = $registry
}

# ===================== JSON =====================
$report | ConvertTo-Json -Depth 10 | Out-File $jsonPath -Encoding UTF8

# ===================== MARKDOWN =====================
$md = @()

$md += "# MSI Scan Report"
$md += ""
$md += "**Application:** $msiName"
$md += ""
$md += "**Generated:** $(Get-Date)"
$md += ""

$md += "## Summary"
$md += ""
$md += "| Section | Count |"
$md += "|--------|-------|"
foreach ($k in $report.Keys) {
    if ($report[$k] -is [System.Collections.IEnumerable]) {
        $md += "| $k | $($report[$k].Count) |"
    }
}
$md += ""

$md += "## Red Flags"
$md += ""

$md += "### Services (All)"
$md += Write-MarkdownTable "Services" $services

$md += "### Drivers"
$md += Write-MarkdownTable "Drivers" $drivers

$md += "### HKLM Registry Writes"
$md += Write-MarkdownTable "Registry" ($registry | Where-Object Root -eq "HKLM")

$md += "## Add-ins / COM Classes"
$md += Write-MarkdownTable "AddIns" $addins

$md -join "`r`n" | Out-File $mdPath -Encoding UTF8

Write-Host "Scan completed successfully"
Write-Host "JSON report     : $jsonPath"
Write-Host "Markdown report : $mdPath"
