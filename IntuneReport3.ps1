param(
    [int]$MaxDevices = 20000,
    [int]$PageSize = 500,
    [int]$TopModels = 10,
    [string[]]$OperatingSystemFilter
)

# ── Environment Setup ──
$requirements = @(
    @{ Command = 'Connect-MgGraph'; Module = 'Microsoft.Graph' }
    @{ Command = 'Export-Excel'; Module = 'ImportExcel' }
)

try {
    if (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue) {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
    }
}
catch {}

foreach ($req in $requirements) {
    if (-not (Get-Module -Name $req.Module -ListAvailable -ErrorAction SilentlyContinue)) {
        try {
            Write-Host "Installing missing dependency '$($req.Module)'... This may take a moment." -ForegroundColor Cyan
            Install-Module -Name $req.Module -Scope CurrentUser -Force -AllowClobber -Confirm:$false -SkipPublisherCheck -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to install module '$($req.Module)'. Script may fail."
        }
    }
}

Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "DeviceManagementApps.Read.All", "User.Read" -ErrorAction Stop

$allColumns = @(
    'complianceState', 'userPrincipalName', 'managedDeviceName', 'serialNumber',
    'lastSyncDateTime', 'enrolledDateTime', 'manufacturer', 'model',
    'operatingSystem', 'osVersion', 'imei', 'deviceCategoryDisplayName',
    'managedDeviceOwnerType', 'enrollmentProfileName', 'isSupervised',
    'azureADDeviceId', 'deviceEnrollmentType'
)

$filter = $null
if ($OperatingSystemFilter) {
    $escaped = $OperatingSystemFilter | ForEach-Object { "'$(($_ -replace "'", "''"))'" }
    $filter = "operatingSystem in ($($escaped -join ', '))"
}

$textColumns = $allColumns | Where-Object { $_ -notmatch 'DateTime' }
$deviceQuery = @{ All = $true; Property = $allColumns; PageSize = $PageSize }
if ($filter) { $deviceQuery.Filter = $filter }

$devices = @(Get-MgDeviceManagementManagedDevice @deviceQuery | Select-Object -First $MaxDevices -Property $allColumns)
if (-not $devices) { throw "No devices returned; check filters or permissions." }
if ($devices.Count -eq $MaxDevices) { Write-Warning "Device list truncated at $MaxDevices; increase -MaxDevices to fetch more." }

$exportPath = ".\Export\Compliance_$(Get-Date -format dd-MM-yyyy_HH.mm).xlsx"
$exportDir = Split-Path -Parent $exportPath
if ($exportDir) { New-Item -ItemType Directory -Force -Path $exportDir | Out-Null }

function Convert-ToCountList {
    param([hashtable]$Counts)
    $Counts.GetEnumerator() |
    Sort-Object Value -Descending |
    ForEach-Object { [PSCustomObject]@{ Name = $_.Key; Count = $_.Value } }
}

function Resolve-DateValue {
    param($Value)
    if ($Value -is [datetime]) { return $Value }
    $parsed = $null
    if ([datetime]::TryParse($Value, [ref]$parsed)) { return $parsed }
    return $null
}

# ── Gather User Context ──
$currentAccount = (Get-MgContext).Account
$currentUser = Get-MgUser -UserId $currentAccount | Select-Object -ExpandProperty DisplayName
if ([string]::IsNullOrWhiteSpace($currentUser)) { $currentUser = "Unknown User" }

# ── Data Sets ──
$totalDevices = 0
$numCompliant = 0
$numNonCompliant = 0
$modelCountsHash = @{}
$osCountsHash = @{}
$catCountsHash = @{}
$ownCountsHash = @{}
$serialBuckets = @{}

$now = Get-Date
$ageBuckets = @(0, 0, 0, 0)  # 0-7, 7-14, 14-28, 28+

foreach ($d in $devices) {
    $totalDevices++

    switch ($d.complianceState) {
        'compliant'    { $numCompliant++ }
        'noncompliant' { $numNonCompliant++ }
    }

    $modelKey = if ([string]::IsNullOrWhiteSpace($d.model)) { 'Unknown Model' } else { $d.model }
    $modelCountsHash[$modelKey]++

    $osKey = if ([string]::IsNullOrWhiteSpace($d.operatingSystem)) { 'Unknown OS' } else { $d.operatingSystem }
    $osCountsHash[$osKey]++

    $catKey = if ([string]::IsNullOrWhiteSpace($d.deviceCategoryDisplayName)) { 'Uncategorized' } else { $d.deviceCategoryDisplayName }
    $catCountsHash[$catKey]++

    $ownerKey = switch -Regex ($d.managedDeviceOwnerType) {
        'company'  { 'Corporate'; break }
        'personal' { 'Personal'; break }
        default    { 'Unknown' }
    }
    $ownCountsHash[$ownerKey]++

    if (-not [string]::IsNullOrWhiteSpace($d.serialNumber)) {
        if (-not $serialBuckets.ContainsKey($d.serialNumber)) { $serialBuckets[$d.serialNumber] = [System.Collections.Generic.List[object]]::new() }
        $serialBuckets[$d.serialNumber].Add($d)
    }

    $syncTime = Resolve-DateValue $d.lastSyncDateTime
    if ($syncTime) {
        $days = ($now - $syncTime).TotalDays
        if     ($days -gt 28) { $ageBuckets[3]++ }
        elseif ($days -gt 14) { $ageBuckets[2]++ }
        elseif ($days -gt 7)  { $ageBuckets[1]++ }
        else                  { $ageBuckets[0]++ }
    }
}

$knownComplianceTotal = $numCompliant + $numNonCompliant
$compPct = if ($knownComplianceTotal -gt 0) { [math]::Round(($numCompliant / $knownComplianceTotal) * 100, 1) } else { 0 }
$modelCounts = Convert-ToCountList -Counts $modelCountsHash | Select-Object -First $TopModels
$osCounts = Convert-ToCountList -Counts $osCountsHash
$catCounts = Convert-ToCountList -Counts $catCountsHash
$ownCounts = Convert-ToCountList -Counts $ownCountsHash

$duplicateDevices = $serialBuckets.GetEnumerator() |
Where-Object { $_.Value.Count -gt 1 } |
ForEach-Object { $_.Value } |
Sort-Object serialNumber, enrolledDateTime

$checkinCounts = @(
    [PSCustomObject]@{ Name = 'Last 7 Days';  Count = $ageBuckets[0] }
    [PSCustomObject]@{ Name = '7 - 14 Days';  Count = $ageBuckets[1] }
    [PSCustomObject]@{ Name = '14 - 28 Days'; Count = $ageBuckets[2] }
    [PSCustomObject]@{ Name = '> 28 Days';    Count = $ageBuckets[3] }
)

# ── Fetch Last 10 Deleted Devices from Audit Logs ──
$deletedDevicesList = [System.Collections.Generic.List[object]]::new()
try {
    $auditEvents = @(Get-MgDeviceManagementAuditEvent -Filter "activityType eq 'Delete ManagedDevice'" -Top 10 -Sort "activityDateTime desc" -ErrorAction Stop)
    foreach ($evt in $auditEvents) {
        $devName = ($evt.Resources | Select-Object -First 1).DisplayName
        if ([string]::IsNullOrWhiteSpace($devName)) { $devName = 'Unknown' }
        $deletedBy = $evt.Actor.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($deletedBy)) { $deletedBy = 'System' }
        $delTime = if ($evt.ActivityDateTime -is [datetime]) { $evt.ActivityDateTime.ToString('dd/MM/yy HH:mm') } else { 'N/A' }
        $deletedDevicesList.Add([PSCustomObject]@{ DeviceName = $devName; DeletedOn = $delTime; DeletedBy = $deletedBy })
    }
}
catch {
    Write-Warning "Could not fetch audit events for deleted devices: $_"
}

# ── 10 Most Recent Enrollments ──
$recentEnrollments = @($devices |
    Where-Object { $_.enrolledDateTime } |
    Sort-Object enrolledDateTime -Descending |
    Select-Object -First 10 |
    ForEach-Object {
        $enrTime = Resolve-DateValue $_.enrolledDateTime
        [PSCustomObject]@{
            DeviceName   = if ([string]::IsNullOrWhiteSpace($_.managedDeviceName)) { 'Unknown' } else { $_.managedDeviceName }
            SerialNumber = if ([string]::IsNullOrWhiteSpace($_.serialNumber)) { 'N/A' } else { $_.serialNumber }
            EnrolledOn   = if ($enrTime) { $enrTime.ToString('dd/MM/yy HH:mm') } else { 'N/A' }
        }
    })

# ══════════════════════════════════════════════════════════════
# REPORT — Clean Professional Theme
# ══════════════════════════════════════════════════════════════
[PSCustomObject]@{ X = '' } | Export-Excel -Path $exportPath -WorksheetName 'Report'

$pkg = Open-ExcelPackage -Path $exportPath
try {
    $ws = $pkg.Workbook.Worksheets['Report']
    $ws.Cells.Clear()
    $ws.Cells.Style.Font.Name = 'Segoe UI Semibold'
    $ws.View.ShowGridLines = $false

    # ── Colours ──
    $white     = [System.Drawing.Color]::White
    $faintGray = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $lightGray = [System.Drawing.Color]::FromArgb(241, 243, 245)
    $midGray   = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $darkText  = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $subText   = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $blue      = [System.Drawing.Color]::FromArgb(13, 110, 253)
    $green     = [System.Drawing.Color]::FromArgb(25, 135, 84)
    $redC      = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $yellow    = [System.Drawing.Color]::FromArgb(202, 138, 4)

    $blueBg    = [System.Drawing.Color]::FromArgb(219, 234, 254)
    $greenBg   = [System.Drawing.Color]::FromArgb(209, 250, 229)
    $redBg     = [System.Drawing.Color]::FromArgb(254, 226, 226)
    $yellowBg  = [System.Drawing.Color]::FromArgb(254, 249, 195)

    $tableAccent = $blue

    # ── Column widths (repeating pattern: gutter=3, indent=6, name=18, name=18, count=8) ──
    @(3,6,18,18,8, 3,6,18,18,8, 3,6,18,18,8, 3,6,18,18,8) |
        ForEach-Object -Begin { $i = 1 } -Process { $ws.Column($i).Width = $_; $i++ }

    Set-ExcelRange -Worksheet $ws -Range "A1:U40" -BackgroundColor $white

    # ══════════════════════════════════════════════════════
    # HEADER BANNER (Rows 1-4)
    # ══════════════════════════════════════════════════════
    $ws.Row(1).Height = 10
    $ws.Row(2).Height = 50
    $ws.Cells["B2:T2"].Merge = $true
    $ws.Cells["B2"].Value = "Device Compliance Report"
    $ws.Cells["B2"].Style.Font.Size = 22
    $ws.Cells["B2"].Style.Font.Bold = $true
    $ws.Cells["B2"].Style.Font.Color.SetColor($darkText)
    $ws.Cells["B2"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Bottom

    $ws.Row(3).Height = 5

    $ws.Row(4).Height = 20
    $ws.Cells["B4:T4"].Merge = $true
    $ws.Cells["B4"].Value = "Generated $($now.ToString('dd MMM yyyy HH:mm')) by $currentUser"
    $ws.Cells["B4"].Style.Font.Size = 9
    $ws.Cells["B4"].Style.Font.Color.SetColor($midGray)
    $ws.Cells["B4"].Style.Font.Italic = $true
    $ws.Cells["B4"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

    $ws.Row(5).Height = 22

    # ══════════════════════════════════════════════════════
    # KPI CARDS (Rows 6-9)
    # ══════════════════════════════════════════════════════
    $ws.Row(6).Height = 20
    $ws.Row(7).Height = 46
    $ws.Row(8).Height = 12
    $ws.Row(9).Height = 8

    $kpis = @(
        @{ Label = "Total Devices";   Value = $totalDevices;    Sub = ""; Fg = $blue;   Bg = $blueBg;   Range = "B6:E8" },
        @{ Label = "Compliant";       Value = $numCompliant;    Sub = ""; Fg = $green;  Bg = $greenBg;  Range = "G6:J8" },
        @{ Label = "Non compliant";   Value = $numNonCompliant; Sub = ""; Fg = $redC;   Bg = $redBg;    Range = "L6:O8" },
        @{ Label = "Compliance Rate"; Value = ($compPct / 100); Sub = ""; Fg = $yellow; Bg = $yellowBg; Range = "Q6:T8"; Format = "0.0%" }
    )

    foreach ($kpi in $kpis) {
        $r = $kpi.Range
        $col    = $r.Split(':')[0].Substring(0, 1)
        $endCol = $r.Split(':')[1].Substring(0, 1)

        Set-ExcelRange -Worksheet $ws -Range $r -BackgroundColor $kpi.Bg

        $ws.Cells["${col}6:${endCol}6"].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${col}6:${endCol}6"].Style.Border.Top.Color.SetColor($lightGray)
        $ws.Cells["${col}8:${endCol}8"].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${col}8:${endCol}8"].Style.Border.Bottom.Color.SetColor($lightGray)
        $ws.Cells["${col}6:${col}8"].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${col}6:${col}8"].Style.Border.Left.Color.SetColor($lightGray)
        $ws.Cells["${endCol}6:${endCol}8"].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${endCol}6:${endCol}8"].Style.Border.Right.Color.SetColor($lightGray)

        # Label
        $ws.Cells["${col}6:${endCol}6"].Merge = $true
        $ws.Cells["${col}6"].Value = $kpi.Label.ToUpper()
        $ws.Cells["${col}6"].Style.Font.Size = 9
        $ws.Cells["${col}6"].Style.Font.Bold = $true
        $ws.Cells["${col}6"].Style.Font.Color.SetColor($subText)
        $ws.Cells["${col}6"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $ws.Cells["${col}6"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Bottom

        # Big number
        $ws.Cells["${col}7:${endCol}7"].Merge = $true
        $ws.Cells["${col}7"].Value = $kpi.Value
        if ($kpi.ContainsKey('Format')) { $ws.Cells["${col}7"].Style.Numberformat.Format = $kpi.Format }
        $ws.Cells["${col}7"].Style.Font.Size = 36
        $ws.Cells["${col}7"].Style.Font.Bold = $true
        $ws.Cells["${col}7"].Style.Font.Color.SetColor($kpi.Fg)
        $ws.Cells["${col}7"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $ws.Cells["${col}7"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

        # Sub label
        $ws.Cells["${col}8:${endCol}8"].Merge = $true
        $ws.Cells["${col}8"].Value = "  $($kpi.Sub)"
        $ws.Cells["${col}8"].Style.Font.Size = 8
        $ws.Cells["${col}8"].Style.Font.Color.SetColor($midGray)
        $ws.Cells["${col}8"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Top
    }

    $ws.Row(10).Height = 18
    $ws.Row(11).Height = 10

    # ══════════════════════════════════════════════════════
    # DATA TABLES (Row 11+)
    # ══════════════════════════════════════════════════════

    function Add-Table {
        param(
            [string]$Title, [string]$Col1, [string]$Col2, [string]$ColEnd,
            $Items, [System.Drawing.Color]$Accent, [int]$StartRow,
            $Worksheet,
            [System.Drawing.Color]$FaintGray,
            [System.Drawing.Color]$LightGray,
            [System.Drawing.Color]$MidGray,
            [System.Drawing.Color]$DarkText,
            [int]$Total = 0
        )

        $hdrRange = "${Col1}${StartRow}:${ColEnd}${StartRow}"
        $Worksheet.Cells[$hdrRange].Merge = $true
        $Worksheet.Cells["${Col1}${StartRow}"].Value = $Title
        $Worksheet.Cells["${Col1}${StartRow}"].Style.Font.Size = 11
        $Worksheet.Cells["${Col1}${StartRow}"].Style.Font.Bold = $true
        $Worksheet.Cells["${Col1}${StartRow}"].Style.Font.Color.SetColor($DarkText)
        $Worksheet.Cells["${Col1}${StartRow}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $Worksheet.Cells["${Col1}${StartRow}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        Set-ExcelRange -Worksheet $Worksheet -Range $hdrRange -BackgroundColor $FaintGray
        $Worksheet.Row($StartRow).Height = [math]::Max($Worksheet.Row($StartRow).Height, 28)
        $tb = $Worksheet.Cells[$hdrRange].Style.Border.Bottom
        $tb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tb.Color.SetColor($LightGray)

        $r = $StartRow + 1
        $Worksheet.Row($r).Height = [math]::Max($Worksheet.Row($r).Height, 6)

        $r++
        foreach ($item in $Items) {
            $Worksheet.Row($r).Height = [math]::Max($Worksheet.Row($r).Height, 24)
            Set-ExcelRange -Worksheet $Worksheet -Range "${Col1}${r}:${ColEnd}${r}" -BackgroundColor $FaintGray

            $nameText = "  $($item.Name)"
            if ($Total -gt 0) { $nameText += "  ·  $([math]::Round(($item.Count / $Total) * 100, 1))%" }
            $Worksheet.Cells["${Col1}${r}:${Col2}${r}"].Merge = $true
            $Worksheet.Cells["${Col1}${r}"].Value = $nameText
            $Worksheet.Cells["${Col1}${r}"].Style.Font.Size = 10
            $Worksheet.Cells["${Col1}${r}"].Style.Font.Color.SetColor($DarkText)
            $Worksheet.Cells["${Col1}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            $Worksheet.Cells["${Col1}${r}"].Style.WrapText = $false

            $Worksheet.Cells["${ColEnd}${r}"].Value = $item.Count
            $Worksheet.Cells["${ColEnd}${r}"].Style.Font.Size = 12
            $Worksheet.Cells["${ColEnd}${r}"].Style.Font.Bold = $true
            $Worksheet.Cells["${ColEnd}${r}"].Style.Font.Color.SetColor($Accent)
            $Worksheet.Cells["${ColEnd}${r}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
            $Worksheet.Cells["${ColEnd}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            $Worksheet.Cells["${ColEnd}${r}"].Style.Indent = 1

            $rb = $Worksheet.Cells["${Col1}${r}:${ColEnd}${r}"].Style.Border.Bottom
            $rb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
            $rb.Color.SetColor($LightGray)

            $r++
        }
        return $r
    }

    $tableArgs = @{ Worksheet = $ws; FaintGray = $faintGray; LightGray = $lightGray; MidGray = $midGray; DarkText = $darkText; Total = $totalDevices }

    $maxDataRow = 14 + $ownCounts.Count + $checkinCounts.Count + 3
    $maxSlots = $maxDataRow - 14 + 1

    function Limit-TableItems {
        param($Items, [int]$MaxSlots)
        if ($Items.Count -le $MaxSlots) { return $Items }
        $kept = @($Items | Select-Object -First ($MaxSlots - 1))
        $rest = @($Items | Select-Object -Skip ($MaxSlots - 1))
        $kept += [PSCustomObject]@{ Name = 'Other'; Count = ($rest | Measure-Object -Property Count -Sum).Sum }
        return $kept
    }

    $modelCounts = Limit-TableItems -Items $modelCounts -MaxSlots $maxSlots
    $catCounts   = Limit-TableItems -Items $catCounts   -MaxSlots $maxSlots
    $osCounts    = Limit-TableItems -Items $osCounts    -MaxSlots $maxSlots

    $null = Add-Table -Title "Device Models"     -Col1 "B" -Col2 "D" -ColEnd "E" -Items $modelCounts -Accent $tableAccent -StartRow 12 @tableArgs
    $null = Add-Table -Title "Device Categories" -Col1 "G" -Col2 "I" -ColEnd "J" -Items $catCounts   -Accent $tableAccent -StartRow 12 @tableArgs
    $null = Add-Table -Title "Operating Systems" -Col1 "L" -Col2 "N" -ColEnd "O" -Items $osCounts    -Accent $tableAccent -StartRow 12 @tableArgs
    $rOwn = Add-Table -Title "Device Ownership"  -Col1 "Q" -Col2 "S" -ColEnd "T" -Items $ownCounts   -Accent $tableAccent -StartRow 12 @tableArgs

    $nextRowFarRight = $rOwn + 2
    $null = Add-Table -Title "Device Check-in Age" -Col1 "Q" -Col2 "S" -ColEnd "T" -Items $checkinCounts -Accent $tableAccent -StartRow $nextRowFarRight @tableArgs

    # ══════════════════════════════════════════════════════
    # DETAIL TABLES — Deleted Devices & Recent Enrollments
    # ══════════════════════════════════════════════════════

    $allEndRows = @(
        (14 + $modelCounts.Count),
        (14 + $catCounts.Count),
        (14 + $osCounts.Count),
        ($nextRowFarRight + 2 + $checkinCounts.Count)
    )
    $detailStartRow = ($allEndRows | Measure-Object -Maximum).Maximum + 3

    function Add-DetailTable {
        param(
            [string]$Title,
            [string[]]$ColStarts,
            [string[]]$ColEnds,
            [string[]]$Headers,
            $Rows,
            [int]$StartRow,
            $Worksheet,
            [System.Drawing.Color]$FaintGray,
            [System.Drawing.Color]$LightGray,
            [System.Drawing.Color]$SubText,
            [System.Drawing.Color]$DarkText
        )

        $firstCol = $ColStarts[0]
        $lastCol = $ColEnds[$ColEnds.Count - 1]

        # Title row
        $hdrRange = "${firstCol}${StartRow}:${lastCol}${StartRow}"
        $Worksheet.Cells[$hdrRange].Merge = $true
        $Worksheet.Cells["${firstCol}${StartRow}"].Value = $Title
        $Worksheet.Cells["${firstCol}${StartRow}"].Style.Font.Size = 11
        $Worksheet.Cells["${firstCol}${StartRow}"].Style.Font.Bold = $true
        $Worksheet.Cells["${firstCol}${StartRow}"].Style.Font.Color.SetColor($DarkText)
        $Worksheet.Cells["${firstCol}${StartRow}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
        $Worksheet.Cells["${firstCol}${StartRow}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        Set-ExcelRange -Worksheet $Worksheet -Range $hdrRange -BackgroundColor $FaintGray
        $Worksheet.Row($StartRow).Height = 28
        $tb = $Worksheet.Cells[$hdrRange].Style.Border.Bottom
        $tb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tb.Color.SetColor($LightGray)

        # Column headers
        $r = $StartRow + 1
        $Worksheet.Row($r).Height = 20
        Set-ExcelRange -Worksheet $Worksheet -Range "${firstCol}${r}:${lastCol}${r}" -BackgroundColor $FaintGray
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $cs = $ColStarts[$i]; $ce = $ColEnds[$i]
            $Worksheet.Cells["${cs}${r}:${ce}${r}"].Merge = $true
            $Worksheet.Cells["${cs}${r}"].Value = $Headers[$i]
            $Worksheet.Cells["${cs}${r}"].Style.Font.Size = 9
            $Worksheet.Cells["${cs}${r}"].Style.Font.Bold = $true
            $Worksheet.Cells["${cs}${r}"].Style.Font.Color.SetColor($SubText)
            $Worksheet.Cells["${cs}${r}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
            $Worksheet.Cells["${cs}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        }
        $hb = $Worksheet.Cells["${firstCol}${r}:${lastCol}${r}"].Style.Border.Bottom
        $hb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $hb.Color.SetColor($LightGray)

        # Data rows
        $r++
        foreach ($row in $Rows) {
            $Worksheet.Row($r).Height = 24
            Set-ExcelRange -Worksheet $Worksheet -Range "${firstCol}${r}:${lastCol}${r}" -BackgroundColor $FaintGray

            for ($i = 0; $i -lt $row.Count; $i++) {
                $cs = $ColStarts[$i]; $ce = $ColEnds[$i]
                $Worksheet.Cells["${cs}${r}:${ce}${r}"].Merge = $true
                $Worksheet.Cells["${cs}${r}"].Value = $row[$i]
                $Worksheet.Cells["${cs}${r}"].Style.Font.Size = 10
                $Worksheet.Cells["${cs}${r}"].Style.Font.Color.SetColor($DarkText)
                $Worksheet.Cells["${cs}${r}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
                $Worksheet.Cells["${cs}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
                $Worksheet.Cells["${cs}${r}"].Style.WrapText = $false
            }

            $rb = $Worksheet.Cells["${firstCol}${r}:${lastCol}${r}"].Style.Border.Bottom
            $rb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
            $rb.Color.SetColor($LightGray)
            $r++
        }
        return $r
    }

    $detailArgs = @{ Worksheet = $ws; FaintGray = $faintGray; LightGray = $lightGray; SubText = $subText; DarkText = $darkText }

    $deletedRows = if ($deletedDevicesList.Count -gt 0) {
        $deletedDevicesList | ForEach-Object { , @($_.DeviceName, $_.DeletedBy, $_.DeletedOn) }
    } else { , @('No deleted devices found', '', '') }

    $null = Add-DetailTable -Title 'Recent Deletions' `
        -ColStarts @('B', 'E', 'I') -ColEnds @('D', 'H', 'J') `
        -Headers @('Device Name', 'Deleted By', 'Deleted On') `
        -Rows $deletedRows -StartRow $detailStartRow @detailArgs

    $enrollRows = if ($recentEnrollments.Count -gt 0) {
        $recentEnrollments | ForEach-Object { , @($_.DeviceName, $_.SerialNumber, $_.EnrolledOn) }
    } else { , @('No enrollments found', '', '') }

    $rAfterEnroll = Add-DetailTable -Title 'Recent Enrollments' `
        -ColStarts @('L', 'O', 'S') -ColEnds @('N', 'R', 'T') `
        -Headers @('Device Name', 'Serial Number', 'Enrolled On') `
        -Rows $enrollRows -StartRow $detailStartRow @detailArgs

    # ── Print Setup ──
    $lastDataRow = [math]::Max($rAfterEnroll, 40)
    $ws.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
    $ws.PrinterSettings.FitToPage = $true
    $ws.PrinterSettings.FitToWidth = 1
    $ws.PrinterSettings.FitToHeight = 0
    $ws.PrinterSettings.PrintArea = $ws.Cells["A1:U${lastDataRow}"]

    # ── Export Data Sheets ──
    $exportCommon = @{
        ExcelPackage       = $pkg
        TableStyle         = 'Medium2'
        NoNumberConversion = $textColumns
        AutoSize           = $true
        FreezeTopRow       = $true
        BoldTopRow         = $true
        PassThru           = $true
    }

    $devices | Export-Excel @exportCommon -WorksheetName 'All Devices' -TableName 'AllDevices' | Out-Null

    $devices | Where-Object complianceState -eq 'compliant' |
    Export-Excel @exportCommon -WorksheetName 'Compliant' -TableName 'Compliant' | Out-Null

    $devices | Where-Object complianceState -eq 'noncompliant' |
    Export-Excel @exportCommon -WorksheetName 'Noncompliant' -TableName 'Noncompliant' | Out-Null

    if ($duplicateDevices) {
        $duplicateDevices | Export-Excel @exportCommon -WorksheetName 'Duplicates' -TableName 'Duplicates' | Out-Null
    }

    # OS-specific sheets — data-driven to avoid repetition
    $naRow = [ordered]@{}; foreach ($col in $allColumns) { $naRow[$col] = 'N/A' }
    $naPlaceholder = [PSCustomObject]$naRow

    @(
        @{ Name = 'Windows';                   Table = 'Windows';                Filter = { $_.operatingSystem -eq 'Windows' } }
        @{ Name = 'iOS';                        Table = 'iOS';                    Filter = { $_.operatingSystem -eq 'iOS' -or $_.operatingSystem -eq 'iPadOS' } }
        @{ Name = 'Android Fully Managed';      Table = 'AndroidFullyManaged';    Filter = { $_.operatingSystem -eq 'AndroidEnterprise' -and $_.deviceEnrollmentType -ne 'androidEnterpriseCorporateWorkProfile' } }
        @{ Name = 'Android Corp Work Profile';  Table = 'AndroidCorpWorkProfile'; Filter = { $_.operatingSystem -eq 'AndroidEnterprise' -and $_.deviceEnrollmentType -eq 'androidEnterpriseCorporateWorkProfile' } }
        @{ Name = 'Personal Android OS';        Table = 'PersonalAndroidOS';      Filter = { $_.operatingSystem -eq 'AndroidForWork' } }
    ) | ForEach-Object {
        $filtered = @($devices | Where-Object $_.Filter)
        if (-not $filtered) { $filtered = $naPlaceholder }
        $filtered | Export-Excel @exportCommon -WorksheetName $_.Name -TableName $_.Table | Out-Null
    }

    # Final pass — tab colours, coloured header rows, centre alignment
    $sheetColors = @{
        'Report'                    = [System.Drawing.Color]::FromArgb(165, 180, 252)
        'All Devices'               = [System.Drawing.Color]::FromArgb(214, 211, 209)
        'Compliant'                 = [System.Drawing.Color]::FromArgb(110, 216, 153)
        'Noncompliant'              = [System.Drawing.Color]::FromArgb(248, 143, 143)
        'Duplicates'                = [System.Drawing.Color]::FromArgb(250, 181, 105)
        'Windows'                   = [System.Drawing.Color]::FromArgb(125, 211, 252)
        'iOS'                       = [System.Drawing.Color]::FromArgb(203, 213, 225)
        'Android Fully Managed'     = [System.Drawing.Color]::FromArgb(167, 243, 208)
        'Android Corp Work Profile' = [System.Drawing.Color]::FromArgb(253, 216, 136)
        'Personal Android OS'       = [System.Drawing.Color]::FromArgb(147, 197, 253)
    }

    foreach ($sheet in $pkg.Workbook.Worksheets) {
        $color = $sheetColors[$sheet.Name]
        if ($color) { $sheet.TabColor = $color }

        if ($sheet.Name -ne 'Report' -and $sheet.Dimension) {
            Set-ExcelRange -Worksheet $sheet -Range $sheet.Dimension.Address -HorizontalAlignment Center

            if ($color) {
                $lastColLetter = $sheet.Dimension.End.Address -replace '\d+', ''
                Set-ExcelRange -Worksheet $sheet -Range "A1:${lastColLetter}1" `
                    -BackgroundColor $color `
                    -FontColor ([System.Drawing.Color]::FromArgb(33, 37, 41)) `
                    -Bold
            }
        }
    }
}
finally {
    Close-ExcelPackage -ExcelPackage $pkg
}

# Object Properties List: https://learn.microsoft.com/en-us/graph/api/resources/intune-devices-manageddevice?view=graph-rest-1.0
