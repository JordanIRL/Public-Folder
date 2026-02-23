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

# Ensure PSRepository is trusted to avoid interactive prompts
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

# Granular imports are faster and more stable than the full rollup
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.DeviceManagement -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

Connect-MgGraph -Scopes "DeviceManagementManagedDevices.Read.All", "DeviceManagementApps.Read.All", "User.Read" -ErrorAction Stop

$allColumns = @(
    'complianceState',
    'userPrincipalName',
    'managedDeviceName',
    'serialNumber',
    'lastSyncDateTime',
    'enrolledDateTime',
    'manufacturer',
    'model',
    'operatingSystem',
    'osVersion',
    'imei',
    'deviceCategoryDisplayName',
    'managedDeviceOwnerType',
    'enrollmentProfileName',
    'isSupervised',
    'azureADDeviceId',
    'deviceEnrollmentType'
)

# Build optional server-side filters to limit payload in large tenants
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

# Ensure export path exists before writing
$exportPath = ".\Export\Compliance_$(Get-Date -format dd-MM-yyyy_HH.mm).xlsx"
$exportDir = Split-Path -Parent $exportPath
if ($exportDir) {
    New-Item -ItemType Directory -Force -Path $exportDir | Out-Null
}

function Convert-ToCountList {
    param([hashtable]$Counts)
    $Counts.GetEnumerator() |
    Sort-Object Value -Descending |
    ForEach-Object { [PSCustomObject]@{ Name = $_.Key; Count = $_.Value } }
}

# ── Gather User Context ──
$currentAccount = (Get-MgContext).Account
$currentUser = Get-MgUser -UserId $currentAccount | Select-Object -ExpandProperty DisplayName
if ([string]::IsNullOrWhiteSpace($currentUser)) {
    $currentUser = "Unknown User"
}

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
$age0to7 = 0
$age7to14 = 0
$age14to28 = 0
$age28Plus = 0

foreach ($d in $devices) {
    $totalDevices++

    switch ($d.complianceState) {
        'compliant' { $numCompliant++ }
        'noncompliant' { $numNonCompliant++ }
    }

    $modelKey = if ([string]::IsNullOrWhiteSpace($d.model)) { 'Unknown Model' } else { $d.model }
    if ($modelCountsHash.ContainsKey($modelKey)) { $modelCountsHash[$modelKey]++ } else { $modelCountsHash[$modelKey] = 1 }

    $osKey = if ([string]::IsNullOrWhiteSpace($d.operatingSystem)) { 'Unknown OS' } else { $d.operatingSystem }
    if ($osCountsHash.ContainsKey($osKey)) { $osCountsHash[$osKey]++ } else { $osCountsHash[$osKey] = 1 }

    $catKey = if ([string]::IsNullOrWhiteSpace($d.deviceCategoryDisplayName)) { 'Uncategorized' } else { $d.deviceCategoryDisplayName }
    if ($catCountsHash.ContainsKey($catKey)) { $catCountsHash[$catKey]++ } else { $catCountsHash[$catKey] = 1 }

    $ownerKey = switch -Regex ($d.managedDeviceOwnerType) {
        'company' { 'Corporate'; break }
        'personal' { 'Personal'; break }
        default { 'Unknown' }
    }
    if ($ownCountsHash.ContainsKey($ownerKey)) { $ownCountsHash[$ownerKey]++ } else { $ownCountsHash[$ownerKey] = 1 }

    if (-not [string]::IsNullOrWhiteSpace($d.serialNumber)) {
        if (-not $serialBuckets.ContainsKey($d.serialNumber)) { $serialBuckets[$d.serialNumber] = [System.Collections.Generic.List[object]]::new() }
        $serialBuckets[$d.serialNumber].Add($d)
    }

    $syncTime = $null
    $parsedDate = $null
    if ($d.lastSyncDateTime -is [datetime]) {
        $syncTime = $d.lastSyncDateTime
    }
    elseif ([datetime]::TryParse($d.lastSyncDateTime, [ref]$parsedDate)) {
        $syncTime = $parsedDate
    }

    if ($syncTime) {
        $days = ($now - $syncTime).TotalDays

        if ($days -gt 28) { $age28Plus++ }
        elseif ($days -gt 14) { $age14to28++ }
        elseif ($days -gt 7) { $age7to14++ }
        else { $age0to7++ }
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
    [PSCustomObject]@{ Name = 'Last 7 Days'; Count = $age0to7 }
    [PSCustomObject]@{ Name = '7 - 14 Days'; Count = $age7to14 }
    [PSCustomObject]@{ Name = '14 - 28 Days'; Count = $age14to28 }
    [PSCustomObject]@{ Name = '> 28 Days'; Count = $age28Plus }
)

# ── Fetch Last 10 Deleted Devices from Audit Logs ──
$deletedDevicesList = @()
try {
    $auditEvents = @(Get-MgDeviceManagementAuditEvent -Filter "activityType eq 'Delete ManagedDevice'" -Top 10 -Sort "activityDateTime desc" -ErrorAction Stop)
    foreach ($evt in $auditEvents) {
        $devName = ($evt.Resources | Select-Object -First 1).DisplayName
        if ([string]::IsNullOrWhiteSpace($devName)) { $devName = 'Unknown' }
        $deletedBy = $evt.Actor.UserPrincipalName
        if ([string]::IsNullOrWhiteSpace($deletedBy)) { $deletedBy = 'System' }
        $delTime = if ($evt.ActivityDateTime -is [datetime]) { $evt.ActivityDateTime.ToString('dd/MM/yy HH:mm') } else { 'N/A' }
        $deletedDevicesList += [PSCustomObject]@{ DeviceName = $devName; DeletedOn = $delTime; DeletedBy = $deletedBy }
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
        $enrTime = if ($_.enrolledDateTime -is [datetime]) { $_.enrolledDateTime.ToString('dd/MM/yy HH:mm') } else {
            $parsed = $null
            if ([datetime]::TryParse($_.enrolledDateTime, [ref]$parsed)) { $parsed.ToString('dd/MM/yy HH:mm') } else { 'N/A' }
        }
        [PSCustomObject]@{
            DeviceName   = if ([string]::IsNullOrWhiteSpace($_.managedDeviceName)) { 'Unknown' } else { $_.managedDeviceName }
            SerialNumber = if ([string]::IsNullOrWhiteSpace($_.serialNumber)) { 'N/A' } else { $_.serialNumber }

            EnrolledOn   = $enrTime
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
    $white = [System.Drawing.Color]::White
    $faintGray = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $lightGray = [System.Drawing.Color]::FromArgb(241, 243, 245)
    $midGray = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $darkText = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $subText = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $blue = [System.Drawing.Color]::FromArgb(13, 110, 253)
    $green = [System.Drawing.Color]::FromArgb(25, 135, 84)
    $redC = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $yellow = [System.Drawing.Color]::FromArgb(202, 138, 4)

    $blueBg = [System.Drawing.Color]::FromArgb(219, 234, 254)
    $greenBg = [System.Drawing.Color]::FromArgb(209, 250, 229)
    $redBg = [System.Drawing.Color]::FromArgb(254, 226, 226)
    $yellowBg = [System.Drawing.Color]::FromArgb(254, 249, 195)

    $tableAccent = $blue   # blue — matching Total Devices KPI

    # ── Column widths ──
    $ws.Column(1).Width = 3      # A gutter
    $ws.Column(2).Width = 6      # B
    $ws.Column(3).Width = 18     # C names
    $ws.Column(4).Width = 18     # D names
    $ws.Column(5).Width = 8      # E count
    $ws.Column(6).Width = 3      # F gap
    $ws.Column(7).Width = 6      # G
    $ws.Column(8).Width = 18     # H names
    $ws.Column(9).Width = 18     # I names
    $ws.Column(10).Width = 8     # J count
    $ws.Column(11).Width = 3     # K gap
    $ws.Column(12).Width = 6     # L
    $ws.Column(13).Width = 18    # M names
    $ws.Column(14).Width = 18    # N names
    $ws.Column(15).Width = 8     # O count
    $ws.Column(16).Width = 3     # P gap
    $ws.Column(17).Width = 6     # Q
    $ws.Column(18).Width = 18    # R names
    $ws.Column(19).Width = 18    # S names
    $ws.Column(20).Width = 8     # T count

    # White background over rows 1-40
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

    # ── Subtitle ──
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
    # Card 1: B-E   Card 2: G-J   Card 3: L-O
    # ══════════════════════════════════════════════════════
    $ws.Row(6).Height = 20  # label
    $ws.Row(7).Height = 46  # big number
    $ws.Row(8).Height = 12  # sublabel
    $ws.Row(9).Height = 8   # spacer

    $kpis = @(
        @{ Label = "Total Devices"; Value = $totalDevices; Sub = ""; Fg = $blue; Bg = $blueBg; Range = "B6:E8" },
        @{ Label = "Compliant"; Value = $numCompliant; Sub = ""; Fg = $green; Bg = $greenBg; Range = "G6:J8" },
        @{ Label = "Non compliant"; Value = $numNonCompliant; Sub = ""; Fg = $redC; Bg = $redBg; Range = "L6:O8" },
        @{ Label = "Compliance Rate"; Value = ($compPct / 100); Format = "0.0%"; Sub = ""; Fg = $yellow; Bg = $yellowBg; Range = "Q6:T8" }
    )

    foreach ($kpi in $kpis) {
        $r = $kpi.Range
        $firstCell = $r.Split(':')[0]
        $col = $firstCell.Substring(0, 1)
        $endCol = $r.Split(':')[1].Substring(0, 1)

        Set-ExcelRange -Worksheet $ws -Range $r -BackgroundColor $kpi.Bg

        # Outer border only to avoid inner white lines
        $ws.Cells["${col}6:${endCol}6"].Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${col}6:${endCol}6"].Style.Border.Top.Color.SetColor($lightGray)
        
        $ws.Cells["${col}8:${endCol}8"].Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${col}8:${endCol}8"].Style.Border.Bottom.Color.SetColor($lightGray)
        
        $ws.Cells["${col}6:${col}8"].Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${col}6:${col}8"].Style.Border.Left.Color.SetColor($lightGray)
        
        $ws.Cells["${endCol}6:${endCol}8"].Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $ws.Cells["${endCol}6:${endCol}8"].Style.Border.Right.Color.SetColor($lightGray)

        # Label row
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
        if ($kpi.ContainsKey('Format')) {
            $ws.Cells["${col}7"].Style.Numberformat.Format = $kpi.Format
        }
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
    # Models: B-E   Categories: G-J   Ownership: L-O
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

        # Header
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
        # Bottom border under title
        $tb = $Worksheet.Cells[$hdrRange].Style.Border.Bottom
        $tb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tb.Color.SetColor($LightGray)

        # Spacer row between title and data
        $r = $StartRow + 1
        $Worksheet.Row($r).Height = [math]::Max($Worksheet.Row($r).Height, 6)

        $r++
        foreach ($item in $Items) {
            $Worksheet.Row($r).Height = [math]::Max($Worksheet.Row($r).Height, 24)

            # Uniform background for items
            Set-ExcelRange -Worksheet $Worksheet -Range "${Col1}${r}:${ColEnd}${r}" -BackgroundColor $FaintGray

            # Name with percentage appended
            $nameText = "  $($item.Name)"
            if ($Total -gt 0) {
                $pct = [math]::Round(($item.Count / $Total) * 100, 1)
                $nameText += "  ·  $pct%"
            }
            $Worksheet.Cells["${Col1}${r}:${Col2}${r}"].Merge = $true
            $Worksheet.Cells["${Col1}${r}"].Value = $nameText
            $Worksheet.Cells["${Col1}${r}"].Style.Font.Size = 10
            $Worksheet.Cells["${Col1}${r}"].Style.Font.Color.SetColor($DarkText)
            $Worksheet.Cells["${Col1}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            $Worksheet.Cells["${Col1}${r}"].Style.WrapText = $false

            # Count (right-aligned, bold, colored)
            $Worksheet.Cells["${ColEnd}${r}"].Value = $item.Count
            $Worksheet.Cells["${ColEnd}${r}"].Style.Font.Size = 12
            $Worksheet.Cells["${ColEnd}${r}"].Style.Font.Bold = $true
            $Worksheet.Cells["${ColEnd}${r}"].Style.Font.Color.SetColor($Accent)
            $Worksheet.Cells["${ColEnd}${r}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Right
            $Worksheet.Cells["${ColEnd}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
            $Worksheet.Cells["${ColEnd}${r}"].Style.Indent = 1

            # Thin bottom border per row
            $rb = $Worksheet.Cells["${Col1}${r}:${ColEnd}${r}"].Style.Border.Bottom
            $rb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
            $rb.Color.SetColor($LightGray)

            $r++
        }
        return $r
    }

    $tableArgs = @{ Worksheet = $ws; FaintGray = $faintGray; LightGray = $lightGray; MidGray = $midGray; DarkText = $darkText; Total = $totalDevices }

    # ── Cap left/middle tables to match right column height ──
    # Right column: Ownership (title+spacer+items) + 2-row gap + Check-in Age (title+spacer+4 items)
    # All tables start at row 12, data starts at row 14.
    # Bottom row of right column = 14 + ownCounts.Count - 1 + 2 + 1 + 1 + checkinCounts.Count
    #                             = 14 + ownCounts.Count + checkinCounts.Count + 3
    $maxDataRow = 14 + $ownCounts.Count + $checkinCounts.Count + 3
    $maxSlots = $maxDataRow - 14 + 1   # items that fit starting from row 14

    function Limit-TableItems {
        param($Items, [int]$MaxSlots)
        if ($Items.Count -le $MaxSlots) { return $Items }
        $kept = @($Items | Select-Object -First ($MaxSlots - 1))
        $rest = @($Items | Select-Object -Skip ($MaxSlots - 1))
        $otherCount = ($rest | Measure-Object -Property Count -Sum).Sum
        $kept += [PSCustomObject]@{ Name = 'Other'; Count = $otherCount }
        return $kept
    }

    $modelCounts = Limit-TableItems -Items $modelCounts -MaxSlots $maxSlots
    $catCounts = Limit-TableItems -Items $catCounts   -MaxSlots $maxSlots
    $osCounts = Limit-TableItems -Items $osCounts    -MaxSlots $maxSlots

    # Models (Left Column - B)
    $null = Add-Table -Title "Device Models" -Col1 "B" -Col2 "D" -ColEnd "E" `
        -Items $modelCounts -Accent $tableAccent -StartRow 12 @tableArgs

    # Categories (Middle Left Column - G - Under Compliant)
    $null = Add-Table -Title "Device Categories" -Col1 "G" -Col2 "I" -ColEnd "J" `
        -Items $catCounts -Accent $tableAccent -StartRow 12 @tableArgs

    # Operating Systems (Middle Right Column - L - Under Noncompliant)
    $null = Add-Table -Title "Operating Systems" -Col1 "L" -Col2 "N" -ColEnd "O" `
        -Items $osCounts -Accent $tableAccent -StartRow 12 @tableArgs

    # Ownership and Check-in Age (Far Right Column - Q - Under Compliance Rate)
    $rOwn = Add-Table -Title "Device Ownership" -Col1 "Q" -Col2 "S" -ColEnd "T" `
        -Items $ownCounts -Accent $tableAccent -StartRow 12 @tableArgs

    $nextRowFarRight = $rOwn + 2

    $null = Add-Table -Title "Device Check-in Age" -Col1 "Q" -Col2 "S" -ColEnd "T" `
        -Items $checkinCounts -Accent $tableAccent -StartRow $nextRowFarRight @tableArgs

    # ══════════════════════════════════════════════════════
    # DETAIL TABLES — Deleted Devices & Recent Enrollments
    # Placed side-by-side matching the grid layout above
    # ══════════════════════════════════════════════════════

    # Find the max row used by existing tables so we can start below
    $allEndRows = @(
        (12 + 2 + $modelCounts.Count),
        (12 + 2 + $catCounts.Count),
        (12 + 2 + $osCounts.Count),
        ($nextRowFarRight + 2 + $checkinCounts.Count)
    )
    $maxRow = ($allEndRows | Measure-Object -Maximum).Maximum
    $detailStartRow = $maxRow + 3

    function Add-DetailTable {
        param(
            [string]$Title,
            [string[]]$ColStarts,     # e.g. @('B','E','H')
            [string[]]$ColEnds,       # e.g. @('D','G','J')
            [string[]]$Headers,
            $Rows,                    # array of arrays
            [int]$StartRow,
            $Worksheet,
            [System.Drawing.Color]$FaintGray,
            [System.Drawing.Color]$LightGray,
            [System.Drawing.Color]$SubText,
            [System.Drawing.Color]$DarkText
        )

        $firstCol = $ColStarts[0]
        $lastCol = $ColEnds[$ColEnds.Count - 1]

        # Title row — same style as existing Add-Table
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

        # Column headers row
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

    # ── Last 10 Deleted Devices (Left: B-J) ──
    $deletedRows = @()
    if ($deletedDevicesList.Count -gt 0) {
        foreach ($d in $deletedDevicesList) {
            $deletedRows += , @($d.DeviceName, $d.DeletedBy, $d.DeletedOn)
        }
    }
    else {
        $deletedRows += , @('No deleted devices found', '', '')
    }
    $null = Add-DetailTable -Title 'Recent Deletions' `
        -ColStarts @('B', 'E', 'I') -ColEnds @('D', 'H', 'J') `
        -Headers @('Device Name', 'Deleted By', 'Deleted On') `
        -Rows $deletedRows -StartRow $detailStartRow @detailArgs

    # ── 10 Most Recent Enrollments (Right: L-T) ──
    $enrollRows = @()
    if ($recentEnrollments.Count -gt 0) {
        foreach ($e in $recentEnrollments) {
            $enrollRows += , @($e.DeviceName, $e.SerialNumber, $e.EnrolledOn)
        }
    }
    else {
        $enrollRows += , @('No enrollments found', '', '')
    }
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
    $devices | Export-Excel -ExcelPackage $pkg -WorksheetName 'All Devices' -TableName 'AllDevices' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $devices | Where-Object complianceState -eq 'compliant' |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'Compliant' -TableName 'Compliant' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $devices | Where-Object complianceState -eq 'noncompliant' |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'Noncompliant' -TableName 'Noncompliant' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    if ($duplicateDevices) {
        $duplicateDevices |
        Export-Excel -ExcelPackage $pkg -WorksheetName 'Duplicates' -TableName 'Duplicates' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null
    }

    $windowsOS = $devices | Where-Object operatingSystem -eq 'Windows'
    if (-not $windowsOS) {
        $naProps = [ordered]@{}
        foreach ($col in $allColumns) { $naProps[$col] = 'N/A' }
        $windowsOS = [PSCustomObject]$naProps
    }
    $windowsOS |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'Windows' -TableName 'Windows' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $iOSOS = $devices | Where-Object { $_.operatingSystem -eq 'iOS' -or $_.operatingSystem -eq 'iPadOS' }
    if (-not $iOSOS) {
        $naProps = [ordered]@{}
        foreach ($col in $allColumns) { $naProps[$col] = 'N/A' }
        $iOSOS = [PSCustomObject]$naProps
    }
    $iOSOS |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'iOS' -TableName 'iOS' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $androidFullyManaged = $devices | Where-Object { $_.operatingSystem -eq 'AndroidEnterprise' -and $_.deviceEnrollmentType -ne 'androidEnterpriseCorporateWorkProfile' }
    if (-not $androidFullyManaged) {
        $naProps = [ordered]@{}
        foreach ($col in $allColumns) { $naProps[$col] = 'N/A' }
        $androidFullyManaged = [PSCustomObject]$naProps
    }
    $androidFullyManaged |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'Android Fully Managed' -TableName 'AndroidFullyManaged' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $androidCorpWorkProfile = $devices | Where-Object { $_.operatingSystem -eq 'AndroidEnterprise' -and $_.deviceEnrollmentType -eq 'androidEnterpriseCorporateWorkProfile' }
    if (-not $androidCorpWorkProfile) {
        $naProps = [ordered]@{}
        foreach ($col in $allColumns) { $naProps[$col] = 'N/A' }
        $androidCorpWorkProfile = [PSCustomObject]$naProps
    }
    $androidCorpWorkProfile |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'Android Corp Work Profile' -TableName 'AndroidCorpWorkProfile' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    $androidWorkProfile = $devices | Where-Object operatingSystem -eq 'AndroidForWork'
    if (-not $androidWorkProfile) {
        $naProps = [ordered]@{}
        foreach ($col in $allColumns) { $naProps[$col] = 'N/A' }
        $androidWorkProfile = [PSCustomObject]$naProps
    }
    $androidWorkProfile |
    Export-Excel -ExcelPackage $pkg -WorksheetName 'Personal Android OS' -TableName 'PersonalAndroidOS' -TableStyle Medium2 -NoNumberConversion $textColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru | Out-Null

    # Final pass — tab colours, coloured header rows, centre alignment
    $sheetMeta = @{
        'Report'                    = @{ Tab = [System.Drawing.Color]::FromArgb(165, 180, 252); HdrBg = $null }
        'All Devices'               = @{ Tab = [System.Drawing.Color]::FromArgb(214, 211, 209); HdrBg = [System.Drawing.Color]::FromArgb(214, 211, 209) }
        'Compliant'                 = @{ Tab = [System.Drawing.Color]::FromArgb(110, 216, 153); HdrBg = [System.Drawing.Color]::FromArgb(110, 216, 153) }
        'Noncompliant'              = @{ Tab = [System.Drawing.Color]::FromArgb(248, 143, 143); HdrBg = [System.Drawing.Color]::FromArgb(248, 143, 143) }
        'Duplicates'                = @{ Tab = [System.Drawing.Color]::FromArgb(250, 181, 105); HdrBg = [System.Drawing.Color]::FromArgb(250, 181, 105) }
        'Windows'                   = @{ Tab = [System.Drawing.Color]::FromArgb(125, 211, 252); HdrBg = [System.Drawing.Color]::FromArgb(125, 211, 252) }
        'iOS'                       = @{ Tab = [System.Drawing.Color]::FromArgb(203, 213, 225); HdrBg = [System.Drawing.Color]::FromArgb(203, 213, 225) }
        'Android Fully Managed'     = @{ Tab = [System.Drawing.Color]::FromArgb(167, 243, 208); HdrBg = [System.Drawing.Color]::FromArgb(167, 243, 208) }
        'Android Corp Work Profile' = @{ Tab = [System.Drawing.Color]::FromArgb(253, 216, 136); HdrBg = [System.Drawing.Color]::FromArgb(253, 216, 136) }
        'Personal Android OS'       = @{ Tab = [System.Drawing.Color]::FromArgb(147, 197, 253); HdrBg = [System.Drawing.Color]::FromArgb(147, 197, 253) }
    }

    foreach ($sheet in $pkg.Workbook.Worksheets) {
        $meta = $sheetMeta[$sheet.Name]
        if ($meta) { $sheet.TabColor = $meta.Tab }

        if ($sheet.Name -ne 'Report' -and $sheet.Dimension) {
            Set-ExcelRange -Worksheet $sheet -Range $sheet.Dimension.Address -HorizontalAlignment Center

            # Coloured header row with dark bold text for readability against pastel backgrounds
            if ($meta -and $meta.HdrBg) {
                $lastColLetter = $sheet.Dimension.End.Address -replace '\d+', ''
                $hdrRange = "A1:${lastColLetter}1"
                Set-ExcelRange -Worksheet $sheet -Range $hdrRange `
                    -BackgroundColor $meta.HdrBg `
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