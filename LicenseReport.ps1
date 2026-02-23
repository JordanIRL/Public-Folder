param(
    [int]$PageSize = 999
)

# ── Environment Setup ──
$requirements = @(
    @{ Module = 'Microsoft.Graph' }
    @{ Module = 'ImportExcel' }
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
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop

Connect-MgGraph -Scopes "User.Read.All", "Directory.Read.All" -ErrorAction Stop

# ══════════════════════════════════════════════════════════════
# LICENCE DEFINITIONS
# ══════════════════════════════════════════════════════════════
$licenceLookup = @{
    '06ebc4ee-1bb5-47dd-8120-11324bc54e06' = 'Microsoft 365 E5'
    '6fd2c87f-b296-42f0-b197-1e91e994b900' = 'Office 365 E3'
    '3271cf8e-2be5-4a09-a549-70fd05baaa17' = 'Microsoft 365 E5 EEA (no Teams)'
    'd711d25a-a21c-492f-bd19-aae1e8ebaf30' = 'Office 365 E3 EEA (no Teams)'
    '05e9a617-0261-4cee-bb44-138d3ef5d965' = 'Microsoft 365 E3'
    'c2fe850d-fbbb-4858-b67d-bd0c6e746da3' = 'Microsoft 365 E3 EEA (no Teams)'
    '66b55226-6b4f-492c-910c-a3b7a3c9d993' = 'Microsoft 365 F3'
}

# Groups for rule evaluation
$premiumSkus = @(
    '06ebc4ee-1bb5-47dd-8120-11324bc54e06'  # M365 E5
    '6fd2c87f-b296-42f0-b197-1e91e994b900'  # O365 E3
    '3271cf8e-2be5-4a09-a549-70fd05baaa17'  # M365 E5 EEA (no Teams)
    'd711d25a-a21c-492f-bd19-aae1e8ebaf30'  # O365 E3 EEA (no Teams)
    '05e9a617-0261-4cee-bb44-138d3ef5d965'  # M365 E3
    'c2fe850d-fbbb-4858-b67d-bd0c6e746da3'  # M365 E3 EEA (no Teams)
)

$o365E3Skus = @(
    '6fd2c87f-b296-42f0-b197-1e91e994b900'  # O365 E3
    'd711d25a-a21c-492f-bd19-aae1e8ebaf30'  # O365 E3 EEA (no Teams)
)

$m365E3Skus = @(
    '05e9a617-0261-4cee-bb44-138d3ef5d965'  # M365 E3
    'c2fe850d-fbbb-4858-b67d-bd0c6e746da3'  # M365 E3 EEA (no Teams)
)

$f3Sku = '66b55226-6b4f-492c-910c-a3b7a3c9d993'

$trackedSkus = $premiumSkus + @($f3Sku)

# ══════════════════════════════════════════════════════════════
# FETCH USERS & LICENCES
# ══════════════════════════════════════════════════════════════
Write-Host "Fetching licensed users..." -ForegroundColor Cyan

$allUsers = @(Get-MgUser -All -PageSize $PageSize `
        -Property 'id', 'displayName', 'userPrincipalName', 'assignedLicenses', 'accountEnabled', 'department', 'jobTitle' `
        -Filter "assignedLicenses/`$count ne 0" -ConsistencyLevel eventual -CountVariable userCount `
        -ErrorAction Stop)

if (-not $allUsers) { throw "No licensed users returned; check permissions." }
Write-Host "Retrieved $($allUsers.Count) licensed users." -ForegroundColor Green

# ══════════════════════════════════════════════════════════════
# ANALYSE LICENCE ASSIGNMENTS
# ══════════════════════════════════════════════════════════════
$overprovisionedUsers = [System.Collections.Generic.List[object]]::new()
$licenceCountsHash = @{}

$totalWithTracked = 0
$totalMultiplePremium = 0
$totalF3WithE5 = 0

foreach ($user in $allUsers) {
    if (-not $user.AssignedLicenses) { continue }
    
    $assignedSkuIds = $user.AssignedLicenses.SkuId
    $matchedSkus = [System.Collections.Generic.List[string]]::new()
    
    foreach ($sku in $assignedSkuIds) {
        if ($sku -in $trackedSkus) {
            $matchedSkus.Add($sku)
            $licenceCountsHash[$licenceLookup[$sku]] += 1
        }
    }

    if ($matchedSkus.Count -eq 0) { continue }
    $totalWithTracked++

    $hasPremium = $matchedSkus.Where({ $_ -in $premiumSkus })
    $hasE3 = $matchedSkus.Where({ $_ -in ($o365E3Skus + $m365E3Skus) })
    $hasF3 = $f3Sku -in $matchedSkus

    $violations = [System.Collections.Generic.List[string]]::new()

    # Rule 1: Users should only have ONE of the premium licences (1-4)
    if ($hasPremium.Count -gt 1) {
        $names = ($hasPremium | ForEach-Object { $licenceLookup[$_] }) -join ' + '
        $violations.Add("Multiple E3/E5")
    }

    # Rule 2: F3 users may only pair with O365 E3 variants, not E5 or M365 E3
    if ($hasF3) {
        $otherPremium = @($hasPremium.Where({ $_ -notin $o365E3Skus }))
        if ($otherPremium.Count -gt 0) {
            $conflictNames = ($otherPremium | ForEach-Object { $licenceLookup[$_] }) -join ' + '
            $violations.Add("F3 + E5 conflict")
            $totalF3WithE5++ 
        }
    }

    $matchedNames = ($matchedSkus | ForEach-Object { $licenceLookup[$_] }) -join ', '

    $record = [PSCustomObject]@{
        DisplayName       = $user.DisplayName
        UserPrincipalName = $user.UserPrincipalName
        AccountEnabled    = $user.AccountEnabled
        AssignedLicences  = $matchedNames
        Violations        = ($violations -join '; ')
        ViolationCount    = $violations.Count
    }

    if ($violations.Count -gt 0) {
        $overprovisionedUsers.Add($record)
    }
}

Write-Host "`nAnalysis complete." -ForegroundColor Green
Write-Host "  Users with tracked licences : $totalWithTracked"
Write-Host "  Overprovisioned             : $($overprovisionedUsers.Count)" -ForegroundColor $(if ($overprovisionedUsers.Count -gt 0) { 'Red' } else { 'Green' })

# ══════════════════════════════════════════════════════════════
# PREPARE SUMMARY DATA
# ══════════════════════════════════════════════════════════════
function Convert-ToCountList {
    param([hashtable]$Counts)
    $Counts.GetEnumerator() |
    Sort-Object Value -Descending |
    ForEach-Object { [PSCustomObject]@{ Name = $_.Key; Count = $_.Value } }
}

# Licence distribution counts
$licenceCounts = Convert-ToCountList -Counts $licenceCountsHash

# Violation type counts
$violationCountsHash = @{}
foreach ($u in $overprovisionedUsers) {
    foreach ($v in ($u.Violations -split '; ')) {
        $violationCountsHash[$v] += 1
    }
}
$violationCounts = Convert-ToCountList -Counts $violationCountsHash



# ── Gather User Context ──
$currentAccount = (Get-MgContext).Account
$currentUser = Get-MgUser -UserId $currentAccount | Select-Object -ExpandProperty DisplayName
if ([string]::IsNullOrWhiteSpace($currentUser)) {
    $currentUser = "Unknown User"
}

$now = Get-Date

# ══════════════════════════════════════════════════════════════
# EXCEL EXPORT
# ══════════════════════════════════════════════════════════════
$exportPath = ".\Export\LicenceOverprovisioning_$(Get-Date -Format dd-MM-yyyy_HH.mm).xlsx"
New-Item -ItemType Directory -Force -Path (Split-Path -Parent $exportPath) | Out-Null

# ── Report Sheet Initialization ──
$pkg = [PSCustomObject]@{ Status = 'Initializing' } | Export-Excel -Path $exportPath -WorksheetName 'Report' -PassThru
try {
    $ws = $pkg.Workbook.Worksheets['Report']
    $ws.Cells.Clear()
    $ws.Cells.Style.Font.Name = 'Segoe UI Semibold'
    $ws.View.ShowGridLines = $false

    $white = [System.Drawing.Color]::White
    $faintGray = [System.Drawing.Color]::FromArgb(250, 251, 252)
    $lightGray = [System.Drawing.Color]::FromArgb(241, 243, 245)
    $midGray = [System.Drawing.Color]::FromArgb(173, 181, 189)
    $darkText = [System.Drawing.Color]::FromArgb(33, 37, 41)
    $subText = [System.Drawing.Color]::FromArgb(108, 117, 125)
    $blue = [System.Drawing.Color]::FromArgb(13, 110, 253)
    $redC = [System.Drawing.Color]::FromArgb(220, 53, 69)
    $purple = [System.Drawing.Color]::FromArgb(111, 66, 193)
    $orange = [System.Drawing.Color]::FromArgb(253, 126, 20)
    $blueBg = [System.Drawing.Color]::FromArgb(219, 234, 254)
    $redBg = [System.Drawing.Color]::FromArgb(254, 226, 226)
    $purpleBg = [System.Drawing.Color]::FromArgb(237, 233, 254)
    $orangeBg = [System.Drawing.Color]::FromArgb(255, 237, 213)
    $tableAccent = $blue

    # ── Column widths ──
    $ws.Column(1).Width = 3      # A gutter
    $ws.Column(2).Width = 6      # B
    $ws.Column(3).Width = 15     # C
    $ws.Column(4).Width = 15     # D
    $ws.Column(5).Width = 8      # E
    $ws.Column(6).Width = 3      # F gap
    $ws.Column(7).Width = 6      # G
    $ws.Column(8).Width = 15     # H
    $ws.Column(9).Width = 15     # I
    $ws.Column(10).Width = 8      # J
    $ws.Column(11).Width = 3      # K gap
    $ws.Column(12).Width = 6      # L
    $ws.Column(13).Width = 15     # M
    $ws.Column(14).Width = 15     # N
    $ws.Column(15).Width = 8      # O
    $ws.Column(16).Width = 3      # P gap
    $ws.Column(17).Width = 6      # Q
    $ws.Column(18).Width = 15     # R
    $ws.Column(19).Width = 15     # S
    $ws.Column(20).Width = 8      # T

    # White background
    Set-ExcelRange -Worksheet $ws -Range "A1:U50" -BackgroundColor $white

    # ══════════════════════════════════════════════════════
    # HEADER BANNER (Rows 1-4)
    # ══════════════════════════════════════════════════════
    $ws.Row(1).Height = 10
    $ws.Row(2).Height = 50
    $ws.Cells["B2:T2"].Merge = $true
    $ws.Cells["B2"].Value = "Licence Overprovisioning Report"
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
    # ══════════════════════════════════════════════════════
    $ws.Row(6).Height = 20
    $ws.Row(7).Height = 46
    $ws.Row(8).Height = 12
    $ws.Row(9).Height = 8

    $overprovisionedPct = if ($totalWithTracked -gt 0) { $overprovisionedUsers.Count / $totalWithTracked } else { 0 }

    $kpis = @(
        @{ Label = "Tracked Users"; Value = $totalWithTracked; Fg = $blue; Bg = $blueBg; Range = "B6:E8" },
        @{ Label = "Overprovisioned"; Value = $overprovisionedUsers.Count; Fg = $redC; Bg = $redBg; Range = "G6:J8" },
        @{ Label = "Percentage Overprovisioned"; Value = $overprovisionedPct; Format = "0.0%"; Fg = $purple; Bg = $purpleBg; Range = "L6:O8" }
    )

    foreach ($kpi in $kpis) {
        $r = $kpi.Range
        $firstCell = $r.Split(':')[0]
        $col = $firstCell.Substring(0, 1)
        $endCol = $r.Split(':')[1].Substring(0, 1)

        Set-ExcelRange -Worksheet $ws -Range $r -BackgroundColor $kpi.Bg

        # Outer border
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

        # Sub label row
        $ws.Cells["${col}8:${endCol}8"].Merge = $true
    }

    $ws.Row(10).Height = 18
    $ws.Row(11).Height = 10

    # ══════════════════════════════════════════════════════
    # DATA TABLES (Row 12+)
    # Licence Distribution: B-E   Violation Types: G-J   Departments: L-O
    # ══════════════════════════════════════════════════════
    function Add-Table {
        param(
            [string]$Title, [string]$Col1, [string]$Col2, [string]$ColEnd,
            $Items, [System.Drawing.Color]$Accent, [int]$StartRow,
            $Worksheet,
            [System.Drawing.Color]$FaintGray,
            [System.Drawing.Color]$LightGray,
            [System.Drawing.Color]$DarkText
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
        $Worksheet.Row($StartRow).Height = 28
        $tb = $Worksheet.Cells[$hdrRange].Style.Border.Bottom
        $tb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $tb.Color.SetColor($LightGray)

        # Spacer
        $r = $StartRow + 1
        $Worksheet.Row($r).Height = 6

        $r++
        foreach ($item in $Items) {
            $Worksheet.Row($r).Height = 24
            Set-ExcelRange -Worksheet $Worksheet -Range "${Col1}${r}:${ColEnd}${r}" -BackgroundColor $FaintGray

            # Name
            $Worksheet.Cells["${Col1}${r}:${Col2}${r}"].Merge = $true
            $Worksheet.Cells["${Col1}${r}"].Value = "  $($item.Name)"
            $Worksheet.Cells["${Col1}${r}"].Style.Font.Size = 10
            $Worksheet.Cells["${Col1}${r}"].Style.Font.Color.SetColor($DarkText)
            $Worksheet.Cells["${Col1}${r}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

            # Count
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

    $tableArgs = @{ Worksheet = $ws; FaintGray = $faintGray; LightGray = $lightGray; DarkText = $darkText }

    # Licence Distribution (B-E)
    $null = Add-Table -Title "Licence Distribution" -Col1 "B" -Col2 "D" -ColEnd "E" `
        -Items $licenceCounts -Accent $tableAccent -StartRow 12 @tableArgs

    # Violation Types (G-J)
    $null = Add-Table -Title "Violation Types" -Col1 "G" -Col2 "I" -ColEnd "J" `
        -Items $violationCounts -Accent $redC -StartRow 12 @tableArgs

    # ── Rules Reference (L-O) ──
    $rulesStartRow = 12
    $hdrRange = "L${rulesStartRow}:O${rulesStartRow}"
    $ws.Cells[$hdrRange].Merge = $true
    $ws.Cells["L${rulesStartRow}"].Value = "Validation Rules"
    $ws.Cells["L${rulesStartRow}"].Style.Font.Size = 11
    $ws.Cells["L${rulesStartRow}"].Style.Font.Bold = $true
    $ws.Cells["L${rulesStartRow}"].Style.Font.Color.SetColor($darkText)
    $ws.Cells["L${rulesStartRow}"].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $ws.Cells["L${rulesStartRow}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    Set-ExcelRange -Worksheet $ws -Range $hdrRange -BackgroundColor $faintGray
    $ws.Row($rulesStartRow).Height = 28
    $tbb = $ws.Cells[$hdrRange].Style.Border.Bottom
    $tbb.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
    $tbb.Color.SetColor($lightGray)

    $rr = $rulesStartRow + 1
    $ws.Row($rr).Height = 6
    $rr++

    $rules = @(
        "One per user: E5, E5 EEA, O365 E3, O365 E3 EEA"
        "F3 users may pair with O365 E3 or O365 E3 EEA"
        "F3 + E5 variants = Overprovisioned"
    )
    foreach ($rule in $rules) {
        $ws.Row($rr).Height = 24
        Set-ExcelRange -Worksheet $ws -Range "L${rr}:O${rr}" -BackgroundColor $faintGray
        $ws.Cells["L${rr}:O${rr}"].Merge = $true
        $ws.Cells["L${rr}"].Value = "  $rule"
        $ws.Cells["L${rr}"].Style.Font.Size = 9
        $ws.Cells["L${rr}"].Style.Font.Color.SetColor($subText)
        $ws.Cells["L${rr}"].Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
        $rb2 = $ws.Cells["L${rr}:O${rr}"].Style.Border.Bottom
        $rb2.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Hair
        $rb2.Color.SetColor($lightGray)
        $rr++
    }

    # ── Print Setup ──
    $ws.PrinterSettings.Orientation = [OfficeOpenXml.eOrientation]::Landscape
    $ws.PrinterSettings.FitToPage = $true
    $ws.PrinterSettings.FitToWidth = 1
    $ws.PrinterSettings.FitToHeight = 0
    $ws.PrinterSettings.PrintArea = $ws.Cells["A1:U50"]

    # ── Export Data Sheets ──
    $allExportColumns = @('DisplayName', 'UserPrincipalName', 'AccountEnabled', 'AssignedLicences', 'Violations', 'ViolationCount')

    $exportData = if ($overprovisionedUsers.Count -eq 0) {
        @( [PSCustomObject]@{
                DisplayName       = 'N/A'
                UserPrincipalName = 'N/A'
                AccountEnabled    = 'N/A'
                AssignedLicences  = 'N/A'
                Violations        = 'N/A'
                ViolationCount    = 0
            } )
    }
    else {
        $overprovisionedUsers
    }

    $null = $exportData | Export-Excel -ExcelPackage $pkg -WorksheetName 'Overprovisioned' -TableName 'Overprovisioned' `
        -TableStyle Medium2 -NoNumberConversion $allExportColumns -AutoSize -FreezeTopRow -BoldTopRow -PassThru

    # ── Tab colours & header formatting ──
    $sheetColors = @{
        'Report'          = [System.Drawing.Color]::FromArgb(165, 180, 252)
        'Overprovisioned' = [System.Drawing.Color]::FromArgb(248, 143, 143)
    }

    foreach ($sheet in $pkg.Workbook.Worksheets) {
        $color = $sheetColors[$sheet.Name]
        if ($color) { $sheet.TabColor = $color }

        if ($sheet.Name -ne 'Report' -and $sheet.Dimension) {
            Set-ExcelRange -Worksheet $sheet -Range $sheet.Dimension.Address -HorizontalAlignment Center

            if ($color) {
                $lastColLetter = $sheet.Dimension.End.Address -replace '\d+', ''
                Set-ExcelRange -Worksheet $sheet -Range "A1:${lastColLetter}1" `
                    -BackgroundColor $color -FontColor $darkText -Bold
            }
        }
    }
}
catch {
    Write-Warning "An error occurred during report generation: $($_.Exception.Message)"
    throw $_
}
finally {
    if ($null -ne $pkg) { 
        Write-Host "Finalizing file..." -ForegroundColor Gray
        Close-ExcelPackage -ExcelPackage $pkg 
    }
}

Write-Host "`nReport saved to: $exportPath" -ForegroundColor Green
