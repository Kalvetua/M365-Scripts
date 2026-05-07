# Kräver: Microsoft.Graph (Invoke-MgGraphRequest) + ImportExcel
# Connect-MgGraph -Scopes "User.Read.All","Organization.Read.All","Directory.Read.All"

# -----------------------------------------------------------------------------
# SKU -> läsbart namn. Utöka vid behov.
# Fullständig lista: learn.microsoft.com -> "Product names and service plan identifiers for licensing"
$skuFriendly = @{
    'ENTERPRISEPACK'           = 'Office 365 E3'
    'ENTERPRISEPREMIUM'        = 'Office 365 E5'
    'STANDARDPACK'             = 'Office 365 E1'
    'SPE_E3'                   = 'Microsoft 365 E3'
    'SPE_E5'                   = 'Microsoft 365 E5'
    'SPE_F1'                   = 'Microsoft 365 F1'
    'SPE_F3'                   = 'Microsoft 365 F3'
    'SPB'                      = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_PREMIUM'    = 'Microsoft 365 Business Standard'
    'O365_BUSINESS_ESSENTIALS' = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'            = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'       = 'Microsoft 365 Apps for Enterprise'
    'EXCHANGESTANDARD'         = 'Exchange Online (Plan 1)'
    'EXCHANGEENTERPRISE'       = 'Exchange Online (Plan 2)'
    'EXCHANGEDESKLESS'         = 'Exchange Online Kiosk'
    'EXCHANGE_S_DESKLESS'      = 'Exchange Online Kiosk'
    'EMS'                      = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'               = 'Enterprise Mobility + Security E5'
    'AAD_PREMIUM'              = 'Microsoft Entra ID P1'
    'AAD_PREMIUM_P2'           = 'Microsoft Entra ID P2'
    'INTUNE_A'                 = 'Microsoft Intune Plan 1'
    'POWER_BI_STANDARD'        = 'Power BI (Free)'
    'POWER_BI_PRO'             = 'Power BI Pro'
    'PROJECTPROFESSIONAL'      = 'Project Plan 3'
    'PROJECTPREMIUM'           = 'Project Plan 5'
    'VISIOCLIENT'              = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'        = 'Visio Plan 1'
    'TEAMS_EXPLORATORY'        = 'Microsoft Teams Exploratory'
    'FLOW_FREE'                = 'Power Automate (Free)'
    'POWERAPPS_VIRAL'          = 'Power Apps (Free)'
    'MCOMEETADV'               = 'Microsoft 365 Audio Conferencing'
    'MCOEV'                    = 'Microsoft Teams Phone Standard'
    'MCOPSTN1'                 = 'Microsoft Teams Domestic Calling Plan'
    'MCOPSTN2'                 = 'Microsoft Teams Domestic & International Calling'
    'PHONESYSTEM_VIRTUALUSER'  = 'Microsoft Teams Phone Resource Account'
    'WIN10_PRO_ENT_SUB'        = 'Windows 10/11 Enterprise E3'
    'WIN10_VDA_E5'             = 'Windows 10/11 Enterprise E5'
    'DEFENDER_ENDPOINT_P1'     = 'Defender for Endpoint P1'
    'WIN_DEF_ATP'              = 'Defender for Endpoint P2'
    'ATP_ENTERPRISE'           = 'Defender for Office 365 P1'
    'THREAT_INTELLIGENCE'      = 'Defender for Office 365 P2'
    'IDENTITY_THREAT_PROTECTION' = 'Microsoft 365 E5 Security'
    'M365_F1'                  = 'Microsoft 365 F1'
    'DESKLESSPACK'             = 'Office 365 F3'
}

function Get-FriendlyName {
    param([string]$sku)
    if ($skuFriendly.ContainsKey($sku)) { return $skuFriendly[$sku] }
    return $sku
}

function ToJaNej($v) { if ($v) { 'Ja' } else { 'Nej' } }

# Hjälpare för paginerade Graph-anrop
function Invoke-GraphPaged {
    param([string]$Uri)
    $all = @()
    $next = $Uri
    while ($next) {
        $resp = Invoke-MgGraphRequest -Method GET -Uri $next
        if ($resp.value) { $all += $resp.value }
        $next = $resp.'@odata.nextLink'
    }
    return $all
}

# -----------------------------------------------------------------------------
Write-Host "Hämtar prenumerationer..." -ForegroundColor Cyan
$skus = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/subscribedSkus').value

$skuLookup = @{}
foreach ($sku in $skus) { $skuLookup[$sku.skuId] = $sku }

$summary = @(
    foreach ($sku in $skus) {
        $totalt  = [int]$sku.prepaidUnits.enabled
        $tilldel = [int]$sku.consumedUnits
        $pausade = [int]$sku.prepaidUnits.suspended
        $kvar    = $totalt - $tilldel

        if ($totalt -eq 0 -and $tilldel -eq 0 -and $pausade -eq 0) { continue }

        $status = switch ($true) {
            ($totalt -eq 0 -and $pausade -gt 0) { 'Pausad / utgången' }
            ($totalt -eq 0 -and $tilldel -gt 0) { 'Inga köpta – tilldelade ändå' }
            ($kvar -lt 0)                       { 'Övertilldelad' }
            ($kvar -eq 0)                       { 'Slut – inga lediga' }
            ($kvar -le 2)                       { 'Nästan slut' }
            default                             { 'OK' }
        }

        [PSCustomObject]@{
            'Licens'        = Get-FriendlyName $sku.skuPartNumber
            'SKU-namn'      = $sku.skuPartNumber
            'Totalt köpta'  = $totalt
            'Tilldelade'    = $tilldel
            'Tillgängliga'  = $kvar
            'Pausade'       = $pausade
            'Status'        = $status
        }
    }
)

# -----------------------------------------------------------------------------
Write-Host "Hämtar användare och licensdetaljer..." -ForegroundColor Cyan
$select = 'id,userPrincipalName,displayName,accountEnabled,department,jobTitle,usageLocation,assignedLicenses'
$users  = Invoke-GraphPaged "https://graph.microsoft.com/v1.0/users?`$select=$select&`$top=999"

$licensedUsers = $users | Where-Object {
    $_.assignedLicenses -and @($_.assignedLicenses).Count -gt 0
}

$userReport = @(
    foreach ($u in $licensedUsers) {
        $namn = @(
            foreach ($l in $u.assignedLicenses) {
                if ($skuLookup.ContainsKey($l.skuId)) {
                    Get-FriendlyName $skuLookup[$l.skuId].skuPartNumber
                } else {
                    $l.skuId
                }
            }
        )

        [PSCustomObject]@{
            'Användare'      = $u.userPrincipalName
            'Visningsnamn'   = $u.displayName
            'Aktiverad'      = ToJaNej $u.accountEnabled
            'Avdelning'      = $u.department
            'Befattning'     = $u.jobTitle
            'Plats'          = $u.usageLocation
            'Antal licenser' = $namn.Count
            'Licenser'       = ($namn | Sort-Object -Unique) -join '; '
        }
    }
)

Write-Host "Bygger rapport: $($summary.Count) licenstyper, $($userReport.Count) användare med licens" -ForegroundColor Cyan
if ($summary.Count -eq 0 -and $userReport.Count -eq 0) {
    Write-Warning "Inget att exportera."
    return
}

# -----------------------------------------------------------------------------
$org    = (Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization').value | Select-Object -First 1
$tenant = ($org.verifiedDomains | Where-Object { $_.isDefault } | Select-Object -First 1).name
if (-not $tenant) { $tenant = 'tenant' }

$timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$xlsxPath  = Join-Path $PWD "Licensrapport_${tenant}_${timestamp}.xlsx"
if (Test-Path $xlsxPath) { Remove-Item $xlsxPath -Force }

# Sortera på allvarsgrad så det som kräver åtgärd hamnar överst
$statusOrder = @{
    'Övertilldelad'                = 0
    'Inga köpta – tilldelade ändå' = 1
    'Slut – inga lediga'           = 2
    'Pausad / utgången'            = 3
    'Nästan slut'                  = 4
    'OK'                           = 5
}
$summarySorted = $summary | Sort-Object @{Expression={$statusOrder[$_.Status]}}, Licens

# Blad 1: Sammanställning
$excel = $summarySorted | Export-Excel -Path $xlsxPath `
    -WorksheetName 'Sammanställning' `
    -TableName    'Sammanstallning' `
    -TableStyle   Medium2 `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow `
    -PassThru

$wsSum = $excel.Workbook.Worksheets['Sammanställning']
$sumStart = 2
$sumEnd   = $wsSum.Dimension.End.Row
$sumLast  = $wsSum.Dimension.End.Address -replace '\d'

if ($sumEnd -ge $sumStart) {
    # Gul: nästan slut eller pausad (kolumn G = Status)
    Add-ConditionalFormatting -Worksheet $wsSum `
        -Range           "A${sumStart}:${sumLast}${sumEnd}" `
        -RuleType        Expression `
        -ConditionValue  "=OR(`$G${sumStart}=`"Nästan slut`",`$G${sumStart}=`"Pausad / utgången`")" `
        -BackgroundColor '#FFF4CE'

    # Röd: slut, övertilldelad, eller tilldelade utan inköp
    Add-ConditionalFormatting -Worksheet $wsSum `
        -Range           "A${sumStart}:${sumLast}${sumEnd}" `
        -RuleType        Expression `
        -ConditionValue  "=OR(`$G${sumStart}=`"Slut – inga lediga`",`$G${sumStart}=`"Övertilldelad`",`$G${sumStart}=`"Inga köpta – tilldelade ändå`")" `
        -BackgroundColor '#F4B6B6' `
        -Bold
}

# Blad 2: Användare och licenser
$excel = $userReport | Sort-Object Visningsnamn | Export-Excel `
    -ExcelPackage $excel `
    -WorksheetName 'Användare och licenser' `
    -TableName    'AnvandareLicenser' `
    -TableStyle   Medium2 `
    -AutoSize `
    -FreezeTopRow `
    -BoldTopRow `
    -PassThru

$wsU = $excel.Workbook.Worksheets['Användare och licenser']

# Licenskolumnen kan bli lång – sätt rimlig bredd, ingen radbryt
$wsU.Column(8).Width = 70
$wsU.Column(8).Style.WrapText = $false

$uStart = 2
$uEnd   = $wsU.Dimension.End.Row
$uLast  = $wsU.Dimension.End.Address -replace '\d'

if ($uEnd -ge $uStart) {
    # Grå: kontot är inaktiverat men har licens (kolumn C = Aktiverad)
    Add-ConditionalFormatting -Worksheet $wsU `
        -Range           "A${uStart}:${uLast}${uEnd}" `
        -RuleType        Expression `
        -ConditionValue  "=`$C${uStart}=`"Nej`"" `
        -BackgroundColor '#E0E0E0'

    # Gul: ingen plats satt (krävs för vissa tjänster, t.ex. Teams Phone)
    Add-ConditionalFormatting -Worksheet $wsU `
        -Range           "A${uStart}:${uLast}${uEnd}" `
        -RuleType        Expression `
        -ConditionValue  "=`$F${uStart}=`"`"" `
        -BackgroundColor '#FFF4CE'
}

# Blad 3: Förklaring
$explain = $excel.Workbook.Worksheets.Add('Förklaring')
$explain.Cells[1,1].Value = "Licensrapport – $tenant"
$explain.Cells[1,1].Style.Font.Bold = $true
$explain.Cells[1,1].Style.Font.Size = 14
$explain.Cells[2,1].Value = "Genererad $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
$explain.Cells[2,1].Style.Font.Italic = $true

$rows = @(
    @('Bladet "Sammanställning"', 'Visar varje licens som är köpt eller tilldelad i tenanten, sorterat så att det som kräver åtgärd hamnar överst.')
    @('Licens',                   'Det läsbara namnet, t.ex. "Microsoft 365 Business Premium". Är SKU:n okänd visas det råa namnet – komplettera tabellen i toppen av skriptet.')
    @('SKU-namn',                 'Det interna namnet hos Microsoft (t.ex. SPB, ENTERPRISEPACK). Bra vid felsökning eller automation.')
    @('Totalt köpta',             'Antal licenser ni har betalt för.')
    @('Tilldelade',               'Antal licenser som är aktivt utdelade till användare.')
    @('Tillgängliga',             'Köpta minus tilldelade. Negativt = övertilldelat (Microsoft tillåter detta kort tid men det måste rättas).')
    @('Pausade',                  'Licenser i suspenderat tillstånd – oftast efter avslutad provperiod eller utebliven betalning.')
    @('Status',                   'OK = bra. Nästan slut = 1–2 lediga. Slut = inga lediga. Övertilldelad = fler tilldelade än köpta. Pausad / utgången = inga aktiva licenser kvar i prenumerationen.')
    @('',                         '')
    @('Bladet "Användare och licenser"', 'Alla användare som har minst en licens tilldelad. Användare utan licenser tas inte med.')
    @('Aktiverad',                'Ja = inloggning är möjlig. Nej = kontot är blockerat men licensen är ändå tilldelad – ofta ett tecken på att licensen kan frigöras.')
    @('Plats',                    'Krävs för vissa tjänster (t.ex. Teams Phone). Tom plats markeras gult.')
    @('Antal licenser',           'Antal olika licenser tilldelade till användaren.')
    @('Licenser',                 'Lista över alla licenser användaren har, separerade med semikolon.')
)
for ($i = 0; $i -lt $rows.Count; $i++) {
    $r = $i + 4
    $explain.Cells[$r,1].Value = $rows[$i][0]
    $explain.Cells[$r,1].Style.Font.Bold = $true
    $explain.Cells[$r,2].Value = $rows[$i][1]
    $explain.Cells[$r,2].Style.WrapText = $true
}
$explain.Column(1).Width = 36
$explain.Column(2).Width = 90

$legendRow = $rows.Count + 6
$explain.Cells[$legendRow,1].Value = 'Färgkoder: Röd = slut/övertilldelad. Gul = nästan slut, pausad eller saknad plats. Grå = inaktiverat konto med licens.'
$explain.Cells[$legendRow,1].Style.Font.Italic = $true

Close-ExcelPackage $excel

Write-Host "Exporterade $($summary.Count) licenstyper och $($userReport.Count) användare till $xlsxPath" -ForegroundColor Green
