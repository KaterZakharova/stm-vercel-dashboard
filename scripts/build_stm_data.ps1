$ErrorActionPreference = "Stop"

$ScriptDir  = Split-Path $MyInvocation.MyCommand.Path -Parent
$ProjectDir = Split-Path $ScriptDir -Parent
$IndexPath  = Join-Path $ProjectDir "index.html"

$KeysPath      = "C:\Users\eklementeva\.codex\skills\odata-proxy\keys.env"
$DataDir       = Join-Path $ProjectDir "data"
$MatchedCsv    = Join-Path $DataDir "stm_matched_orders.csv"
$ClassifierCsv = Join-Path $DataDir "1c_classifier_probe.csv"
$PlanDocNumber = "ЦБ-00000218"

$MayStart  = "2026-05-01T00:00:00"
$JuneStart = "2026-06-01T00:00:00"
$JuneEnd   = "2026-07-01T00:00:00"

# ─── helpers ────────────────────────────────────────────────────────────────
function Read-Keys {
    # env vars take priority (GitHub Actions), fall back to local keys file
    $envUrl = $env:ODATA_PROXY_URL
    $envKey = $env:ODATA_PROXY_READ_KEY
    if ($envUrl -and $envKey) { return @{ Url = $envUrl; Key = $envKey } }
    $txt = Get-Content -LiteralPath $KeysPath -Raw
    @{
        Url = ([regex]::Match($txt, 'ODATA_PROXY_URL="([^"]+)"')).Groups[1].Value
        Key = ([regex]::Match($txt, 'ODATA_PROXY_READ_KEY="([^"]+)"')).Groups[1].Value
    }
}
$OData = Read-Keys
$ODataHeaders = @{ "X-API-Key" = $OData.Key }

function Invoke-OData($entity, $query, $timeout = 120) {
    $uri = "$($OData.Url)/odata/${entity}?`$format=json&$query"
    Invoke-RestMethod -Method Get -Uri $uri -Headers $ODataHeaders -TimeoutSec $timeout
}

function Invoke-ODataPaged($entity, $filterEncoded, $selectRaw, $pageSize = 1000) {
    $all = [System.Collections.Generic.List[object]]::new()
    $skip = 0
    do {
        $q = "`$select=$selectRaw&`$filter=$filterEncoded&`$top=$pageSize&`$skip=$skip"
        $r = Invoke-OData $entity $q 180
        $page = @($r.value)
        $all.AddRange($page)
        $skip += $pageSize
    } while ($page.Count -eq $pageSize)
    return $all.ToArray()
}

function Escape-OData($v) { ([string]$v).Replace("'", "''") }

function To-Num($v) {
    $s = ([string]$v).Trim().Replace(" ","").Replace([string][char]160,"").Replace(",",".")
    $d = 0.0
    if ($s -and [double]::TryParse($s, [System.Globalization.NumberStyles]::Any,
        [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    0.0
}

function Normalize-Sku($sku) {
    $s = ([string]$sku).Trim().ToUpperInvariant().Replace([string][char]0xFEFF, "")
    if ($s.StartsWith("BK") -and $s.Length -gt 2) { return $s.Substring(2) }
    $s
}

# ─── nomenclature cache ──────────────────────────────────────────────────────
$nomByRef = @{}; $nomByArticle = @{}

function Get-NomByRef($refKey) {
    $k = ([string]$refKey).Trim()
    if ($k -eq "" -or $k -eq "00000000-0000-0000-0000-000000000000") { return $null }
    if ($nomByRef.ContainsKey($k)) { return $nomByRef[$k] }
    $f = [uri]::EscapeDataString("Ref_Key eq guid'$k' and DeletionMark eq false")
    try {
        $r = Invoke-OData "Catalog_Номенклатура" "`$select=Ref_Key,Артикул,Description&`$filter=$f&`$top=1"
        $item = if ($r.value.Count -gt 0) { $r.value[0] } else { $null }
        $nomByRef[$k] = $item
        if ($item -and $item.Артикул) { $nomByArticle[(Normalize-Sku $item.Артикул)] = $item }
        return $item
    } catch { $nomByRef[$k] = $null; return $null }
}

function Get-NomByArticle($skuNorm) {
    if ($nomByArticle.ContainsKey($skuNorm)) { return $nomByArticle[$skuNorm] }
    $safe = Escape-OData $skuNorm
    $f = [uri]::EscapeDataString("Артикул eq '$safe' and DeletionMark eq false")
    try {
        $r = Invoke-OData "Catalog_Номенклатура" "`$select=Ref_Key,Артикул,Description&`$filter=$f&`$top=1"
        $item = if ($r.value.Count -gt 0) { $r.value[0] } else { $null }
        $nomByRef[$item.Ref_Key] = $item
        $nomByArticle[$skuNorm] = $item
        return $item
    } catch { $nomByArticle[$skuNorm] = $null; return $null }
}

# ─── analytics + manager caches ─────────────────────────────────────────────
$analyticsCache = @{}
function Get-AnalyticsKeys($nomRefKey) {
    if ($analyticsCache.ContainsKey($nomRefKey)) { return $analyticsCache[$nomRefKey] }
    $f = [uri]::EscapeDataString("Номенклатура_Key eq guid'$nomRefKey'")
    try {
        $r = Invoke-OData "Catalog_КлючиАналитикиУчетаНоменклатуры" "`$select=Ref_Key&`$filter=$f&`$top=200"
        $keys = @($r.value | Select-Object -ExpandProperty Ref_Key)
        $analyticsCache[$nomRefKey] = $keys; return $keys
    } catch { $analyticsCache[$nomRefKey] = @(); return @() }
}

$analyticsToNomCache = @{}
function Get-NomFromAnalyticsKey($analyticsKey) {
    $k = ([string]$analyticsKey).Trim()
    if ($k -eq "" -or $k -eq "00000000-0000-0000-0000-000000000000") { return $null }
    if ($analyticsToNomCache.ContainsKey($k)) { return $analyticsToNomCache[$k] }
    $f = [uri]::EscapeDataString("Ref_Key eq guid'$k'")
    try {
        $r = Invoke-OData "Catalog_КлючиАналитикиУчетаНоменклатуры" "`$select=Ref_Key,Номенклатура_Key&`$filter=$f&`$top=1"
        $nomKey = if ($r.value.Count -gt 0) { [string]$r.value[0].Номенклатура_Key } else { $null }
        $analyticsToNomCache[$k] = $nomKey; return $nomKey
    } catch { $analyticsToNomCache[$k] = $null; return $null }
}

$managerCache = @{}
function Get-ManagerName($managerKey) {
    $k = ([string]$managerKey).Trim()
    if ($k -eq "" -or $k -eq "00000000-0000-0000-0000-000000000000") { return "" }
    if ($managerCache.ContainsKey($k)) { return $managerCache[$k] }
    $f = [uri]::EscapeDataString("Ref_Key eq guid'$k'")
    try {
        $r = Invoke-OData "Catalog_Пользователи" "`$select=Ref_Key,Description&`$filter=$f&`$top=1"
        $name = if ($r.value.Count -gt 0) { [string]$r.value[0].Description } else { "" }
        $managerCache[$k] = $name; return $name
    } catch { $managerCache[$k] = ""; return "" }
}

# ─── step 1: production plan ─────────────────────────────────────────────────
Write-Host "=== Шаг 1: план производства $PlanDocNumber ===" -ForegroundColor Cyan
$safe = Escape-OData $PlanDocNumber
$f = [uri]::EscapeDataString("Number eq '$safe' and DeletionMark eq false")
$docResp = Invoke-OData "Document_ПланПроизводства" "`$filter=$f&`$top=1" 180
if ($docResp.value.Count -eq 0) { throw "Документ ПланПроизводства '$PlanDocNumber' не найден" }
$planDoc = $docResp.value[0]

$planByMonth = @{ "2026-05" = @{}; "2026-06" = @{} }
foreach ($line in @($planDoc.Продукция)) {
    if ($line.Отменено -eq $true) { continue }
    $qty = To-Num $line.Количество
    if ($qty -le 0) { continue }
    $n = Get-NomByRef $line.Номенклатура_Key
    if ($null -eq $n) { continue }
    $skuNorm = Normalize-Sku $n.Артикул
    $rd = if ($line.ДатаВыпуска) { try { [datetime]$line.ДатаВыпуска } catch { $null } } else { $null }
    $monthCode = if ($rd -and $rd.Year -gt 1900) { $rd.ToString("yyyy-MM") } else { "" }
    if ($monthCode -notin @("2026-05", "2026-06")) { continue }
    if (-not $planByMonth[$monthCode].ContainsKey($skuNorm)) {
        $planByMonth[$monthCode][$skuNorm] = [pscustomobject]@{
            Sku = [string]$n.Артикул; SkuNorm = $skuNorm
            Name = [string]$n.Description; PlanQty = 0.0; NomRef = [string]$n.Ref_Key
        }
    }
    $planByMonth[$monthCode][$skuNorm].PlanQty += $qty
}
Write-Host "  Май: $($planByMonth['2026-05'].Count) SKU, Июнь: $($planByMonth['2026-06'].Count) SKU"

# ─── step 2: actual release (Выпуск продукции) ───────────────────────────────
Write-Host "=== Шаг 2: фактический выпуск (ВыпускПродукции) ===" -ForegroundColor Cyan
$relFilter = [uri]::EscapeDataString(
    "Period ge datetime'$MayStart' and Period lt datetime'$JuneEnd' and Active eq true"
)
$relRecords = @()
try {
    $relRecords = Invoke-ODataPaged "AccumulationRegister_ВыпускПродукции_RecordType" $relFilter "Period,АналитикаУчетаНоменклатуры_Key,Количество"
    Write-Host "  Получено записей: $($relRecords.Count)"
} catch {
    Write-Warning "Не удалось загрузить ВыпускПродукции: $_"
}

# aggregate by analytics key first (minimizes catalog lookups)
$releaseByAnalyticsKey = @{ "2026-05" = @{}; "2026-06" = @{} }
foreach ($rec in $relRecords) {
    $dt = try { [datetime]$rec.Period } catch { continue }
    $mc = $dt.ToString("yyyy-MM")
    if ($mc -notin @("2026-05", "2026-06")) { continue }
    $ak = [string]$rec.АналитикаУчетаНоменклатуры_Key
    if (-not $releaseByAnalyticsKey[$mc].ContainsKey($ak)) { $releaseByAnalyticsKey[$mc][$ak] = 0.0 }
    $releaseByAnalyticsKey[$mc][$ak] += To-Num $rec.Количество
}

# resolve analytics key → Номенклатура_Key → skuNorm
$releaseBySkuNorm = @{ "2026-05" = @{}; "2026-06" = @{} }
foreach ($mc in @("2026-05", "2026-06")) {
    foreach ($ak in $releaseByAnalyticsKey[$mc].Keys) {
        $nomRefKey = Get-NomFromAnalyticsKey $ak; if ($null -eq $nomRefKey) { continue }
        $n = Get-NomByRef $nomRefKey; if ($null -eq $n) { continue }
        $skuNorm = Normalize-Sku $n.Артикул
        if (-not $releaseBySkuNorm[$mc].ContainsKey($skuNorm)) { $releaseBySkuNorm[$mc][$skuNorm] = 0.0 }
        $releaseBySkuNorm[$mc][$skuNorm] += $releaseByAnalyticsKey[$mc][$ak]
    }
    Write-Host "  $mc release SKUs: $($releaseBySkuNorm[$mc].Count)"
}

# ─── step 3: prices & STM classifier ─────────────────────────────────────────
Write-Host "=== Шаг 3: классификатор и заказы ===" -ForegroundColor Cyan
$classRows = @(Import-Csv -LiteralPath $ClassifierCsv -Delimiter ";")
$matchedRows = @(Import-Csv -LiteralPath $MatchedCsv)

$stmSkus = @{}; $classPrices = @{}
foreach ($row in $classRows) {
    $sn = Normalize-Sku $row."Артикул"
    if ($row."Группировка бренда 1С" -eq "СТМ" -or $row."Группировка бренда Excel" -eq "СТМ") {
        $stmSkus[$sn] = $true
        $p = To-Num $row."Плановая цена 1С"
        if ($p -eq 0) { $p = To-Num $row."Плановая цена Excel" }
        $classPrices[$sn] = $p
    }
}
foreach ($row in $matchedRows) {
    $sn = Normalize-Sku $row.SkuNorm
    if ($sn) { $stmSkus[$sn] = $true }
}
Write-Host "  STM SKU в классификаторе: $($stmSkus.Count)"

function Build-PriceMap($month) {
    $map = @{}
    foreach ($grp in ($matchedRows | Where-Object { $_.ShipMonth -eq $month } | Group-Object SkuNorm)) {
        $items = @($grp.Group)
        $qty = ($items | Measure-Object QtyGoal -Sum).Sum; if (-not $qty) { $qty = 0 }
        $ft = 0.0; $ct = 0.0
        foreach ($i in $items) { $q = To-Num $i.QtyGoal; $ft += $q * (To-Num $i.FactoryPriceNoVat); $ct += $q * (To-Num $i.ClientPriceNoVat) }
        $map[$grp.Name] = [pscustomobject]@{
            Manager = (($items | Select-Object -ExpandProperty Manager -Unique) -join "; ")
            Orders  = (($items | Select-Object -ExpandProperty OrderNumber -Unique) -join "; ")
            Nomenclature = ($items | Select-Object -First 1 -ExpandProperty Nomenclature)
            IsNew   = if (($items | Where-Object { $_.IsNew -eq "Да" }).Count -gt 0) { "Да" } else { ($items | Select-Object -First 1).IsNew }
            FactoryPrice = if ($qty -ne 0) { $ft / $qty } else { 0 }
            ClientPrice  = if ($qty -ne 0) { $ct / $qty } else { 0 }
            IsMonthMatch = $true
        }
    }
    foreach ($grp in ($matchedRows | Where-Object { $_.ShipMonth -ne $month } | Group-Object SkuNorm)) {
        if ($map.ContainsKey($grp.Name)) { continue }
        $items = @($grp.Group)
        $qty = ($items | Measure-Object QtyGoal -Sum).Sum; if (-not $qty) { $qty = 0 }
        $ft = 0.0; $ct = 0.0
        foreach ($i in $items) { $q = To-Num $i.QtyGoal; $ft += $q * (To-Num $i.FactoryPriceNoVat); $ct += $q * (To-Num $i.ClientPriceNoVat) }
        $map[$grp.Name] = [pscustomobject]@{
            Manager = (($items | Select-Object -ExpandProperty Manager -Unique) -join "; ")
            Orders  = (($items | Select-Object -ExpandProperty OrderNumber -Unique) -join "; ")
            Nomenclature = ($items | Select-Object -First 1 -ExpandProperty Nomenclature)
            IsNew   = if (($items | Where-Object { $_.IsNew -eq "Да" }).Count -gt 0) { "Да" } else { ($items | Select-Object -First 1).IsNew }
            FactoryPrice = if ($qty -ne 0) { $ft / $qty } else { 0 }
            ClientPrice  = if ($qty -ne 0) { $ct / $qty } else { 0 }
            IsMonthMatch = $false
        }
    }
    return $map
}

$mayPriceMap  = Build-PriceMap "2026-05"
$junePriceMap = Build-PriceMap "2026-06"

# ─── step 4: sales facts ─────────────────────────────────────────────────────
Write-Host "=== Шаг 4: факт продаж ===" -ForegroundColor Cyan
$salesCache = @{}
function Get-SalesFact($skuNorm, $monthStart, $monthEnd) {
    $key = "$skuNorm|$monthStart"
    if ($salesCache.ContainsKey($key)) { return $salesCache[$key] }
    $empty = [pscustomobject]@{ ShippedQty=0.0; Revenue=0.0; Cost=0.0; Managers=""; Dates=""; Rows=0; Operation="" }
    $n = Get-NomByArticle $skuNorm
    if ($null -eq $n) { $salesCache[$key] = $empty; return $empty }
    $analyticsKeys = @(Get-AnalyticsKeys $n.Ref_Key)
    if ($analyticsKeys.Count -eq 0) { $salesCache[$key] = $empty; return $empty }
    $allMoves = [System.Collections.Generic.List[object]]::new()
    foreach ($akey in $analyticsKeys) {
        $f = [uri]::EscapeDataString("АналитикаУчетаНоменклатуры_Key eq guid'$akey' and Period ge datetime'$monthStart' and Period lt datetime'$monthEnd' and Active eq true and Сторно eq false")
        $sel = "Period,Менеджер_Key,Количество,СуммаВыручкиБезНДС,СтоимостьБезНДС,ХозяйственнаяОперация"
        try {
            $r = Invoke-OData "AccumulationRegister_ВыручкаИСебестоимостьПродаж_RecordType" "`$select=$sel&`$filter=$f&`$top=500"
            foreach ($m in $r.value) { $allMoves.Add($m) }
        } catch {}
    }
    $arr = $allMoves.ToArray()
    $qty  = ($arr | Measure-Object Количество       -Sum).Sum
    $rev  = ($arr | Measure-Object СуммаВыручкиБезНДС -Sum).Sum
    $cost = ($arr | Measure-Object СтоимостьБезНДС   -Sum).Sum
    $managers = @(); foreach ($mk in ($arr | Select-Object -ExpandProperty Менеджер_Key -Unique)) { $nm = Get-ManagerName $mk; if ($nm) { $managers += $nm } }
    $dates = @($arr | ForEach-Object { try { ([datetime]$_.Period).ToString("dd.MM.yyyy") } catch {} } | Select-Object -Unique)
    $ops   = @($arr | Select-Object -ExpandProperty ХозяйственнаяОперация -Unique | Where-Object { $_ })
    $result = [pscustomobject]@{
        ShippedQty = if ($qty)  { [double]$qty }  else { 0.0 }
        Revenue    = if ($rev)  { [double]$rev }  else { 0.0 }
        Cost       = if ($cost) { [double]$cost } else { 0.0 }
        Managers   = ($managers | Select-Object -Unique) -join "; "
        Dates      = $dates -join "; "
        Rows       = $arr.Count
        Operation  = $ops -join "; "
    }
    $salesCache[$key] = $result; return $result
}

# ─── step 5: build DATA rows ──────────────────────────────────────────────────
Write-Host "=== Шаг 5: сборка строк ===" -ForegroundColor Cyan
$dataRows = [System.Collections.Generic.List[object]]::new()

function Build-MonthRows($monthCode, $monthName, $priceMap) {
    $planMonth    = $planByMonth[$monthCode]
    $releaseMonth = $releaseBySkuNorm[$monthCode]
    $mStart = if ($monthCode -eq "2026-05") { $MayStart  } else { $JuneStart }
    $mEnd   = if ($monthCode -eq "2026-05") { $JuneStart } else { $JuneEnd   }
    $included = @{}

    # rows from production plan
    foreach ($skuNorm in ($planMonth.Keys | Sort-Object)) {
        if (-not $stmSkus.ContainsKey($skuNorm)) { continue }
        $included[$skuNorm] = $true
        $pd  = $planMonth[$skuNorm]
        $rel = if ($releaseMonth.ContainsKey($skuNorm)) { $releaseMonth[$skuNorm] } else { 0.0 }
        $price = $priceMap[$skuNorm]
        $sales = Get-SalesFact $skuNorm $mStart $mEnd

        # factory price: actual cost/unit from 1C first, CSV as fallback
        $fp_1c  = if ($sales.ShippedQty -gt 0 -and $sales.Cost -gt 0) { $sales.Cost / $sales.ShippedQty } else { 0.0 }
        $fp_csv = if ($price) { [double]$price.FactoryPrice } else { 0.0 }
        if ($fp_csv -eq 0 -and $classPrices.ContainsKey($skuNorm)) { $fp_csv = [double]$classPrices[$skuNorm] }
        $fp = if ($fp_1c -gt 0) { $fp_1c } else { $fp_csv }

        # client price: actual revenue/unit from 1C first, CSV as fallback
        $acp = if ($sales.ShippedQty -ne 0) { $sales.Revenue / $sales.ShippedQty } else { 0.0 }
        $cp_csv = if ($price) { [double]$price.ClientPrice } else { 0.0 }
        $cp = if ($acp -ne 0) { $acp } elseif ($cp_csv -ne 0) { $cp_csv } else { 0.0 }

        $planQty        = [math]::Round($pd.PlanQty, 3)
        $releaseQty     = [math]::Round([math]::Max($rel, 0), 3)
        $plannedRevenue = [math]::Round($planQty * $cp, 2)
        $plannedGp      = [math]::Round($planQty * ($cp - $fp), 2)
        # actual GP: use direct cost from 1C register when available
        $actualGp = if ($sales.ShippedQty -gt 0 -and $sales.Cost -gt 0) {
            [math]::Round($sales.Revenue - $sales.Cost, 2)
        } else {
            [math]::Round($sales.Revenue - ($sales.ShippedQty * $fp), 2)
        }
        $margin = if ($sales.Revenue -ne 0) { [math]::Round($actualGp / $sales.Revenue, 4) } else { 0 }

        $ps = if ($fp_1c -gt 0 -or $acp -gt 0) { "Факт 1С" } `
              elseif ($price) { if ($price.IsMonthMatch) { "Заказы месяца" } else { "Заказы другого месяца" } } `
              elseif ($fp_csv -gt 0 -or $cp_csv -gt 0) { "CSV" } else { "Цена не найдена" }

        $script:dataRows.Add([ordered]@{
            month           = $monthName
            sku             = $pd.Sku
            isNew           = if ($price) { $price.IsNew } else { "" }
            name            = $pd.Name
            order           = if ($price) { $price.Orders } else { "" }
            planQty         = $planQty
            releaseQty      = $releaseQty
            shippedQty      = [math]::Round($sales.ShippedQty, 3)
            remainingQty    = [math]::Round($planQty - $releaseQty, 3)
            manager         = if ($price -and $price.Manager) { $price.Manager } else { $sales.Managers }
            factoryPrice    = [math]::Round($fp,  2)
            clientPrice     = [math]::Round($cp,  2)
            actualClientPrice = [math]::Round($acp, 2)
            plannedRevenue  = $plannedRevenue
            actualRevenue   = [math]::Round($sales.Revenue, 2)
            plannedGp       = $plannedGp
            actualGp        = $actualGp
            margin          = $margin
            shipDate        = $sales.Dates
            priceSource     = $ps
            operation       = $sales.Operation
            salesRows       = $sales.Rows
        })
        Write-Host "  $monthName | $($pd.Sku) | план=$planQty выпуск=$releaseQty отгружено=$([math]::Round($sales.ShippedQty,0))"
    }

    # rows with actual sales but not in plan
    foreach ($skuNorm in $priceMap.Keys) {
        if ($included.ContainsKey($skuNorm)) { continue }
        if (-not $stmSkus.ContainsKey($skuNorm)) { continue }
        $sales = Get-SalesFact $skuNorm $mStart $mEnd
        if ($sales.ShippedQty -eq 0 -and $sales.Revenue -eq 0) { continue }
        $included[$skuNorm] = $true
        $price = $priceMap[$skuNorm]
        $fp_1c  = if ($sales.ShippedQty -gt 0 -and $sales.Cost -gt 0) { $sales.Cost / $sales.ShippedQty } else { 0.0 }
        $fp_csv = if ($price) { [double]$price.FactoryPrice } else { 0.0 }
        if ($fp_csv -eq 0 -and $classPrices.ContainsKey($skuNorm)) { $fp_csv = [double]$classPrices[$skuNorm] }
        $fp  = if ($fp_1c -gt 0) { $fp_1c } else { $fp_csv }
        $acp = if ($sales.ShippedQty -ne 0) { $sales.Revenue / $sales.ShippedQty } else { 0.0 }
        $cp_csv = if ($price) { [double]$price.ClientPrice } else { 0.0 }
        $cp  = if ($acp -ne 0) { $acp } elseif ($cp_csv -ne 0) { $cp_csv } else { 0.0 }
        $actualGp = if ($sales.ShippedQty -gt 0 -and $sales.Cost -gt 0) {
            [math]::Round($sales.Revenue - $sales.Cost, 2)
        } else {
            [math]::Round($sales.Revenue - ($sales.ShippedQty * $fp), 2)
        }
        $margin = if ($sales.Revenue -ne 0) { [math]::Round($actualGp / $sales.Revenue, 4) } else { 0 }
        $n = Get-NomByArticle $skuNorm
        $name = if ($price -and $price.Nomenclature) { $price.Nomenclature } elseif ($n) { $n.Description } else { "" }

        $script:dataRows.Add([ordered]@{
            month           = $monthName
            sku             = $skuNorm
            isNew           = if ($price) { $price.IsNew } else { "" }
            name            = $name
            order           = if ($price) { $price.Orders } else { "" }
            planQty         = 0.0
            releaseQty      = 0.0
            shippedQty      = [math]::Round($sales.ShippedQty, 3)
            remainingQty    = 0.0
            manager         = if ($price -and $price.Manager) { $price.Manager } else { $sales.Managers }
            factoryPrice    = [math]::Round($fp,  2)
            clientPrice     = [math]::Round($cp,  2)
            actualClientPrice = [math]::Round($acp, 2)
            plannedRevenue  = 0.0
            actualRevenue   = [math]::Round($sales.Revenue, 2)
            plannedGp       = 0.0
            actualGp        = $actualGp
            margin          = $margin
            shipDate        = $sales.Dates
            priceSource     = "Факт продаж"
            operation       = $sales.Operation
            salesRows       = $sales.Rows
        })
        Write-Host "  $monthName | $skuNorm | вне плана отгружено=$([math]::Round($sales.ShippedQty,0))"
    }
}

Build-MonthRows "2026-05" "Май"   $mayPriceMap
Build-MonthRows "2026-06" "Июнь"  $junePriceMap
Write-Host "Всего строк: $($dataRows.Count)"

# ─── step 6: потенциал из Google Sheets ──────────────────────────────────────
Write-Host "=== Шаг 6: потенциал из Google Sheets ===" -ForegroundColor Cyan
$PotentialUrl = "https://docs.google.com/spreadsheets/d/1OxBfhc5Mmnfo2o_-1_k_lP8dSMWdt9FC4hLnO3E48EA/export?format=csv&gid=2092893482"
$monthMap = @{ "май"="Май"; "июнь"="Июнь"; "июль"="Июль"; "август"="Август"; "сентябрь"="Сентябрь"; "октябрь"="Октябрь"; "ноябрь"="Ноябрь"; "декабрь"="Декабрь" }
$potentialRows = [System.Collections.Generic.List[object]]::new()
try {
    $wc = New-Object System.Net.WebClient; $wc.Encoding = [System.Text.Encoding]::UTF8
    $csvText = $wc.DownloadString($PotentialUrl)
    $csvLines = $csvText -split "`n" | Where-Object { $_.Trim() -ne "" }
    $header = $csvLines[0] -split ","
    foreach ($line in ($csvLines | Select-Object -Skip 1)) {
        # RFC 4180 CSV split: respect quoted fields
        $fields = [System.Collections.Generic.List[string]]::new()
        $inQuote = $false; $cur = [System.Text.StringBuilder]::new()
        foreach ($ch in $line.ToCharArray()) {
            if ($ch -eq '"') { $inQuote = -not $inQuote }
            elseif ($ch -eq ',' -and -not $inQuote) { $fields.Add($cur.ToString()); $cur = [System.Text.StringBuilder]::new() }
            else { $cur.Append($ch) | Out-Null }
        }
        $fields.Add($cur.ToString())
        $monthRaw = $fields[0].Trim().ToLower()
        $monthRu  = if ($monthMap.ContainsKey($monthRaw)) { $monthMap[$monthRaw] } else { continue }
        $potentialRows.Add([ordered]@{
            month        = $monthRu
            sku          = $fields[1].Trim()
            name         = $fields[2].Trim()
            client       = $fields[3].Trim()
            manager      = $fields[4].Trim()
            qty          = To-Num $fields[5]
            factoryPrice = To-Num $fields[6]
            clientPrice  = To-Num $fields[7]
            revenue      = To-Num $fields[8]
            gp           = To-Num $fields[9]
            margin       = To-Num $fields[10]
            status       = if ($fields.Count -gt 11) { $fields[11].Trim() } else { "" }
            comment      = if ($fields.Count -gt 12) { $fields[12].Trim() } else { "" }
            source       = "Ручной потенциал"
        })
    }
    Write-Host "  Потенциал: $($potentialRows.Count) строк"
} catch {
    Write-Warning "Не удалось загрузить потенциал из Google Sheets: $_"
}

# ─── step 7: inject into index.html ──────────────────────────────────────────
Write-Host "=== Шаг 7: обновляем index.html ===" -ForegroundColor Cyan
$jsonStr      = ($dataRows.ToArray()      | ConvertTo-Json -Compress -Depth 5)
$jsonPotential = ($potentialRows.ToArray() | ConvertTo-Json -Compress -Depth 5)
$html = Get-Content -LiteralPath $IndexPath -Raw -Encoding UTF8
if ($html -notmatch '(?s)const DATA = \[.*?\];') { throw "Паттерн 'const DATA = [...]' не найден в index.html" }
$newHtml = $html -replace '(?s)const DATA = \[.*?\];', "const DATA = $jsonStr;"
if ($potentialRows.Count -gt 0 -and $newHtml -match '(?s)const MANUAL_POTENTIAL = \[.*?\];') {
    $newHtml = $newHtml -replace '(?s)const MANUAL_POTENTIAL = \[.*?\];', "const MANUAL_POTENTIAL = $jsonPotential;"
}
if ($newHtml -ceq $html) {
    Write-Host "Данные не изменились, файл не перезаписан | $(Get-Date -Format 'dd.MM.yyyy HH:mm')" -ForegroundColor Yellow
} else {
    [System.IO.File]::WriteAllText($IndexPath, $newHtml, [System.Text.UTF8Encoding]::new($false))
    Write-Host "Готово! DATA: $($dataRows.Count) строк, Потенциал: $($potentialRows.Count) строк | $(Get-Date -Format 'dd.MM.yyyy HH:mm')" -ForegroundColor Green
}
