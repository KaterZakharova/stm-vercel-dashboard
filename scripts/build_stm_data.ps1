$ErrorActionPreference = "Stop"

$ScriptDir  = Split-Path $MyInvocation.MyCommand.Path -Parent
$ProjectDir = Split-Path $ScriptDir -Parent
$IndexPath  = Join-Path $ProjectDir "index.html"

$KeysPath      = "C:\Users\eklementeva\.codex\skills\odata-proxy\keys.env"
$DataDir       = Join-Path $ProjectDir "data"
$MatchedCsv    = Join-Path $DataDir "stm_matched_orders.csv"
$ClassifierCsv = Join-Path $DataDir "1c_classifier_probe.csv"
# Источники плана: ЦБ-00000218 — Май; ЦБ-00000220 — Июнь/Июль/Август (создан 29.05.2026)
$PlanDocNumbers = @("ЦБ-00000218", "ЦБ-00000220")

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
    $maxAttempts = 5
    $attempt = 0
    while ($true) {
        $attempt++
        try {
            return Invoke-RestMethod -Method Get -Uri $uri -Headers $ODataHeaders -TimeoutSec $timeout
        } catch {
            $msg = $_.Exception.Message
            $isUpstream = $msg -match 'upstream error' -or $msg -match '50\d' -or $msg -match 'timeout' -or $msg -match 'timed out'
            if (-not $isUpstream -or $attempt -ge $maxAttempts) { throw }
            $delay = [Math]::Min(60, [Math]::Pow(2, $attempt) * 2)
            Write-Host "  ! upstream error на $entity (попытка $attempt/$maxAttempts), retry через $delay сек..." -ForegroundColor Yellow
            Start-Sleep -Seconds $delay
        }
    }
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

# ─── 1С price resolvers: цена фабрики (прайс-лист) + цена клиента (Заказ давальца) ───
$FactoryPriceTypeKey = "70152faf-6704-11ec-b0bf-00155d640300"  # ВидЦены "Цена Фабрики"
$factoryPriceCache = @{}
function Get-OneCFactoryPrice($nomRefKey) {
    $k = ([string]$nomRefKey).Trim()
    if ($k -eq "" -or $k -eq "00000000-0000-0000-0000-000000000000") { return 0.0 }
    if ($factoryPriceCache.ContainsKey($k)) { return $factoryPriceCache[$k] }
    $f = [uri]::EscapeDataString("Номенклатура_Key eq guid'$k' and ВидЦены_Key eq guid'$FactoryPriceTypeKey'")
    try {
        $r = Invoke-OData "InformationRegister_ЦеныНоменклатуры_RecordType" "`$select=Period,Цена&`$filter=$f&`$orderby=Period desc&`$top=10"
        $val = 0.0
        foreach ($rec in @($r.value)) { $p = To-Num $rec.Цена; if ($p -gt 0) { $val = $p; break } }
        $factoryPriceCache[$k] = $val; return $val
    } catch { $factoryPriceCache[$k] = 0.0; return 0.0 }
}

$clientPriceCache = @{}
function Get-OneCClientPrice($nomRefKey) {
    $k = ([string]$nomRefKey).Trim()
    if ($k -eq "" -or $k -eq "00000000-0000-0000-0000-000000000000") { return 0.0 }
    if ($clientPriceCache.ContainsKey($k)) { return $clientPriceCache[$k] }
    # Цена клиента = последний оформленный Заказ давальца (БЕЗ НДС).
    # У нас два типа документа: Document_ЗаказДавальца (старый) и Document_ЗаказДавальца2_5 (новый, 2026+).
    # Все заказы Мая 2026 — в 2_5. Ищем сначала в нём, потом в старом.
    # У обоих ЦенаВключаетНДС=True, поэтому net = (СуммаСНДС - СуммаНДС) / Количество.
    $f = [uri]::EscapeDataString("Номенклатура_Key eq guid'$k' and Отменено eq false")
    $val = 0.0
    foreach ($docTable in @("Document_ЗаказДавальца2_5_Продукция", "Document_ЗаказДавальца_Продукция")) {
        try {
            $r = Invoke-OData $docTable "`$select=Цена,Количество,СуммаНДС,СуммаСНДС,ДатаОтгрузки&`$filter=$f&`$orderby=ДатаОтгрузки desc&`$top=1"
            if (@($r.value).Count -gt 0) {
                $line = $r.value[0]
                $qty   = To-Num $line.Количество
                $gross = To-Num $line.СуммаСНДС
                $vat   = To-Num $line.СуммаНДС
                if ($qty -gt 0 -and $gross -gt 0) {
                    $val = ($gross - $vat) / $qty
                } else {
                    $val = (To-Num $line.Цена) / 1.2
                }
                if ($val -gt 0) { break }
            }
        } catch {}
    }
    $clientPriceCache[$k] = $val; return $val
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
Write-Host "=== Шаг 1: план производства ($($PlanDocNumbers -join ', ')) ===" -ForegroundColor Cyan
$planByMonth = @{ "2026-05" = @{}; "2026-06" = @{} }
foreach ($planNum in $PlanDocNumbers) {
    $safe = Escape-OData $planNum
    # ищем по подстроке (Number eq иногда падает с AUTOORDER ошибкой), затем дочитываем по Ref_Key
    $tail = ($planNum -replace '[^\d]', '')
    $f = [uri]::EscapeDataString("substringof('$tail',Number) and DeletionMark eq false")
    $listResp = Invoke-OData "Document_ПланПроизводства" "`$filter=$f&`$select=Ref_Key,Number" 180
    $match = @($listResp.value) | Where-Object { $_.Number -eq $planNum } | Select-Object -First 1
    if (-not $match) {
        Write-Warning "Документ ПланПроизводства '$planNum' не найден — пропускаю"
        continue
    }
    # читаем сам документ с табличной частью Продукция через entity-key URL
    $planDoc = Invoke-RestMethod -Method Get -Uri "$($OData.Url)/odata/Document_ПланПроизводства(guid'$($match.Ref_Key)')?`$format=json" -Headers $ODataHeaders -TimeoutSec 180
    $added = @{ "2026-05" = 0; "2026-06" = 0 }
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
        $added[$monthCode] += 1
    }
    Write-Host "  $planNum : +$($added['2026-05']) строк Мая, +$($added['2026-06']) строк Июня"
}
Write-Host "  Итого Май: $($planByMonth['2026-05'].Count) SKU, Июнь: $($planByMonth['2026-06'].Count) SKU"

# ─── step 2: actual release (Выпуск продукции) ───────────────────────────────
Write-Host "=== Шаг 2: фактический выпуск (ВыпускПродукции) ===" -ForegroundColor Cyan
$relFilter = [uri]::EscapeDataString(
    "Period ge datetime'$MayStart' and Period lt datetime'$JuneEnd' and Active eq true"
)
$relRecords = @()
try {
    $relRecords = Invoke-ODataPaged "AccumulationRegister_ВыпускПродукции_RecordType" $relFilter "Period,АналитикаУчетаНоменклатуры_Key,Количество,Recorder,Recorder_Type"
    Write-Host "  Получено записей: $($relRecords.Count)"
} catch {
    Write-Warning "Не удалось загрузить ВыпускПродукции: $_"
}

# aggregate by analytics key first (minimizes catalog lookups)
$releaseByAnalyticsKey = @{ "2026-05" = @{}; "2026-06" = @{} }
# also track latest Recorder (Этап производства) per analytics key per month — нужно для трассировки до Заказа давальца
$latestRecorderByAk = @{ "2026-05" = @{}; "2026-06" = @{} }
foreach ($rec in $relRecords) {
    $dt = try { [datetime]$rec.Period } catch { continue }
    $mc = $dt.ToString("yyyy-MM")
    if ($mc -notin @("2026-05", "2026-06")) { continue }
    $ak = [string]$rec.АналитикаУчетаНоменклатуры_Key
    if (-not $releaseByAnalyticsKey[$mc].ContainsKey($ak)) { $releaseByAnalyticsKey[$mc][$ak] = 0.0 }
    $releaseByAnalyticsKey[$mc][$ak] += To-Num $rec.Количество
    # сохраняем наиболее свежий регистратор (Этап производства) для этой аналитики
    $cur = if ($latestRecorderByAk[$mc].ContainsKey($ak)) { $latestRecorderByAk[$mc][$ak] } else { $null }
    if ($null -eq $cur -or [datetime]$rec.Period -gt [datetime]$cur.Period) {
        $latestRecorderByAk[$mc][$ak] = [pscustomobject]@{ Period = $rec.Period; Recorder = [string]$rec.Recorder; RecorderType = [string]$rec.Recorder_Type }
    }
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

# ─── step 2.6: цепочка выпуск → Этап → Заказ-на-произв. → Заказ давальца2_5 ──
# Для каждой записи выпуска находим конкретный Заказ давальца2_5, по которому
# был сделан выпуск. Это важно, потому что один и тот же SKU может быть в
# нескольких заказах с разной комплектацией и ценой.
Write-Host "=== Шаг 2.6: цепочка выпуск → Заказ давальца2_5 ===" -ForegroundColor Cyan
$etapToZakazProd = @{}     # Recorder (Этап) Ref → Заказ на производство Ref
$zakazProdToBase = @{}     # Заказ на производство Ref → @{ OrderRef, OrderType }
$skuToOrderRef = @{ "2026-05" = @{}; "2026-06" = @{} }  # skuNorm → ЗаказДавальца Ref (с типом)

# собираем уникальные Recorder Refs из всех месяцев
$uniqueRecorders = @{}
foreach ($mc in @("2026-05", "2026-06")) {
    foreach ($ak in $latestRecorderByAk[$mc].Keys) {
        $r = $latestRecorderByAk[$mc][$ak]
        if ($r.Recorder -and $r.RecorderType -like "*ЭтапПроизводства*") {
            $uniqueRecorders[$r.Recorder] = $true
        }
    }
}
Write-Host "  Уникальных Этапов производства: $($uniqueRecorders.Count)"

# Для каждого Этапа — получить Распоряжение_Key (Заказ на производство)
foreach ($etapRef in $uniqueRecorders.Keys) {
    try {
        $f = [uri]::EscapeDataString("Ref_Key eq guid'$etapRef'")
        $r = Invoke-OData "Document_ЭтапПроизводства2_2" "`$select=Распоряжение_Key&`$filter=$f&`$top=1"
        if (@($r.value).Count -gt 0) {
            $rasporKey = [string]$r.value[0].Распоряжение_Key
            if ($rasporKey -and $rasporKey -ne "00000000-0000-0000-0000-000000000000") {
                $etapToZakazProd[$etapRef] = $rasporKey
            }
        }
    } catch {}
}
Write-Host "  Этапов с Распоряжением: $($etapToZakazProd.Count)"

# Для каждого уникального Заказа на производство — получить ДокументОснование (Заказ давальца)
$uniqueZakazProd = @{}
foreach ($v in $etapToZakazProd.Values) { $uniqueZakazProd[$v] = $true }
foreach ($zpRef in $uniqueZakazProd.Keys) {
    try {
        $f = [uri]::EscapeDataString("Ref_Key eq guid'$zpRef'")
        $r = Invoke-OData "Document_ЗаказНаПроизводство2_2" "`$select=ДокументОснование,ДокументОснование_Type&`$filter=$f&`$top=1"
        if (@($r.value).Count -gt 0) {
            $baseRef = [string]$r.value[0].ДокументОснование
            $baseType = [string]$r.value[0].ДокументОснование_Type
            if ($baseRef -and $baseRef -ne "00000000-0000-0000-0000-000000000000") {
                $zakazProdToBase[$zpRef] = [pscustomobject]@{ Ref = $baseRef; Type = $baseType }
            }
        }
    } catch {}
}
$withZakazDav = ($zakazProdToBase.Values | Where-Object { $_.Type -match "ЗаказДавальца" }).Count
Write-Host "  Заказов на производство с Заказом давальца: $withZakazDav из $($zakazProdToBase.Count)"

# Собираем SkuNorm → Ref Заказа давальца (по последнему выпуску в месяце)
foreach ($mc in @("2026-05", "2026-06")) {
    foreach ($ak in $latestRecorderByAk[$mc].Keys) {
        $rec = $latestRecorderByAk[$mc][$ak]
        $etapRef = $rec.Recorder
        if (-not $etapToZakazProd.ContainsKey($etapRef)) { continue }
        $zpRef = $etapToZakazProd[$etapRef]
        if (-not $zakazProdToBase.ContainsKey($zpRef)) { continue }
        $base = $zakazProdToBase[$zpRef]
        if ($base.Type -notmatch "ЗаказДавальца") { continue }
        # разрешаем аналитику → skuNorm
        $nomRefKey = Get-NomFromAnalyticsKey $ak; if (-not $nomRefKey) { continue }
        $n = Get-NomByRef $nomRefKey; if (-not $n) { continue }
        $skuNorm = Normalize-Sku $n.Артикул
        $skuToOrderRef[$mc][$skuNorm] = $base
    }
    Write-Host "  $mc : SKU с разрешённым заказом давальца: $($skuToOrderRef[$mc].Count)"
}

# Получаем имя контрагента (юр. лица заказчика) для каждого уникального заказа давальца
$orderToCounterpartyName = @{}
$counterpartyCache = @{}
$uniqueOrders = @{}
foreach ($mc in @("2026-05", "2026-06")) {
    foreach ($oi in $skuToOrderRef[$mc].Values) {
        $uniqueOrders["$($oi.Type)|$($oi.Ref)"] = $oi
    }
}
foreach ($entry in $uniqueOrders.GetEnumerator()) {
    $oi = $entry.Value
    $docTable = if ($oi.Type -match "ЗаказДавальца2_5") { "Document_ЗаказДавальца2_5" }
                elseif ($oi.Type -match "ЗаказДавальца") { "Document_ЗаказДавальца" }
                else { $null }
    if (-not $docTable) { continue }
    try {
        $f = [uri]::EscapeDataString("Ref_Key eq guid'$($oi.Ref)'")
        $r = Invoke-OData $docTable "`$select=Контрагент_Key&`$filter=$f&`$top=1"
        if (@($r.value).Count -gt 0) {
            $cpRef = [string]$r.value[0].Контрагент_Key
            if ($cpRef -and $cpRef -ne "00000000-0000-0000-0000-000000000000") {
                if (-not $counterpartyCache.ContainsKey($cpRef)) {
                    try {
                        $rf = [uri]::EscapeDataString("Ref_Key eq guid'$cpRef'")
                        $rc = Invoke-OData "Catalog_Контрагенты" "`$select=Description&`$filter=$rf&`$top=1"
                        if (@($rc.value).Count -gt 0) {
                            $counterpartyCache[$cpRef] = [string]$rc.value[0].Description
                        }
                    } catch {}
                }
                if ($counterpartyCache.ContainsKey($cpRef)) {
                    $orderToCounterpartyName[$oi.Ref] = $counterpartyCache[$cpRef]
                }
            }
        }
    } catch {}
}
Write-Host "  Заказов давальца с контрагентом: $($orderToCounterpartyName.Count)"

# Функция: получить цену клиента БЕЗ НДС из конкретного Заказа давальца по конкретной номенклатуре
$orderPriceCache = @{}
function Get-PriceFromOrder($orderRef, $orderType, $nomRefKey) {
    $cacheKey = "$orderRef|$nomRefKey"
    if ($orderPriceCache.ContainsKey($cacheKey)) { return $orderPriceCache[$cacheKey] }
    $val = 0.0
    $table = if ($orderType -match "ЗаказДавальца2_5") { "Document_ЗаказДавальца2_5_Продукция" }
             elseif ($orderType -match "ЗаказДавальца") { "Document_ЗаказДавальца_Продукция" }
             else { $null }
    if (-not $table) { $orderPriceCache[$cacheKey] = 0.0; return 0.0 }
    $f = [uri]::EscapeDataString("Ref_Key eq guid'$orderRef' and Номенклатура_Key eq guid'$nomRefKey' and Отменено eq false")
    try {
        $r = Invoke-OData $table "`$select=Цена,Количество,СуммаНДС,СуммаСНДС&`$filter=$f&`$top=1"
        if (@($r.value).Count -gt 0) {
            $line = $r.value[0]
            $qty = To-Num $line.Количество; $gross = To-Num $line.СуммаСНДС; $vat = To-Num $line.СуммаНДС
            if ($qty -gt 0 -and $gross -gt 0) { $val = ($gross - $vat) / $qty }
            else { $val = (To-Num $line.Цена) / 1.2 }
        }
    } catch {}
    $orderPriceCache[$cacheKey] = $val; return $val
}

# ─── step 2.5: реальные остатки на складах СТМ и ПГП ──────────────────────────
# Берём ВНаличииBalance из AccumulationRegister_ТоварыНаСкладах/Balance (текущий момент).
# Склад СТМ        = 391dda2d-1423-11ee-b0cd-00155d640e00 (основной склад готовой продукции СТМ)
# Склад ПГП        = c371ec92-aa22-11ea-b087-00155d640300 (полуфабрикаты готовой продукции)
Write-Host "=== Шаг 2.5: остатки на складах СТМ и ПГП (Balance) ===" -ForegroundColor Cyan
$StmWhKey = '391dda2d-1423-11ee-b0cd-00155d640e00'
$PgpWhKey = 'c371ec92-aa22-11ea-b087-00155d640300'
$stockStmByNomRef = @{}
$stockPgpByNomRef = @{}
foreach ($whKey in @($StmWhKey, $PgpWhKey)) {
    $whName = if ($whKey -eq $StmWhKey) { 'СТМ' } else { 'ПГП' }
    $target = if ($whKey -eq $StmWhKey) { $stockStmByNomRef } else { $stockPgpByNomRef }
    $filter = [uri]::EscapeDataString("Склад_Key eq guid'$whKey'")
    try {
        $r = Invoke-OData "AccumulationRegister_ТоварыНаСкладах/Balance" "`$select=Номенклатура_Key,ВНаличииBalance&`$filter=$filter" 240
        foreach ($rec in @($r.value)) {
            $nk = [string]$rec.Номенклатура_Key
            if (-not $nk -or $nk -eq '00000000-0000-0000-0000-000000000000') { continue }
            $q = To-Num $rec.ВНаличииBalance
            if (-not $target.ContainsKey($nk)) { $target[$nk] = 0.0 }
            $target[$nk] += $q
        }
        Write-Host "  Склад $whName : $($target.Count) SKU с положительным остатком"
    } catch {
        Write-Warning "Не удалось получить остатки склада $whName : $_"
    }
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

function Get-MayCarryTag($skuNorm) {
    if (-not $script:mayCarryByNorm -or -not $script:mayCarryByNorm.ContainsKey($skuNorm)) { return "" }
    $c = $script:mayCarryByNorm[$skuNorm]
    $parts = @()
    if ($c.Deficit  -gt 0) { $parts += "ушло из мая: -$($c.Deficit) шт" }
    if ($c.Leftover -gt 0) { $parts += "выпущено в мае: +$($c.Leftover) шт" }
    return ($parts -join "; ")
}

function Build-MonthRows($monthCode, $monthName, $priceMap) {
    $planMonth    = $planByMonth[$monthCode]
    $releaseMonth = $releaseBySkuNorm[$monthCode]
    $mStart = if ($monthCode -eq "2026-05") { $MayStart  } else { $JuneStart }
    $mEnd   = if ($monthCode -eq "2026-05") { $JuneStart } else { $JuneEnd   }
    $included = @{}

    # rows from production plan — план ЦБ-00000218 сам источник истины по СТМ,
    # не отсеиваем по статическому классификатору (он неполный/устаревший)
    foreach ($skuNorm in ($planMonth.Keys | Sort-Object)) {
        $included[$skuNorm] = $true
        $pd  = $planMonth[$skuNorm]
        $rel = if ($releaseMonth.ContainsKey($skuNorm)) { $releaseMonth[$skuNorm] } else { 0.0 }
        $price = $priceMap[$skuNorm]
        $sales = Get-SalesFact $skuNorm $mStart $mEnd

        # factory price: ПЛАНОВАЯ цена фабрики (как у куба аналитики).
        # Cost из 1С (cost/qty по реальным отгрузкам) НЕ используется в валовке —
        # он зачастую недостоверен (битые проводки) и не отражает плановую себестоимость.
        # Приоритеты: CSV plan → классификатор → прайс-лист 1С ("Цена Фабрики") → fp_1c (last resort)
        $fp_1c  = if ($sales.ShippedQty -gt 0 -and $sales.Cost -gt 0) { $sales.Cost / $sales.ShippedQty } else { 0.0 }
        $fp_csv = if ($price) { [double]$price.FactoryPrice } else { 0.0 }
        if ($fp_csv -eq 0 -and $classPrices.ContainsKey($skuNorm)) { $fp_csv = [double]$classPrices[$skuNorm] }
        $fp = if ($fp_csv -gt 0) { $fp_csv } else { 0.0 }
        if ($fp -eq 0) { $pf = Get-OneCFactoryPrice $pd.NomRef; if ($pf -gt 0) { $fp = $pf } }
        if ($fp -eq 0 -and $fp_1c -gt 0) { $fp = $fp_1c }   # last resort

        # client price — приоритет:
        # 1) order-specific (трассировка выпуск → Заказ давальца2_5) — точная цена комплектации
        # 2) CSV (плановая, но может быть от старого заказа)
        # 3) Get-OneCClientPrice (последний оформленный заказ 2_5 или старый)
        # 4) фактическая по продажам (acp) — на случай если нет ничего планового
        $cp_order = 0.0; $cpFromOrder = $false
        if ($skuToOrderRef[$monthCode].ContainsKey($skuNorm)) {
            $oi = $skuToOrderRef[$monthCode][$skuNorm]
            $cp_order = Get-PriceFromOrder $oi.Ref $oi.Type $pd.NomRef
            if ($cp_order -gt 0) { $cpFromOrder = $true }
        }
        $acp = if ($sales.ShippedQty -ne 0) { $sales.Revenue / $sales.ShippedQty } else { 0.0 }
        $cp_csv = if ($price) { [double]$price.ClientPrice } else { 0.0 }
        $cpFromDavalec = $false
        $cp_fallback = 0.0
        if ($cp_order -eq 0 -and $cp_csv -eq 0) {
            $cp_fallback = Get-OneCClientPrice $pd.NomRef
            if ($cp_fallback -gt 0) { $cpFromDavalec = $true }
        }
        # эффективная цена клиента (для отображения, выручки факт)
        $cp = if ($acp -ne 0) { $acp } `
              elseif ($cp_order -gt 0) { $cp_order } `
              elseif ($cp_csv -gt 0) { $cp_csv } `
              else { $cp_fallback }

        $planQty        = [math]::Round($pd.PlanQty, 3)
        $releaseQty     = [math]::Round([math]::Max($rel, 0), 3)
        # plan client price — точная цена из конкретного заказа давальца, привязанного к выпуску
        $pcPlan         = if ($cp_order -gt 0) { $cp_order } `
                          elseif ($cp_csv -gt 0) { $cp_csv } `
                          elseif ($cp_fallback -gt 0) { $cp_fallback } `
                          else { $cp }
        $hasSellPrice   = ($pcPlan -gt 0)
        $plannedRevenue = if ($hasSellPrice) { [math]::Round($planQty * $pcPlan, 2) } else { 0.0 }
        $plannedGp      = if ($hasSellPrice) { [math]::Round($planQty * ($pcPlan - $fp), 2) } else { 0.0 }
        # actualGp = плановая валовка по факт. отгруженным шт (как считает куб):
        # ShippedQty × (планЦенаКлиента − ценаФабрики). Совпадает с отчётом аналитиков.
        $actualGp = if ($sales.ShippedQty -gt 0 -and $hasSellPrice) {
            [math]::Round($sales.ShippedQty * ($pcPlan - $fp), 2)
        } else { 0.0 }
        $marginBase = if ($sales.ShippedQty -gt 0 -and $hasSellPrice) { $sales.ShippedQty * $pcPlan } else { 0 }
        $margin = if ($marginBase -ne 0) { [math]::Round($actualGp / $marginBase, 4) } else { 0 }

        $ps = if ($cpFromOrder) { "Заказ давальца → выпуск" } `
              elseif ($cpFromDavalec) { "Заказ давальца (последний)" } `
              elseif (-not $hasSellPrice) { "Цена не найдена" } `
              elseif ($price) { if ($price.IsMonthMatch) { "Заказы месяца" } else { "Заказы другого месяца" } } `
              elseif ($fp_csv -gt 0) { "CSV/классификатор" } `
              elseif ($fp -gt 0) { "Прайс-лист 1С" } `
              else { "Факт 1С (last resort)" }

        $stmStockQty = if ($stockStmByNomRef.ContainsKey([string]$pd.NomRef)) { [math]::Round($stockStmByNomRef[[string]$pd.NomRef], 0) } else { 0 }
        $pgpStockQty = if ($stockPgpByNomRef.ContainsKey([string]$pd.NomRef)) { [math]::Round($stockPgpByNomRef[[string]$pd.NomRef], 0) } else { 0 }
        $counterparty = ""
        if ($skuToOrderRef[$monthCode].ContainsKey($skuNorm)) {
            $oref = $skuToOrderRef[$monthCode][$skuNorm].Ref
            if ($orderToCounterpartyName.ContainsKey($oref)) { $counterparty = $orderToCounterpartyName[$oref] }
        }
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
            stockStm        = $stmStockQty
            stockPgp        = $pgpStockQty
            counterparty    = $counterparty
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
            mayCarry        = if ($monthCode -eq "2026-06") { Get-MayCarryTag $skuNorm } else { "" }
        })
        Write-Host "  $monthName | $($pd.Sku) | план=$planQty выпуск=$releaseQty отгружено=$([math]::Round($sales.ShippedQty,0)) СТМ=$stmStockQty ПГП=$pgpStockQty"
    }

    # rows with actual sales but not in plan
    foreach ($skuNorm in $priceMap.Keys) {
        if ($included.ContainsKey($skuNorm)) { continue }
        if (-not $stmSkus.ContainsKey($skuNorm)) { continue }
        $sales = Get-SalesFact $skuNorm $mStart $mEnd
        if ($sales.ShippedQty -eq 0 -and $sales.Revenue -eq 0) { continue }
        $included[$skuNorm] = $true
        $price = $priceMap[$skuNorm]
        # factory price — ТОЛЬКО плановая (CSV/классиф/прайс-лист), cost из 1С — last resort
        $fp_1c  = if ($sales.ShippedQty -gt 0 -and $sales.Cost -gt 0) { $sales.Cost / $sales.ShippedQty } else { 0.0 }
        $fp_csv = if ($price) { [double]$price.FactoryPrice } else { 0.0 }
        if ($fp_csv -eq 0 -and $classPrices.ContainsKey($skuNorm)) { $fp_csv = [double]$classPrices[$skuNorm] }
        $fp = if ($fp_csv -gt 0) { $fp_csv } else { 0.0 }
        if ($fp -eq 0) {
            $n_tmp = Get-NomByArticle $skuNorm
            if ($n_tmp) { $pf = Get-OneCFactoryPrice $n_tmp.Ref_Key; if ($pf -gt 0) { $fp = $pf } }
        }
        if ($fp -eq 0 -and $fp_1c -gt 0) { $fp = $fp_1c }   # last resort
        # client price: трассировка → CSV → fallback по последнему заказу
        $cp_order = 0.0
        if ($skuToOrderRef[$monthCode].ContainsKey($skuNorm)) {
            $n_tmp2 = Get-NomByArticle $skuNorm
            if ($n_tmp2) {
                $oi = $skuToOrderRef[$monthCode][$skuNorm]
                $cp_order = Get-PriceFromOrder $oi.Ref $oi.Type $n_tmp2.Ref_Key
            }
        }
        $acp = if ($sales.ShippedQty -ne 0) { $sales.Revenue / $sales.ShippedQty } else { 0.0 }
        $cp_csv = if ($price) { [double]$price.ClientPrice } else { 0.0 }
        $cp_fallback = 0.0
        if ($cp_order -eq 0 -and $cp_csv -eq 0) {
            $n_tmp3 = Get-NomByArticle $skuNorm
            if ($n_tmp3) { $cp_fallback = Get-OneCClientPrice $n_tmp3.Ref_Key }
        }
        $cp = if ($acp -ne 0) { $acp } `
              elseif ($cp_order -gt 0) { $cp_order } `
              elseif ($cp_csv -gt 0) { $cp_csv } `
              else { $cp_fallback }
        $pcPlan = if ($cp_order -gt 0) { $cp_order } `
                  elseif ($cp_csv -gt 0) { $cp_csv } `
                  elseif ($cp_fallback -gt 0) { $cp_fallback } `
                  else { $cp }
        $actualGp = if ($sales.ShippedQty -gt 0 -and $pcPlan -gt 0 -and $fp -gt 0) {
            [math]::Round($sales.ShippedQty * ($pcPlan - $fp), 2)
        } else { 0.0 }
        $marginBase = if ($sales.ShippedQty -gt 0 -and $pcPlan -gt 0) { $sales.ShippedQty * $pcPlan } else { 0 }
        $margin = if ($marginBase -ne 0) { [math]::Round($actualGp / $marginBase, 4) } else { 0 }
        $n = Get-NomByArticle $skuNorm
        $name = if ($price -and $price.Nomenclature) { $price.Nomenclature } elseif ($n) { $n.Description } else { "" }
        $nrk = if ($n) { [string]$n.Ref_Key } else { "" }
        $stmStockQty = if ($nrk -and $stockStmByNomRef.ContainsKey($nrk)) { [math]::Round($stockStmByNomRef[$nrk], 0) } else { 0 }
        $pgpStockQty = if ($nrk -and $stockPgpByNomRef.ContainsKey($nrk)) { [math]::Round($stockPgpByNomRef[$nrk], 0) } else { 0 }
        $counterparty = ""
        if ($skuToOrderRef[$monthCode].ContainsKey($skuNorm)) {
            $oref = $skuToOrderRef[$monthCode][$skuNorm].Ref
            if ($orderToCounterpartyName.ContainsKey($oref)) { $counterparty = $orderToCounterpartyName[$oref] }
        }

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
            stockStm        = $stmStockQty
            stockPgp        = $pgpStockQty
            counterparty    = $counterparty
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
            priceSource     = if ($cp_order -gt 0) { "Заказ давальца → выпуск" } elseif ($fp_csv -gt 0) { "Факт продаж (план CSV)" } elseif ($fp -gt 0) { "Факт продаж (прайс 1С)" } else { "Факт продаж" }
            operation       = $sales.Operation
            salesRows       = $sales.Rows
            mayCarry        = if ($monthCode -eq "2026-06") { Get-MayCarryTag $skuNorm } else { "" }
        })
        Write-Host "  $monthName | $skuNorm | вне плана отгружено=$([math]::Round($sales.ShippedQty,0))"
    }

    # carry-over из Мая для Июня — SKU с дефицитом/остатком, которых нет ни в плане Июня,
    # ни среди фактических продаж Июня. Показываем со statQty=0, чтобы пользователь видел
    # позиции, которые сами «вылезли» из Мая и требуют решения.
    if ($monthCode -eq "2026-06" -and $script:mayCarryByNorm) {
        foreach ($skuNorm in $script:mayCarryByNorm.Keys) {
            if ($included.ContainsKey($skuNorm)) { continue }
            $n = Get-NomByArticle $skuNorm
            if (-not $n) { continue }
            $nrk = [string]$n.Ref_Key
            $stmQ = if ($stockStmByNomRef.ContainsKey($nrk)) { [math]::Round($stockStmByNomRef[$nrk], 0) } else { 0 }
            $pgpQ = if ($stockPgpByNomRef.ContainsKey($nrk)) { [math]::Round($stockPgpByNomRef[$nrk], 0) } else { 0 }
            $tag = Get-MayCarryTag $skuNorm
            $price = $priceMap[$skuNorm]
            $script:dataRows.Add([ordered]@{
                month           = $monthName
                sku             = $skuNorm
                isNew           = if ($price) { $price.IsNew } else { "" }
                name            = [string]$n.Description
                order           = if ($price) { $price.Orders } else { "" }
                planQty         = 0.0
                releaseQty      = 0.0
                shippedQty      = 0.0
                remainingQty    = 0.0
                stockStm        = $stmQ
                stockPgp        = $pgpQ
                counterparty    = ""
                manager         = if ($price -and $price.Manager) { $price.Manager } else { "" }
                factoryPrice    = 0.0
                clientPrice     = 0.0
                actualClientPrice = 0.0
                plannedRevenue  = 0.0
                actualRevenue   = 0.0
                plannedGp       = 0.0
                actualGp        = 0.0
                margin          = 0
                shipDate        = ""
                priceSource     = "carry-over из Мая"
                operation       = ""
                salesRows       = 0
                mayCarry        = $tag
            })
            $included[$skuNorm] = $true
            Write-Host "  $monthName | $skuNorm | carry-over: $tag"
        }
    }
}

Build-MonthRows "2026-05" "Май"   $mayPriceMap

# carry-over из Мая: для каждого SKU считаем дефицит производства и складской остаток,
# который пойдёт в Июнь
$script:mayCarryByNorm = @{}
foreach ($r in $script:dataRows) {
    if ($r.month -ne "Май") { continue }
    $norm = Normalize-Sku $r.sku
    $plan = [double]$r.planQty
    $rel  = [double]$r.releaseQty
    $ship = [double]$r.shippedQty
    $stockTotal = ([double]$r.stockStm) + ([double]$r.stockPgp)
    $deficit  = [math]::Max(0.0, $plan - $rel)
    # «выпущено в мае, но не ушло» — берём фактический складской остаток (он надёжнее, чем release-ship,
    # потому что часть могла быть списана/перемещена)
    $leftover = [math]::Round($stockTotal, 0)
    if ($deficit -gt 0 -or $leftover -gt 0) {
        $script:mayCarryByNorm[$norm] = [pscustomobject]@{
            Deficit  = [math]::Round($deficit, 0)
            Leftover = $leftover
            MayPlan  = $plan
            MayRel   = $rel
            MayShip  = $ship
        }
    }
}
Write-Host "  Carry-over из Мая: $($script:mayCarryByNorm.Count) SKU"

Build-MonthRows "2026-06" "Июнь"  $junePriceMap
Write-Host "Всего строк: $($dataRows.Count)"

# ─── step 6: прогноз из открытых Заказов давальца2_5 ─────────────────────────
# Заказ давальца2_5 → потенциал производства/выручки. Сверяем с планом:
#   планQty ≈ orderQty (±5%) → в плане полностью → не показываем
#   планQty < orderQty       → частично в плане (uncertain=true) → показываем разницу
#   планQty = 0              → нет в плане (РС не готова) → показываем целиком
Write-Host "=== Шаг 6: прогноз из ЗаказДавальца2_5 ===" -ForegroundColor Cyan
$potentialRows = [System.Collections.Generic.List[object]]::new()
$monthNameMap = @{ "2026-05"="Май"; "2026-06"="Июнь"; "2026-07"="Июль" }
try {
    # список 2026-х заказов
    $f = [uri]::EscapeDataString("Posted eq true and DeletionMark eq false")
    $listResp = Invoke-OData "Document_ЗаказДавальца2_5" "`$filter=$f&`$select=Ref_Key,Number,Date,Контрагент_Key,Менеджер_Key&`$top=10000" 180
    $y2026 = @($listResp.value) | Where-Object { ([string]$_.Date).StartsWith("2026") }
    Write-Host "  2026 заказов давальца: $($y2026.Count)"

    $aggByKey = @{}
    $i = 0
    foreach ($order in $y2026) {
        $i++
        if ($i % 25 -eq 0) { Write-Host "    обработано $i / $($y2026.Count)" }
        try {
            $doc = Invoke-RestMethod -Method Get -Uri "$($OData.Url)/odata/Document_ЗаказДавальца2_5(guid'$($order.Ref_Key)')?`$format=json" -Headers $ODataHeaders -TimeoutSec 60
        } catch { continue }
        foreach ($line in @($doc.Продукция)) {
            if ($line.Отменено) { continue }
            $rd = try { [datetime]$line.ДатаОтгрузки } catch { $null }
            if (-not $rd -or $rd.Year -lt 1900) { continue }
            $mc = $rd.ToString("yyyy-MM")
            if (-not $monthNameMap.ContainsKey($mc)) { continue }
            $n = Get-NomByRef $line.Номенклатура_Key
            if (-not $n) { continue }
            $skuNorm = Normalize-Sku $n.Артикул
            $qty = To-Num $line.Количество
            if ($qty -le 0) { continue }
            # цена клиента без НДС
            $sumNoVat = (To-Num $line.СуммаСНДС) - (To-Num $line.СуммаНДС)
            $pricePerUnit = if ($qty -gt 0 -and $sumNoVat -gt 0) { $sumNoVat / $qty } else { 0 }

            $key = "$mc|$skuNorm"
            if (-not $aggByKey.ContainsKey($key)) {
                $aggByKey[$key] = [pscustomobject]@{
                    Month = $mc; SkuNorm = $skuNorm; Sku = [string]$n.Артикул; Name = [string]$n.Description
                    NomRef = [string]$n.Ref_Key; OrderQty = 0.0; Price = 0.0
                    OrderNumbers = @(); Managers = @(); Counterparties = @()
                }
            }
            $a = $aggByKey[$key]
            $a.OrderQty += $qty
            if ($pricePerUnit -gt 0) { $a.Price = $pricePerUnit }
            if ($a.OrderNumbers -notcontains $order.Number) { $a.OrderNumbers += $order.Number }
            $mgr = Get-ManagerName $order.Менеджер_Key
            if ($mgr -and ($a.Managers -notcontains $mgr)) { $a.Managers += $mgr }
            $cpRef = [string]$order.Контрагент_Key
            $cpName = if ($counterpartyCache.ContainsKey($cpRef)) { $counterpartyCache[$cpRef] } else {
                try {
                    $rf = [uri]::EscapeDataString("Ref_Key eq guid'$cpRef'")
                    $rc = Invoke-OData "Catalog_Контрагенты" "`$select=Description&`$filter=$rf&`$top=1"
                    if (@($rc.value).Count -gt 0) {
                        $counterpartyCache[$cpRef] = [string]$rc.value[0].Description
                        $counterpartyCache[$cpRef]
                    } else { "" }
                } catch { "" }
            }
            if ($cpName -and ($a.Counterparties -notcontains $cpName)) { $a.Counterparties += $cpName }
        }
    }
    Write-Host "  Агрегатов (sku × месяц): $($aggByKey.Count)"

    foreach ($a in $aggByKey.Values) {
        $planQty = if ($planByMonth.ContainsKey($a.Month) -and $planByMonth[$a.Month].ContainsKey($a.SkuNorm)) {
            $planByMonth[$a.Month][$a.SkuNorm].PlanQty
        } else { 0.0 }

        # допуск 5% — если plan ≈ order, считаем что план уже покрывает
        $tol = $a.OrderQty * 0.05
        if ($planQty -gt 0 -and [math]::Abs($planQty - $a.OrderQty) -le $tol) { continue }
        # план шире заказа — пропускаем (избыточное планирование, не потенциал)
        if ($planQty -ge $a.OrderQty) { continue }

        $potentialQty = $a.OrderQty - $planQty
        $uncertain = $false
        $status = if ($planQty -eq 0) { "нет в плане" } else {
            $uncertain = $true
            "частично в плане: $([math]::Round($planQty,0))/$([math]::Round($a.OrderQty,0))"
        }

        # factory price — плановая (CSV→классификатор→прайс 1С)
        $fp = 0.0
        if ($classPrices.ContainsKey($a.SkuNorm)) { $fp = [double]$classPrices[$a.SkuNorm] }
        if ($fp -eq 0) { $pf = Get-OneCFactoryPrice $a.NomRef; if ($pf -gt 0) { $fp = $pf } }

        $cp = $a.Price
        $revenue = $potentialQty * $cp
        $gp      = $potentialQty * ($cp - $fp)
        $margin  = if ($revenue -ne 0) { [math]::Round($gp / $revenue, 4) } else { 0 }

        $potentialRows.Add([ordered]@{
            month        = $monthNameMap[$a.Month]
            sku          = $a.Sku
            name         = $a.Name
            client       = ($a.Counterparties -join ", ")
            manager      = ($a.Managers -join ", ")
            qty          = [math]::Round($potentialQty, 0)
            factoryPrice = [math]::Round($fp, 2)
            clientPrice  = [math]::Round($cp, 2)
            revenue      = [math]::Round($revenue, 2)
            gp           = [math]::Round($gp, 2)
            margin       = $margin
            status       = $status
            comment      = ""
            source       = "Прогноз"
            uncertain    = $uncertain
            planQty      = [math]::Round($planQty, 0)
            orderQty     = [math]::Round($a.OrderQty, 0)
            orderRefs    = ($a.OrderNumbers -join "; ")
        })
    }
    Write-Host "  Прогноз: $($potentialRows.Count) строк ($(($potentialRows | Where-Object { $_.uncertain }).Count) с сомнением)"
} catch {
    Write-Warning "Не удалось построить прогноз: $_"
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
