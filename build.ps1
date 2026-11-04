# Bay Area multifamily CMBS debt pipeline
# Sources: SEC EDGAR ABS-EE asset-level files (exh_102.xml), FRED DGS10, Census geocoder.
# Run: powershell -ExecutionPolicy Bypass -File build.ps1 [-LookbackDays 90]
#
# Network calls use the Windows curl.exe: Invoke-WebRequest/Invoke-RestMethod hung
# against FRED from this machine, while curl.exe connects.

param(
    [int]$LookbackDays = 90
)

$ErrorActionPreference = 'Stop'
$Root  = $PSScriptRoot
$Cache = Join-Path $Root 'cache'
$Site  = Join-Path $Root 'docs'
$Tmp   = Join-Path $Root 'cache\tmp'
New-Item -ItemType Directory -Force -Path $Cache, $Site, $Tmp | Out-Null

# SEC requires a declared contact in the User-Agent (SEC fair-access policy)
$UA  = 'BayAreaDebtDashboard/0.1 (Bozzuto; jonathan.nasser@bozzuto.com)'
$Inv = [System.Globalization.CultureInfo]::InvariantCulture
$Curl = Join-Path $env:SystemRoot 'System32\curl.exe'

# ---- Criteria (from user) ----
$BayAreaCounties = @('Alameda','Contra Costa','Marin','Napa','San Francisco','San Mateo','Santa Clara','Solano','Sonoma')
$MinBalance      = 10000000
$MaxBalance      = 400000000
$BenchmarkSpread = 0.0225   # current market rate = 10yr UST + 225 bps
$SpreadThreshold = 0.03     # flag if refi gap >= 3.00%
$TopPercentile   = 0.90     # also flag top 10% of the Bay Area set by refi gap

function Get-File([string]$Url, [string]$OutFile) {
    # Retry transient server errors (SEC/FRED return intermittent 5xx); 404s are not retried
    for ($attempt = 1; $attempt -le 4; $attempt++) {
        $code = & $Curl -sS -L --max-time 180 -A $UA -o $OutFile -w "%{http_code}" $Url 2>$null
        if ($code -eq '200') { return }
        if ($code -notlike '5*') { break }   # 4xx (e.g. 404) is not transient
        Start-Sleep -Seconds (3 * $attempt)
    }
    throw "HTTP $code for $Url"
}
function Get-Text([string]$Url) {
    $f = Join-Path $Tmp ([guid]::NewGuid().ToString() + '.txt')
    Get-File $Url $f
    $t = Get-Content -Raw -Encoding UTF8 -Path $f
    Remove-Item $f -ErrorAction SilentlyContinue
    return $t
}
function Get-Num([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $v = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, $Inv, [ref]$v)) { return $v }
    return $null
}
function Get-Date-Mdy([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    try { return [DateTime]::ParseExact($s, 'MM-dd-yyyy', $Inv).ToString('yyyy-MM-dd') } catch { return $null }
}

# ---- 1. Benchmark: 10-year Treasury (FRED DGS10), latest non-missing value ----
Write-Host 'Fetching 10-year Treasury (FRED DGS10)...'
$fredCsv = Get-Text 'https://fred.stlouisfed.org/graph/fredgraph.csv?id=DGS10'
$last = ($fredCsv -split "`n") | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2},[\d.]+' } | Select-Object -Last 1
$dgsDate, $dgsVal = $last.Trim().Split(',')
$ust10    = [double]::Parse($dgsVal, $Inv) / 100
$benchmark = $ust10 + $BenchmarkSpread
Write-Host ("10yr UST {0} = {1:P2}; benchmark = {2:P2}" -f $dgsDate, $ust10, $benchmark)

# ---- 2. Enumerate recent ABS-EE filings via EDGAR full-text search ----
$end   = (Get-Date).ToString('yyyy-MM-dd')
$start = (Get-Date).AddDays(-$LookbackDays).ToString('yyyy-MM-dd')
$filings = @{}
$from = 0; $total = 1
do {
    $url = "https://efts.sec.gov/LATEST/search-index?q=%22Mortgage%20Trust%22&forms=ABS-EE&dateRange=custom&startdt=$start&enddt=$end&from=$from"
    $r = Get-Text $url | ConvertFrom-Json
    $total = [int]$r.hits.total.value
    foreach ($h in $r.hits.hits) {
        $s = $h._source
        if (-not $filings.ContainsKey($s.adsh)) {
            $filings[$s.adsh] = [pscustomobject]@{ adsh = $s.adsh; date = $s.file_date; ciks = @($s.ciks); names = @($s.display_names) }
        }
    }
    $from += 100
    Start-Sleep -Milliseconds 200
} while ($from -lt $total)
Write-Host "ABS-EE filings found: $($filings.Count) (search total reported: $total)"

# ---- 3. Keep the latest filing per trust ----
$trusts = @{}
foreach ($f in $filings.Values) {
    $tn = $f.names | Where-Object { $_ -match 'Trust' } | Select-Object -First 1
    if (-not $tn) { $tn = $f.names | Select-Object -First 1 }
    $key = ($tn -replace '\s*\(CIK \d+\)\s*', '').Trim()
    if (-not $trusts.ContainsKey($key) -or [string]$trusts[$key].date -lt [string]$f.date) { $trusts[$key] = $f }
}
Write-Host "Distinct trusts (latest filing each): $($trusts.Count)"

# ---- 4. Download exh_102 (asset-level) and keep Bay Area loans ----
$rows = New-Object System.Collections.Generic.List[object]
$i = 0; $missing = 0
foreach ($kv in $trusts.GetEnumerator()) {
    $i++
    $t = $kv.Value
    $xmlPath = Join-Path $Cache ("{0}.xml" -f $t.adsh)
    if (-not (Test-Path $xmlPath)) {
        # Filenames differ by filer (exh_102.xml vs ex102.xml), so read the filing index
        # and take the asset-level exhibit from it. Try the trust CIK(s) first.
        $ok = $false
        $acc = $t.adsh -replace '-', ''
        $cands = @($t.ciks) + @($t.names | ForEach-Object { if ($_ -match 'CIK (\d+)') { $Matches[1] } }) | ForEach-Object { [int64]$_ } | Select-Object -Unique
        foreach ($c in $cands) {
            $idxUrl = "https://www.sec.gov/Archives/edgar/data/$c/$acc/$($t.adsh)-index.htm"
            try {
                $idx = Get-Text $idxUrl
                $m = [regex]::Matches($idx, 'href="(/Archives/edgar/data/\d+/\d+/[^"]*?(?:ex|exh_?)102\.xml)"', 'IgnoreCase')
                if ($m.Count -gt 0) {
                    Get-File ("https://www.sec.gov" + $m[0].Groups[1].Value) $xmlPath
                    $ok = $true; break
                }
            } catch { Remove-Item $xmlPath -ErrorAction SilentlyContinue }
            Start-Sleep -Milliseconds 150
        }
        if (-not $ok) { $missing++; Write-Warning "No asset file for $($t.adsh) ($($kv.Key))"; continue }
        Start-Sleep -Milliseconds 150
    }
    [xml]$doc = Get-Content -Raw -Encoding UTF8 -Path $xmlPath
    foreach ($a in $doc.assetData.assets) {
        $props = @($a.property)
        $bay = $props | Where-Object { $_.propertyState -eq 'CA' -and $BayAreaCounties -contains $_.propertyCounty } | Select-Object -First 1
        if (-not $bay) { continue }
        $bal  = Get-Num $a.reportPeriodEndActualBalanceAmount
        $rate = Get-Num $a.reportPeriodInterestRatePercentage
        if ($null -eq $bal -or $null -eq $rate) { continue }
        $rows.Add([pscustomobject]@{
            trust           = $kv.Key
            trustCik        = [int64]$t.ciks[0]
            filingAdsh      = $t.adsh
            filedDate       = $t.date
            reportingDate   = Get-Date-Mdy $a.reportPeriodEndDate
            loanId          = [string]$a.assetNumber
            propertyType    = [string]$bay.propertyTypeCode
            propertyName    = [string]$bay.propertyName
            address         = [string]$bay.propertyAddress
            city            = [string]$bay.propertyCity
            county          = [string]$bay.propertyCounty
            zip             = [string]$bay.propertyZip
            units           = Get-Num $bay.unitsBedsRoomsNumber
            yearBuilt       = Get-Num $bay.yearBuiltNumber
            sqft            = Get-Num $bay.netRentableSquareFeetNumber
            propertyCount   = $props.Count
            balance         = $bal
            couponRate      = $rate
            maturityDate    = Get-Date-Mdy $a.maturityDate
            dscr            = Get-Num $a.debtServiceCoverageNetCashFlowSecuritizationPercentage
            valuation       = Get-Num $bay.valuationSecuritizationAmount
            occupancy       = Get-Num $bay.physicalOccupancySecuritizationPercentage
            paymentStatus   = [string]$a.paymentStatusLoanCode
            workoutStrategy = [string]$a.workoutStrategyCode
        })
    }
    if ($i % 25 -eq 0) { Write-Host "  processed $i / $($trusts.Count) trusts" }
}
Write-Host "Trusts without asset file: $missing"
Write-Host "Bay Area loans (all property types, all sizes): $($rows.Count)"

# ---- 5. Criteria: multifamily (MF) and retail (RT), loan balance $10M-$400M ----
# Same loan often appears in several trusts (pari passu notes). Combine them by
# property + address + maturity + coupon, summing balances, before applying the size band.
$mfAll = @($rows | Where-Object { $_.propertyType -in @('MF','RT') })
$loanGroups = $mfAll | Group-Object { "{0}|{1}|{2}|{3}|{4}" -f $_.propertyType, ($_.propertyName.Trim().ToUpper()), ($_.address.Trim().ToUpper()), $_.maturityDate, ([math]::Round($_.couponRate, 6)) }
$mf = @($loanGroups | ForEach-Object {
    $g = $_.Group
    $first = $g[0]
    $bal = ($g | Measure-Object balance -Sum).Sum
    if ($bal -lt $MinBalance -or $bal -gt $MaxBalance) { return }
    $o = $first.PSObject.Copy()
    $o.balance = $bal
    $o | Add-Member -NotePropertyName trust -NotePropertyValue (($g | ForEach-Object { $_.trust } | Select-Object -Unique) -join '; ') -Force
    $o | Add-Member -NotePropertyName noteCount -NotePropertyValue $g.Count -Force
    $o
})
Write-Host "MF/RT rows: $($mfAll.Count) -> $(@($mf).Count) loans within size band (MF: $(@($mf | Where-Object propertyType -eq 'MF').Count), RT: $(@($mf | Where-Object propertyType -eq 'RT').Count))"

# ---- 6. Geocode (Census), cached ----
$geoCachePath = Join-Path $Cache 'geocode.json'
$geo = @{}
if (Test-Path $geoCachePath) { (Get-Content -Raw $geoCachePath | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $geo[$_.Name] = $_.Value } }
foreach ($r in $mf) {
    $street = (($r.address -split ' and ')[0] -split ';')[0].Trim()
    $key = "$street|$($r.city)|$($r.zip)"
    if ($geo.ContainsKey($key)) { continue }
    $oneLine = "$street, $($r.city), CA $($r.zip)"
    $q = "https://geocoding.geo.census.gov/geocoder/locations/onelineaddress?address=$([uri]::EscapeDataString($oneLine))&benchmark=Public_AR_Current&format=json"
    try {
        $g = Get-Text $q | ConvertFrom-Json
        $m = $g.result.addressMatches | Select-Object -First 1
        if ($m) { $geo[$key] = [pscustomobject]@{ lat = $m.coordinates.y; lon = $m.coordinates.x } } else { $geo[$key] = $null }
    } catch { $geo[$key] = $null }
}
$geo | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 $geoCachePath
Write-Host "Geocoded addresses: $(@($geo.Keys | Where-Object { $geo[$_] }).Count) of $($geo.Count)"

# ---- 7. Derived fields and flags ----
# Top-10% cutoff is computed separately for each property type (MF and RT differ in rate profile)
$p90ByType = @{}
foreach ($type in @('MF','RT')) {
    $gaps = @($mf | Where-Object propertyType -eq $type | ForEach-Object { $benchmark - $_.couponRate } | Sort-Object)
    if ($gaps.Count -gt 0) { $p90ByType[$type] = $gaps[[math]::Min([math]::Floor($TopPercentile * ($gaps.Count - 1)), $gaps.Count - 1)] }
}

$out = foreach ($r in $mf) {
    $street = (($r.address -split ' and ')[0] -split ';')[0].Trim()
    $g = $geo["$street|$($r.city)|$($r.zip)"]
    $gap = $benchmark - $r.couponRate
    $p90 = $p90ByType[$r.propertyType]
    $flagSpread = $gap -ge $SpreadThreshold
    $flagTop    = ($null -ne $p90) -and ($gap -ge $p90)
    # Special-servicing proxy from Reg AB codes (see notes; verify code tables)
    $ss = $false
    if ($r.workoutStrategy -and $r.workoutStrategy -notin @('98','99')) { $ss = $true }
    if ($r.paymentStatus -in @('4','5','A','B')) { $ss = $true }
    $monthsToMat = $null
    if ($r.maturityDate) { $monthsToMat = [math]::Round(([DateTime]::Parse($r.maturityDate, $Inv) - (Get-Date)).TotalDays / 30.4, 1) }
    [pscustomobject]@{
        id               = "$($r.trustCik)-$($r.loanId)-$($r.filingAdsh)"
        trust            = $r.trust
        loanId           = $r.loanId
        propertyName     = $r.propertyName
        address          = $r.address
        city             = $r.city
        county           = $r.county
        zip              = $r.zip
        propertyType     = $r.propertyType
        units            = $r.units
        sqft             = $r.sqft
        yearBuilt        = $r.yearBuilt
        propertyCount    = $r.propertyCount
        balance          = $r.balance
        couponRate       = $r.couponRate
        benchmarkRate    = $benchmark
        refiGap          = [math]::Round($gap, 5)
        maturityDate     = $r.maturityDate
        monthsToMaturity = $monthsToMat
        dscr             = $r.dscr
        valuation        = $r.valuation
        occupancy        = $r.occupancy
        paymentStatus    = $r.paymentStatus
        workoutStrategy  = $r.workoutStrategy
        specialServicing = $ss
        flagSpread       = $flagSpread
        flagTop10        = $flagTop
        flagged          = ($flagSpread -or $flagTop)
        lat              = if ($g) { $g.lat } else { $null }
        lon              = if ($g) { $g.lon } else { $null }
        reportingDate    = $r.reportingDate
        filedDate        = $r.filedDate
        secFiling        = "https://www.sec.gov/Archives/edgar/data/$($r.trustCik)/$($r.filingAdsh -replace '-','')/$($r.filingAdsh)-index.htm"
    }
}

$payload = [pscustomobject]@{
    generatedAt  = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    lookbackDays = $LookbackDays
    benchmark    = [pscustomobject]@{ ust10yr = $ust10; ust10yrDate = $dgsDate; spread = $BenchmarkSpread; rate = $benchmark }
    criteria     = [pscustomobject]@{ minBalance = $MinBalance; maxBalance = $MaxBalance; spreadThreshold = $SpreadThreshold; topPercentile = $TopPercentile; p90GapMF = $p90ByType['MF']; p90GapRT = $p90ByType['RT'] }
    counts       = [pscustomobject]@{ trusts = $trusts.Count; bayAreaLoansAllTypes = $rows.Count; mfInBand = $mf.Count; flagged = @($out | Where-Object flagged).Count }
    loans        = @($out)
}
# data.js (not JSON) so the page also works when opened directly from disk
$json = $payload | ConvertTo-Json -Depth 6
Set-Content -Encoding UTF8 -Path (Join-Path $Site 'data.js') -Value ("window.LOAN_DATA = " + $json + ";")
Write-Host "Wrote data.js with $(@($out).Count) loans, $(@($out | Where-Object flagged).Count) flagged."
