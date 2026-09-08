# =====================================================
# 1. Список фиксированных URL-адресов (35 штук)
# =====================================================
$fixedUrls = @(
    "https://cbr.ru/vfs/finmarkets/files/supervision/list_aofr.xlsx",
    "https://www.cbr.ru/vfs/finmarkets/files/supervision/reg_bki.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_brokers.xlsx",
    "http://www.cbr.ru/vfs/finmarkets/files/supervision/list_cliring.xlsx",
    "http://www.cbr.ru/vfs/finmarkets/files/supervision/list_kra.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_fpikra.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_skpk.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_KPK_gov.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_dealers.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_depositaries.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_specdepositaries.xlsx",
    "https://www.cbr.ru/vfs/finmarkets/files/supervision/list_financial_platform_op.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_forex_dealers.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_ZHNK.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_ssd.xlsx",
    "http://www.cbr.ru/vfs/finmarkets/files/supervision/list_ossd.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/List_is.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_if.xlsx",
    "https://www.cbr.ru/vfs/finmarkets/files/supervision/list_invest_platform_op.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_PS.xlsx",
    "https://cbr.ru/vfs/finmarkets/files/supervision/list_MFO.xlsx",
    "https://www.cbr.ru/vfs/finmarkets/files/supervision/list_ois.xlsx",
    "http://www.cbr.ru/vfs/registries/admissionfinmarket/list_oocfa.xlsx",
    "https://www.cbr.ru/vfs/finmarkets/files/supervision/list_opp.xlsx",
    "https://cbr.ru/vfs/finmarkets/files/supervision/List_OSR.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_npf_Valid.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_reestersavers.xlsx",
    "http://www.cbr.ru/vfs/finmarkets/files/supervision/list_repositaries.xlsx",
    "https://cbr.ru/vfs/finmarkets/files/supervision/list_PIF.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_sro_actuarials.xlsx",
    "http://www.cbr.ru/vfs/finmarkets/files/supervision/list_exchange.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_managementcompanies.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_ukso.xlsx",
    "http://cbr.ru/vfs/finmarkets/files/supervision/list_trust.xlsx",
    "https://data.economy.gov.ru/files/sonko_organizations.xlsx"
)

# =====================================================
# 2. Папка для сохранения
# =====================================================
$destination = "C:\files"
if (-not (Test-Path $destination)) {
    New-Item -ItemType Directory -Path $destination -Force | Out-Null
    Write-Host "Created folder: $destination" -ForegroundColor Cyan
}

# =====================================================
# 3. Вспомогательная функция: получить абсолютный URL из относительного
# =====================================================
function Get-AbsoluteUrl($baseUrl, $relativeUrl) {
    if ($relativeUrl -match '^https?://') {
        return $relativeUrl
    }
    if ($relativeUrl -match '^//') {
        return "https:$relativeUrl"
    }
    # Используем класс Uri для правильного объединения
    $base = $baseUrl -replace '/[^/]*$', '/'  # убираем последний сегмент
    if ($relativeUrl -match '^/') {
        # абсолютный путь относительно домена
        $uri = [System.Uri]::new($baseUrl)
        $full = $uri.Scheme + "://" + $uri.Host + $relativeUrl
        return $full
    } else {
        # относительный путь внутри текущего каталога
        return $base + $relativeUrl
    }
}

# =====================================================
# 4. Функция: поиск самого свежего файла на странице по маске
# =====================================================
function Get-LatestFileByPattern($pageUrl, $pattern) {
    Write-Host "Searching for files matching pattern '$pattern' on $pageUrl ..." -ForegroundColor Cyan
    try {
        $response = Invoke-WebRequest -Uri $pageUrl -UseBasicParsing -ErrorAction Stop
        $links = $response.Links | Where-Object { $_.href -match $pattern }
        if (-not $links) {
            Write-Host "No matching files found." -ForegroundColor Yellow
            return $null
        }

        # Извлекаем даты из найденных ссылок (предполагаем 8 цифр после последнего подчёркивания)
        $found = @()
        foreach ($link in $links) {
            $href = $link.href
            if ($href -match $pattern) {
                $dateStr = $matches[1]  # группа из 8 цифр
                $found += [PSCustomObject]@{ Date = $dateStr; Href = $href }
            }
        }
        if ($found.Count -eq 0) {
            Write-Host "No dates extracted." -ForegroundColor Yellow
            return $null
        }
        # Выбираем максимальную дату (лексикографически, т.к. формат YYYYMMDD или DDMMYYYY – но для YYYYMMDD работает)
        $max = $found | Sort-Object Date -Descending | Select-Object -First 1
        $fullUrl = Get-AbsoluteUrl -baseUrl $pageUrl -relativeUrl $max.Href
        Write-Host "Latest file found: $fullUrl (date: $($max.Date))" -ForegroundColor Green
        return $fullUrl
    }
    catch {
        Write-Host "Error while parsing page: $_" -ForegroundColor Red
        return $null
    }
}

# =====================================================
# 5. Получаем динамические ссылки
# =====================================================
$foreignUrl = Get-LatestFileByPattern -pageUrl "https://www.cbr.ru/registries/infrastr/" -pattern 'list_foreign_org_(\d{8})\.xlsx'
$mfoUrl = Get-LatestFileByPattern -pageUrl "https://cbr.ru/registries/microfinance" -pattern 'list_MFO_p_(\d{8})\.xlsx'

# =====================================================
# 6. Формируем итоговый список и скачиваем
# =====================================================
$allUrls = $fixedUrls
if ($foreignUrl) { $allUrls += $foreignUrl }
if ($mfoUrl)     { $allUrls += $mfoUrl }

Write-Host "`nStarting download of $($allUrls.Count) files..." -ForegroundColor Cyan

foreach ($url in $allUrls) {
    $fileName = [System.IO.Path]::GetFileName($url)
    $filePath = Join-Path $destination $fileName

    Write-Host "Downloading: $fileName ..." -ForegroundColor Cyan
    try {
        Invoke-WebRequest -Uri $url -OutFile $filePath -ErrorAction Stop
        Write-Host "Saved: $filePath" -ForegroundColor Green
    } catch {
        Write-Host "Error downloading $fileName : $_" -ForegroundColor Red
    }
}

Write-Host "`nAll downloads completed!" -ForegroundColor Yellow