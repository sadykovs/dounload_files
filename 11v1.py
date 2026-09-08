import os
import re
import requests
import urllib3
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# Отключаем предупреждения о небезопасном запросе (verify=False)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --------------------- 1. Фиксированные файлы (url, filename) ---------------------
FIXED_FILES = [
    ("https://cbr.ru/vfs/finmarkets/files/supervision/list_aofr.xlsx", "list_aofr.xlsx"),
    ("https://www.cbr.ru/vfs/finmarkets/files/supervision/reg_bki.xlsx", "reg_bki.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_brokers.xlsx", "list_brokers.xlsx"),
    ("http://www.cbr.ru/vfs/finmarkets/files/supervision/list_cliring.xlsx", "list_cliring.xlsx"),
    ("http://www.cbr.ru/vfs/finmarkets/files/supervision/list_kra.xlsx", "list_kra.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_fpikra.xlsx", "list_fpikra.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_skpk.xlsx", "list_skpk.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_KPK_gov.xlsx", "list_KPK_gov.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_dealers.xlsx", "list_dealers.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_depositaries.xlsx", "list_depositaries.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_specdepositaries.xlsx", "list_specdepositaries.xlsx"),
    ("https://www.cbr.ru/vfs/finmarkets/files/supervision/list_financial_platform_op.xlsx", "list_financial_platform_op.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_forex_dealers.xlsx", "list_forex_dealers.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_ZHNK.xlsx", "list_ZHNK.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_ssd.xlsx", "list_ssd.xlsx"),
    ("http://www.cbr.ru/vfs/finmarkets/files/supervision/list_ossd.xlsx", "list_ossd.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/List_is.xlsx", "List_is.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_if.xlsx", "list_if.xlsx"),
    ("https://www.cbr.ru/vfs/finmarkets/files/supervision/list_invest_platform_op.xlsx", "list_invest_platform_op.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_PS.xlsx", "list_PS.xlsx"),
    ("https://cbr.ru/vfs/finmarkets/files/supervision/list_MFO.xlsx", "list_MFO.xlsx"),
    ("https://www.cbr.ru/vfs/finmarkets/files/supervision/list_ois.xlsx", "list_ois.xlsx"),
    ("http://www.cbr.ru/vfs/registries/admissionfinmarket/list_oocfa.xlsx", "list_oocfa.xlsx"),
    ("https://www.cbr.ru/vfs/finmarkets/files/supervision/list_opp.xlsx", "list_opp.xlsx"),
    ("https://cbr.ru/vfs/finmarkets/files/supervision/List_OSR.xlsx", "List_OSR.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_npf_Valid.xlsx", "list_npf_Valid.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_reestersavers.xlsx", "list_reestersavers.xlsx"),
    ("http://www.cbr.ru/vfs/finmarkets/files/supervision/list_repositaries.xlsx", "list_repositaries.xlsx"),
    ("https://cbr.ru/vfs/finmarkets/files/supervision/list_PIF.xlsx", "list_PIF.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_sro_actuarials.xlsx", "list_sro_actuarials.xlsx"),
    ("http://www.cbr.ru/vfs/finmarkets/files/supervision/list_exchange.xlsx", "list_exchange.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_managementcompanies.xlsx", "list_managementcompanies.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_ukso.xlsx", "list_ukso.xlsx"),
    ("http://cbr.ru/vfs/finmarkets/files/supervision/list_trust.xlsx", "list_trust.xlsx"),
    ("https://data.economy.gov.ru/files/sonko_organizations.xlsx", "sonko_organizations.xlsx")
]

DEST_DIR = r"C:\files"

# --------------------- 2. Заголовки для обхода 403 ---------------------
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1"
}

# --------------------- 3. Вспомогательные функции ---------------------
def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"Created folder: {path}")

def download_file(url, dest_path, referer=None, verify_ssl=True):
    """Скачивает файл, добавляя заголовки, и сохраняет."""
    try:
        headers = HEADERS.copy()
        if referer:
            headers["Referer"] = referer
        # Для Минюста отключаем проверку SSL
        if "reestrs.minjust.gov.ru" in url:
            verify_ssl = False
        response = requests.get(url, headers=headers, timeout=120, verify=verify_ssl)
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(response.content)
        print(f"Saved: {dest_path}")
        return True
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

def get_latest_file_by_pattern(page_url, pattern):
    print(f"Searching for pattern '{pattern}' on {page_url} ...")
    try:
        response = requests.get(page_url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')
        links = soup.find_all('a', href=True)
        found = []
        for a in links:
            href = a['href']
            m = re.search(pattern, href)
            if m:
                date_str = m.group(1)
                full_url = urljoin(page_url, href)
                found.append((date_str, full_url))
        if not found:
            print("No matching files found.")
            return None
        latest = max(found, key=lambda x: x[0])
        filename = os.path.basename(latest[1])
        print(f"Latest file found: {latest[1]} (date: {latest[0]})")
        return (latest[1], filename)
    except Exception as e:
        print(f"Error parsing page: {e}")
        return None

def get_minjust_file_info(page_url):
    """
    Парсит страницу Минюста, извлекает ID реестра и базовый URL API,
    формирует ссылку для скачивания.
    Возвращает кортеж (url, filename)
    """
    print(f"Fetching page: {page_url} ...")
    try:
        response = requests.get(page_url, headers=HEADERS, timeout=30)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        registry_id = None
        base_url = None

        scripts = soup.find_all('script')
        for script in scripts:
            if script.string:
                # Ищем let id = '...'
                id_match = re.search(r"let id = '([a-f0-9-]+)';", script.string)
                if id_match:
                    registry_id = id_match.group(1)
                # Ищем ExternalApi.setBaseUrl('...')
                base_match = re.search(r"ExternalApi\.setBaseUrl\('([^']+)'\);", script.string)
                if base_match:
                    base_url = base_match.group(1)

        if not registry_id or not base_url:
            raise Exception("ID реестра или базовый URL не найдены")

        download_url = f"{base_url}/rest/registry/{registry_id}/export"
        print(f"Ссылка для скачивания: {download_url}")
        return (download_url, "reestr_nko.xlsx")

    except Exception as e:
        print(f"Ошибка при парсинге страницы Минюста: {e}")
        return None

# --------------------- 4. Основная логика ---------------------
def main():
    ensure_dir(DEST_DIR)

    # Собираем список файлов (кортежи url, filename)
    files_to_download = FIXED_FILES.copy()

    # 1. ЦБ – иностранные организации
    foreign = get_latest_file_by_pattern(
        "https://www.cbr.ru/registries/infrastr/",
        r'list_foreign_org_(\d{8})\.xlsx'
    )
    if foreign:
        files_to_download.append(foreign)

    # 2. ЦБ – микрофинансовые организации (MFO p)
    mfo = get_latest_file_by_pattern(
        "https://cbr.ru/registries/microfinance",
        r'list_MFO_p_(\d{8})\.xlsx'
    )
    if mfo:
        files_to_download.append(mfo)

    # 3. Минюст – реестр НКО
    minjust = get_minjust_file_info("https://minjust.gov.ru/ru/pages/reestr-nko-ispolnitelej-obshestvenno-poleznyh-uslug/")
    if minjust:
        files_to_download.append(minjust)

    print(f"\nStarting download of {len(files_to_download)} files...\n")

    for url, filename in files_to_download:
        dest_path = os.path.join(DEST_DIR, filename)
        print(f"Downloading: {filename} ...")
        # Для sonko_organizations.xlsx укажем Referer
        referer = None
        if "data.economy.gov.ru" in url:
            referer = "https://data.economy.gov.ru/"
        download_file(url, dest_path, referer=referer)

    print("\nAll downloads completed!")

if __name__ == "__main__":
    main()