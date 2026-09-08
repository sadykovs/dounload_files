import os
import re
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# --------------------- 1. Фиксированные URL ---------------------
FIXED_URLS = [
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
]

DEST_DIR = r"C:\files"

# --------------------- 2. Вспомогательные функции ---------------------
def ensure_dir(path):
    """Создаёт папку, если её нет."""
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"Created folder: {path}")

def download_file(url, dest_path):
    """Скачивает файл по URL и сохраняет по указанному пути."""
    try:
        response = requests.get(url, timeout=30)
        response.raise_for_status()
        with open(dest_path, 'wb') as f:
            f.write(response.content)
        print(f"Saved: {dest_path}")
        return True
    except Exception as e:
        print(f"Error downloading {url}: {e}")
        return False

def get_latest_file_by_pattern(page_url, pattern):
    """
    Парсит страницу, ищет все ссылки, соответствующие regex-паттерну,
    извлекает дату (первая группа из 8 цифр) и возвращает URL самого свежего файла.
    Если ничего не найдено, возвращает None.
    """
    print(f"Searching for pattern '{pattern}' on {page_url} ...")
    try:
        resp = requests.get(page_url, timeout=30)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        links = soup.find_all('a', href=True)
        found = []
        for a in links:
            href = a['href']
            m = re.search(pattern, href)
            if m:
                date_str = m.group(1)  # 8 цифр
                # Преобразуем относительный URL в абсолютный
                full_url = urljoin(page_url, href)
                found.append((date_str, full_url))
        if not found:
            print("No matching files found.")
            return None
        # Находим максимальную дату (лексикографически для YYYYMMDD)
        latest = max(found, key=lambda x: x[0])
        print(f"Latest file found: {latest[1]} (date: {latest[0]})")
        return latest[1]
    except Exception as e:
        print(f"Error parsing page: {e}")
        return None

# --------------------- 3. Основная логика ---------------------
def main():
    ensure_dir(DEST_DIR)

    # Собираем все URL для скачивания
    urls_to_download = FIXED_URLS.copy()

    # Динамические файлы
    foreign_url = get_latest_file_by_pattern(
        "https://www.cbr.ru/registries/infrastr/",
        r'list_foreign_org_(\d{8})\.xlsx'
    )
    if foreign_url:
        urls_to_download.append(foreign_url)

    mfo_url = get_latest_file_by_pattern(
        "https://cbr.ru/registries/microfinance",
        r'list_MFO_p_(\d{8})\.xlsx'
    )
    if mfo_url:
        urls_to_download.append(mfo_url)

    print(f"\nStarting download of {len(urls_to_download)} files...\n")
    for url in urls_to_download:
        filename = os.path.basename(url)
        if not filename:
            filename = "unknown.xlsx"
        dest_path = os.path.join(DEST_DIR, filename)
        print(f"Downloading: {filename} ...")
        download_file(url, dest_path)

    print("\nAll downloads completed!")

if __name__ == "__main__":
    main()