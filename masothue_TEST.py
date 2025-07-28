import requests
from bs4 import BeautifulSoup
import pandas as pd
import random
from fake_useragent import UserAgent
import time
import json

print("Starting...")

session = requests.Session()

def get_headers():
    ua = UserAgent()
    return {
        'User-Agent': ua.random,
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
        'Accept-Language': 'en-US,en;q=0.9',
        'Cache-Control': 'max-age=0',
        'Referer': 'https://masothue.com/',
        'Connection': 'keep-alive'
    }

def get_cookies():
    url = "https://masothue.com/"
    response = session.get(url, headers=get_headers())
    return session.cookies.get_dict() if response.status_code == 200 else {}

def get_slug_from_search(tax_code):
    token = 'Y94zySWsC0'  # fixed
    ajax_url = "https://masothue.com/Ajax/Search"
    headers = {
        'User-Agent': get_headers()['User-Agent'],
        'Referer': 'https://masothue.com/',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'X-Requested-With': 'XMLHttpRequest'
    }
    data = {
        'q': tax_code,
        'type': 'auto',
        'token': token,
        'force-search': '0'
    }

    # Bước 1: Gửi request AJAX để lấy slug
    search_response = session.post(ajax_url, headers=headers, data=data)

    if search_response.status_code != 200:
        print(f"Search failed for {tax_code}, status code: {search_response.status_code}")
        return None

    try:
        results = search_response.json()
        if results.get("success") == "1" and results.get("url") and results.get("url") != "\\/":
            slug = results["url"].replace("\\/", "/")
            full_url = f"https://masothue.com{slug}"
            print(f"\n[REDIRECTING TO]: {full_url}")

            # Bước 2: Gửi GET tới URL đó để lấy headers thật
            detail_response = session.get(full_url, headers=headers, allow_redirects=True)

            # In response headers tại URL đích
            print(f"\n[RESPONSE HEADERS at {full_url}]:")
            for key, value in detail_response.headers.items():
                print(f"{key}: {value}")

            return slug
        else:
            print(f"No result found for {tax_code}")
            return None
    except Exception as e:
        print(f"Error parsing JSON: {e}")
        return None


def get_company_data(slug_url, original_tax_code):
    url = f"https://masothue.com{slug_url}"
    headers = get_headers()
    cookies = get_cookies()

    try:
        response = session.get(url, headers=headers, cookies=cookies, timeout=10)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, "html.parser")
    except requests.exceptions.RequestException as e:
        print(f"Request failed: {e}")
        return {
            "Tax code": original_tax_code,
            "Company name": "Not found",
            "Address": "Not found"
        }

    try:
        name = soup.select_one("th[itemprop='name'] span.copy").text.strip()
    except:
        name = "Not found"

    try:
        address = soup.select_one("td[itemprop='address'] span.copy").text.strip()
    except:
        address = "Not found"

    return {
        "Tax code": original_tax_code,
        "Company name": name,
        "Address": address
    }

# Load Excel
df = pd.read_excel("Tracuunnt/Tracuunnt.xlsx", dtype=str)
tax_codes = df.iloc[:, 0].dropna().tolist()

output = []
for i, tax_code in enumerate(tax_codes, 1):
    print(f"Processing {tax_code}...")
    slug = get_slug_from_search(tax_code)
    if slug:
        data = get_company_data(slug, tax_code)
    else:
        data = {"Tax code": tax_code, "Company name": "Not found", "Address": "Not found"}
    
    output.append(data)

    if i % 21 == 0:
        print("Sleeping 30 seconds after 21 requests...")
        time.sleep(30)
    else:
        time.sleep(random.uniform(1, 3))

# Save output
output_file_name = "output_request_Main_FixedToken.xlsx"
pd.DataFrame(output).to_excel(output_file_name, index=False)
print(f"Done. File saved: {output_file_name}")
