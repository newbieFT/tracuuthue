import requests
import pandas as pd
import base64
import time
from bs4 import BeautifulSoup

# Constants and file paths
INPUT_FILE = r"Tracuunnt\Masothue_TEST.xlsx"
OUTPUT_FILE = r"Tracuunnt\Tracuunnt_output.xlsx"
WEBSITE_URL = "https://tracuunnt.gdt.gov.vn/tcnnt/mstdn.jsp"
API_KEY = '337439d35b1cab249ca44ffe5e1940ca'  # Your 2Captcha API Key

def solve_captcha(session, img_url):
    response = session.get(img_url)
    
    # Convert image content to base64
    img_base64 = base64.b64encode(response.content).decode('utf-8')
    
    # Prepare data for 2Captcha
    data = {
        'key': API_KEY,
        'method': 'base64',
        'body': img_base64,
        'json': 1
    }
    response = requests.post('http://2captcha.com/in.php', data=data)
    json_response = response.json()
    
    if json_response.get('status') == 1:
        captcha_id = json_response.get('request')
        time.sleep(20)  # Wait for CAPTCHA to be solved
        
        response_res = requests.get(f'https://2captcha.com/res.php?key={API_KEY}&action=get&id={captcha_id}&json=1')
        json_dua_res = response_res.json()
        if json_dua_res.get('status') == 1:
            return json_dua_res.get('request')
    
    return None

def process_tax_code(session, tax_code, headers, cookies):
    response = session.get(WEBSITE_URL, headers=headers, cookies=cookies)
    soup = BeautifulSoup(response.content, 'html.parser')
    
    # Solve CAPTCHA
    captcha_img = soup.find('img', {'src': '/tcnnt/captcha.png?uid='})
    if not captcha_img:
        print(f"Failed to find CAPTCHA for tax code {tax_code}")
        return {"Tax code": tax_code, "Company name": "Not found", "Address": "Not found"}
    
    img_url = 'https://tracuunnt.gdt.gov.vn' + captcha_img['src']
    captcha_solution = solve_captcha(session, img_url)
    if captcha_solution is None:
        return {"Tax code": tax_code, "Company name": "Captcha error", "Address": "Captcha error"}
    
    # Submit form with tax code and CAPTCHA solution
    form_data = {
        'cm': 'cm',
        'mst': tax_code,
        'fullname': '',
        'address': '',
        'cmt': '',
        'captcha': captcha_solution
    }
    submit_response = session.post(WEBSITE_URL, data=form_data, headers=headers, cookies=cookies)
    submit_soup = BeautifulSoup(submit_response.content, 'html.parser')
    
    # Extract data
    company_name = submit_soup.find('td', {'id': 'TEN_NNT'}).get_text(strip=True)
    address = submit_soup.find('td', {'id': 'DIA_CHI_TRU_SO'}).get_text(strip=True)

    return {
        "Tax code": tax_code,
        "Company name": company_name if company_name else "Not found",
        "Address": address if address else "Not found"
    }

def main():
    session = requests.Session()
    cookies = {
        'f5_cspm': '1234',
        'f5avraaaaaaaaaaaaaaaa_session_': 'PALAEKAKLNOKBPJNJIJIBDFOELOEGOBECMGKNKJNBDDCFLDILNPFHMPHLIHIEIGHEGODGLEBOOJBHLEMODGADBGIGPEFKBAJANOOHILILOKAJNBAMHEDEELPMCHBLNMI',
        'f5avrbbbbbbbbbbbbbbbb': 'JAMCIOJFJAEHCIGFENPFJGODCAKFEPGKGCCKFOKKJKEPOLCLFCGJBIOHBBOCKLNIIBDIFMMOIBKDJLEPHEEJJCMLBHLAGHBLBPLLMCOFNBMMMONCKPNNBEKHFBEAMBHA',
        'JGhfjdFAdk': '2722316298.31011.0000',
        'Merry-Christmas': '4097982474.20480.0000',
        '_ga': 'GA1.3.1582689975.1753177652',
        '_gtpk_testcookie..undefined': '1',
        '_gtpk_testcookie.295.7f87': '1',
        '_gid': 'GA1.3.1997948474.1753408991',
        'f5avrbbbbbbbbbbbbbbbb': 'ICNOELEECHFHDEMFJMMFLNJEAPOCNIIHGPBPJPGGINAOCPNNEGCLHOMMMAAIMDDNMLDPMIGOLLCDPCIGHBOHPHMPFKCACLLHEFHDCMDJICPKBMAKJOLPMBOPAGBMKFNF',
        'f5avraaaaaaaaaaaaaaaa_session_': 'GGBKMPLDOPJNJKPGKAFIEMMBFEGKDFINGAJDOABGBCIOCGGMFGMELIMPANIIEJPBCJODOKDMEFLKNOKEKDMACLKIJOLCGCJNJBAHBBNCEFAGLHGJCOAOGMNGJGPNPOOD',
        'JSESSIONID': '0000011h50DtpYGbSIVYPZQ2E18:19ciqlo0c',
        '_gtpk_ses.295.7f87': '1',
        '_gat': '1',
        '_gtpk_id.295.7f87': '49e4851294fce48b.1753177676.4.1753415660.1753415354.',
        'TS0133099e': '01dc12c85ed5de58fc0cf20bfa28344f83a7fe173fa49352ee45233a3194ee98046c6c11db2565208f1e32632195ff03ca2d05c557e146fe1984cb36fd630f287d1a30faa4b2b050289eb170e3a22905e232076fc8d3c5deb7f361736a176459dcf120ac49f8862e2efc240b590867305e418e805a8aa977d44632ebf0eefbd9b5eab53c7ddcf552048395c0ebee79daeaae719dbbf28759d23caa185430259e483df476adffa6823dc39c1921e22f30050cf47657b6bedcb92a81cc52260323b974a4c43cc8ec95947e4cdddf5a98a7f7af5853d0',
        '_ga_F9DJVRC3NS': 'GS2.3.s1753415353$o4$g1$t1753415660$j49$l0$h0',
        'f5avr1796159269aaaaaaaaaaaaaaaa': 'MOHAGAJCEILDEFHFCKFBFJIDLDNNBKNHCMENMPIFILJIGLNIFCMNNHFBBEPONDKELFKPFDBGPNACIAMBMBFBJOFECDNAOPGEBBCPPLEJNGCPGPGFADBBIILKGOENJECD',
        'f5avr1075019721aaaaaaaaaaaaaaaa_cspm_': 'OGKJCNJLOJPJCKJLHADAOKDDEEFJAIAPPIFIJBPKJODCGKHCNIMPHGNCPFHPPODCECNCPGNNBBNGPEAOCHGABGCPADGAPKAMIONMBMOPPHKGBBGIKJAGIKOJPOIPNMOO',
    }

    headers = {
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7',
        'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
        'Cache-Control': 'max-age=0',
        'Connection': 'keep-alive',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'none',
        'Sec-Fetch-User': '?1',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36',
        'sec-ch-ua': '"Not)A;Brand";v="8", "Chromium";v="138", "Google Chrome";v="138"',
        'sec-ch-ua-mobile': '?0',
        'sec-ch-ua-platform': '"Windows"',
    }
    
    # Load tax codes
    df = pd.read_excel(INPUT_FILE, dtype=str)
    tax_codes = df.iloc[:, 0].dropna().tolist()

    results = []
    for tax_code in tax_codes:
        if not tax_code:
            continue
        
        result = process_tax_code(session, tax_code, headers, cookies)
        results.append(result)
        
        # Sleep to avoid being blocked
        time.sleep(2)
    
    # Save results
    results_df = pd.DataFrame(results)
    results_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Scraping completed. Results saved to {OUTPUT_FILE}")
    
if __name__ == "__main__":
    main()
