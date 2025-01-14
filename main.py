import cloudscraper
import pandas as pd
import os
import json
import time
from datetime import datetime

def read_data():
    df = pd.read_excel('Book4.xlsx')
    df.columns = df.columns.str.strip()
    emiten_list = df['Emiten'].dropna().unique()
    return emiten_list

def download_file(file_url, year, kodeEmiten):
    scraper = cloudscraper.create_scraper(
        browser={
            'browser': 'firefox',
            'platform': 'windows',
            'mobile': False
        },
        delay=10
    )

    # get valid cookies
    main_url = "https://www.idx.co.id/id/perusahaan-tercatat/laporan-keuangan-dan-tahunan/"
    scraper.get(main_url)

    time.sleep(2)

    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/115.0',
        'Accept': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Accept-Language': 'en-US,en;q=0.5',
        'Accept-Encoding': 'gzip, deflate, br',
        'Connection': 'keep-alive',
        'Referer': 'https://www.idx.co.id/id/perusahaan-tercatat/laporan-keuangan-dan-tahunan/',
        'Sec-Fetch-Dest': 'document',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-Site': 'same-origin',
        'Pragma': 'no-cache',
        'Cache-Control': 'no-cache'
    }

    try:
        print("Memulai download...")
        response = scraper.get(file_url, headers=headers, stream=True)
        
        if response.status_code == 200:
            
            if 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in response.headers.get('Content-Type', ''):
                filename = f"response_file/{year}/FinancialStatement-{year}-Tahunan-{kodeEmiten}.xlsx"
                with open(filename, 'wb') as f:
                    for chunk in response.iter_content(chunk_size=8192):
                        if chunk:
                            f.write(chunk)
                print(f"File berhasil didownload: {filename}")
                print(f"Ukuran file: {os.path.getsize(filename)} bytes")
            else:
                print("Response bukan file Excel")
                print(f"Content-Type: {response.headers.get('Content-Type')}")
                print(f"Response preview: {response.text[:200]}")
        else:
            print(f"Gagal download file. Status code: {response.status_code}")
            print(f"Response headers: {dict(response.headers)}")
            print(f"Response: {response.text[:500]}")

    except Exception as e:
        print(f"Error terjadi: {str(e)}")

def main():
    scraper = cloudscraper.create_scraper(delay=10)
    emiten_list = read_data()
    for kodeEmiten in emiten_list:
        year = 2021
        
        url = f"https://www.idx.co.id/primary/ListedCompany/GetFinancialReport?indexFrom=1&pageSize=12&year={year}&reportType=rdf&EmitenType=s&periode=audit&kodeEmiten={kodeEmiten}&"

        headers = {
            'User-Agent': 'Mozilla/5.0 (X11; Ubuntu; Linux x86_64; rv:134.0) Gecko/20100101 Firefox/134.0',
            'Accept': 'application/json',
            'Accept-Language': 'en-US,en;q=0.5',
            'Accept-Encoding': 'identity',
            'Connection': 'keep-alive',
            'Referer': 'https://www.idx.co.id/id/perusahaan-tercatat/laporan-keuangan-dan-tahunan/'
        }

        try:
            response = scraper.get(url, headers=headers)
            print("Status Code:", response.status_code)
            
            if response.status_code == 200:
                try:
                    json_data = response.json()
                    
                    filename = f"response_json/{year}/idx_financial_report_{kodeEmiten}.json"
                    
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(json_data, f, indent=2, ensure_ascii=False)
                    
                    print(f"\nData berhasil disimpan ke file: {filename}")
                    time.sleep(2)
                    
                except json.JSONDecodeError as e:
                    print("\nGagal parse JSON:", str(e))
                    print("\nRaw Response:", response.text)
                
        except Exception as e:
            print("Error occurred:", str(e))


if __name__ == "__main__":
    main()