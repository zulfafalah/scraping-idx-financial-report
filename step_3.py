import pandas as pd
from  main import read_data
import json
from datetime import datetime

def format_date(date):
    format_dengan_microdetik = "%Y-%m-%dT%H:%M:%S.%f"
    
    format_tanpa_microdetik = "%Y-%m-%dT%H:%M:%S"
    
    try:
        tanggal_obj = datetime.strptime(date, format_dengan_microdetik)
    except ValueError:
        try:
            tanggal_obj = datetime.strptime(date, format_tanpa_microdetik)
        except ValueError:
            return ""
    
    tanggal_baru = tanggal_obj.strftime("%d/%m/%Y")
    
    return tanggal_baru

def get_lk_data(year, kode_emiten):
    file_path = f"response_json/{year}/idx_financial_report_{kode_emiten}.json"
    try:
        with open(file_path, 'r') as file:
            data = json.load(file)
            result = data['Results']
            if len(result) <= 0:
                return ""
            return result[0]['File_Modified']

    except Exception as e:
        return ""    
    

def get_hau_data(year, kode_emiten):
    file_path = f"response_file/{year}/FinancialStatement-{year}-Tahunan-{kode_emiten}.xlsx"

    try:
        df = pd.read_excel(file_path, sheet_name="1000000")
        target_rows = df[df['[1000000] General information'].str.contains('Jumlah Hal Audit Utama', na=False)]["Unnamed: 1"]

        result = target_rows.to_string(index=False)
        if result == "NaN":
            return ""
        
        return result

    except Exception as e:
        return ""    
    
def get_paragraf_hau_data(year, kode_emiten):
    file_path = f"response_file/{year}/FinancialStatement-{year}-Tahunan-{kode_emiten}.xlsx"

    try:
        df = pd.read_excel(file_path, sheet_name="1000000")
        target_rows = df[df['[1000000] General information'].str.contains('Paragraf Hal Audit Utama', na=False)]["Unnamed: 1"]

        result = target_rows.to_string(index=False)
        if result == "NaN":
            return ""
        
        return result

    except Exception as e:
        return ""    
    
def get_auditor_tahun_berjalan(year, kode_emiten):
    file_path = f"response_file/{year}/FinancialStatement-{year}-Tahunan-{kode_emiten}.xlsx"

    try:
        df = pd.read_excel(file_path, sheet_name="1000000")
        target_rows = df[df['[1000000] General information'].str.contains('Auditor tahun berjalan', na=False)]["Unnamed: 1"]

        result = target_rows.to_string(index=False)
        if result == "NaN":
            return ""
        
        return result

    except Exception as e:
        return ""    
    
def get_auditor_tahun_sebelumnya(year, kode_emiten):
    file_path = f"response_file/{year}/FinancialStatement-{year}-Tahunan-{kode_emiten}.xlsx"

    try:
        df = pd.read_excel(file_path, sheet_name="1000000")
        target_rows = df[df['[1000000] General information'].str.contains('Auditor tahun sebelumnya', na=False)]["Unnamed: 1"]

        result = target_rows.to_string(index=False)
        if result == "NaN":
            return ""
        
        return result

    except Exception as e:
        return ""    
    

def main_lk():
    emiten_list = read_data()
    year = 2023
    df = pd.read_excel('Book4.xlsx')
    for kode_emiten in emiten_list:
        date = get_lk_data(year, kode_emiten)
        formated_date = format_date(date)

        # Lihat data awal
        print("Data Awal:\n", df)

        # Contoh manipulasi: Update stok item_code 'A002' menjadi 60
        df.loc[df['Emiten'] == kode_emiten, 'LK 2023'] = formated_date


        # Simpan kembali ke Excel
    df.to_excel('Book4.xlsx', index=False)

    print("\nData setelah dimanipulasi dan disimpan ke 'Book4.xlsx'")

def main_hau():
    year = 2023
    emiten_list = read_data()
    df = pd.read_excel('Book4.xlsx')

    for kode_emiten in emiten_list:
        hau = get_hau_data(year, kode_emiten)
        df.loc[df['Emiten'] == kode_emiten, 'Jumlah Hal Audit Utama 2023'] = hau
        print(f"{kode_emiten} : {hau}")

        paragraf_hau = get_paragraf_hau_data(year, kode_emiten)
        df.loc[df['Emiten'] == kode_emiten, 'Paragraf HAU 2023'] = paragraf_hau
        print(f"{kode_emiten} : {paragraf_hau}")

        auditor_tahun_berjalan = get_auditor_tahun_berjalan(year, kode_emiten)
        df.loc[df['Emiten'] == kode_emiten, 'Auditor Tahun Berjalan 2023'] = auditor_tahun_berjalan
        print(f"{kode_emiten} : {auditor_tahun_berjalan}")

        auditor_tahun_sebelumnya = get_auditor_tahun_sebelumnya(year, kode_emiten)
        df.loc[df['Emiten'] == kode_emiten, 'Auditor tahun sebelumnya (2022)'] = auditor_tahun_sebelumnya
        print(f"{kode_emiten} : {auditor_tahun_sebelumnya}")


    df.to_excel('Book4.xlsx', index=False)


if __name__ == '__main__':
    main_lk()
    main_hau()