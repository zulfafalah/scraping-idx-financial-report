import os
import json
import re
from main import download_file


def process_attachment(attachment, year):
    pattern = r'.*\.(xlsx|xls)$'
    if re.match(pattern, attachment["File_Name"], re.IGNORECASE):
        print(attachment["File_Path"])
        download_file(
            file_url=f"""https://www.idx.co.id/{attachment["File_Path"]}""",
            year=year,
            kodeEmiten=attachment["Emiten_Code"]
        )

def main():
    year = 2021
    folder_path = f'response_json/{year}'
    
    # Use list comprehension and generator for better memory efficiency
    json_files = (os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.json'))
    
    for file_path in json_files:
        with open(file_path, 'r') as file:
            data = json.load(file)
            result = data.get('Results', [])
            
            if result and result[0].get('Attachments'):
                for attachment in result[0]['Attachments']:
                    process_attachment(attachment, year)


if __name__ == '__main__':
    main()