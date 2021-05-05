import requests
from openpyxl import load_workbook

TOKEN = "c143937605392a452de7bbf7838e829e"
MAIN_URL =  f'https://apidata.mos.ru/v1'
payload = {
    'api_key': TOKEN
}
LIST_OF_TITLES_PATH = './list_of_titles.xlsx'

def update_list_of_titles(list_of_titles=LIST_OF_TITLES_PATH, quantity=3500):

    workbook = load_workbook(list_of_titles)
    print('Workbook was opened')
    sheet = workbook['page1']
    print(f'Start cycle from 1 to {quantity}')
    counter = 1

    for i in range(1, quantity):
        reuuest = requests.get(f'{MAIN_URL}/datasets/{i}/', params=payload)
        statuscode = reuuest.status_code

        if statuscode == 200:
            json_data = reuuest.json()
            counter += 1
            sheet[f'A{counter}'] = json_data['Id']
            sheet[f'B{counter}'] = json_data['IdentificationNumber']
            sheet[f'C{counter}'] = json_data['CategoryId']
            sheet[f'D{counter}'] = json_data['CategoryCaption']
            sheet[f'E{counter}'] = json_data['DepartmentId']
            sheet[f'F{counter}'] = json_data['DepartmentCaption']
            sheet[f'G{counter}'] = json_data['Caption']
            sheet[f'H{counter}'] = json_data['Description']
    print(f'Total values: {counter-1}')
    print('Workbook in the process of saving...')
    workbook.save(LIST_OF_TITLES_PATH)
    print('DONE')

update_list_of_titles(LIST_OF_TITLES_PATH, 5000)