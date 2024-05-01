from openpyxl.workbook import Workbook
from openpyxl import load_workbook
import requests
from bs4 import BeautifulSoup


book_name="cert-validity.xlsx"
cert_col='A'
valid_col='B'  
start_row=2
wb=Workbook()
wb=load_workbook(book_name)
ws=wb.active
end_row = ws.max_row


def retrieve_current_until(cert_id):
    try:
        url = f'https://www.redhat.com/rhtapps/verify/?certId={cert_id}'
        response = requests.get(url)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, 'html.parser')
            table = soup.find('table', {'style': 'width: 40%; white-space: nowrap'})
            rows = table.find_all('tr')
            for row in rows:
                cells = row.find_all('td')
                if len(cells) == 2 and cells[0].text.strip() == 'Current Until:':
                    current_until = cells[1].text.strip()
                    return current_until
            return "invalid id"
        else:
            print("Cannot able to load the page please check internet connection.")
    except Exception as e:
        print("ID is wrong")
        return "invalid id"
        
for row in range(2, end_row+1):
    print(f'Executing {row-1} id')
    try:
        cell=f'{cert_col}{row}'
        cert_id=ws[cell].value
        valid_cur_row=f'{valid_col}{row}'
        if cert_id is None:
            ws[valid_cur_row]='invalid id'
            continue
        valid_date=retrieve_current_until(cert_id)
        ws[valid_cur_row]=valid_date
    except Exception as e:
        print(e)
        wb.save(book_name)

wb.save(book_name)
    



