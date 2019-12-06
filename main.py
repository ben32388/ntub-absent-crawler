import requests

from bs4 import BeautifulSoup

from getpass import getpass

from openpyxl import Workbook

BASE_URL = 'http://ntcbadm1.ntub.edu.tw'


def get_score(username, password):
    client = requests.Session()
    res = client.get(BASE_URL)
    soup = BeautifulSoup(res.text, 'html.parser')

    data = {}
    for input_tag in soup.find_all('input'):
        name = input_tag.get('name')
        value = input_tag.get('value','')
        if name is None:
            continue
        
        data[name] = value
    
    data.update({
        'UserID':username,
        'PWD':password,
    })

    res = client.post(BASE_URL, data)
    if res.url != BASE_URL + '/Portal/indexSTD.aspx':
        print('Login fail')
        return


    res = client.get(BASE_URL + '/ACAD/STDWEB/GRD_GRDQry_All.aspx')
    if '請重新登入' in res.text:
        print('Login fail')
        return

    soup = BeautifulSoup(res.text,'html.parser')
    header, *rows = soup.select('#ctl00_ContentPlaceHolder1_GRD tr')
    
    header = [h.text for h in header.select('th')]
    # rows = [[c.text.replace('\n','') for c in row.select('td')] for row in rows]

    result = []
    for row in rows:
        c = []
        for col in row.select('td'):
            c.append(col.text.replace('\n',''))
        
        result.append(c)

    return header, result

def write_file(header, scores):
    wb = Workbook()
    sheet = wb['Sheet']

    sheet.append(header)
    for row in scores:
        sheet.append(row)

    wb.save('a.xlsx')
def main():
    username = input('Student ID:')
    password = getpass('Password: ')
    result = get_score(username, password)
    if result is None:
        return
    header, scores = result   
    write_file(header, scores)

if __name__ == '__main__':
    main()

