import requests
import json
import time
from datetime import datetime

#neosintez_id = '993f811a-c563-ec11-911b-005056b6948b'
url = 'http://construction.irkutskoil.ru/'
path = 'path.txt' # путь к директории, в которой должен быть файл roots.txt и в которую будут сохраняться выгрузки
roots_file = 'roots.txt'

def authentification(url, aut_string):  # функция возвращает токен для атуентификации. Учетные данные берет из файла
    req_url = url + 'connect/token'
    payload = aut_string  # строка вида grant_type=password&username=????&password=??????&client_id=??????&client_secret=??????
    headers = {
        'Content-Type': 'application/x-www-form-urlencoded'
    }
    response = requests.post(req_url, data=payload, headers=headers)
    if response.status_code == 200:
        token = json.loads(response.text)['access_token']
    else:
        token = ''
    return token

def export_classes(url, token, neosintez_id):
    req_url = url + f'api/objects/export/classes'  # id сущности, в которой меняем атрибут
    payload = json.dumps([neosintez_id])  # тело запроса в виде списка/словаря
    headers = {
        'Accept': 'application/json',
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json-patch+json',
        'Referer': f'{url}objects?id={neosintez_id}]',
        'X-HTTP-Method-Override': 'GET',
        'X-Requested-With': 'XMLHttpRequest'
    }
    response = requests.post(req_url, headers=headers, data=payload)
    if response.status_code != 200:
        # print(req_body)
        # print(response.text)
        pass
    return response


def export(url, token, neosintez_id, part):  # для получения id сессии для скачивания файла
    req_url = url + f'api/objects/export'
    payload = json.dumps({
        "IncludeContent": False,
        "IncludeRoots": False,
        "FileName": "",
        "GroupColumnsByClass": False,
        "Roots": [
        neosintez_id
         ],
         "OutputFormat": 1,
         "Classes": part
    })  # тело запроса в виде списка/словаря

    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Encoding': 'gzip, deflate, br',
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json;charset=UTF-8',
        'Referer': f'{url}objects?id={neosintez_id}',
        'X-Requested-With': 'XMLHttpRequest'
    }
    response = requests.post(req_url, headers=headers, data=payload)
    if response.status_code != 200:
        # print(payload)
        # print(response.text)
        pass
    return response


def negotiate(url, token, id):  # запрос post перед запросом на скачивание файла нужен обязательно
    req_url = url + f'api/objects/export/{id}'
    headers = {
        'Accept': 'application/json',
        'Host': 'construction.irkutskoil.ru',
        'Authorization': f'Bearer {token}'

    }
    requests.post(req_url, headers=headers)


def download(url, token, id):  # скачать файл по id сессии
    req_url = url + f'api/objects/export/{id}'
    headers = {
        'Accept': 'application/json',
        'Host': 'construction.irkutskoil.ru',
        'Authorization': f'Bearer {token}'

    }
    response = requests.get(req_url, headers=headers)

    return response


def get_excel(neosintez_id, name):
    # получение классов
    classes = export_classes(url,token,neosintez_id)
    classes = json.loads(classes.text)
    # получение id сесии экспорта
    response = export(url,token,neosintez_id,classes)
    download_id = json.loads(response.text)['Id']
    # получение файла и скачивание его
    negotiate(url, token, download_id) #  переговоры. Запрос перед get запросом
    time.sleep(120) # ожидание подготовки файла ожидание или ожидания после переговоров
    excel = download(url, token, download_id)
    print(excel.headers)
    if 'spreadsheetml' in excel.headers['Content-Type']:
        message = f'Успешный эскпорт {name} из узла {neosintez_id}'
    else:
        message = f'Ошибка экспорта {name} из узла {neosintez_id}'
    file_name = f'{name}_{datetime.now().strftime("%Y-%m-%d_%H.%M.%S")}.xlsx'
    file_path = path + file_name
    with open(file_path, 'wb') as f:
        f.write(excel.content)
    return message


#  путь к директории для сохранения файлов
with open(path, encoding='utf-8') as p:
    path = p.read()

# токен для авторизации
with open('auth_data.txt') as f:
    aut_string = f.read()
token = authentification(url=url, aut_string=aut_string)

#  чтение файла roots для определения объектов для экспорта
roots_file_path = path + roots_file
with open(roots_file_path, encoding='utf-8') as f:
    roots_list = [tuple(r.split(';')) for r in f.read().split('\n')] #  создание списка кортежей названия узлов и их id


for root in roots_list:
    name = root[0]
    neosintez_id = root[1]
    print(get_excel(neosintez_id=neosintez_id, name=name))