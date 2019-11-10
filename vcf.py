#! /usr/bin/env python
# -*- coding: utf-8 -*-
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials


# Файл, полученный в Google Developer Console
CREDENTIALS_FILE = 'creds.json'
# ID Google Sheets документа (его можно взять из URL)
spreadsheet_id = '12345678901234567890123456789012345678901234'

# Авторизуемся и получаем service — экземпляр доступа к API
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    CREDENTIALS_FILE,
    ['https://www.googleapis.com/auth/spreadsheets',
     'https://www.googleapis.com/auth/drive'])
httpAuth = credentials.authorize(httplib2.Http())
service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)


# Получаем массивы, содержащие ФИО и номер телефона
def get_rows(our_range):
    values = service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=our_range,
        majorDimension='ROWS'
    ).execute()
    return values.get('values')


# Записываем контакты в файл формата vcf
def vcf_writer(result_file_path, selected_range):
    new_file = open(result_file_path, 'w')
    for person in get_rows(selected_range):
        new_file.write("BEGIN:VCARD\n"
                       "VERSION:2.1\n"
                       "FN:%s %s\n"
                       "TEL;TYPE=WORK:%s\n"
                       "END:VCARD\n"
                       % ('First event', person[0], person[1].strip("\n")))
    new_file.close()


vcf_writer(result_file_path='C:/Users/USERNAME/Desktop/result.vcf', selected_range='A1:B3')
