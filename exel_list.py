import gspread # Импортируем gspread — библиотека для работы с Google Sheets
from google.oauth2.service_account import Credentials #Импортируем класс Credentials из Google API Используется для авторизации через файл credentials.json, который ты скачал при создании сервисного аккаунта.
from datetime import datetime
import os, json


scopes = ['https://www.googleapis.com/auth/spreadsheets'] #Указываем область доступа (scopes) Это означает: «даю доступ к Google Sheets полностью».
creds = Credentials.from_service_account_file('/etc/secrets/credentials.json', scopes=scopes) #Загружаем credentials.json и создаём объект creds

client = gspread.authorize(creds) #Авторизуемся в gspread
sheet_id = '1UmZ6CD6xN5Rbt9DaEE2vZEqM1QTL0gjghEAiCo4ac-A' #Указываем ID твоей таблицы
workbook = client.open_by_key(sheet_id) #Открываем саму таблицу

sheet = workbook.worksheet('Лист2')

def write_to_sheet_1(position):
    header_row = sheet.row_values(1)
    new_col_index = len(header_row) + 1
    # обновляем дату
    today = datetime.today().strftime("%d.%m.%Y")
    sheet.update_cell(1, new_col_index, today)
    # ——— Записываем позиции построчно ———
    for row, pos in enumerate(position, start=2):  # строки начинаются со 2
        sheet.update_cell(row, new_col_index, pos)

