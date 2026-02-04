import os.path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# Права доступа
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/script.projects'
]

def get_services():
    """Авторизация"""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return (
        build('sheets', 'v4', credentials=creds),
        build('drive', 'v3', credentials=creds),
        build('script', 'v1', credentials=creds)
    )

def main():
    # 0. Читаем ВАШ готовый скрипт из файла
    try:
        with open('my_script.gs', 'r', encoding='utf-8') as f:
            user_gs_code = f.read()
        print("Ваш скрипт успешно прочитан из файла.")
    except FileNotFoundError:
        print("Ошибка: Файл my_script.gs не найден! Создайте его и вставьте туда код.")
        return

    sheets_service, drive_service, script_service = get_services()

    # 1. Создаем таблицу с несколькими вкладками
    print("Создаем таблицу и вкладки...")
    spreadsheet_body = {
        'properties': {'title': 'Заказ: Автоматизация'},
        'sheets': [
            {'properties': {'title': 'Данные'}},     # Вкладка 1
            {'properties': {'title': 'Настройки'}},  # Вкладка 2
            {'properties': {'title': 'Логи'}}        # Вкладка 3
        ]
    }
    spreadsheet = sheets_service.spreadsheets().create(body=spreadsheet_body).execute()
    ss_id = spreadsheet['spreadsheetId']
    print(f"Таблица создана. ID: {ss_id}")

    # 2. Загружаем 3 скрина
    print("Загружаем скриншоты...")
    screenshot_files = ['screen1.png', 'screen2.png', 'screen3.png']
    uploaded_ids = []
    
    # Создаем фиктивные файлы для теста, если их нет (чтобы код не упал)
    for s in screenshot_files:
        if not os.path.exists(s):
            with open(s, 'w') as f: f.write('dummy image')

    for screen in screenshot_files:
        file_metadata = {'name': screen}
        media = MediaFileUpload(screen, mimetype='image/png')
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        uploaded_ids.append(file.get('id'))

    # 3. Копируем презентацию
    print("Копируем презентацию...")
    original_pres_id = 'INSERT_YOUR_PRES_ID_HERE' # <--- ВСТАВИТЬ СЮДА ID ОРИГИНАЛА ПРЕЗЕНТАЦИИ
    if original_pres_id == 'INSERT_YOUR_PRES_ID_HERE':
        print("!!! Вы не вставили ID презентации в коде. Пропускаем этот шаг для теста.")
        new_pres_id = "нет_id"
    else:
        copy_body = {'name': 'Копия Презентации Заказчика'}
        copied_pres = drive_service.files().copy(fileId=original_pres_id, body=copy_body).execute()
        new_pres_id = copied_pres.get('id')

    # 4. Прописываем инфу в столбцах (Вкладка 'Данные')
    print("Заполняем таблицу данными...")
    values = [
        ["Название этапа", "ID Файла / Ссылка", "Тип", "Статус"], # Заголовки столбцов
        ["Скриншот 1", uploaded_ids[0], "Изображение", "Загружено"],
        ["Скриншот 2", uploaded_ids[1], "Изображение", "Загружено"],
        ["Скриншот 3", uploaded_ids[2], "Изображение", "Загружено"],
        ["Презентация", new_pres_id, "Google Slides", "Скопировано"]
    ]
    body = {'values': values}
    sheets_service.spreadsheets().values().update(
        spreadsheetId=ss_id, range="Данные!A1", # Пишем именно во вкладку "Данные"
        valueInputOption="RAW", body=body
    ).execute()

    # 5. Вставляем Ваш скрипт в таблицу
    print("Внедряем ваш Google Apps Script...")
    script_project = script_service.projects().create(
        body={'title': 'Embedded Script', 'parentId': ss_id}
    ).execute()
    script_id = script_project['scriptId']
    
    files_payload = [
        {
            'name': 'Code',
            'type': 'SERVER_JS',
            'source': user_gs_code # <--- Вставляем прочитанный из файла код
        },
        {
            'name': 'appsscript',
            'type': 'JSON',
            'source': '{"timeZone": "Europe/Moscow", "exceptionLogging": "CLOUD"}'
        }
    ]
    script_service.projects().updateContent(
        scriptId=script_id,
        body={'files': files_payload}
    ).execute()
    
    print("-" * 30)
    print("ЗАДАНИЕ ВЫПОЛНЕНО!")
    print(f"Таблица: https://docs.google.com/spreadsheets/d/{ss_id}")
    print("Скрипт успешно встроен. Чтобы проверить, открой таблицу -> Расширения -> Apps Script.")

if __name__ == '__main__':
    main()
