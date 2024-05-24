import os
import re
import requests
import pandas as pd
from bs4 import BeautifulSoup
from telegram import Update
from telegram.ext import ApplicationBuilder, CommandHandler, ContextTypes
from datetime import datetime

TOKEN = '6939312353:AAGXqguJv2n0xksnW5zNqtBOyhOjJF_vA_U'


# Функции для скачивания и обработки файлов расписания
def fetch_excel_file_urls():
    base_url = "https://pkps-perm.ru"
    response = requests.get(f"{base_url}/students/raspisanie/fakultet_dizayna_i_servisa_ul_chernyshevskogo_11/")
    document = BeautifulSoup(response.text, 'html.parser')
    links = document.select('a[href$=".xls"]')
    schedule_links = [(base_url + link['href'], link.get_text()) for link in links if "Изменения" in link.get_text()]
    return schedule_links


def sanitize_file_name(file_name):
    sanitized = file_name.replace(" ", "_")
    sanitized = re.sub(r"[^\w\d._-]", "", sanitized)
    if not sanitized.lower().endswith('.xls'):
        sanitized += '.xls'
    return sanitized


def download_file(url, date_str):
    response = requests.get(url)
    if response.status_code == 200:
        files_xls_dir = os.path.join(os.path.dirname(__file__), 'files_xls')
        os.makedirs(files_xls_dir, exist_ok=True)
        file_name = f"schedule_{date_str}.xls"
        file_path = os.path.join(files_xls_dir, file_name)
        with open(file_path, 'wb') as file:
            file.write(response.content)
        return file_path
    else:
        return None


# Функция для чтения расписания из файла Excel
def read_schedule(file_path, group_name):
    xls = pd.ExcelFile(file_path)
    schedule = []
    found = False

    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        group_col = None
        print(f"Checking sheet: {sheet_name}")  # Debug print

        # Попытка найти строку с названиями групп
        for row in range(5, 11):  # Ищем с 6 по 10 строку
            for col in df.columns:
                if df.iloc[row, col] == group_name:
                    group_col = col
                    found = True
                    break
            if found:
                break

        if group_col is not None:
            for row in range(row + 1, len(df)):  # Начинаем с следующей строки после найденной
                time = df.iloc[row, 1]
                subject = df.iloc[row, group_col]
                if pd.notna(subject):
                    schedule.append(f"{time}: {subject}")

            if found:
                return "\n".join(schedule)

    if not found:
        print(f"Group {group_name} not found in any sheet")  # Debug print

    return None


# Обработчик команды /start
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "Привет! Введите команду /schedule <дата в формате ДД.ММ> <название группы>, чтобы получить расписание. "
        "Пример: /schedule 24.05 ИП-22-9К"
    )


# Обработчик команды /schedule
async def schedule(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        if len(context.args) < 2:
            await update.message.reply_text("Пожалуйста, укажите дату и название группы.")
            return

        date_str = context.args[0]
        group_name = context.args[1]

        # Автоподстановка года
        if len(date_str) == 5:  # Формат ДД.ММ
            current_year = datetime.now().year
            date_str += f".{current_year}"

        # Скачиваем файлы расписаний
        urls = fetch_excel_file_urls()
        file_path = None
        for url, text in urls:
            if date_str in text:
                date_match = re.search(r'\d{2}\.\d{2}\.\d{4}', text)
                if date_match:
                    formatted_date_str = date_match.group(0).replace('.', '-')
                    downloaded_file = download_file(url, formatted_date_str)
                    if downloaded_file:
                        file_path = downloaded_file
                        break

        if not file_path:
            await update.message.reply_text(f"Не удалось найти расписание на дату {date_str}.")
            return

        schedule_text = read_schedule(file_path, group_name)
        if schedule_text:
            await update.message.reply_text(f"Расписание для группы {group_name} на {date_str}:\n\n{schedule_text}")
        else:
            await update.message.reply_text(f"Группа {group_name} не найдена.")
    except Exception as e:
        await update.message.reply_text(f"Произошла ошибка при чтении файла: {e}")


def main():
    application = ApplicationBuilder().token(TOKEN).build()

    start_handler = CommandHandler('start', start)
    schedule_handler = CommandHandler('schedule', schedule)

    application.add_handler(start_handler)
    application.add_handler(schedule_handler)

    application.run_polling()


if __name__ == '__main__':
    main()
