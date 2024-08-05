import os
from datetime import timedelta, date

from dotenv import load_dotenv
import httpx
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def create_csv(key, work_sheet):
    link = f"https://docs.google.com/spreadsheets/d/{
        key}/gviz/tq?tqx=out:csv&sheet={work_sheet}"
    request = httpx.get(link)
    csv_file = request.read().decode("utf-8").split("\n")
    return csv_file


def sort_list(csv_file):
    titles = list(map(lambda x: x.replace(
        '"', ''), csv_file.pop(0).split(",")))

    application_index = titles.index("Заявка до")

    hurry_up = [titles]
    today = date.today()

    for event in csv_file:
        event = list(map(lambda x: x.replace('"', ''), event.split(",")))
        try:
            day, month, year = list(
                map(int, event[application_index].split()[0].split(".")))
            event_date = date(year, month, day)
            if today >= event_date - timedelta(days=7) and not event_date < today:
                hurry_up.append(event)
        except (ValueError, IndexError):
            continue
    return hurry_up


def send_to_google(table_id, to_work_list, sorted_list):
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        "credentials.json", scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_key(table_id)
    sheet.worksheet(to_work_list).clear()
    sheet.values_append(
        f"{to_work_list}!A1",
        params={'valueInputOption': 'RAW'},
        body={'values': sorted_list}
    )


def main():
    load_dotenv(override=True)

    # ID Электронной таблицы
    sheet_key = os.getenv("ID_TABLE")
    # Лист откуда взять мероприятия
    from_work_sheet_name = os.getenv("FROM_SPREADSHEET").replace(" ", "%20")
    # Лист куда сохранить мероприятия
    to_work_sheet_name = os.getenv("TO_SPREADSHEET").replace(" ", "%20")
    send_to_google(sheet_key, to_work_sheet_name, sort_list(
        create_csv(sheet_key, from_work_sheet_name)))


if __name__ == "__main__":
    main()
