import requests
import json
import time
from datetime import datetime, timedelta, date as date_type, time as time_type
from zoneinfo import ZoneInfo
import tkinter as tk
from tkcalendar import DateEntry
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment


# === КОНФИГУРАЦИЯ ===

def load_webhook():
    try:
        with open("webhook.txt", "r", encoding="utf-8") as f:
            url = f.read().strip()
            if not url:
                raise ValueError("Файл webhook.txt пустой")
            if not url.endswith("/"):
                url += "/"
            return url
    except FileNotFoundError:
        print("Файл webhook.txt не найден. Создайте его в папке с проектом.")
        exit()
    except Exception as e:
        print(f"Ошибка чтения webhook.txt: {e}")
        exit()


WEBHOOK_URL = load_webhook()
PORTAL_URL = WEBHOOK_URL.split("/rest/")[0]

USER_TZ = ZoneInfo("Asia/Vladivostok")

WORK_START = time_type(9, 0)
WORK_END = time_type(18, 0)

WORK_INTERVALS = [
    (time_type(9, 0), time_type(11, 0)),
    (time_type(11, 15), time_type(13, 0)),
    (time_type(14, 0), time_type(16, 0)),
    (time_type(16, 15), time_type(18, 0)),
]

LONG_TASK_THRESHOLD = 4 * 60  # 4 часа = 240 минут

TASK_TITLE_PREFIX = "Запрос на ЦП"


def get_task_url(task_id, user_id):
    """Формирует ссылку на задачу"""
    return f"{PORTAL_URL}/company/personal/user/{user_id}/tasks/task/view/{task_id}/"


def parse_date(date_str):
    """Парсит дату из ISO формата, сохраняя часовой пояс"""
    if not date_str:
        return None
    try:
        return datetime.fromisoformat(date_str)
    except Exception:
        return None


def format_minutes(minutes):
    """Форматирует минуты в читаемый вид"""
    if minutes <= 0:
        return "0 мин"
    hours = minutes // 60
    mins = minutes % 60
    if hours > 0:
        return f"{hours} ч {mins} мин"
    return f"{mins} мин"


# === РАСЧЁТ РАБОЧЕГО ВРЕМЕНИ ===

def is_working_day(d):
    """Проверяет, является ли день рабочим (пн-пт)"""
    return d.weekday() < 5


def time_to_minutes(t):
    """Переводит время в минуты от начала дня"""
    return t.hour * 60 + t.minute


def working_minutes_in_interval(start_t, end_t, interval_start, interval_end):
    """Считает рабочие минуты в пересечении двух интервалов"""
    s = max(time_to_minutes(start_t), time_to_minutes(interval_start))
    e = min(time_to_minutes(end_t), time_to_minutes(interval_end))
    return max(0, e - s)


def working_minutes_in_day(d, start_t=None, end_t=None):
    """Считает рабочие минуты в конкретном дне"""
    if not is_working_day(d):
        return 0

    if start_t is None:
        start_t = WORK_START
    if end_t is None:
        end_t = WORK_END

    if time_to_minutes(start_t) >= time_to_minutes(WORK_END):
        return 0
    if time_to_minutes(end_t) <= time_to_minutes(WORK_START):
        return 0

    start_t = max(start_t, WORK_START)
    end_t = min(end_t, WORK_END)

    total = 0
    for interval_start, interval_end in WORK_INTERVALS:
        total += working_minutes_in_interval(start_t, end_t, interval_start, interval_end)

    return total


def calculate_working_minutes(start_dt, end_dt):
    """Главная функция — считает рабочие минуты между двумя datetime"""
    start_vl = start_dt.astimezone(USER_TZ)
    end_vl = end_dt.astimezone(USER_TZ)

    if start_vl >= end_vl:
        return 0

    start_date = start_vl.date()
    end_date = end_vl.date()

    if start_date == end_date:
        return working_minutes_in_day(start_date, start_vl.time(), end_vl.time())

    total = 0

    total += working_minutes_in_day(start_date, start_vl.time(), WORK_END)

    current = start_date + timedelta(days=1)
    while current < end_date:
        total += working_minutes_in_day(current)
        current += timedelta(days=1)

    total += working_minutes_in_day(end_date, WORK_START, end_vl.time())

    return total


# === ВЫГРУЗКА ИЗ БИТРИКС24 ===

def get_all_tasks():
    """Выгружает все задачи из Битрикс24"""
    all_tasks = []
    start = 0

    while True:
        payload = {
            "select": [
                "ID", "TITLE", "STATUS", "RESPONSIBLE_ID",
                "CLOSED_DATE", "CREATED_DATE"
            ],
            "start": start
        }

        try:
            response = requests.post(
                f"{WEBHOOK_URL}tasks.task.list.json",
                json=payload,
                timeout=30
            )
        except Exception as e:
            print(f"Ошибка подключения: {e}")
            break

        if response.status_code != 200:
            print(f"HTTP ошибка: {response.status_code}")
            break

        try:
            result = response.json()
        except json.JSONDecodeError:
            break

        if "error" in result:
            print(f"Ошибка API: {result['error']}")
            break

        tasks = result.get("result", {}).get("tasks", [])
        if not tasks:
            break

        all_tasks.extend(tasks)
        print(f"Загружено {len(all_tasks)} задач...")

        next_start = result.get("next")
        if not next_start:
            break

        start = next_start
        time.sleep(0.2)

    return all_tasks


def filter_tasks(tasks, date_from, date_to):
    """Фильтрует задачи по периоду и названию (с учётом часовых поясов)"""
    filtered = []

    date_from_dt = datetime.combine(date_from, time_type.min).replace(tzinfo=USER_TZ)
    date_to_dt = datetime.combine(date_to, time_type.max).replace(tzinfo=USER_TZ)

    for task in tasks:
        title = task.get("title", "")

        if not title.startswith(TASK_TITLE_PREFIX):
            continue

        created = parse_date(task.get("createdDate"))
        if not created:
            continue

        created_vl = created.astimezone(USER_TZ)

        if not (date_from_dt <= created_vl <= date_to_dt):
            continue

        filtered.append(task)

    return filtered


def is_task_closed(task):
    """Проверяет, закрыта ли задача (статус 5)"""
    return str(task.get("status")) == "5"


def save_to_excel(tasks, filename="tasks_report.xlsx"):
    """Сохраняет задачи в Excel с форматированием"""
    wb = Workbook()
    ws = wb.active
    ws.title = "Задачи"

    headers = ["Название", "Дата создания", "Дата закрытия", "Время выполнения", "Ссылка"]
    ws.append(headers)

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")

    for task in tasks:
        title = task.get("title", "")
        closed = parse_date(task.get("closedDate"))
        created = parse_date(task.get("createdDate"))
        url = get_task_url(task["id"], task.get("responsibleId", "1"))

        created_vl = created.astimezone(USER_TZ) if created else None
        closed_vl = closed.astimezone(USER_TZ) if closed else None

        created_str = created_vl.strftime("%d.%m.%Y %H:%M") if created_vl else "N/A"

        # Для закрытых и открытых задач — разная логика
        if is_task_closed(task) and closed_vl:
            # Закрытая задача
            closed_str = closed_vl.strftime("%d.%m.%Y %H:%M")
            working_mins = calculate_working_minutes(created, closed)
            working_str = format_minutes(working_mins)
            is_long_task = working_mins > LONG_TASK_THRESHOLD
        else:
            # Открытая задача
            closed_str = "В процессе"
            working_str = "Выполняется"
            is_long_task = False  # Открытые не красим

        row = [title, created_str, closed_str, working_str, url]
        ws.append(row)

        if is_long_task:
            for cell in ws[ws.max_row]:
                cell.fill = yellow_fill

    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 70

    wb.save(filename)
    print(f"Сохранено в {filename}")


def start_export():
    """Запускает процесс выгрузки после выбора дат"""
    date_from = cal_from.get_date()
    date_to = cal_to.get_date()

    print(f"\nПериод: с {date_from} по {date_to}")
    print(f"Часовой пояс: {USER_TZ}")
    print(f"Порог для жёлтого: {format_minutes(LONG_TASK_THRESHOLD)} рабочего времени")
    print("Начинаю выгрузку задач...\n")

    all_tasks = get_all_tasks()
    print(f"\nВсего задач в системе: {len(all_tasks)}")

    filtered = filter_tasks(all_tasks, date_from, date_to)
    print(f"Отфильтровано задач: {len(filtered)}")

    open_count = sum(1 for t in filtered if not is_task_closed(t))
    closed_count = len(filtered) - open_count
    print(f"Из них: закрытых — {closed_count}, открытых — {open_count}")

    if filtered:
        save_to_excel(filtered, "tasks_report.xlsx")
    else:
        print("Нет задач, подходящих под условия")


# === GUI ===
root = tk.Tk()
root.title("Тест")
root.geometry("550x350")

title_label = tk.Label(root, text="Выберите период", font=("Arial", 14, "bold"))
title_label.pack(pady=10)

frame_from = tk.Frame(root)
frame_from.pack(pady=5)
tk.Label(frame_from, text="С даты: ", width=10, anchor="w").pack(side=tk.LEFT)
cal_from = DateEntry(frame_from, width=15, date_pattern="yyyy-mm-dd")
cal_from.pack(side=tk.LEFT)

frame_to = tk.Frame(root)
frame_to.pack(pady=5)
tk.Label(frame_to, text="По дату: ", width=10, anchor="w").pack(side=tk.LEFT)
cal_to = DateEntry(frame_to, width=15, date_pattern="yyyy-mm-dd")
cal_to.pack(side=tk.LEFT)

btn = tk.Button(root, text="Выгрузить задачи", command=start_export,
                bg="#4472C4", fg="white", font=("Arial", 12, "bold"),
                padx=20, pady=10)
btn.pack(pady=20)

info = tk.Label(root,
                text="ООО Современная Школа",
                fg="gray", justify="center")
info.pack()

root.mainloop()