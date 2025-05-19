import logging
import math
import os
from datetime import date, timedelta 

import requests 
from dateutil import parser as date_parser
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv

import gspread
from google.oauth2.service_account import Credentials 

WEEEK_API_BASE_URL: str = "https://api.weeek.net/public/v1/"
API_TOKEN_ENV_NAME: str = "WEEEK_API_TOKEN"
ENV_FILE_PATH: str = ".env"

GOOGLE_SHEET_ID_ENV_NAME: str = "GOOGLE_SHEET_ID"
GOOGLE_CREDENTIALS_FILENAME_ENV_NAME: str = "GOOGLE_CREDENTIALS_FILENAME"
SHEET_NAME: str = "Отчет WEEEK"

PRIORITY_MAP: dict[int, str] = {
    0: "Низкий", 1: "Средний", 2: "Высокий", 3: "Замороженный"
}
PRIORITY_COLORS_GSHEETS: dict[str, dict[str, float]] = {
    "Замороженный": {"red": 0.36, "green": 0.73, "blue": 1},
    "Высокий": {"red": 1.0, "green": 0.6, "blue": 0.0},
    "Средний": {"red": 1.0, "green": 0.9, "blue": 0.4},
    "Низкий": {"red": 0.7, "green": 0.85, "blue": 0.5},
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(module)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

def get_env_variable(var_name: str, is_critical: bool = True) -> str | None:
    value = os.getenv(var_name)
    if not value and is_critical:
        logger.error(f"Критическая переменная окружения '{var_name}' не установлена или пуста.")
        logger.info(f"Убедитесь, что она определена в вашем окружении или в файле '{ENV_FILE_PATH}'.")
        exit(1)
    elif not value:
        logger.warning(f"Необязательная переменная окружения '{var_name}' не установлена.")
    return value

def make_api_request(endpoint: str, token: str, params: dict | None = None) -> dict | None:
    headers = {"Authorization": f"Bearer {token}"}
    target_url = WEEEK_API_BASE_URL + endpoint
    try:
        response = requests.get(target_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        logger.error(f"HTTP ошибка для {target_url}: {http_err}")
        logger.error(f"Тело ответа: {response.text if hasattr(response, 'text') else 'N/A'}")
    except requests.exceptions.Timeout:
        logger.error(f"Таймаут запроса для {target_url}")
    except requests.exceptions.RequestException as req_err:
        logger.error(f"Ошибка запроса для {target_url}: {req_err}")
    except ValueError as json_err:
        logger.error(f"Ошибка декодирования JSON от {target_url}: {json_err}")
        content = response.text if 'response' in locals() and hasattr(response, 'text') else 'N/A'
        logger.error(f"Тело ответа: {content[:500]}...")
    return None

def parse_date_string(date_str: str | None) -> date | None:
    if not date_str: return None
    try: return date_parser.parse(date_str).date()
    except (ValueError, TypeError):
        logger.warning(f"Не удалось преобразовать строку в дату: '{date_str}'")
        return None

def get_reporting_periods(current_date: date) -> tuple[tuple[date, date], tuple[date, date]]:
    day_of_month = current_date.day
    p1_start_day, p1_end_day = 6, 20
    p2_start_day, p2_end_day = 21, 5
    p1_base_date = current_date if day_of_month > 5 else current_date - relativedelta(months=1)
    period1_start = date(p1_base_date.year, p1_base_date.month, p1_start_day)
    period1_end = date(p1_base_date.year, p1_base_date.month, p1_end_day)
    p2_base_date = p1_base_date 
    period2_start = date(p2_base_date.year, p2_base_date.month, p2_start_day)
    p2_end_month_base = p2_base_date + relativedelta(months=1)
    period2_end = date(p2_end_month_base.year, p2_end_month_base.month, p2_end_day)
    return (period1_start, period1_end), (period2_start, period2_end)

def fetch_weeek_data(token: str) -> tuple[dict[str, str], list[dict], dict[tuple[int, int], str]]:
    logger.info("Начало загрузки данных из WEEEK API...")
    members_map = fetch_workspace_members(token)
    tasks_data = fetch_all_tasks(token)
    board_column_names_map: dict[tuple[int, int], str] = {}
    if tasks_data:
        unique_board_ids = {
            task.get("boardId") for task in tasks_data if task.get("boardId") is not None
        }
        for board_id in unique_board_ids:
            if not isinstance(board_id, int):
                try:
                    board_id_int = int(board_id)
                except (ValueError, TypeError):
                    logger.warning(f"Обнаружен некорректный boardId: {board_id} (тип: {type(board_id)}), который не может быть преобразован в int. Пропуск.")
                    continue
            else:
                board_id_int = board_id

            columns = fetch_board_columns_for_board(board_id_int, token)
            for column in columns:
                col_id = column.get("id")
                if isinstance(col_id, int): 
                    column_name = column.get("name", str(col_id))
                    board_column_names_map[(board_id_int, col_id)] = column_name
                elif col_id is not None:
                     logger.warning(f"ID колонки {col_id} для доски {board_id_int} не является числом. Пропуск колонки.")

    logger.info("Загрузка данных из WEEEK API завершена.")
    return members_map, tasks_data, board_column_names_map

def fetch_workspace_members(token: str) -> dict[str, str]:
    logger.info("Загрузка участников рабочей области...")
    data = make_api_request("ws/members", token)
    member_map: dict[str, str] = {}
    if data and data.get("success") and "members" in data:
        for member in data["members"]:
            member_id, first_name = member.get("id"), member.get("firstName")
            if member_id:
                member_map[member_id] = first_name if first_name else member_id
    else:
        logger.error("Не удалось загрузить участников рабочей области.")
    return member_map

def fetch_board_columns_for_board(board_id: int | str, token: str) -> list[dict]:
    data = make_api_request(f"tm/board-columns?boardId={board_id}", token)
    columns_list: list[dict] = []
    if data and data.get("success") and "boardColumns" in data:
        columns_list = data["boardColumns"]
    else:
        logger.warning(f"Не удалось загрузить колонки для доски ID: {board_id}.")
    return columns_list

def fetch_all_tasks(token: str) -> list[dict]:
    logger.info("Загрузка списка задач...")
    all_tasks_list: list[dict] = []
    cursor: str | None = None
    page_num = 0
    while True:
        page_num += 1
        params = {"cursor": cursor} if cursor else {}
        data = make_api_request("tm/tasks", token, params=params)
        if data and data.get("success") and "tasks" in data:
            tasks_on_page = data["tasks"]
            all_tasks_list.extend(tasks_on_page)
            if data.get("hasMore") and data.get("cursor"):
                cursor = data["cursor"]
            else:
                break
        else:
            logger.error(f"Ошибка при загрузке {page_num}-й страницы задач.")
            break
    return all_tasks_list

def format_value_for_sheet(value: any) -> str | int | float:
    if isinstance(value, date): return value.isoformat()
    if value is None: return ""
    return value

def process_tasks_to_sheet_rows(
    tasks_data: list[dict],
    members_map: dict[str, str],
    board_cols_map: dict[tuple[int, int], str],
    period1_dates: tuple[date, date],
    period2_dates: tuple[date, date]
) -> list[list[str | int | float]]:
    sheet_rows: list[list[str | int | float]] = []
    p1_start, p1_end = period1_dates
    p2_start, p2_end = period2_dates

    for task in tasks_data:
        priority_val = task.get("priority")
        priority_str = PRIORITY_MAP.get(priority_val, str(priority_val)) if priority_val is not None else ""

        user_id = task.get("userId")
        executor = members_map.get(user_id, user_id if user_id else "")

        task_board_id_raw = task.get("boardId")
        task_col_id_raw = task.get("boardColumnId")
        status = ""
        task_board_id = int(task_board_id_raw) if isinstance(task_board_id_raw, str) and task_board_id_raw.isdigit() else task_board_id_raw
        task_col_id = int(task_col_id_raw) if isinstance(task_col_id_raw, str) and task_col_id_raw.isdigit() else task_col_id_raw

        if isinstance(task_board_id, int) and isinstance(task_col_id, int):
             status = board_cols_map.get((task_board_id, task_col_id), str(task_col_id))
        elif isinstance(task_col_id, int) :
            status = str(task_col_id)
        elif task_col_id is not None: 
            status = str(task_col_id)


        created_at = parse_date_string(task.get("createdAt"))
        date_ended = None
        if task.get("isCompleted"):
            time_entries = task.get("timeEntries", []) or task.get("workloads", [])
            valid_entries = [entry for entry in time_entries if entry.get("date")]
            if valid_entries:
                latest_entry = sorted(valid_entries, key=lambda x: date_parser.parse(x["date"]).date(), reverse=True)[0]
                date_ended = parse_date_string(latest_entry["date"])
            elif task.get("updatedAt"):
                date_ended = parse_date_string(task.get("updatedAt"))

        est_minutes = task.get("duration")
        est_hours = math.ceil(est_minutes / 60) if est_minutes and est_minutes > 0 else None

        time_p1_min, time_p2_min = 0, 0
        comments_p1, comments_p2 = [], []
        for workload in task.get("workloads", []):
            workload_date = parse_date_string(workload.get("date"))
            duration = workload.get("duration", 0) or 0
            comment = workload.get("comment", "")
            if workload_date and duration > 0:
                if p1_start <= workload_date <= p1_end:
                    time_p1_min += duration
                    if comment: comments_p1.append(comment)
                if p2_start <= workload_date <= p2_end:
                    time_p2_min += duration
                    if comment: comments_p2.append(comment)
        
        comments_p1_str = "\n".join(f"{i+1}. {c}" for i, c in enumerate(comments_p1) if c)
        comments_p2_str = "\n".join(f"{i+1}. {c}" for i, c in enumerate(comments_p2) if c)
        time_p1_hours = math.ceil(time_p1_min / 60) if time_p1_min > 0 else None
        time_p2_hours = math.ceil(time_p2_min / 60) if time_p2_min > 0 else None

        row = [
            f"{task.get('title', '')} ({task.get('id', '')})",
            priority_str, executor, status,
            created_at, date_ended,
            est_hours, time_p1_hours, comments_p1_str,
            time_p2_hours, comments_p2_str,
            None, None
        ]
        sheet_rows.append([format_value_for_sheet(val) for val in row])
        
    return sheet_rows

def get_gspread_client(credentials_file: str) -> gspread.Client | None:
    logger.info(f"Авторизация в Google Sheets API...")
    try:
        creds = Credentials.from_service_account_file(
            credentials_file,
            scopes=['https://www.googleapis.com/auth/spreadsheets']
        )
        client = gspread.authorize(creds)
        logger.info("Авторизация в Google Sheets API прошла успешно.")
        return client
    except FileNotFoundError:
        logger.error(f"Файл учетных данных Google '{credentials_file}' не найден.")
    except Exception as e:
        logger.error(f"Ошибка авторизации в Google Sheets API: {e}")
    return None

def update_google_sheet(
    client: gspread.Client,
    sheet_id: str,
    sheet_name_target: str,
    headers_list: list[str],
    data_to_write: list[list[str | int | float]]
) -> None:
    try:
        spreadsheet = client.open_by_key(sheet_id)
        try:
            worksheet = spreadsheet.worksheet(sheet_name_target)
            worksheet.clear()
        except gspread.exceptions.WorksheetNotFound:
            logger.info(f"Лист '{sheet_name_target}' не найден, создается новый...")
            worksheet = spreadsheet.add_worksheet(
                title=sheet_name_target,
                rows=str(max(100, len(data_to_write) + 1)),
                cols=str(max(20, len(headers_list)))
            )

        all_sheet_data = [headers_list] + data_to_write
        worksheet.update(all_sheet_data, value_input_option='USER_ENTERED')

        format_requests_batch = []
        sheet_id_gid = worksheet.id

        format_requests_batch.append({
            "repeatCell": {
                "range": {"sheetId": sheet_id_gid, "startRowIndex": 0, "endRowIndex": 1,
                          "startColumnIndex": 0, "endColumnIndex": len(headers_list)},
                "cell": {"userEnteredFormat": {
                    "textFormat": {"bold": True, "fontFamily": "Times New Roman", "fontSize": 12},
                    "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE"}},
                "fields": "userEnteredFormat(textFormat,horizontalAlignment,verticalAlignment)"}})
        
        header_map = {name: index for index, name in enumerate(headers_list)}
        priority_col_index = header_map.get("Приоритет")
        
        tracked_time_headers = [h for h in headers_list if h.startswith("Затрекано (")]
        col_h_index = col_j_index = col_l_index = salary_col_index = None
        
        if len(tracked_time_headers) >= 1 and tracked_time_headers[0] in header_map:
            col_h_index = header_map[tracked_time_headers[0]]
        if len(tracked_time_headers) >= 2 and tracked_time_headers[1] in header_map:
            col_j_index = header_map[tracked_time_headers[1]]
        if "Стоимость в час" in header_map: col_l_index = header_map["Стоимость в час"]
        if "Расчет зп" in header_map: salary_col_index = header_map["Расчет зп"]

        for row_num_based, row_data in enumerate(data_to_write, start=2):
            row_index_based = row_num_based - 1

            if salary_col_index is not None and col_h_index is not None and \
               col_j_index is not None and col_l_index is not None:
                h_letter = gspread.utils.rowcol_to_a1(1, col_h_index + 1)[:-1] 
                j_letter = gspread.utils.rowcol_to_a1(1, col_j_index + 1)[:-1]
                l_letter = gspread.utils.rowcol_to_a1(1, col_l_index + 1)[:-1]
                
                formula = f"=(N({h_letter}{row_num_based}) + N({j_letter}{row_num_based})) * N({l_letter}{row_num_based})"
                worksheet.update_acell(gspread.utils.rowcol_to_a1(row_num_based, salary_col_index + 1), f"={formula}")
                
                format_requests_batch.append({
                    "repeatCell": {
                        "range": {"sheetId": sheet_id_gid, "startRowIndex": row_index_based, "endRowIndex": row_index_based + 1,
                                  "startColumnIndex": salary_col_index, "endColumnIndex": salary_col_index + 1},
                        "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0.00"}}},
                        "fields": "userEnteredFormat.numberFormat"}})
            
            if priority_col_index is not None and priority_col_index < len(row_data):
                priority_text = str(row_data[priority_col_index])
                bg_color = PRIORITY_COLORS_GSHEETS.get(priority_text)
                if bg_color:
                    format_requests_batch.append({
                        "repeatCell": {
                            "range": {"sheetId": sheet_id_gid, "startRowIndex": row_index_based, "endRowIndex": row_index_based + 1,
                                      "startColumnIndex": priority_col_index, "endColumnIndex": priority_col_index + 1},
                            "cell": {"userEnteredFormat": {"backgroundColor": bg_color}},
                            "fields": "userEnteredFormat.backgroundColor"}})
            
            format_requests_batch.append({
                "repeatCell": {
                    "range": {"sheetId": sheet_id_gid, "startRowIndex": row_index_based, "endRowIndex": row_index_based + 1,
                              "startColumnIndex": 0, "endColumnIndex": len(headers_list)},
                    "cell": {"userEnteredFormat": {
                        "textFormat": {"fontFamily": "Times New Roman", "fontSize": 12},
                        "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP"}},
                    "fields": "userEnteredFormat(textFormat,verticalAlignment,wrapStrategy)"}})

            for col_index_based, header_name in enumerate(headers_list):
                horz_align = "LEFT" if header_name == "Задача (Номер)" or header_name.startswith("Комментарии") else "CENTER"
                format_requests_batch.append({
                    "repeatCell": {
                        "range": {"sheetId": sheet_id_gid, "startRowIndex": row_index_based, "endRowIndex": row_index_based + 1,
                                  "startColumnIndex": col_index_based, "endColumnIndex": col_index_based + 1},
                        "cell": {"userEnteredFormat": {"horizontalAlignment": horz_align}},
                        "fields": "userEnteredFormat.horizontalAlignment"}})
                
                cell_value = row_data[col_index_based]
                if (header_name == "Оценка времени" or header_name.startswith("Затрекано (")) and \
                   (salary_col_index is None or col_index_based != salary_col_index): 
                    if isinstance(cell_value, (int, float)): 
                        format_requests_batch.append({
                            "repeatCell": {
                                "range": {"sheetId": sheet_id_gid, "startRowIndex": row_index_based, "endRowIndex": row_index_based + 1,
                                          "startColumnIndex": col_index_based, "endColumnIndex": col_index_based + 1},
                                "cell": {"userEnteredFormat": {"numberFormat": {"type": "NUMBER", "pattern": "0"}}},
                                "fields": "userEnteredFormat.numberFormat"}})
        
        if headers_list:
            format_requests_batch.append({
                "autoResizeDimensions": {
                    "dimensions": {"sheetId": sheet_id_gid, "dimension": "COLUMNS",
                                   "startIndex": 0, "endIndex": len(headers_list)}}})

        if format_requests_batch:
            try:
                spreadsheet.batch_update({"requests": format_requests_batch})
            except gspread.exceptions.APIError as e_api:
                err_details = e_api.response.json() if hasattr(e_api, 'response') and hasattr(e_api.response, 'json') else str(e_api)
                logger.error(f"Ошибка API Google Sheets при пакетном обновлении: {err_details}")
            except Exception as e_gen:
                logger.error(f"Непредвиденная ошибка при пакетном обновлении: {e_gen}")

        logger.info(f"Данные успешно записаны и отформатированы в Google Таблице.")

    except gspread.exceptions.SpreadsheetNotFound:
        logger.error(f"Google Таблица не найдена. Проверьте ID и права доступа сервисного аккаунта.")
    except gspread.exceptions.APIError as e:
        err_details = e.response.json() if hasattr(e, 'response') and hasattr(e.response, 'json') else str(e)
        logger.error(f"Общая ошибка Google Sheets API: {err_details}")
    except Exception as e:
        logger.error(f"Непредвиденная ошибка при работе с Google Таблицей: {e}", exc_info=True)

def main() -> None:
    logger.info("Запуск скрипта генерации отчета WEEEK.")
    if not load_dotenv(ENV_FILE_PATH):
        logger.info(f"Файл '{ENV_FILE_PATH}' не найден. Используются системные переменные окружения.")
    
    weeek_api_token = get_env_variable(API_TOKEN_ENV_NAME)
    google_sheet_id = get_env_variable(GOOGLE_SHEET_ID_ENV_NAME)
    google_creds_file = get_env_variable(GOOGLE_CREDENTIALS_FILENAME_ENV_NAME)
    
    if not (weeek_api_token and google_sheet_id and google_creds_file):
        logger.error("Одна или несколько критических переменных окружения отсутствуют. Завершение работы.")
        return

    gs_client = get_gspread_client(google_creds_file)
    if not gs_client:
        logger.error("Не удалось инициализировать клиент Google Sheets. Завершение работы.")
        return
    
    current_day = date.today()
    (p1_start, p1_end), (p2_start, p2_end) = get_reporting_periods(current_day)

    members, tasks, board_cols_map = fetch_weeek_data(weeek_api_token)

    if not tasks:
        logger.warning("Данные по задачам не загружены. Формирование отчета невозможно.")
        return
    
    sheet_rows_data = process_tasks_to_sheet_rows(
        tasks, members, board_cols_map, (p1_start, p1_end), (p2_start, p2_end)
    )

    report_headers = [
        "Задача (Номер)", "Приоритет", "Исполнитель", "Статус",
        "Дата начала", "Дата окончания", "Оценка времени",
        f"Затрекано ({p1_start:%d.%m} - {p1_end:%d.%m})", 
        "Комментарии", 
        f"Затрекано ({p2_start:%d.%m} - {p2_end:%d.%m})", 
        "Комментарии", 
        "Стоимость в час", "Расчет зп"
    ]
    
    update_google_sheet(gs_client, google_sheet_id, SHEET_NAME, report_headers, sheet_rows_data)
    logger.info("Скрипт генерации отчета WEEEK успешно завершил работу.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.critical(f"Критическая ошибка в процессе выполнения скрипта: {e}", exc_info=True)