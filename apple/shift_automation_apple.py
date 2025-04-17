import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, jsonify, render_template, request
from flask_cors import CORS
import string
from dotenv import load_dotenv
import re, os
import subprocess

app = Flask(__name__)
CORS(app)

load_dotenv(dotenv_path="apple/.env")

json_path = os.getenv("GOOGLE_JSON_PATH")
credentials = Credentials.from_service_account_file(
    json_path,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
gs = gspread.authorize(credentials)

shift_spreadsheet_id = ""
master_spreadsheet_id = "16y2vkOALlolyPpcZ86_7XLQAARrAMNrH1wS31pqLm5Q"

def url_spreadsheet_id(url):
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    return match.group(1) if match else None

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/reset_last_row', methods=['POST'])
def reset_last_row():
    try:
        if os.path.exists('last_row.txt'):
            os.remove('last_row.txt')
            return jsonify({"message": "最終処理行がリセットされました"}), 200
        else:
            return jsonify({"message": "リセットするファイルがありません"}), 404
    except Exception as e:
        return jsonify({"error": f"ファイルの削除中にエラーが発生しました: {str(e)}"}), 500

@app.route('/set_shift_spreadsheet', methods=['POST'])
def set_shift_spreadsheet():
    global shift_spreadsheet_id
    data = request.get_json()
    url = data.get("spreadsheet_url", "").strip()
    shift_spreadsheet_id = url_spreadsheet_id(url)
    if not shift_spreadsheet_id:
        return jsonify({"error": "スプレッドシートのURLが無効です"}), 400
    return jsonify({"message": "スプレッドシートIDが設定されました"})

@app.route('/process_shift', methods=['POST'])
def process_shift():
    global shift_spreadsheet_id
    if not shift_spreadsheet_id:
        return jsonify({"error": "スプレッドシートが設定されていません"}), 400

    data = request.get_json()
    raw_column_number = data.get("column_number", "").strip()
    column_count = int(data.get("column_count", 0))

    column_number = col_letter_to_number(raw_column_number)
    if not column_number:
        return jsonify({"error": "有効な列番号を入力してください（例: 3 または C）"}), 400

    shift_spreadsheet = gs.open_by_key(shift_spreadsheet_id)
    global shift_worksheet
    shift_worksheet = shift_spreadsheet.get_worksheet(0)

    last_processed_row = get_last_processed_row()
    updated_row = process_shifts(column_number, column_count) or []
    updated_list = [msg.split(" のシフト")[0] for msg in updated_row]


    response = {
        "last_processed_row": last_processed_row,
        "message": updated_row,
        "updated_list": updated_list,
        "info_message": "初回実行として処理します"
    }
    return jsonify(response)

def col_letter_to_number(col):
    col = col.upper()
    if col.isdigit():
        return int(col)
    result = 0
    for char in col:
        if char in string.ascii_uppercase:
            result = result * 26 + (ord(char) - ord('A') + 1)
        else:
            return None
    return result

def get_last_processed_row():
    try:
        with open('last_row.txt', 'r') as file:
            return int(file.read().strip() or 0)
    except FileNotFoundError:
        return 0

def save_last_processed_row(last_row):
    with open('last_row.txt', 'w') as file:
        file.write(str(last_row))

def get_column_letter(col_num):
    if col_num <= 0:
        return ""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = string.ascii_uppercase[remainder] + result
    return result

def process_shifts(column_number, column_count):
    updates = []
    shift_spreadsheet = gs.open_by_key(shift_spreadsheet_id)
    global shift_worksheet
    shift_worksheet = shift_spreadsheet.get_worksheet(0)

    last_processed_row = get_last_processed_row()
    current_now_count = len(shift_worksheet.get_all_values())

    if current_now_count > last_processed_row:
        new_data = shift_worksheet.get_all_values()[last_processed_row:current_now_count]
        master_sheet = gs.open_by_key(master_spreadsheet_id).get_worksheet(0)
        master_data = [[row[3], row[4]] for row in master_sheet.get_all_values()]

        updated_row = []
        for row in new_data:
            shift_id = row[4].strip().replace(' ', '').replace('　', '').upper()
            shift_name = row[3].strip().replace(' ', '').replace('　', '')
            shift_status_list = [row[5 + i] if 5 + i < len(row) else "" for i in range(column_count)]

            for row_num, master_row in enumerate(master_data, start=1):
                master_id = master_row[0].strip().replace(' ', '').replace('　', '').upper()
                master_name = master_row[1].strip().replace(' ', '').replace('　', '')
                if shift_id == master_id and shift_name == master_name:
                    for i in range(column_count):
                        cell = f'{get_column_letter(column_number + i)}{row_num}'
                        updates.append({"range": cell, "values": [[shift_status_list[i]]]})
                    updated_row.append(f"{shift_name}のシフトが更新されました。")
                    break

        save_last_processed_row(last_processed_row + len(new_data))

        if updates:
            master_sheet.batch_update(updates)
            return updated_row

    return None

@app.route('/start_bot', methods=['POST'])
def start_bot():
    try:
        script_path = r"C:\Users\bunta\Downloads\shift-automation\apple\kakutei-automation-apple.py"
        subprocess.Popen(["python", script_path])
        return jsonify({"message": "Discord Bot 起動します"}), 200
    except Exception as e:
        return jsonify({"error": f"Bot起動中にエラーが発生しました: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True)