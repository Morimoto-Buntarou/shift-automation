# shift_kakutei_integration.py

import gspread
from google.oauth2.service_account import Credentials
from flask import Flask, jsonify, render_template, request
from flask_cors import CORS
import string
from dotenv import load_dotenv
import json
import re, os
import asyncio
import discord
import traceback
from threading import Thread, Event
import datetime

app = Flask(__name__)
CORS(app)
load_dotenv()

stop_event = Event()

# 認証情報
service_account_json = os.getenv("GOOGLE_CREDENTIALS")
credentials_info = json.loads(service_account_json)
credentials_info["private_key"] = credentials_info["private_key"].replace('\\n', '\n')
credentials = Credentials.from_service_account_info(
    credentials_info,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
gs = gspread.authorize(credentials)

shift_spreadsheet_id = ""
master_spreadsheet_id = os.getenv("MASTER_SPREADSHEET_ID")

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
    match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    shift_spreadsheet_id = match.group(1) if match else None
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

    shift_worksheet = gs.open_by_key(shift_spreadsheet_id).get_worksheet(0)
    master_sheet = gs.open_by_key(master_spreadsheet_id).get_worksheet(0)

    last_processed_row = get_last_processed_row()
    new_data = shift_worksheet.get_all_values()[last_processed_row:]
    master_data = [[row[3], row[4]] for row in master_sheet.get_all_values()]

    updates = []
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
                    updates.append({"range": cell, "values": [[shift_status_list[i]]]} )
                updated_row.append(f"{shift_name}のシフトが更新されました。")
                break

    save_last_processed_row(last_processed_row + len(new_data))
    if updates:
        master_sheet.batch_update(updates)

    updated_list = [msg.split(" のシフト")[0] for msg in updated_row]
    return jsonify({
        "last_processed_row": last_processed_row,
        "message": updated_row,
        "updated_list": updated_list,
        "info_message": "初回実行として処理します"
    })

@app.route('/start_bot', methods=['POST'])
def start_bot():
    try:
        thread = Thread(target=run_discord_bot)
        thread.start()
        return jsonify({"message": "Discord Bot 起動します"}), 200
    except Exception as e:
        return jsonify({"error": f"Bot起動中にエラーが発生しました: {str(e)}"}), 500

# ----- ユーティリティ -----
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
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = string.ascii_uppercase[remainder] + result
    return result

# ----- Discord Bot（関数化）-----
def run_discord_bot():
    intents = discord.Intents.default()
    intents.message_content = True
    client = discord.Client(intents=intents)

    def extract_dates(text):
        pattern = r'(\d{1,2})月((?:\d{1,2},?)+)日'
        matches = re.findall(pattern, text)
        dates = []
        for month_str, day_str in matches:
            month = int(month_str)
            for d in day_str.split(","):
                try:
                    dates.append(f"{month}/{int(d)}")
                except ValueError:
                    continue
        return dates

    def clean_username(text):
        text = text.replace('確定', '')
        text = re.sub(r'\d{1,2}月\d{1,2}(?:日)?', '', text)
        return re.sub(r'[、,\s　]', '', text)

    @client.event
    async def on_ready():
        print(f'{client.user} が起動しました。')
        for guild in client.guilds:
            for channel in guild.text_channels:
                if channel.name == "確定監視bot":
                    await channel.send("起動しました")
                    return

    @client.event
    async def on_message(message):
        if message.author.bot:
            return

        if message.content.strip() == "終了":
            await message.channel.send("Botを終了します")
            await client.close()
            stop_event.set()
            return

        if not message.content.startswith("@command"):
            return

        try:
            parts = message.content.split()
            if len(parts) not in [3, 4]:
                raise ValueError
            _, channel_name, start_date = parts[:3]
            end_date = parts[3] if len(parts) == 4 else start_date

            channel = discord.utils.get(message.guild.channels, name=channel_name)
            if not channel:
                await message.channel.send(f"チャンネル {channel_name} が見つかりません")
                return

            start_dt = datetime.datetime.strptime(start_date, "%Y-%m-%d")
            end_dt = datetime.datetime.strptime(end_date, "%Y-%m-%d") + datetime.timedelta(days=1)

            sheet = gs.open_by_key(master_spreadsheet_id).get_worksheet(0)
            headers = sheet.row_values(2)
            date_map = {v.strip(): i+1 for i, v in enumerate(headers) if re.match(r"\d{1,2}/\d{1,2}", v)}
            master_data = sheet.get_all_values()

            updates = []
            updated_date = []
            async for msg in channel.history(after=start_dt, before=end_dt, oldest_first=True):
                dates = extract_dates(msg.content)
                username = clean_username(msg.content)
                for row_num, row in enumerate(master_data, start=1):
                    if row_num < 3 or len(row) < 5:
                        continue
                    master_name = row[4].replace(" ", "").replace("　", "").strip()
                    if username == master_name:
                        for date in dates:
                            if date in date_map:
                                cell = gspread.utils.rowcol_to_a1(row_num, date_map[date])
                                updates.append({"range": cell, "values": [["確定"]]})
                                updated_date.append(date)
                        if updated_date:
                            await message.channel.send(f"{username}: {', '.join(updated_date)} の確定を更新しました。")
                            updated_date.clear()
                        break

            if updates:
                sheet.batch_update(updates)
                await message.channel.send("処理が完了しました。終了します。")
            else:
                await message.channel.send("処理対象がありませんでした。")
            await client.close()
            stop_event.set()

        except Exception as e:
            print(traceback.format_exc())
            await message.channel.send("エラーが発生しました。詳細はログを確認してください。")

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    try:
        loop.run_until_complete(client.start(os.getenv("DISCORD_TOKEN")))
    except Exception as e:
        print(f"Bot実行中にエラー発生: {e}")
    finally:
        if not client.is_closed():
            loop.run_until_complete(client.close())
        stop_event.set()
        loop.close()

if __name__ == '__main__':
    app.run(debug=True)
