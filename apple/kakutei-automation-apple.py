import discord
import datetime
import re
import gspread
from google.oauth2.service_account import Credentials
import asyncio
import traceback
from dotenv import load_dotenv
import os

# Load .env variables
load_dotenv()

# Setup Discord client
intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)

# Setup Google Sheets API
json_path = os.getenv("GOOGLE_JSON_PATH")
credentials = Credentials.from_service_account_file(
    json_path, scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
gs = gspread.authorize(credentials)

# 起動確認
@client.event
async def on_ready():
    print(f'{client.user} が起動しました。')
    target_channel_name = "確定監視bot"
    for guild in client.guilds:
        for channel in guild.text_channels:
            if channel.name == target_channel_name:
                await channel.send("起動しました")
                break

# 日付抽出関数
def extract_dates(text):
    today = datetime.date.today()
    current_year = today.year
    pattern = r'(\d{1,2})月((?:\d{1,2},?)+)日'
    matches = re.findall(pattern, text)
    dates = []
    for month_str, day_str in matches:
        month = int(month_str)
        day_list = day_str.split(",")
        for d in day_list:
            try:
                date_obj = datetime.date(current_year, month, int(d))
                dates.append(f"{date_obj.month}/{date_obj.day}")
            except ValueError:
                continue
    return dates

# ユーザー名整形関数
def clean_username(text):
    text = text.replace('確定', '')
    text = re.sub(r'\d{1,2}月\d{1,2}(?:日)?', '', text)
    return re.sub(r'[、,\s　]', '', text)

@client.event
async def on_message(message):
    print(f"受信メッセージ: {message.content}")
    if message.content.strip() == "終了":
        await message.channel.send("Botを終了します")
        await client.close()
        print("Bot手動終了")
        return
    
    if message.author.bot or not message.content.startswith("@command"):
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

        start_datetime = datetime.datetime.strptime(start_date, "%Y-%m-%d")
        end_datetime = datetime.datetime.strptime(end_date, "%Y-%m-%d") + datetime.timedelta(days=1)

        # Google Sheets 読み込み
        master_spreadsheet_id = os.getenv("MASTER_SPREADSHEET_ID")
        master_sheet = gs.open_by_key(master_spreadsheet_id).get_worksheet(0)
        header_row = master_sheet.row_values(2)
        date_column_map = {
            cell.strip(): idx for idx, cell in enumerate(header_row, start=1)
            if re.match(r"\d{1,2}/\d{1,2}", cell.strip())
        }
        print(f"ヘッダー列対応マップ: {date_column_map}")
        master_data = master_sheet.get_all_values()

        updates = []
        updated_date=[]
        async for msg in channel.history(after=start_datetime, before=end_datetime, oldest_first=True):
            print(f"対処メッセージ: {msg.content}")
            dates = extract_dates(msg.content)
            username = clean_username(msg.content)
            print(f"抽出日付: {dates}, 整形ユーザー名: {username}")

            for row_num, row in enumerate(master_data, start=1):
                if row_num < 3 or len(row) < 3:
                    continue
                master_name = row[4].replace(" ", "").replace("　", "").strip()

                # print(f"照合対象マスタ名: '{row[4]}' → 整形後: '{master_name}'")
                if username == master_name:
                    一致=True
                    for date in dates:
                        if date in date_column_map:
                            cell = gspread.utils.rowcol_to_a1(row_num, date_column_map[date])
                            col=date_column_map[date]
                            updates.append({"range": cell, "values": [["確定"]]})
                            updated_date.append(date)
                            # print(f"{username} の {date} に確定を書き込みます → セル: R{row_num}C{col}")
                    if 一致 and updated_date:
                        unique_dates = sorted(set(updated_date), key=updated_date.index)
                        await message.channel.send(f"{username}: {', '.join(unique_dates)} の確定を更新しました。")
                        updated_date.clear()
                    break

       

        if updates:
            master_sheet.batch_update(updates)
            print("すべての確定処理を終了しました")

        if not updates:
            await message.channel.send(f"チャンネル {channel_name} に {start_date}~{end_date} の処理対象はありませんでした。")

        await message.channel.send("処理が完了しました。終了します。")
        await client.close()
        print("bot終了")

    except ValueError:
        await message.channel.send("コマンドの形式が正しくありません。\n 例: '@command 確定チャット 2025-04-14 2025-04-17'")
    except Exception as e:
        print(f"予期しないエラー: {e}")
        print(traceback.format_exc())
        await message.channel.send("エラーが発生しました。詳細はログを確認してください。")

async def main():
    token = os.getenv("DISCORD_TOKEN")
    if not token:
        print("DISCORD_TOKENが設定されていません。.envファイルを確認してください。")
        return
    await client.start(token)

if __name__ == "__main__":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    asyncio.run(main())
