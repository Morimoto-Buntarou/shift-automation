import discord
import datetime
import re
import gspread
from google.oauth2.service_account import Credentials
import asyncio
import string
from dotenv import load_dotenv
import os
import traceback

load_dotenv()

intents = discord.Intents.default()
intents.message_content = True
client = discord.Client(intents=intents)

#認証情報設定
load_dotenv()
json_path=os.getenv("GOOGLE_JSON_PATH")

credentials=Credentials.from_service_account_file(
    json_path,
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)

#google sheets apiに認証してアクセス
gs=gspread.authorize(credentials)

#起動確認
@client.event
async def on_ready(): 
    print(f'{client.user} が起動しました。')

    target_channel_name="確定監視bot"

    for guild in client.guilds:
        for channel in guild.text_channels:
            if channel.name==target_channel_name:
                await channel.send("起動しました")
                break



@client.event
async def on_message(message):
    print(f"受信メッセージ{message.content}")
    if message.author.bot:
        return
    
    if message.content.strip() == "終了":
        await message.channel.send("Botを終了します。")
        await client.close()
        print("Bot手動終了")
        return


    #"@command"から始まるメッセージにのみ反応するようにする
    if not message.content.startswith("@command"):
        return

    try:

        #コマンドの形式が2つなら1日指定、3つなら範囲指定
        command_date_list=message.content.split()
        if len(command_date_list) not in [3,4]: 
            raise ValueError #コマンド形式エラーとして処理

        _, channel_name,start_date=command_date_list[0:3] #最初の3つを取得
        end_date=start_date if len(command_date_list)==3 else command_date_list[3] #4つめがあれば範囲指定

      

        #指定したチャンネルを取得
        channel=discord.utils.get(message.guild.channels,name=channel_name)

        #指定したチャンネルがなかった場合
        if channel is None:
            await message.channel.send(f"チャンネル{channel_name}が見つかりません")
            return

        #日付を指定 (フォーマット:YYYY-MM-DD)
        start_datetime=datetime.datetime.strptime(start_date,"%Y-%m-%d")
        end_datetime=datetime.datetime.strptime(end_date,"%Y-%m-%d")+datetime.timedelta(days=1)

      
        def hiduke_date(text):
            today=datetime.date.today()
            current_year=today.year
            pattern=r'(\d{1,2})月((?:\d{1,2},?)+)日'
            matches=re.findall(pattern,text)
            dates=[]
            for month_str,day_str in matches:
                month=int(month_str)
                day_list=day_str.split(",")
                for day_str in day_list:
                    try:
                        day=int(day_str)
                        dt=datetime.date(current_year,month,day)
                        dates.append(f"{dt.month}/{dt.day}")
                    except ValueError:
                        continue
            return dates


        #指定範囲内のメッセージ履歴を取得してマスタに書き込む
        messages=[]

         #デバック確認用
        # print(f"直した後ユザネ{username}")

        #マスタのスプレッドシートを開く
        master_spreadsheet=gs.open_by_key('16y2vkOALlolyPpcZ86_7XLQAARrAMNrH1wS31pqLm5Q')

        #マスタシートの取得
        master_worksheet=master_spreadsheet.get_worksheet(0)

            

        header_row=master_worksheet.row_values(2)
            
        date_column_map={}
        for col_index,cell in enumerate(header_row,start=1):
            if re.match(r"\d{1,2}/\d{1,2}",cell.strip()):
                date_column_map[cell.strip()]=col_index
        print(f"ヘッダー列対応マップ{date_column_map}")
        #マスタシートのデータを取得
        master_data=master_worksheet.get_all_values()

        updates=[]
        async for msg in channel.history(after=start_datetime,before=end_datetime,oldest_first=True):
            print(f"対処メッセージ{msg.content}")
            messages.append(f"{msg.content}")
            username=msg.content

            
        
            #デバック確認用
            # print(f"治す前ユザネ{username}")

            # #ユーザー情報取得テスト
            # member=message.guild.get_member(msg.author.id)

            # #初期化
            # server_nickname=None

            # if member is not None:
            #     #サーバー専用のニックネームを取得
            #     server_nickname=member.nick
            
            # #サーバー内表示名が取得できたかどうか確認
            # if server_nickname is not None:
            #     print(f"サーバー内表示名:{server_nickname}")

            # #サーバー内表示名がない場合、グローバル表示名を使用
            # else:
            #     print(f"グローバル表示名{username}")
            dates=hiduke_date(username)
            print(f"抽出日付{dates}")
            
            #usernameに格納されたデータを名前のみにする
            username = username.replace('確定', '')  # 「確定」の削除
            username = re.sub(r'\d{1,2}月\d{1,2}(?:日)?', '', username)  # 「4月20日」「4月20」など削除
            username = re.sub(r'[、,\s　]', '', username)  # スペースや句読点を削除
            print(f"ユーザー整形後{username}")
            
           


            #デバック確認用
            # print(f"直した後ユザネ{username}")

            #マスタのスプレッドシートを開く
            master_spreadsheet=gs.open_by_key('16y2vkOALlolyPpcZ86_7XLQAARrAMNrH1wS31pqLm5Q')

            #マスタシートの取得
            master_worksheet=master_spreadsheet.get_worksheet(0)

            

            header_row=master_worksheet.row_values(2)
            
            date_column_map={}
            for col_index,cell in enumerate(header_row,start=1):
                if re.match(r"\d{1,2}/\d{1,2}",cell.strip()):
                    date_column_map[cell.strip()]=col_index
            print(f"ヘッダー列対応マップ{date_column_map}")
            #マスタシートのデータを取得
            master_data=master_worksheet.get_all_values()

            #必要な列だけ取得
            # master_data=[row[4].strip().replace(' ', '').replace('　', '') for row in master_data]

            #マスタシートの行をループして、学籍番号、氏名が一致する行を探す

            一致=False
            for row_num,master_row in enumerate(master_data,start=1):
                if row_num<5:
                    continue

                if len(master_row)<5:
                    continue
                master_name=master_row[4].replace(" ", "").replace("　","").strip()
                print(f"元のマスタ名{master_row[4]}")

                #デバック用
                print(f"比較中: シフト名={username} vs マスタ名={master_name}")

                
                #学籍番号、氏名が一致する場合
                if  username==master_name:
                    一致=True
                    for date in dates:
                        if date in date_column_map:
                            col=date_column_map[date]
                            cell_label=gspread.utils.rowcol_to_a1(row_num,col)
                            updates.append({
                                "range":cell_label,
                                "values":[["確定"]]
                            })
                            # print(f"使用予定日付列: {date} → 列番号: {col if date in date_column_map else '未登録'}")
                            
                            print(f"{username} の {date} に確定を書き込みます → セル: R{row_num}C{col}")
                    #一致する行にシフト状態(確定)を書き込む
                           
                            print(f"{username}のシフトが更新されました。")
                    break
            if not 一致:
                print(f"一致しませんでした{username}")
                
                
        if updates:
            master_worksheet.batch_update(updates)
            print("すべての確定処理を終了しました")

        # #デバック用格納確認
        # print(messages)

        #指定範囲にメッセージがなかった場合
        if not messages:
            await message.channel.send(f"チャンネル{channel_name}に{start_date}~{end_date}のメッセージはありません。")
            return
        
        # #デバック用メッセージを一括取得して送信
        # all_messages="\n".join(messages)

    #コマンドが間違ってた場合
    except ValueError:
        await message.channel.send("コマンドの形式が正しくありません。\n 1日指定 '@command 確定チャット 2025-02-02 23' \n 範囲指定 '@command 確定チャット 2025-02-02 2025-02-05 23'")
    except Exception as e:
        print(f"予期しないエラー:{e}")
        print(traceback.format_exc())
        await message.channel.send("エラーが発生しました。")
    

    await message.channel.send("処理が完了しました。終了します。")
    await client.close()
    print("bot終了")


async def main():
    token=os.getenv("DISCORD_TOKEN")
    if not token:
        print("DICORD_TOKENが設定されていません。.envファイルを確認してください。")
    try:
        await client.start(token)
    finally:
        await client.close()
        print("終了")

if __name__=="__main__":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())
    loop=asyncio.new_event_loop()
    asyncio.set_event_loop(loop)

    try:
        loop.run_until_complete(main())

    finally:
        loop.close()