import gspread
from google.oauth2.service_account import Credentials
from flask import Flask,jsonify,render_template,request
from flask_cors import CORS
import string
import re,os
import subprocess

#flaskアプリのインスタンス作成
app=Flask(__name__)#自動でフォルダを指定
CORS(app) #全てのエンドポイントに対してCORSを許可

#認証情報設定
credentials=Credentials.from_service_account_file(
    "C:/Users/bunta/Downloads/shift-automation/shift-automation-450015-5b4b75779314.json", #jsonファイルのパス
    scopes=["https://www.googleapis.com/auth/spreadsheets"] #APIのアクセススコープ(認証におけるアクセスリソースの定義)
)

#google sheets apiに認証してアクセス
gs=gspread.authorize(credentials)

#シフトスプレッドシートのID
shift_spreadsheet_id=""

#マスタスプレッドシートのID
master_spreadsheet="1-AdOvvYrqgSgVK_p-0JUNx1G6fN6UoUomTMqZ2kH0o0"

#urlからidを抽出する
def url_spreadsheet_id(url):
    match=re.search(r"/d/([a-zA-Z0-9-_]+)",url)
    if match:
        return match.group(1) if match else None
    return None #idが見つからない場合

#最終処理行を記録したファイルを削除するエンドポイント
@app.route('/reset_last_row',methods=['POST'])
def reset_last_row():
    try:
        if os.path.exists('last_row.txt'): #ファイルがあるかチェック
            os.remove('last_row.txt') #ファイルがあったら削除
            return jsonify({"message":"最終処理行がリセットされました"}),200
        else:
            return jsonify({"message":"リセットするファイルがありません"}),404
    except Exception as e:
        return jsonify({"error":f"ファイルの削除中にエラーが発生しました:{str(e)}"}),500

#アクセスした際に表示される内容
@app.route('/')
def home():
    return render_template('index.html') #HTMLページを返す

#スプレッドシートIDを設定するエンドポイント
@app.route('/set_shift_spreadsheet',methods=['POST'])
def set_shift_spreadsheet():
    global shift_spreadsheet_id
    data=request.get_json()
    url=data.get("spreadsheet_url","").strip()
    shift_spreadsheet_id=url_spreadsheet_id(url)
    if not shift_spreadsheet_id:
        return jsonify({"error":"スプレッドシートのurlが無効です"}),400
    return jsonify({"message":"スプレッドシートIDが設定されました"})

#実行ボタンで呼ばれるエンドポイント
@app.route('/process_shift',methods=['POST'])
def process_shift():
    global shift_spreadsheet_id
    if not shift_spreadsheet_id:
        return jsonify({"error":"スプレッドシートが設定されていません"}),400
    
    data=request.get_json()
    column_number=int(data.get("column_number",0)) #入力された列番号を取得
    column_count=int(data.get("column_count",0)) #入力された開催日数を取得

    shift_spreadsheet=gs.open_by_key(shift_spreadsheet_id)

    global shift_worksheet
    shift_worksheet=shift_spreadsheet.get_worksheet(0)

    response=shift_worksheet.get_all_values()
    # print(f"取得したデータ:{response}")

    #開催日の列を取得するために、5行目から指定された列数だけデータを取得
    shift_data=shift_worksheet.get_all_values()[4:4+column_count]

    info_message="初回実行として処理します"
    last_processd_row=get_last_processed_row()
    updated_row=process_shifts(column_number,column_count)or[]
    # print(info_message)

    #更新された学籍番号をリストとして格納
    updated_list=[msg.split("学籍番号")[1].split("のシフト")[0]for msg in updated_row]

    response={
        "last_processed_row":last_processd_row,
        "message":updated_row,
        "updated_list":updated_list,
        "info_message":info_message
    }
    return jsonify(response)

    # if updated_row:
    #     print(f"Success:{response}")
    #     return jsonify(response),200
    # else:
    #     print(f"Not Found:{response}")
    #     return jsonify(response),404


#前回取得時点での最後の保存
def get_last_processed_row():

    try:
        with open ('last_row.txt','r') as file:
            last_processed_row=int(file.read().strip() or 0)
            print(f"前回の最後の行処理:{last_processed_row}") #デバック用出力
    except FileNotFoundError:
        last_processed_row=0 #初回実行の場合一行目を取得する
        print("初回実行として処理します。")
    
    return last_processed_row

#処理した行を保存
def save_last_processed_row(last_row):
    
    with open('last_row.txt','w')as file:
        file.write(str(last_row))

# #現在のシフトシートの行数を取得
# current_now_count=shift_worksheet.row_count

# #新しく追加された行数を取得
# current_new_count=current_now_count-last_processed_row
# # print(f"current_now_count: {current_now_count}, last_processed_row: {last_processed_row}, current_new_count: {current_new_count}") #デバック用

def get_colum_letter(col_num):
    if col_num<=0:
        return ""
    result=""
    while col_num>0:
        col_num,remainder=divmod(col_num-1,26)
        result=string.ascii_uppercase[remainder]+result
    return result




#シフトデータを取得後マスタデータと照合して更新する関数
def process_shifts(column_number,column_count):
    print(f"受け取ったnumber:{column_number},count:{column_count}")#デバッグ用

    #更新情報を格納するリスト
    updates=[]
    
    
    #シフトスプレッドシートを開く
    shift_spreadsheet=gs.open_by_key(shift_spreadsheet_id)

    #シフトシート1の取得
    global shift_worksheet
    shift_worksheet=shift_spreadsheet.get_worksheet(0)

    last_processed_row=get_last_processed_row()
    current_now_count=len(shift_worksheet.get_all_values())

    #新しく追加された行のデータを取得
    if current_now_count>last_processed_row:
        # last_processed_row-=100 #なぜか100
        # current_now_count-=1 #なぜか1
        new_data=shift_worksheet.get_all_values()[last_processed_row:current_now_count] 
        # print(f"前回最後: {last_processed_row}") #デバック用
        # print(f"現在: {current_now_count}") #デバック用
        all_values=shift_worksheet.get_all_values()
        # print(f"シフトシート全データ: {all_values}") #デバック用
        # print(f"新しいデータ: {new_data}") #デバック用
        # print(f"全シートデータ: {shift_worksheet.get_all_values()}") #デバック用
        # print(f"new_data 取得範囲: {shift_worksheet.get_all_values()[last_processed_row:current_now_count]}") #デバック用

        #マスタのスプレッドシートを開く
        master_spreadsheet=gs.open_by_key('1-AdOvvYrqgSgVK_p-0JUNx1G6fN6UoUomTMqZ2kH0o0')

        #マスタシート1の取得
        master_worksheet=master_spreadsheet.get_worksheet(0)

        #マスタシートのデータを取得
        master_data=master_worksheet.get_all_values()

        #必要な列だけ取得
        master_data=[[row[3],row[4]]for row in master_data]

        #更新情報を格納するリスト
        updated_row=[]
        

        #シフトスプレッドシートから新しいデータを処理
        for row in new_data:
            shift_id=row[4].strip().replace(' ', '').replace('　', '').upper() #学籍番号
            shift_name=row[3].strip().replace(' ', '').replace('　', '') #氏名
            shift_status_list = []  # 取得した〇×を格納するリスト

            for i in range(column_count):  # 開催日数分ループ
                if 5 + i >= len(row):
                    shift_status_list.append("")  # データがない場合は空欄を追加
                else:
                    shift_status_list.append(row[5 + i])  # シフトデータをリストに追加
           
    
            #マスタシートの行をループして、学籍番号、氏名が一致する行を探す
            for row_num,master_row in enumerate(master_data,start=1):
                master_id=master_row[0].strip().replace(' ', '').replace('　', '').upper() #学籍番号
                master_name=master_row[1].strip().replace(' ', '').replace('　', '') #氏名

                # #デバック用
                # print(f"比較中: シフトID={shift_id}, シフト名={shift_name} vs マスタID={master_id}, マスタ名={master_name}")


                #学籍番号、氏名が一致する場合
                if shift_id==master_id and shift_name==master_name:
                    for i in range(column_count):
                        column_letter=get_colum_letter(column_number+i)
                        update_range=f'{column_letter}{row_num}'
                        updates.append({
                            'range':update_range,
                            'values':[[shift_status_list[i]]]
                        })
                    # updates.append({
                    #     'range':f'{chr(64+column_number)}{row_num}',
                    #     'values':[[shift_status]]
                    # })
                    #一致する行にシフト状態(〇or×)を書き込む
                    # master_worksheet.update_cell(row_num,23,shift_status)
                    updated_row.append(f"学籍番号{shift_id}のシフトが更新されました。")
                    break
                
                # #デバック用
                # else:
                #     print(f"一致しません:学籍番号{shift_id},氏名{shift_name}")
                
        
        #処理した行番号を更新して保存
        last_processed_row+=len(new_data) #新しい行数分だけ更新
        save_last_processed_row(last_processed_row)

    if updates:
        #  # 🔹 **デバッグ用に書き込むデータを出力**
        # print("=== 📝 更新予定のデータ一覧 ===")
        # for update in updates:
        #     print(f"範囲: {update['range']} → 書き込む値: {update['values']}")

        master_worksheet.batch_update(updates)

        return updated_row

   

    # if updates:
    #     print(f"Updates: {updates}") #デバック用
    #     print(f"newdata:{new_data}")#デバック用
    #     print(f"更新する範囲: {update_range}")#デバック用
    #     print(f"Column Number: {column_number}, Column Count: {column_count}")#デバック用
    #     print(shift_spreadsheet_id)

    return updated_row if updates else None

@app.route('/start_bot',methods=['POST'])
def start_bot():
    try:
        script_path=r"C:\Users\bunta\Downloads\shift-automation\apple\kakutei-automation-apple.py"
        #discordbotを起動する処理
        process=subprocess.Popen(["python",script_path])
        return jsonify({"message":"Discord Bot が起動しました"}),200
    except Exception as e:
        return jsonify({"error":f"Bot起動中にエラーが発生しました:{str(e)}"}),500
    

if __name__=='__main__':
    #flaskアプリを実行(デバック)
    app.run(debug=True)