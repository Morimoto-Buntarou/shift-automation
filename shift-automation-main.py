import gspread
from google.oauth2.service_account import Credentials

#認証情報設定
credentials=Credentials.from_service_account_file(
    "C:/Users/bunta/Downloads/shift-automation/shift-automation-450015-5b4b75779314.json", #jsonファイルのパス
    scopes=["https://www.googleapis.com/auth/spreadsheets"] #APIのアクセススコープ(認証におけるアクセスリソースの定義)
)

#google sheets apiに認証してアクセス
gs=gspread.authorize(credentials)

#シフト回答のスプレッドシートを開く
shift_spreadsheet=gs.open_by_key('1Yntx0Ev6yYb7INyRgdE2Gul-gp8bbNrusrrQ7k5lcmg')

#シフトシートの取得
shift_worksheet=shift_spreadsheet.get_worksheet(0)

#前回取得時点での最後の保存
try:
    with open ('last_row.txt','r') as file:
        last_processed_row=int(file.read())
        print(f"前回の最後の行処理:{last_processed_row}") #デバック用出力
except FileNotFoundError:
    last_processed_row=1 #初回実行の場合一行目を取得する
    print("初回実行として処理します。")

#現在のシフトシートの行数を取得
current_now_count=shift_worksheet.row_count

#新しく追加された行数を取得
current_new_count=current_now_count-last_processed_row
# print(f"current_now_count: {current_now_count}, last_processed_row: {last_processed_row}, current_new_count: {current_new_count}") #デバック用

#新しく追加された行のデータを取得
if current_new_count>0:
    last_processed_row-=100 #なぜか100
    current_now_count-=1 #なぜか1
    new_data=shift_worksheet.get_all_values()[last_processed_row:current_now_count] 
    # print(f"前回最後: {last_processed_row}") #デバック用
    # print(f"現在: {current_now_count}") #デバック用
    all_values=shift_worksheet.get_all_values()
    # print(f"シフトシート全データ: {all_values}") #デバック用
    # print(f"新しいデータ: {new_data}") #デバック用
    # print(f"全シートデータ: {shift_worksheet.get_all_values()}") #デバック用
    # print(f"new_data 取得範囲: {shift_worksheet.get_all_values()[last_processed_row:current_now_count]}") #デバック用

    #マスタのスプレッドシートを開く
    master_spreadsheet=gs.open_by_key('1H_FUDG9YCypi6H6XkyjYxt4MHsLsb-tuaUVF_AvKU3c')

    #マスタシートの取得
    master_worksheet=master_spreadsheet.get_worksheet(0)

    #マスタシートのデータを取得
    master_data=master_worksheet.get_all_values()

    #必要な列だけ取得
    master_data=[[row[3],row[4]]for row in master_data]

    #シフトスプレッドシートから新しいデータを処理
    for row in new_data:
        shift_id=row[4].strip().replace(' ', '').replace('　', '').upper() #学籍番号
        shift_name=row[3].strip().replace(' ', '').replace('　', '') #氏名
        shift_status=row[5] #シフト状態(〇or×)
   
        #マスタシートの行をループして、学籍番号、氏名が一致する行を探す
        for row_num,master_row in enumerate(master_data,start=1):
            master_id=master_row[0].strip().replace(' ', '').replace('　', '').upper() #学籍番号
            master_name=master_row[1].strip().replace(' ', '').replace('　', '') #氏名

            # #デバック用
            # print(f"比較中: シフトID={shift_id}, シフト名={shift_name} vs マスタID={master_id}, マスタ名={master_name}")


            #学籍番号、氏名が一致する場合
            if shift_id==master_id and shift_name==master_name:
                #一致する行にシフト状態(〇or×)を書き込む
                master_worksheet.update_cell(row_num,23,shift_status)
                print(f"学籍番号{shift_id}のシフトが更新されました。")
            
            # #デバック用
            # else:
            #     print(f"一致しません:学籍番号{shift_id},氏名{shift_name}")
            
    #処理した行番号を更新して保存
    last_processed_row+=100
    current_now_count+=1
    with open('last_row.txt','w')as file:
        file.write(str(current_now_count))

#新しいデータがなかった場合
else:
    print("新しいデータはありません")