import json

# サービスアカウントのjsonファイルのパス
with open("C:\\Users\\bunta\\Downloads\\shift-automation\\apple\\shift-automation-450015-5b4b75779314.json", "r", encoding="utf-8") as f:
    data = json.load(f)

# 改行を \n → \\n に変換（1行にする）
data["private_key"] = data["private_key"].replace("\n", "\\n")

# 1行のJSONに変換
one_line = json.dumps(data)
print(one_line)
