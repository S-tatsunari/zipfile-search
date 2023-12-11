import requests
import pandas as pd
import time

# Excelファイルの読み込み
file_path = 'xx' # 使用するファイルパス
df = pd.read_excel(file_path, usecols="X") 

# ZIPファイルの存在をチェックする関数
def check_zip_file(url):
    try:
        # HEADリクエストを使用して実際にファイルをダウンロードしない
        response = requests.head(url)
        # ステータスコードが200（OK）の場合、ファイルは存在する
        if response.status_code == 200:
            return 'Exists'
        else:
            # 他のステータスコードの場合は、そのコードを記録する
            return f'Not found (Status code: {response.status_code})'
    except requests.RequestException as e:
        return f'Error (Exception: {e})'

# 結果を保存するリスト
results = []

# 各URLに対してZIPファイルの存在を確認
for url in df.iloc[:, 0]:
    result = check_zip_file(url)
    results.append(result)
    # サーバーへの負荷を考慮して少し待機
    time.sleep(1)

# 結果をDataFrameに追加
df['ZIP File Status'] = results

# 結果を新しいExcelファイルに保存
df.to_excel('zip_file_check_results.xlsx', index=False)
