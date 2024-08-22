# スプレッドシートを使ってGoogle検索の順位を読み取り書き込む処理

import datetime
import gspread
import os
from oauth2client.service_account import ServiceAccountCredentials
import requests
from bs4 import BeautifulSoup
from dotenv import load_dotenv

# .envファイルを読み込む
load_dotenv('.env') 

SHPREADSHEET_KEY = os.environ.get("SHPREADSHEET_KEY")

# Googleスプレッドシートにアクセス
def auth_gspread(json_file_name, shpreadsheet_key):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file_name, scope)
    gc = gspread.authorize(credentials)
    worksheet = gc.open_by_key(shpreadsheet_key).sheet1
    return worksheet


# worksheet = auth_gspread(json_file_name, shpreadsheet_key)
# keyword = worksheet.col_values(1)
# print(keyword)

# Google検索
def google_search(keyword, domain, num_results=30):
    google_search_url = "https://www.google.com/search"
    headers = {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    }
    params = {
        "q": keyword,
        "num": num_results,
    }
    response = requests.get(google_search_url, headers=headers, params=params)
    soup = BeautifulSoup(response.text, "html.parser")
    search_results = soup.select("div.yuRUbf > div > span > a")

    for index, result in enumerate(search_results):
        href = result.get("href")
        if domain in href:
            return index + 1
    return "圏外です"

# メインの処理
json_file_name = '/content/drive/MyDrive/Engineering/03_blog/ferrous-marking-431123-d9-ee2ea4870813.json'
shpreadsheet_key = SHPREADSHEET_KEY # 環境変数からスプレッドシートのキーを使用

worksheet = auth_gspread(json_file_name, shpreadsheet_key)
keywords = worksheet.col_values(1)[1:]  # 最初の行をスキップ
domain = "https://kick.tokyo/"


# 日付の処理
  # if(値があれば上書き) else (値がなければ要素+1で新しい列に書く)
today = datetime.datetime.now().strftime('%Y-%m-%d')
# today = str(datetime.datetime.now())
col_index = worksheet.row_values(1).index(today) + 1 if today in worksheet.row_values(1) else len(worksheet.row_values(1)) + 1
worksheet.update_cell(1, col_index, today)


for i, keyword in enumerate(keywords, start=2):  # 最初の行をスキップしてインデックス調整
    rank = google_search(keyword, domain)
    print(f"キーワード「{keyword}」における「{domain}」のランキング：{rank}")
    worksheet.update_cell(i, col_index, rank)
