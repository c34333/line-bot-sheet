from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage

import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Flaskアプリ
app = Flask(__name__)

# LINE APIの設定（環境変数から取得）
LINE_CHANNEL_ACCESS_TOKEN = '/FkVb/dexhEuFsB74xXlu81VY6v1eAI7HdKLDcMQpVQ1QQYtDvqNOCoVlfL0jx5Y3rHbtivI9gslRT/njwTihE6Ru/iXhGyNMzLvvxnFSuJsy3HgW5egfmnx8R9ydFWbaKdpezv3yIPosAUlU6Vl5AdB04t89/1O/w1cDnyilFU='
LINE_CHANNEL_SECRET = '74a58afe4666258e4a64610534d5eacc'


line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets認証（環境変数から読み込み）
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)

# スプレッドシートを開く（"LINEログ" という名前で用意しておく）
sheet = gc.open('LINEログ').sheet1

# LINE Webhookの受け口
@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)

    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'

# メッセージ受信イベント
@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    text = event.message.text
    sheet.append_row([user_id, text])
    print(f'📝 {user_id} が "{text}" と送信 → スプレッドシートに追加！')

if __name__ == "__main__":
    app.run(port=5000)
