from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Flaskアプリ
app = Flask(__name__)

# LINE APIの設定（環境変数から取得）
LINE_CHANNEL_ACCESS_TOKEN = os.environ['LINE_CHANNEL_ACCESS_TOKEN']
LINE_CHANNEL_SECRET = os.environ['LINE_CHANNEL_SECRET']

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets認証（環境変数から読み込み）
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)

# スプレッドシートを開く（"LINEログ" という名前で用意しておく）
sheet = gc.open('LINEログ').sheet1

# ユーザーごとの入力状態を保持
user_sessions = {}

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
    text = event.message.text.strip()

    # 新規案件の開始
    if text == "【新規案件】":
        user_sessions[user_id] = {"step": "awaiting_company"}
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="会社名を入力してください")
        )
        return

    # ステップ: 会社名入力中
    if user_id in user_sessions:
        session = user_sessions[user_id]

        if session["step"] == "awaiting_company":
            session["company_name"] = text
            session["step"] = "awaiting_branch"
            # 次ステップ（拠点）でクイックリプライ表示（あとで追加予定）
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="（仮）拠点名を選んでください（あとでクイックリプライ対応）")
            )
            return

    # その他のメッセージはログ＆無視
    print(f"📝 {user_id} が \"{text}\" と送信 → 対応外メッセージ")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
