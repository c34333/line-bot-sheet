from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, QuickReply, QuickReplyButton, MessageAction

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

    # ステップごとの処理
    if user_id in user_sessions:
        session = user_sessions[user_id]

        if session["step"] == "awaiting_company":
            session["company_name"] = text
            session["step"] = "awaiting_branch"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text="拠点名を選んでください",
                    quick_reply=QuickReply(items=[
                        QuickReplyButton(action=MessageAction(label="本社", text="本社")),
                        QuickReplyButton(action=MessageAction(label="関東", text="関東")),
                        QuickReplyButton(action=MessageAction(label="前橋", text="前橋"))
                    ])
                )
            )
            return

        elif session["step"] == "awaiting_branch":
            session["branch"] = f":{text}"
            session["step"] = "awaiting_site"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="現場名を入力してください")
            )
            return

        elif session["step"] == "awaiting_site":
            session["site"] = text
            session["step"] = "awaiting_month"
            months = [QuickReplyButton(action=MessageAction(label=f"{i}月", text=f"{i}月")) for i in range(1, 13)]
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text="作業予定月を選択してください",
                    quick_reply=QuickReply(items=months)
                )
            )
            return

        elif session["step"] == "awaiting_month":
            session["month"] = f"2025年{text}"
            session["step"] = "awaiting_worker"
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(
                    text="対応者を選択してください",
                    quick_reply=QuickReply(items=[
                        QuickReplyButton(action=MessageAction(label="自社", text="自社")),
                        QuickReplyButton(action=MessageAction(label="外注", text="外注"))
                    ])
                )
            )
            return

        elif session["step"] == "awaiting_worker":
            session["worker"] = text
            session["step"] = "done"

            # 空いているB列を探す（1〜2000）
            for row in range(1, 2001):
                if sheet.cell(row, 2).value in [None, ""]:
                    sheet.update_cell(row, 2, "新規追加")  # B列
                    sheet.update_cell(row, 5, session["company_name"])  # E列
                    sheet.update_cell(row, 6, session["branch"])         # F列
                    sheet.update_cell(row, 8, session["site"])           # H列
                    sheet.update_cell(row, 9, session["month"])          # I列
                    sheet.update_cell(row, 10, session["worker"])        # J列
                    line_bot_api.reply_message(
                        event.reply_token,
                        TextSendMessage(text=f"登録完了しました！（案件番号：{row}）")
                    )
                    break
            return

    # その他のメッセージはログ＆無視
    print(f"📝 {user_id} が \"{text}\" と送信 → 対応外メッセージ")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
