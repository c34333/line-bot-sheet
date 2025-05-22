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

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    text = event.message.text.strip()

    if text == "新規":
        user_sessions[user_id] = {"step": "awaiting_company"}
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text="① 会社名を入力してください"))
        return

    if user_id in user_sessions:
        session = user_sessions[user_id]

        if session["step"] == "awaiting_company":
            session["company"] = text
            session["step"] = "awaiting_introducer"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="② 元請・紹介者名を入力してください"))
            return

        elif session["step"] == "awaiting_introducer":
            session["introducer"] = text
            session["step"] = "awaiting_content"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(
                text="③ 内容を入力してください（スキップ可）",
                quick_reply=QuickReply(items=[
                    QuickReplyButton(action=MessageAction(label="スキップ", text="スキップ"))
                ])
            ))
            return

        elif session["step"] == "awaiting_content":
            session["content"] = "" if text == "スキップ" else text
            session["step"] = "awaiting_branch"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(
                text="④ 拠点名を選んでください",
                quick_reply=QuickReply(items=[
                    QuickReplyButton(action=MessageAction(label="本社", text="本社")),
                    QuickReplyButton(action=MessageAction(label="関東", text="関東")),
                    QuickReplyButton(action=MessageAction(label="前橋", text="前橋"))
                ])
            ))
            return

        elif session["step"] == "awaiting_branch":
            session["branch"] = f":{text}"
            session["step"] = "awaiting_site"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text="⑤ 現場名を入力してください"))
            return

        elif session["step"] == "awaiting_site":
            session["site"] = text
            session["step"] = "awaiting_month"
            months = [QuickReplyButton(action=MessageAction(label=f"{i}月", text=f"{i}月")) for i in range(1, 13)]
            months.insert(0, QuickReplyButton(action=MessageAction(label="未定", text="未定")))
            line_bot_api.reply_message(event.reply_token, TextSendMessage(
                text="⑥ 作業予定月を選択してください",
                quick_reply=QuickReply(items=months)
            ))
            return

        elif session["step"] == "awaiting_month":
            session["month"] = "未定" if text == "未定" else f"2025年{text}"
            session["step"] = "awaiting_worker"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(
                text="⑦ 対応者を選択してください",
                quick_reply=QuickReply(items=[
                    QuickReplyButton(action=MessageAction(label="自社", text="自社")),
                    QuickReplyButton(action=MessageAction(label="外注", text="外注"))
                ])
            ))
            return

        elif session["step"] == "awaiting_worker":
            session["worker"] = text
            session["step"] = "awaiting_etc"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(
                text="⑧ その他入力項目があれば入れてください（スキップ可）",
                quick_reply=QuickReply(items=[
                    QuickReplyButton(action=MessageAction(label="スキップ", text="スキップ"))
                ])
            ))
            return

        elif session["step"] == "awaiting_etc":
            session["etc"] = "" if text == "スキップ" else text
            session["step"] = "done"

            for row in range(1, 2001):
                if sheet.cell(row, 2).value in [None, ""]:
                    sheet.update_cell(row, 2, "新規追加")
                    sheet.update_cell(row, 5, session.get("company", ""))
                    sheet.update_cell(row, 6, session.get("branch", ""))
                    sheet.update_cell(row, 8, session.get("site", ""))
                    sheet.update_cell(row, 9, session.get("month", ""))
                    sheet.update_cell(row, 10, session.get("worker", ""))
                    sheet.update_cell(row, 11, session.get("introducer", ""))
                    sheet.update_cell(row, 12, session.get("content", ""))
                    sheet.update_cell(row, 13, session.get("etc", ""))

                    line_bot_api.reply_message(event.reply_token, [
                        TextSendMessage(text=f"登録完了しました！（案件番号：{row}）"),
                        TextSendMessage(text=(
                            f"①会社名：{session.get('company', '')}\n"
                            f"②元請・紹介者名：{session.get('introducer', '')}\n"
                            f"③内容：{session.get('content', '')}\n"
                            f"⑤現場名：{session.get('site', '')}\n"
                            f"⑥作業予定月：{session.get('month', '')}\n"
                            f"⑧その他：{session.get('etc', '')}"
                        ))
                    ])
                    break
            return

    print(f"📝 {user_id} が \"{text}\" と送信 → 対応外メッセージ")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
