import os
import json
import gspread
from flask import Flask, request, abort
from oauth2client.service_account import ServiceAccountCredentials

from linebot.v3.webhook import WebhookHandler
from linebot.v3.messaging import (
    Configuration,
    ApiClient,
    MessagingApi,
    ReplyMessageRequest,
    TextMessage,
    QuickReply,
    QuickReplyItem,
    MessageAction
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent
from linebot.v3.exceptions import InvalidSignatureError

app = Flask(__name__)

# LINEの設定
LINE_CHANNEL_ACCESS_TOKEN = os.environ['LINE_CHANNEL_ACCESS_TOKEN']
LINE_CHANNEL_SECRET = os.environ['LINE_CHANNEL_SECRET']

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
api_client = ApiClient(configuration)
line_bot_api = MessagingApi(api_client)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets 設定
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)
sheet = gc.open('LINEログ').sheet1

# ユーザーごとの入力状況を記録
user_sessions = {}

# スプレッドシートの空いている行を取得（B列が空欄）
def find_next_available_row():
    col_b = sheet.col_values(2)
    for i in range(1, 2001):
        if i >= len(col_b) or col_b[i] == '':
            return i + 1  # 1-indexed
    return None

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent)
def handle_message(event):
    if not isinstance(event.message, TextMessageContent):
        return

    user_id = event.source.user_id
    text = event.message.text.strip()

    if user_id not in user_sessions:
        if text == "新規":
            user_sessions[user_id] = {"step": "status"}
            send_quick_reply(event.reply_token, "① 案件進捗を選んでください", ["新規追加", "3:受注", "4:作業完了", "定期"])
        return

    session = user_sessions[user_id]
    step = session.get("step")

    if step == "status":
        session["status"] = text
        session["step"] = "company"
        reply(event.reply_token, "② 会社名を入力してください")
    elif step == "company":
        session["company"] = text
        session["step"] = "client"
        reply(event.reply_token, "③ 元請・紹介者名を入力してください")
    elif step == "client":
        session["client"] = text
        session["step"] = "site"
        reply(event.reply_token, "④ 現場名を入力してください")
    elif step == "site":
        session["site"] = text
        session["step"] = "branch"
        send_quick_reply(event.reply_token, "⑤ 拠点名を選んでください", ["本社", "関東", "前橋"])
    elif step == "branch":
        session["branch"] = f":{text}"
        session["step"] = "content"
        send_quick_reply(event.reply_token, "⑥ 依頼内容・ポイントを入力してください（スキップ可）", ["スキップ"])
    elif step == "content":
        session["content"] = "" if text == "スキップ" else text
        session["step"] = "month"
        send_quick_reply(event.reply_token, "⑦ 作業予定月を選んでください", ["未定"] + [f"{i}月" for i in range(1,13)])
    elif step == "month":
        session["month"] = f"2025年{text}" if text != "未定" else "未定"
        session["step"] = "type"
        send_quick_reply(event.reply_token, "⑧ 対応者を選んでください", ["自社", "外注"])
    elif step == "type":
        session["type"] = text
        session["step"] = "memo"
        send_quick_reply(event.reply_token, "⑨ その他入力項目があれば入力してください（スキップ可）", ["スキップ"])
    elif step == "memo":
        session["memo"] = "" if text == "スキップ" else text

        # スプレッドシートに転記
        row = find_next_available_row()
        if row:
            sheet.update_cell(row, 2, format_status(session["status"]))
            sheet.update_cell(row, 5, session["company"])
            sheet.update_cell(row, 6, session["branch"])
            sheet.update_cell(row, 8, session["site"])
            sheet.update_cell(row, 9, session["month"])
            sheet.update_cell(row, 10, session["type"])

            # A列の番号（入力済のもの）を取得
            cell_value = sheet.cell(row, 1).value or str(row - 1)

            summary = f"""登録完了しました！（案件番号：{cell_value}）

① 案件進捗：{session['status']}
② 会社名：{session['company']}
③ 現場名：{session['site']}
④ 拠点名：{session['branch']}
⑤ 依頼内容：{session['content']}
⑥ 月：{session['month']}
⑦ その他：{session['memo']}"""
            reply(event.reply_token, summary)
        else:
            reply(event.reply_token, "⚠ スプレッドシートの空きが見つかりませんでした。")
        del user_sessions[user_id]

# 補助関数
def send_quick_reply(token, text, options):
    items = [QuickReplyItem(action=MessageAction(label=opt, text=opt)) for opt in options]
    line_bot_api.reply_message(ReplyMessageRequest(
        reply_token=token,
        messages=[TextMessage(text=text, quick_reply=QuickReply(items=items))]
    ))

def reply(token, text):
    line_bot_api.reply_message(ReplyMessageRequest(
        reply_token=token,
        messages=[TextMessage(text=text)]
    ))

def format_status(status):
    if status == "3:受注":
        return "3:受注"
    elif status == "4:作業完了":
        return "4:作業完了"
    elif status == "定期":
        return "定期"
    else:
        return "新規追加"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
