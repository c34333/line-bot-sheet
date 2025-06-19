# main.py

import os
import json
from datetime import datetime
import gspread
from flask import Flask, request, abort
from oauth2client.service_account import ServiceAccountCredentials
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage, QuickReply, QuickReplyButton, MessageAction
)

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.getenv("LINE_CHANNEL_ACCESS_TOKEN")
LINE_CHANNEL_SECRET = os.getenv("LINE_CHANNEL_SECRET")
report_group_id = "C6736021a0854b9c9526fdea9cf5acfa1"
silent_group_ids = ["C6736021a0854b9c9526fdea9cf5acfa1", "Cac0760acd664e7fdfa7a40975c340351"]

line_bot_api = LineBotApi(LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)

sheet = gc.open_by_key('1Y3-XMKstJIfqBd55JtIdB4Ymx9zs6DFQo2U78fuFyME').worksheet('基本入力')
ref_sheet = gc.open_by_key('1Y3-XMKstJIfqBd55JtIdB4Ymx9zs6DFQo2U78fuFyME').worksheet('参照値')

user_sessions = {}

@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers["X-Line-Signature"]
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    user_id = event.source.user_id
    group_id = getattr(event.source, 'group_id', None)
    text = event.message.text.strip()

    command_aliases = {
        "案件追加": "ん",
        "新規案件": "ん",
        "テスト": "テスト",
        "確認": "テスト",
    }
    command = command_aliases.get(text, text)

    if command in ["リセット", "キャンセル"]:
        if user_id in user_sessions:
            del user_sessions[user_id]
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="リセットしました。再度コマンドを入力してください。")
            )
        else:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="現在、進行中の登録はありません。再度コマンドを入力してください。")
            )
        return

    if group_id in silent_group_ids and user_id not in user_sessions and command not in ["ん", "テスト"]:
        return

    if command == "ん":
        user_sessions[user_id] = {"step": "inputter", "sender_name": get_user_display_name(user_id)}
        send_quick_reply(event.reply_token, "① 入力者を選んでください", ["未定", "諸橋", "酒井", "大塚", "原", "関野", "志賀", "加勢", "藤巻"])
        return
    elif command == "テスト":
        user_sessions[user_id] = {"step": "inputter", "test_mode": True, "sender_name": get_user_display_name(user_id)}
        send_quick_reply(event.reply_token, "① 入力者を選んでください", ["未定", "諸橋", "酒井", "大塚", "原", "関野", "志賀", "加勢", "藤巻"])
        return

    session = user_sessions.get(user_id)
    if not session:
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="「ん」または「テスト」と入力して最初から始めてください。")
        )
        return

    step = session["step"]
    session[step] = text if text != "スキップ" else ""

    step_order = [
        "inputter",
        "status",
        "company",
        "main_contact",
        "site_name",
        "branch",
        "request_details",
        "work_details",
        "work_month",
        "other_notes"
    ]

    next_step_index = step_order.index(step) + 1 if step in step_order else None
    if next_step_index is not None and next_step_index < len(step_order):
        session["step"] = step_order[next_step_index]
        ask_question(event.reply_token, step_order[next_step_index])
    else:
        finalize_and_record(event, session)
        del user_sessions[user_id]

def ask_question(reply_token, step):
    messages = {
        "status": ("② 案件進捗を選んでください", ["新規追加", "受注", "請求待ち", "請求済み"]),
        "company": ("③ 会社名を入力してください", None),
        "main_contact": ("④ 元請担当を入力してください（スキップ可）", None),
        "site_name": ("⑤ 現場名を入力してください", None),
        "branch": ("⑥ 拠点名を入力してください（スキップ可）", None),
        "request_details": ("⑦ 依頼内容を入力してください", None),
        "work_details": ("⑧ 施工内容を入力してください", None),
        "work_month": ("⑨ 作業月を選んでください", ["未定"] + [f"{(datetime.now().month + i - 1) % 12 + 1}月" for i in range(6)]),
        "other_notes": ("⑩ その他を入力してください（スキップ可）", None),
    }
    text, options = messages[step]
    if options:
        send_quick_reply(reply_token, text, options)
    else:
        line_bot_api.reply_message(reply_token, TextSendMessage(text=text))

def finalize_and_record(event, session):
    values = sheet.get_all_values()
    next_row = next(i+1 for i, row in enumerate(values) if not any(cell.strip() for cell in row))

    sheet.update_cell(next_row, 2, session.get("status", ""))
    sheet.update_cell(next_row, 3, session.get("inputter", ""))
    sheet.update_cell(next_row, 6, session.get("company", ""))
    sheet.update_cell(next_row, 7, session.get("branch", ""))
    sheet.update_cell(next_row, 9, session.get("site_name", ""))
    sheet.update_cell(next_row, 10, session.get("work_month", ""))
    sheet.update_cell(next_row, 11, session.get("contractor", ""))
    sheet.update_cell(next_row, 13, session.get("work_details", ""))

    report_to = event.source.user_id if session.get("test_mode") else report_group_id
    summary = f"{session.get('sender_name', session.get('inputter', 'ユーザー'))}さんが案件を登録しました。\n"
    summary += f"入力者：{session.get('inputter', '')}\n現場名：{session.get('site_name', '')}\n作業月：{session.get('work_month', '')}\n対応者：{session.get('contractor', '')}"
    line_bot_api.push_message(report_to, TextSendMessage(text=summary))

def send_quick_reply(token, text, options):
    quick_reply = QuickReply(items=[
        QuickReplyButton(action=MessageAction(label=opt, text=opt)) for opt in options
    ])
    line_bot_api.reply_message(token, TextSendMessage(text=text, quick_reply=quick_reply))

def get_user_display_name(user_id):
    try:
        profile = line_bot_api.get_profile(user_id)
        return profile.display_name
    except:
        return "ユーザー"

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
