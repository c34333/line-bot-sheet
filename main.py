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
        user_sessions[user_id] = {"step": "inputter"}
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

    if session["step"] == "inputter":
        session["inputter"] = text
        session["step"] = "company_initial"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="② 会社名の頭文字（ひらがな1文字）を入力してください。\n新規の場合は「新規」と入力してください。")
        )
        return

    if session["step"] == "company_initial":
        session["company_initial"] = text
        session["step"] = "site_name"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="③ 現場名を入力してください。")
        )
        return

    if session["step"] == "site_name":
        session["site_name"] = text
        session["step"] = "work_month"
        now = datetime.now()
        month_options = ["未定"] + [f"{(now.month + i - 1) % 12 + 1}月" for i in range(6)]
        send_quick_reply(event.reply_token, "④ 作業月を選んでください。", month_options)
        return

    if session["step"] == "work_month":
        session["work_month"] = text
        session["step"] = "contractor"
        send_quick_reply(event.reply_token, "⑤ 対応者を選んでください。", ["自社", "外注", "未定"])
        return

    if session["step"] == "contractor":
        session["contractor"] = text
        session["step"] = "work_details"
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="⑥ 施工内容を入力してください。スキップする場合は「スキップ」と入力してください。")
        )
        return

    if session["step"] == "work_details":
        session["work_details"] = text if text != "スキップ" else ""
        session["step"] = "final"
        finalize_and_record(event, session)
        del user_sessions[user_id]
        return

def finalize_and_record(event, session):
    values = sheet.get_all_values()
    next_row = next(i+1 for i, row in enumerate(values) if not any(cell.strip() for cell in row))
    sheet.update_cell(next_row, 3, session.get("inputter", ""))  # C列
    sheet.update_cell(next_row, 6, "会社名仮")                  # F列（今後の処理で会社名追加）
    sheet.update_cell(next_row, 9, session.get("site_name", ""))  # I列
    sheet.update_cell(next_row, 10, session.get("work_month", ""))# J列
    sheet.update_cell(next_row, 11, session.get("contractor", ""))# K列
    sheet.update_cell(next_row, 13, session.get("work_details", ""))# M列

    # 通知
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
