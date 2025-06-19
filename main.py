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

    # コマンドエイリアス
    command_aliases = {
        "案件追加": "ん",
        "新規案件": "ん",
        "テスト": "テスト",
        "確認": "テスト",
    }
    command = command_aliases.get(text, text)

    # リセット対応
    if command in ["リセット", "キャンセル"]:
        user_sessions.pop(user_id, None)
        line_bot_api.reply_message(
            event.reply_token,
            TextSendMessage(text="リセットしました。再度コマンドを入力してください。")
        )
        return

    # サイレントグループ制御
    if group_id in silent_group_ids and user_id not in user_sessions and command not in ["ん", "テスト"]:
        return

    if command == "ん":
        user_sessions[user_id] = {"step": "inputter"}
        send_quick_reply(event.reply_token, "① 入力者を選んでください", ["未定", "諸橋", "酒井", "大塚", "原", "関野", "志賀", "加勢", "藤巻"])
    elif command == "テスト":
        user_sessions[user_id] = {"step": "inputter", "test_mode": True, "sender_name": get_user_display_name(user_id)}
        send_quick_reply(event.reply_token, "① 入力者を選んでください", ["未定", "諸橋", "酒井", "大塚", "原", "関野", "志賀", "加勢", "藤巻"])
    else:
        session = user_sessions.get(user_id)
        if not session:
            line_bot_api.reply_message(
                event.reply_token,
                TextSendMessage(text="「ん」または「テスト」と入力して最初から始めてください。")
            )
            return

        # 各種ステップ処理（略、元の詳細ステップをここに展開可能）
        # 例： if session["step"] == "inputter": ...

        # 最終ステップまで完了したら：
        # 1. スプレッドシートに転記
        # 2. 通常なら report_group_id に通知 / テストなら自分に通知

        # 最後にセッション削除
        # del user_sessions[user_id]


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
