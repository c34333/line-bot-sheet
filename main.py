# main.py

import os
import json
from datetime import datetime
import gspread
from flask import Flask, request, abort
from oauth2client.service_account import ServiceAccountCredentials
from linebot.v3.webhook import WebhookHandler
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage, QuickReply, QuickReplyItem, MessageAction, PushMessageRequest
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent
from linebot.v3.exceptions import InvalidSignatureError

app = Flask(__name__)

LINE_CHANNEL_ACCESS_TOKEN = os.environ['LINE_CHANNEL_ACCESS_TOKEN']
LINE_CHANNEL_SECRET = os.environ['LINE_CHANNEL_SECRET']
report_group_id = "C6736021a0854b9c9526fdea9cf5acfa1"
silent_group_ids = ["C6736021a0854b9c9526fdea9cf5acfa1", "Cac0760acd664e7fdfa7a40975c340351"]

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
api_client = ApiClient(configuration)
line_bot_api = MessagingApi(api_client)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)

# ✅ URLで指定（open_by_url）に変更
sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1aVg4VIJRkEyyVs7FLik0smlhujtU-0DW/edit#gid=1558050264').worksheet('基本入力')
ref_sheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1aVg4VIJRkEyyVs7FLik0smlhujtU-0DW/edit#gid=1558050264').worksheet('参照値')

user_sessions = {}

def find_next_available_row():
    col_b = sheet.col_values(2)
    for i in range(1, 2001):
        if i >= len(col_b) or col_b[i] == '':
            return i + 1
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
    group_id = getattr(event.source, 'group_id', None)
    text = event.message.text.strip()

    if text in ["リセット", "最初から"]:
        user_sessions[user_id] = {"step": "inputter"}
        send_quick_reply(event.reply_token, "① 入力者を選んでください", ["未定", "諸橋", "酒井", "大塚", "原", "関野", "志賀", "加勢", "藤巻"])
        return

    if group_id in silent_group_ids and user_id not in user_sessions:
        return

    if text in ["あ", "テスト"]:
        user_sessions[user_id] = {"step": "inputter", "test_mode": text == "テスト"}
        send_quick_reply(event.reply_token, "① 入力者を選んでください", ["未定", "諸橋", "酒井", "大塚", "原", "関野", "志賀", "加勢", "藤巻"])
        return

    session = user_sessions.get(user_id)
    if not session:
        reply(event.reply_token, "「あ」または「テスト」と入力して最初から始めてください。")
        return

    step = session.get("step")

    if step == "inputter":
        session["inputter_name"] = text
        session["step"] = "status"
        send_quick_reply(event.reply_token, "② 案件進捗を選んでください", ["新規追加", "3:受注", "4:作業完了", "定期", "キャンセル"])

    elif step == "status":
        session["status"] = text
        session["step"] = "company_head"
        reply(event.reply_token, "③ 会社名の頭文字（ひらがな）を入力してください または『新規』")

    elif step == "company_head":
        if text == "新規":
            session["step"] = "company_head_new"
            reply(event.reply_token, "③-1 新規登録する会社の頭文字を入力してください")
            return
        companies = get_company_list_by_head(text)
        if not companies:
            reply(event.reply_token, f"頭文字 '{text}' に該当する会社が見つかりませんでした")
        else:
            session["step"] = "company_select"
            session["company_head"] = text
            send_quick_reply(event.reply_token, "③-2 該当する会社を選んでください（0で頭文字再選択）", ["0"] + companies[:12])

    elif step == "company_select":
        if text == "0":
            session["step"] = "company_head"
            reply(event.reply_token, "③ 会社名の頭文字を再入力してください")
            return
        session["company"] = text
        session["step"] = "client"
        reply(event.reply_token, "④ 元請担当を入力してください（スキップ可）")

    elif step == "company_head_new":
        session["company_head_new"] = text
        session["step"] = "company_name_new"
        reply(event.reply_token, "③-2 新規会社名を入力してください")

    elif step == "company_name_new":
        new_company = text
        session["company"] = new_company
        session["step"] = "client"
        ref_sheet.append_row([session["company_head_new"], new_company], table_range="P:Q")
        reply(event.reply_token, "④ 元請担当を入力してください（スキップ可）")

# （以下、client～memo→転記＆通知ステップへと続きます）
