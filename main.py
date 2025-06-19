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

sheet = gc.open_by_key('1Y3-XMKstJIfqBd55JtIdB4Ymx9zs6DFQo2U78fuFyME').worksheet('基本入力')
ref_sheet = gc.open_by_key('1Y3-XMKstJIfqBd55JtIdB4Ymx9zs6DFQo2U78fuFyME').worksheet('参照値')

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
        send_quick_reply(event.reply_token, "④ 元請担当を入力してください（スキップ可）", ["スキップ"])

    elif step == "company_head_new":
        session["company_head_new"] = text
        session["step"] = "company_name_new"
        reply(event.reply_token, "③-2 新規会社名を入力してください")

    elif step == "company_name_new":
        new_company = text
        session["company"] = new_company
        session["step"] = "client"
        ref_sheet.append_row([session["company_head_new"], new_company], table_range="P:Q")
        send_quick_reply(event.reply_token, "④ 元請担当を入力してください（スキップ可）", ["スキップ"])

    elif step == "client":
        session["client"] = "" if text == "スキップ" else text
        session["step"] = "site"
        send_quick_reply(event.reply_token, "⑤ 現場名を入力してください（スキップ可）", ["スキップ"])

    elif step == "site":
        session["site"] = "" if text == "スキップ" else text
        session["step"] = "branch"
        send_quick_reply(event.reply_token, "⑥ 拠点名を選んでください", ["本社", "関東", "前橋", "その他"])

    elif step == "branch":
        session["branch"] = text
        session["step"] = "content"
        send_quick_reply(event.reply_token, "⑦ 依頼内容を入力してください（スキップ可）", ["スキップ"])

    elif step == "content":
        session["content"] = "" if text == "スキップ" else text
        session["step"] = "worktype"
        send_quick_reply(event.reply_token, "⑧ 施工内容を選んでください", ["洗浄", "清掃", "調査", "工事", "点検", "塗装", "修理"])

    elif step == "worktype":
        session["worktype"] = text
        session["step"] = "month"
        now = datetime.now()
        months = ["未定"] + [f"{(now.month + i - 1) % 12 + 1}月" for i in range(6)]
        send_quick_reply(event.reply_token, "⑨ 作業月を選んでください", months)

    elif step == "month":
        session["month"] = text
        session["step"] = "memo"
        send_quick_reply(event.reply_token, "⑩ その他特記事項があれば入力してください（スキップ可）", ["スキップ"])

    elif step == "memo":
        session["memo"] = "" if text == "スキップ" else text

        if session.get("test_mode"):
            reply(event.reply_token, "テストモードのためスプレッドシートには転記されません")
        else:
            row = find_next_available_row()
            if row:
                sheet.update_cell(row, 2, session["status"])
                sheet.update_cell(row, 3, session["inputter_name"])
                sheet.update_cell(row, 6, session["company"])
                sheet.update_cell(row, 7, session["branch"])
                sheet.update_cell(row, 8, session["client"])
                sheet.update_cell(row, 9, session["site"])
                sheet.update_cell(row, 10, session["month"])
                sheet.update_cell(row, 11, session["inputter_name"])
                sheet.update_cell(row, 12, session["worktype"])
                sheet.update_cell(row, 13, session["content"])
                sheet.update_cell(row, 14, session["memo"])

            summary = f"{session['inputter_name']}さんが案件を登録しました！\n\n" \
                      f"① 入力者：{session['inputter_name']}\n" \
                      f"② 案件進捗：{session['status']}\n" \
                      f"③ 会社名：{session['company']}\n" \
                      f"④ 元請担当：{session['client']}\n" \
                      f"⑤ 現場名：{session['site']}\n" \
                      f"⑥ 拠点名：{session['branch']}\n" \
                      f"⑦ 依頼内容：{session['content']}\n" \
                      f"⑧ 施工内容：{session['worktype']}\n" \
                      f"⑨ 作業月：{session['month']}\n" \
                      f"⑩ その他：{session['memo']}"

            line_bot_api.push_message(PushMessageRequest(to=report_group_id, messages=[TextMessage(text=summary)]))
            reply(event.reply_token, "✅ 案件を登録しました！")

        del user_sessions[user_id]

def get_company_list_by_head(head):
    values = ref_sheet.get_all_values()
    return [row[1] for row in values if row[0] == head and len(row) > 1]

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
# Flask アプリケーションをポート5000で起動（Render対応）
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
