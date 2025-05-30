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

LINE_CHANNEL_ACCESS_TOKEN = os.environ['LINE_CHANNEL_ACCESS_TOKEN']
LINE_CHANNEL_SECRET = os.environ['LINE_CHANNEL_SECRET']

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
api_client = ApiClient(configuration)
line_bot_api = MessagingApi(api_client)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets setup
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)
sheet = gc.open('LINEログ').sheet1

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
    text = event.message.text.strip()

    # 強制リセットトリガー
    if text in ["あ", "テスト"]:
        user_sessions[user_id] = {
            "step": "inputter",
            "test_mode": text == "テスト",
            "inputter_page": 1
        }
        send_quick_reply(event.reply_token, "👤 入力者を選択してください（1/2）", ["未定", "諸橋", "酒井", "大塚", "原", "次へ ➡"])
        return

    # ID確認
    if text == "あなたのIDは？":
        msg = f"🆔 あなたのユーザーID:\n{user_id}"
        group_id = getattr(event.source, 'group_id', None)
        if group_id:
            msg += f"\n👥 グループID:\n{group_id}"
        reply(event.reply_token, msg)
        return

    # キャンセル処理
    if text == "キャンセル":
        if user_id in user_sessions:
            del user_sessions[user_id]
        reply(event.reply_token, "入力をキャンセルしました。最初からやり直してください。")
        return

    # セッションがない場合（例外）
    if user_id not in user_sessions:
        reply(event.reply_token, "「あ」または「テスト」と入力して最初から始めてください。")
        return

    session = user_sessions[user_id]
    step = session.get("step")

    if step == "inputter":
        if text == "次へ ➡":
            session["inputter_page"] = 2
            send_quick_reply(event.reply_token, "👤 入力者を選択してください（2/2）", ["関野", "志賀", "加勢", "藤巻", "キャンセル"])
            return
        session["inputter_name"] = text
        session["step"] = "status"
        send_quick_reply(event.reply_token, f"{text}さんですね。\n① 案件進捗を選んでください", ["新規追加", "3:受注", "4:作業完了", "定期", "キャンセル"])
    elif step == "status":
        session["status"] = text
        session["step"] = "company"
        reply(event.reply_token, "② 会社名を入力してください（キャンセル可）")
    elif step == "company":
        session["company"] = text
        session["step"] = "client"
        reply(event.reply_token, "③ 元請・紹介者名を入力してください（キャンセル可）")
    elif step == "client":
        session["client"] = text
        session["step"] = "site"
        reply(event.reply_token, "④ 現場名を入力してください（キャンセル可）")
    elif step == "site":
        session["site"] = text
        session["step"] = "branch"
        send_quick_reply(event.reply_token, "⑤ 拠点名を選んでください", ["本社", "関東", "前橋", "キャンセル"])
    elif step == "branch":
        session["branch"] = f":{text}"
        session["step"] = "content"
        send_quick_reply(event.reply_token, "⑥ 依頼内容・ポイントを入力してください（スキップ可）", ["スキップ", "キャンセル"])
    elif step == "content":
        session["content"] = "" if text == "スキップ" else text
        session["step"] = "worktype"
        send_quick_reply(event.reply_token, "⑦ 施工内容を選んでください", ["洗浄", "清掃", "調査", "工事", "点検", "塗装", "修理", "キャンセル"])
    elif step == "worktype":
        session["worktype"] = text
        session["step"] = "month"
        session["month_page"] = 1
        send_quick_reply(event.reply_token, "⑧ 作業予定月を選んでください（1/2）", ["未定", "1月", "2月", "3月", "4月", "5月", "6月", "次へ ➡"])
    elif step == "month":
        if text == "次へ ➡":
            session["month_page"] = 2
            send_quick_reply(event.reply_token, "⑧ 作業予定月を選んでください（2/2）", ["7月", "8月", "9月", "10月", "11月", "12月", "キャンセル"])
            return
        session["month"] = f"2025年{text}" if text != "未定" else "未定"
        session["step"] = "type"
        send_quick_reply(event.reply_token, "⑨ 対応者を選んでください", ["自社", "外注", "未定", "キャンセル"])
    elif step == "type":
        session["type"] = text
        session["step"] = "memo"
        send_quick_reply(event.reply_token, "⑩ その他入力項目があれば入力してください（スキップ可）", ["スキップ", "キャンセル"])
    elif step == "memo":
        session["memo"] = "" if text == "スキップ" else text

        profile = line_bot_api.get_profile(user_id)
        display_name = profile.display_name

        if session.get("test_mode"):
            a_number = "テスト"
        else:
            row = find_next_available_row()
            if row:
                sheet.update_cell(row, 2, format_status(session["status"]))
                sheet.update_cell(row, 3, session["inputter_name"])
                sheet.update_cell(row, 6, session["company"])
                sheet.update_cell(row, 7, session["branch"])
                sheet.update_cell(row, 9, session["site"])
                sheet.update_cell(row, 10, session["month"])
                sheet.update_cell(row, 11, session["type"])
                sheet.update_cell(row, 12, session["worktype"])
                a_number = sheet.cell(row, 1).value or str(row - 1)
            else:
                reply(event.reply_token, "⚠ スプレッドシートの空きが見つかりませんでした。")
                del user_sessions[user_id]
                return

        summary = f"{display_name}さんが案件を登録しました！（案件番号：{a_number}）\n\n" \
                  f"入力者：{session['inputter_name']}\n" \
                  f"① 案件進捗：{session['status']}\n" \
                  f"② 会社名：{session['company']}\n" \
                  f"③ 元請・紹介者名：{session['client']}\n" \
                  f"④ 現場名：{session['site']}\n" \
                  f"⑤ 拠点名：{session['branch']}\n" \
                  f"⑥ 依頼内容・ポイント：{session['content']}\n" \
                  f"⑦ 施工内容：{session['worktype']}\n" \
                  f"⑧ 作業予定月：{session['month']}\n" \
                  f"⑨ 対応者：{session['type']}\n" \
                  f"⑩ その他：{session['memo']}"

        reply(event.reply_token, summary)
        del user_sessions[user_id]

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
    return status if status in ["3:受注", "4:作業完了", "定期"] else "新規追加"

if __name__ == "__main__":
    print(">>> Flask App Starting <<<")
    app.run(host="0.0.0.0", port=5000)
