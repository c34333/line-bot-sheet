from flask import Flask, request, abort
import os
import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from linebot.v3.webhook import WebhookHandler
from linebot.v3.messaging import MessagingApi, ReplyMessageRequest, TextMessage, QuickReply, QuickReplyItem, MessageAction
from linebot.v3.exceptions import InvalidSignatureError

# Flaskアプリ
app = Flask(__name__)

# LINE APIの設定（環境変数から取得）
LINE_CHANNEL_ACCESS_TOKEN = os.environ['LINE_CHANNEL_ACCESS_TOKEN']
LINE_CHANNEL_SECRET = os.environ['LINE_CHANNEL_SECRET']

line_bot_api = MessagingApi(channel_access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# Google Sheets認証
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
credentials_info = json.loads(os.environ['GOOGLE_CREDENTIALS'])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_info, scope)
gc = gspread.authorize(credentials)
sheet = gc.open('LINEログ').sheet1

user_sessions = {}

@app.route("/callback", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(event_type='message')
def handle_message(event):
    user_id = event.source.user_id
    text = event.message.text.strip()

    def reply(reply_token, messages):
        if isinstance(messages, list):
            line_bot_api.reply_message(ReplyMessageRequest(reply_token=reply_token, messages=messages))
        else:
            line_bot_api.reply_message(ReplyMessageRequest(reply_token=reply_token, messages=[messages]))

    def text_msg(text):
        return TextMessage(text=text)

    def quick_reply_msg(prompt, options):
        return TextMessage(
            text=prompt,
            quick_reply=QuickReply(items=[QuickReplyItem(action=MessageAction(label=opt, text=opt)) for opt in options])
        )

    if text == "新規":
        user_sessions[user_id] = {"step": "awaiting_status"}
        reply(event.reply_token, quick_reply_msg("案件進捗を選んでください", ["新規追加", "3:受注", "4:作業完了", "定期"]))
        return

    if user_id in user_sessions:
        session = user_sessions[user_id]

        if session["step"] == "awaiting_status":
            session["status"] = text
            session["step"] = "awaiting_company"
            reply(event.reply_token, text_msg("① 会社名を入力してください"))
            return

        elif session["step"] == "awaiting_company":
            session["company"] = text
            session["step"] = "awaiting_introducer"
            reply(event.reply_token, text_msg("② 元請・紹介者名を入力してください"))
            return

        elif session["step"] == "awaiting_introducer":
            session["introducer"] = text
            session["step"] = "awaiting_site"
            reply(event.reply_token, text_msg("③ 現場名を入力してください"))
            return

        elif session["step"] == "awaiting_site":
            session["site"] = text
            session["step"] = "awaiting_branch"
            reply(event.reply_token, quick_reply_msg("④ 拠点名を選んでください", ["本社", "関東", "前橋"]))
            return

        elif session["step"] == "awaiting_branch":
            session["branch"] = f":{text}"
            session["step"] = "awaiting_content"
            reply(event.reply_token, quick_reply_msg("⑤ 依頼内容・ポイントを入力してください（スキップ可）", ["スキップ"]))
            return

        elif session["step"] == "awaiting_content":
            session["content"] = "" if text == "スキップ" else text
            session["step"] = "awaiting_month"
            reply(event.reply_token, quick_reply_msg("⑥ 作業予定月を選択してください", ["未定"] + [f"{i}月" for i in range(1, 13)]))
            return

        elif session["step"] == "awaiting_month":
            session["month"] = "未定" if text == "未定" else f"2025年{text}"
            session["step"] = "awaiting_worker"
            reply(event.reply_token, quick_reply_msg("⑦ 対応者を選択してください", ["自社", "外注"]))
            return

        elif session["step"] == "awaiting_worker":
            session["worker"] = text
            session["step"] = "awaiting_etc"
            reply(event.reply_token, quick_reply_msg("⑧ その他入力項目があれば入れてください（スキップ可）", ["スキップ"]))
            return

        elif session["step"] == "awaiting_etc":
            session["etc"] = "" if text == "スキップ" else text
            session["step"] = "done"

            for row in range(1, 2001):
                if sheet.cell(row, 2).value in [None, ""]:
                    case_number = sheet.cell(row, 1).value
                    sheet.update_cell(row, 2, session.get("status", ""))
                    sheet.update_cell(row, 5, session.get("company", ""))
                    sheet.update_cell(row, 6, session.get("branch", ""))
                    sheet.update_cell(row, 8, session.get("site", ""))
                    sheet.update_cell(row, 9, session.get("month", ""))
                    sheet.update_cell(row, 10, session.get("worker", ""))
                    reply(event.reply_token, [
                        text_msg(f"登録完了しました！（案件番号：{case_number}）"),
                        text_msg(
                            f"①会社名：{session.get('company', '')}\n"
                            f"②元請・紹介者名：{session.get('introducer', '')}\n"
                            f"③現場名：{session.get('site', '')}\n"
                            f"④拠点名：{session.get('branch', '')}\n"
                            f"⑤依頼内容・ポイント：{session.get('content', '')}\n"
                            f"⑥作業予定月：{session.get('month', '')}\n"
                            f"⑦対応者：{session.get('worker', '')}\n"
                            f"⑧その他：{session.get('etc', '')}"
                        )
                    ])
                    break
            return

    print(f"📝 {user_id} が \"{text}\" と送信 → 対応外メッセージ")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
