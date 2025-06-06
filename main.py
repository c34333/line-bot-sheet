（中略... 既存コード省略）

    elif step == "client":
        session["client"] = text if text != "スキップ" else ""
        session["step"] = "site"
        reply(event.reply_token, "④ 現場名を入力してください（スキップ可）")

    elif step == "site":
        session["site"] = text if text != "スキップ" else ""
        session["step"] = "branch"
        send_quick_reply(event.reply_token, "⑤ 拠点名を選んでください", ["本社", "関東", "前橋", "キャンセル"])

    elif step == "branch":
        session["branch"] = text
        session["step"] = "content"
        send_quick_reply(event.reply_token, "⑥ 依頼内容・ポイントを入力してください（スキップ可）", ["スキップ", "キャンセル"])

    elif step == "content":
        session["content"] = text if text != "スキップ" else ""
        session["step"] = "worktype"
        send_quick_reply(event.reply_token, "⑦ 施工内容を選んでください", ["洗浄", "清掃", "調査", "工事", "点検", "塗装", "修理", "キャンセル"])

    elif step == "worktype":
        session["worktype"] = text
        session["step"] = "month"
        today = datetime.today()
        current_month = today.month
        first_page = ["未定"] + [f"{m}月" for m in range(current_month, current_month + 6 if current_month <= 7 else 13)]
        send_quick_reply(event.reply_token, "⑧ 作業予定月を選んでください（1/2）", first_page + ["次へ ➡"])

    elif step == "month":
        if text == "次へ ➡":
            rest = [f"{m}月" for m in range((current_month + 6) % 12 or 12, (current_month - 1) % 12 + 1)]
            send_quick_reply(event.reply_token, "⑧ 作業予定月を選んでください（2/2）", rest + ["キャンセル"])
            return
        session["month"] = f"2025年{text}" if text != "未定" else "未定"
        session["step"] = "type"
        send_quick_reply(event.reply_token, "⑨ 対応者を選んでください", ["自社", "外注", "未定", "キャンセル"])

    elif step == "type":
        session["type"] = text
        session["step"] = "memo"
        send_quick_reply(event.reply_token, "⑩ その他入力があれば入力してください（スキップ可）", ["スキップ", "キャンセル"])

    elif step == "memo":
        session["memo"] = text if text != "スキップ" else ""

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
                sheet.update_cell(row, 13, session["memo"])
                sheet.update_cell(row, 14, session["client"])
                sheet.update_cell(row, 15, session["content"])
                a_number = sheet.cell(row, 1).value or str(row - 1)
            else:
                reply(event.reply_token, "⚠ スプレッドシートの空きが見つかりませんでした。")
                del user_sessions[user_id]
                return

        summary = f"入力者：{session['inputter_name']}\n{session['inputter_name']}さんが案件を登録しました！（案件番号：{a_number}）\n\n" \
                  f"① 案件進捗：{session['status']}\n" \
                  f"② 会社名：{session['company']}\n" \
                  f"③ 元請担当：{session['client']}\n" \
                  f"④ 現場名：{session['site']}\n" \
                  f"⑤ 拠点名：{session['branch']}\n" \
                  f"⑥ 依頼内容・ポイント：{session['content']}\n" \
                  f"⑦ 施工内容：{session['worktype']}\n" \
                  f"⑧ 作業予定月：{session['month']}\n" \
                  f"⑨ 対応者：{session['type']}\n" \
                  f"⑩ その他：{session['memo']}"

        reply(event.reply_token, summary)

        if not session.get("test_mode"):
            line_bot_api.push_message(PushMessageRequest(
                to=report_group_id,
                messages=[TextMessage(text=summary)]
            ))

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

def get_company_list_by_head(head):
    values = ref_sheet.get_all_values()
    companies = []
    for row in values:
        if len(row) >= 17:
            if row[15] == head:
                companies.append(row[16])
    return companies

if __name__ == "__main__":
    print(">>> Flask App Starting <<<")
    app.run(host="0.0.0.0", port=5000)
