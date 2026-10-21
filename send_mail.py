import pandas as pd
import smtplib
import time
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# =====================================================================
# 【設定項目】
# =====================================================================
MAIL_SERVICE = 'gmail'
SENDER_EMAIL = 'yk081323434@gmail.com'
SENDER_PASSWORD = 'xuub mlmf wjaw nxed'

SEAT_EXCEL_FILE = 'stage_output_seats.xlsx'

# ngrokで発行された外部公開用URL（※ngrokを再起動するたびに変わるので注意してください）
WEB_APP_URL = 'https://legged-subsystem-matchless.ngrok-free.dev/'
# =====================================================================


def find_all_seats_for_user(name, df):
    """指定した代表者名に割り当てられた全座席、同行者情報、受付キーを取得する"""
    df = df.copy()
    df['代表者名'] = df['代表者名'].astype(str).str.strip()
    match = df[df['代表者名'] == name.strip()]

    if match.empty:
        return "", "", []

    seats = match['座席番号'].astype(str).str.strip().tolist()
    token = str(match.iloc[0]['受付キー']).strip()

    members = []
    for _, r in match.iterrows():
        m_name = str(r.get('氏名', '')).strip()
        m_seat = str(r.get('座席番号', '')).strip()
        if m_name and m_name != 'nan':
            members.append({'name': m_name, 'seat': m_seat})

    return ", ".join(seats), token, members


def _seat_cell_html(seat_code, seat_to_name, my_seats_set):
    """1マス分の座席HTMLを返す（メール互換のインラインスタイル）"""
    base_style = (
        "width:36px;height:32px;border:1px solid #333;border-radius:3px;"
        "text-align:center;vertical-align:middle;font-size:9px;font-weight:bold;"
        "padding:1px;line-height:1.1;"
    )

    if seat_code in my_seats_set:
        # 自分のグループの席 → 青くハイライト + 名前表示
        name = seat_to_name.get(seat_code, '')
        short_name = (name[:4] + '…') if len(name) > 4 else name
        return (
            f'<td style="{base_style}background-color:#4a90e2;color:#fff;border-color:#357abd;">'
            f'<div style="font-size:8px;opacity:0.9;">{seat_code}</div>'
            f'<div style="font-size:9px;">{short_name}</div></td>'
        )
    else:
        # 他の席 or 空き
        return (
            f'<td style="{base_style}background-color:#ffffff;color:#555;">'
            f'<div style="font-size:8px;">{seat_code}</div>'
            f'<div style="font-size:9px;color:#aaa;">-</div></td>'
        )


def generate_seat_map_html(member_list):
    """
    stage.py のレイアウトをメール互換のテーブルHTMLに変換する。
    自分のグループの座席だけ青くハイライトし、名前も表示する。
    """
    # 座席コード → 名前 の辞書と、自分の座席セットを作成
    seat_to_name = {m['seat']: m['name'] for m in member_list}
    my_seats_set = set(seat_to_name.keys())

    rows_html = []

    # ---- 奥側ラベル ----
    rows_html.append(
        '<tr><td colspan="20" style="text-align:center;background-color:#eee;'
        'padding:6px;font-weight:bold;border:1px dashed #777;font-size:12px;">'
        '奥側</td></tr>'
    )
    rows_html.append('<tr><td colspan="20" style="height:6px;"></td></tr>')

    # stage.py と同じ並び順（G行が奥・上、A行が手前・下）
    for row_label in ['G', 'F', 'E', 'D', 'C', 'B', 'A']:
        cells = []

        # 行ラベル（左）
        cells.append(
            f'<td style="width:28px;font-weight:bold;text-align:center;font-size:12px;">{row_label}</td>'
        )

        # 座席 1〜18
        for s in range(1, 19):
            seat_code = f'{row_label}-{s}'
            cells.append(_seat_cell_html(seat_code, seat_to_name, my_seats_set))

        # 行ラベル（右）
        cells.append(
            f'<td style="width:28px;font-weight:bold;text-align:center;font-size:12px;">{row_label}</td>'
        )

        rows_html.append('<tr>' + ''.join(cells) + '</tr>')

    # ---- ひな壇エリア（A行側・手前） ----
    rows_html.append('<tr><td colspan="20" style="height:8px;"></td></tr>')
    rows_html.append(
        '<tr><td colspan="20" style="text-align:center;background-color:#eee;'
        'padding:8px;font-weight:bold;border:1px dashed #777;font-size:13px;">'
        '🪑 ひな壇エリア (A行側・手前)</td></tr>'
    )

    # 凡例
    rows_html.append('<tr><td colspan="20" style="height:12px;"></td></tr>')
    rows_html.append(
        '<tr><td colspan="20" style="font-size:11px;color:#555;text-align:center;">'
        '<span style="display:inline-block;width:14px;height:14px;background:#4a90e2;'
        'border:1px solid #357abd;vertical-align:middle;margin-right:4px;"></span>'
        'あなたの座席 &nbsp;&nbsp;'
        '<span style="display:inline-block;width:14px;height:14px;background:#fff;'
        'border:1px solid #333;vertical-align:middle;margin-right:4px;"></span>'
        'その他の座席'
        '</td></tr>'
    )

    table = (
        '<table cellpadding="0" cellspacing="2" border="0" '
        'style="border-collapse:collapse;margin:0 auto;font-family:sans-serif;">'
        + ''.join(rows_html)
        + '</table>'
    )
    return table


# =====================================================================
# メイン処理
# =====================================================================
try:
    print(f"使用するWEBチケットURL: {WEB_APP_URL}")

    df_seats = pd.read_excel(SEAT_EXCEL_FILE)
    df_seats.columns = df_seats.columns.str.strip()

    unique_winners = df_seats.drop_duplicates(subset=['代表者名'])

    print(f"メール送信対象（当選グループ）: {len(unique_winners)}件")

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)

    for _, row in unique_winners.iterrows():
        name = str(row['代表者名']).strip()
        to_email = row['代表者メールアドレス']  # テスト時はご自身のメールアドレスに書き換えてください
        count = row.get('グループ人数', '複数')

        assigned_seats, token_key, member_list = find_all_seats_for_user(name, df_seats)

        # グループメンバーの座席詳細テキストを作成
        member_details_html = ""
        for m in member_list:
            member_details_html += (
                f"<li><b>{m['name']}</b> 様 "
                f"(座席: <span style='color:#2980b9;'>{m['seat']}</span>)</li>"
            )

        # ★ stage.py と同じレイアウトの座席マップを生成
        seat_map_html = generate_seat_map_html(member_list)

        # メールオブジェクトの作成（HTML形式）
        msg = MIMEMultipart('alternative')
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = '【重要】当選通知および当日の受付・座席番号のお知らせ'

        ticket_url = f"{WEB_APP_URL}?token={token_key}"

        # HTML形式の本文（会場マップを追加）
        html_body = f'''
<html>
<body style="font-family: sans-serif; color: #333; line-height: 1.6;">
<p><b>{name} 様</b></p>

<p>この度はご応募いただき誠にありがとうございました。<br>
厳正なる抽選の結果、ご当選となりましたのでお知らせいたします。</p>

<p>当日は以下の専用WEBチケット確認ページにアクセスし、画面に表示される情報を受付にてご提示ください。</p>

<div style="background-color: #fff0f0; border-left: 4px solid #b22222; padding: 12px; margin: 15px 0;">
  <p style="color: #b22222; font-weight: bold; margin: 0;">【重要：入場時の注意点】</p>
  <p style="margin: 5px 0 0 0; font-size: 14px;">
    当日は必ず本人の学生証を持参してください（代表者本人、および同行者がいらっしゃる場合は全員分の学生証が必要です）。<br>
    本人確認ができない場合、ご入場いただくことができません。学生証以外の証明書は無効となりますのでご注意ください。<br>
    開演10分前にならないとサイトは開けません。
  </p>
</div>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<h3 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px;">📍 ホール会場・座席配置のご案内</h3>

<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px;">
  <tr>
    <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9; width: 35%;"><b>公演日時</b></td>
    <td style="padding: 8px; border: 1px solid #ddd;">2026年8月15日(土) 14:00開演 (受付開始 13:30)</td>
  </tr>
  <tr>
    <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><b>お申込人数</b></td>
    <td style="padding: 8px; border: 1px solid #ddd;">{count} 名様</td>
  </tr>
  <tr>
    <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><b>割当座席番号（全体）</b></td>
    <td style="padding: 8px; border: 1px solid #ddd; color: #2980b9; font-weight: bold; font-size: 16px;">{assigned_seats}</td>
  </tr>
  <tr>
    <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><b>参加者・座席内訳</b></td>
    <td style="padding: 8px; border: 1px solid #ddd;">
      <ul style="margin: 0; padding-left: 20px;">
        {member_details_html}
      </ul>
    </td>
  </tr>
  <tr>
    <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9;"><b>受付認証キー</b></td>
    <td style="padding: 8px; border: 1px solid #ddd; font-family: monospace;">{token_key}</td>
  </tr>
</table>

<!-- ★ 会場座席マップ（stage.py のレイアウトをメール互換テーブルで再現） -->
<div style="margin: 25px 0; text-align: center;">
  <p style="font-weight: bold; color: #2c3e50; margin-bottom: 8px;">🎬 会場座席レイアウト（あなたの座席は青色）</p>
  <div style="overflow-x: auto; max-width: 100%;">
    {seat_map_html}
  </div>
</div>

<div style="text-align: center; margin: 25px 0;">
  <a href="{ticket_url}" style="background-color: #3498db; color: white; padding: 12px 24px; text-decoration: none; font-weight: bold; border-radius: 5px; display: inline-block;">
    🎫 WEBチケット確認ページを開く
  </a>
</div>
<p style="font-size: 12px; color: #666; word-break: break-all; text-align: center;">
  うまくボタンから開けない場合は、以下のURLをブラウザに貼り付けてください。<br>
  {ticket_url}
</p>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<p><b>※注意事項</b><br>
1: 当日は上記のWEBチケット確認ページにアクセスして画面をご提示ください。<br>
2: 第三者への譲渡や転売によるトラブルについて、主催者は一切の責任を負いません。</p>

<p>当日のご来場を心よりお待ちしております。</p>
</body>
</html>
'''
        msg.attach(MIMEText(html_body, 'html', 'utf-8'))

        server.send_message(msg)
        print(f"送信完了: {name} 様 ({to_email}) -> 座席: {assigned_seats} [キー: {token_key}]")
        time.sleep(1)

    server.quit()
    print("すべての送信処理が完了しました。")

except Exception as e:
    print(f"エラーが発生しました: {e}")
