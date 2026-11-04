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

# Streamlit Community Cloud の正式URL（https:// で始める・末尾スラッシュなし）
WEB_APP_URL = 'https://ticket-system2-dye2ztspbtagmtnaneocwv.streamlit.app'
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


def _seat_cell_html(seat_code, is_disabled, seat_to_name, my_seats_set):
    """1マス分の座席HTMLを返す（メール互換のインラインスタイル）"""
    base_style = (
        "width:36px;height:32px;border:1px solid #333;border-radius:3px;"
        "text-align:center;vertical-align:middle;font-size:9px;font-weight:bold;"
        "padding:1px;line-height:1.1;"
    )

    if is_disabled:
        # 1列目・18列目などの使用不可席
        return (
            f'<td style="{base_style}background-color:#e0e0e0;color:#aaa;border-style:dashed;">'
            f'<div style="font-size:8px;">{seat_code}</div>'
            f'<div style="font-size:9px;">×</div></td>'
        )
    elif seat_code in my_seats_set:
        name = seat_to_name.get(seat_code, '')
        short_name = (name[:4] + '…') if len(name) > 4 else name
        return (
            f'<td style="{base_style}background-color:#4a90e2;color:#fff;border-color:#357abd;">'
            f'<div style="font-size:8px;opacity:0.9;">{seat_code}</div>'
            f'<div style="font-size:9px;">{short_name}</div></td>'
        )
    else:
        return (
            f'<td style="{base_style}background-color:#ffffff;color:#555;">'
            f'<div style="font-size:8px;">{seat_code}</div>'
            f'<div style="font-size:9px;color:#aaa;">-</div></td>'
        )


def generate_seat_map_html(member_list):
    """
    stage.py のレイアウトに合わせ、E〜G行および1・18列を除外したメール用座席表HTMLを生成する。
    """
    seat_to_name = {m['seat']: m['name'] for m in member_list}
    my_seats_set = set(seat_to_name.keys())

    rows_html = []

    rows_html.append(
        '<tr><td colspan="20" style="text-align:center;background-color:#eee;'
        'padding:6px;font-weight:bold;border:1px dashed #777;font-size:12px;">'
        '奥側</td></tr>'
    )
    rows_html.append('<tr><td colspan="20" style="height:6px;"></td></tr>')

    # 【修正】E〜G行を除外し、A〜D行のみを対象にする
    for row_label in ['D', 'C', 'B', 'A']:
        cells = []

        cells.append(
            f'<td style="width:28px;font-weight:bold;text-align:center;font-size:12px;">{row_label}</td>'
        )

        # 【修正】1〜18列目のうち、1列目と18列目は除外（使用不可）とする
        for s in range(1, 19):
            seat_code = f'{row_label}-{s}'
            is_disabled = (s in [1, 18])
            cells.append(_seat_cell_html(seat_code, is_disabled, seat_to_name, my_seats_set))

        cells.append(
            f'<td style="width:28px;font-weight:bold;text-align:center;font-size:12px;">{row_label}</td>'
        )

        rows_html.append('<tr>' + ''.join(cells) + '</tr>')

    rows_html.append('<tr><td colspan="20" style="height:8px;"></td></tr>')
    rows_html.append(
        '<tr><td colspan="20" style="text-align:center;background-color:#eee;'
        'padding:8px;font-weight:bold;border:1px dashed #777;font-size:13px;">'
        '🪑 ひな壇エリア (A行側・手前)</td></tr>'
    )

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
    base_url = WEB_APP_URL.rstrip('/')
    print(f"使用するWEBチケットURL: {base_url}")

    df_seats = pd.read_excel(SEAT_EXCEL_FILE)
    df_seats.columns = df_seats.columns.str.strip()

    unique_winners = df_seats.drop_duplicates(subset=['代表者名'])

    print(f"メール送信対象（当選グループ）: {len(unique_winners)}件")

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(SENDER_EMAIL, SENDER_PASSWORD)

    for _, row in unique_winners.iterrows():
        name = str(row['代表者名']).strip()

        # --- メールアドレス1とメールアドレス2の両方を安全に収集 ---
        emails = []
        for col_name in ['メールアドレス1', 'メールアドレス2']:
            if col_name in df_seats.columns:
                addr = str(row.get(col_name, '')).strip()
                if addr and addr != 'nan' and '@' in addr:
                    if addr not in emails:
                        emails.append(addr)
        
        if not emails:
            raise ValueError(f"【エラー】代表者「{name}」様のメールアドレスがどちらも見つかりませんでした。実行できませんでした。")

        count = row.get('グループ人数', '複数')

        assigned_seats, token_key, member_list = find_all_seats_for_user(name, df_seats)

        member_details_html = ""
        for m in member_list:
            member_details_html += (
                f"<li><b>{m['name']}</b> 様 "
                f"(座席: <span style='color:#2980b9;'>{m['seat']}</span>)</li>"
            )

        seat_map_html = generate_seat_map_html(member_list)

        ticket_url = f"{base_url}/?token={token_key}"

        # 収集したメールアドレスへそれぞれ個別に送信
        for to_email in emails:
            msg = MIMEMultipart('alternative')
            msg['From'] = SENDER_EMAIL
            msg['To'] = to_email
            msg['Subject'] = '【重要】当選通知および当日の受付・座席番号のお知らせ'

            # ---- テキスト版 ----
            text_body = f"""{name} 様

この度はご応募いただき誠にありがとうございました。
お申し込みいただいた人数分の席をご用意いたしましたのでご案内いたします。

【WEBチケット確認ページ】
下のURLをタップ（またはクリック）してください。
{ticket_url}

受付認証キー: {token_key}
割当座席番号: {assigned_seats}
お申込人数: {count} 名様

※当日は本人の学生証を持参してください。
※解禁日時まではQRコードは表示されません。
"""

            # ---- HTML版 ----
            html_body = f'''
<html>
<body style="font-family: sans-serif; color: #333; line-height: 1.6;">
<p><b>{name} 様</b></p>

<p>この度はご応募いただき誠にありがとうございました。<br>
お申し込みいただいた人数分の席をご用意いたしましたのでご案内いたします。</p>

<p>当日は以下の専用WEBチケット確認ページにアクセスし、画面に表示される情報を受付にてご提示ください。</p>

<div style="background-color: #fff0f0; border-left: 4px solid #b22222; padding: 12px; margin: 15px 0;">
  <p style="color: #b22222; font-weight: bold; margin: 0;">【重要：入場時の注意点】</p>
  <ul style="margin: 8px 0 0 0; padding-left: 20px; font-size: 14px;">
    <li>当日は必ずご本人の学生証をご持参ください（代表者様・同行者様含め全員分が必要です）。</li>
    <li>本人確認ができない場合、ご入場いただけません。学生証以外の証明書は無効です。</li>
  </ul>
</div>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<h3 style="color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 5px;">📍 ホール会場・座席配置のご案内</h3>

<table style="width: 100%; border-collapse: collapse; margin-bottom: 20px; font-size: 14px;">
  <tr>
    <td style="padding: 8px; border: 1px solid #ddd; background: #f9f9f9; width: 35%;"><b>公演日時</b></td>
    <td style="padding: 8px; border: 1px solid #ddd;">2026年11月7日(土)<br>最終整列時間 12:30<br>開場受付 12:40<br>開演 13:00-13:45</td>
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

<div style="margin: 25px 0; text-align: center;">
  <p style="font-weight: bold; color: #2c3e50; margin-bottom: 8px;">🎬 会場座席レイアウト（あなたの座席は青色）</p>
  <div style="overflow-x: auto; max-width: 100%;">
    {seat_map_html}
  </div>
</div>

<div style="text-align: center; margin: 30px 0;">
  <a href="{ticket_url}"
     target="_blank"
     rel="noopener noreferrer"
     style="background-color: #3498db; color: #ffffff; padding: 14px 28px; text-decoration: none; font-weight: bold; border-radius: 6px; display: inline-block; font-size: 16px;">
    🎫 WEBチケット確認ページを開く
  </a>
</div>

<p style="font-size: 14px; text-align: center; margin: 16px 0;">
  ボタンが開けない場合は、下の青いリンクをタップしてください。<br>
  <a href="{ticket_url}" target="_blank" rel="noopener noreferrer"
     style="color: #1a73e8; text-decoration: underline; word-break: break-all; font-size: 13px;">
    {ticket_url}
  </a>
</p>

<hr style="border: none; border-top: 1px solid #ddd; margin: 20px 0;">

<p><b>【注意事項】</b></p>
<ul style="margin: 8px 0; padding-left: 20px;">
  <li>当日は開演10分前までに会場へお越しいただき、上記の「WEBチケット確認ページ」画面をご提示ください。</li>
  <li>チケットの転売・譲渡は禁止です。転売・譲渡されたチケットは無効となり、ご入場をお断りいたします。</li>
  <li>また、転売等により生じたトラブルについて主催者は一切の責任を負いません。</li>
</ul>

<p>当日のご来場を心よりお待ちしております。</p>
</body>
</html>
'''

            msg.attach(MIMEText(text_body, 'plain', 'utf-8'))
            msg.attach(MIMEText(html_body, 'html', 'utf-8'))

            server.send_message(msg)
            print(f"送信完了: {name} 様 ({to_email}) -> 座席: {assigned_seats} [キー: {token_key}]")
            print(f"  リンク: {ticket_url}")
            time.sleep(1)

    server.quit()
    print("すべての送信処理が完了しました。")

except Exception as e:
    print(f"\n❌ エラーが発生しました。実行できませんでした: {e}")
