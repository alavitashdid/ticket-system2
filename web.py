import streamlit as st
import pandas as pd
import qrcode
import io
import streamlit.components.v1 as components
from datetime import datetime, timezone, timedelta

try:
    from zoneinfo import ZoneInfo
    JST = ZoneInfo("Asia/Tokyo")
except Exception:
    JST = timezone(timedelta(hours=9))

st.set_page_config(page_title='入場チケット確認ページ', layout='centered')

# --- スタイリング用のCSS ---
st.markdown("""
<style>
    .ticket-box {
        background-color: #ffffff;
        border: 2px solid #e0e0e0;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 25px;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.05);
    }
</style>
""", unsafe_allow_html=True)

# ----------------------------------------------
# ★ 公開日時（指定時間）の設定【日本時間】
RELEASE_YEAR = 2026
RELEASE_MONTH = 11
RELEASE_DAY = 7
RELEASE_HOUR = 12
RELEASE_MINUTE = 30
RELEASE_SECOND = 0
# ----------------------------------------------

# 表示用の解禁日時文字列（24時間表記）
RELEASE_LABEL = f"{RELEASE_YEAR}年{RELEASE_MONTH}月{RELEASE_DAY}日 {RELEASE_HOUR:02d}:{RELEASE_MINUTE:02d}"

SEAT_EXCEL_FILE = 'stage_output_seats.xlsx'

st.title("🎫 入場チケット・受付QRコード確認画面")

try:
    df_seats = pd.read_excel(SEAT_EXCEL_FILE)
except Exception as e:
    st.error(f"ファイル読み込みエラー: {e}")
    st.stop()

df_seats['受付キー'] = df_seats['受付キー'].astype(str).str.strip()

st.sidebar.header("🔑 チケット認証")
default_token = df_seats['受付キー'].iloc[0] if len(df_seats) > 0 else ""

# メール等のURLパラメータ（?token=...）から安全にトークン取得
query_params = st.query_params
url_token = query_params.get("token", default_token)

input_token = st.sidebar.text_input("受付キー（トークン）を入力", value=url_token).strip()

user_rows = df_seats[df_seats['受付キー'] == input_token].to_dict(orient='records')

if len(user_rows) == 0:
    st.error(f"❌ 入力されたキー「{input_token}」に一致するデータがありません。")
    st.stop()

rep_name = user_rows[0]['代表者名']
st.markdown(f"### 👤 お申込グループ代表: {rep_name} 様")
st.markdown("---")

# ---------------------------------------------------------------------------
# 画面最上部の0.1秒単位のリアルタイム現在時刻時計（日本時間固定）
# ---------------------------------------------------------------------------
top_clock_html = f"""
<div style="font-family: monospace; font-size: 15px; color: #004085; background-color: #cce5ff; padding: 10px; border-radius: 5px; text-align: center; font-weight: bold; margin-bottom: 15px;" id="live-clock">
    🕒 現在時刻: 読み込み中...
</div>

<script>
    function updateTopClock() {{
        const now = new Date();
        const parts = new Intl.DateTimeFormat("ja-JP", {{
            timeZone: "Asia/Tokyo",
            year: "numeric",
            month: "2-digit",
            day: "2-digit",
            hour: "2-digit",
            minute: "2-digit",
            second: "2-digit",
            hourCycle: "h23"
        }}).formatToParts(now);

        const part = (type) => parts.find(item => item.type === type).value;
        const year = part("year");
        const month = part("month");
        const day = part("day");
        const hours = part("hour");
        const minutes = part("minute");
        const seconds = part("second");
        const deciseconds = Math.floor(now.getMilliseconds() / 100);

        document.getElementById('live-clock').innerText = "🕒 現在時刻: " + year + "年" + month + "月" + day + "日 " + hours + ":" + minutes + ":" + seconds + "." + deciseconds;
    }}
    setInterval(updateTopClock, 100);
    updateTopClock();
</script>
"""
components.html(top_clock_html, height=55)

# Python側の時間判定（日本時間で比較 → CloudのUTCでも正しく動く）
now = datetime.now(JST)
release_dt = datetime(
    RELEASE_YEAR, RELEASE_MONTH, RELEASE_DAY,
    RELEASE_HOUR, RELEASE_MINUTE, RELEASE_SECOND,
    tzinfo=JST
)

if now < release_dt:
    st.warning(f"🔒 指定時間までロックされています。\n\n**解禁日時: {RELEASE_LABEL}**\n\n各チケット上のカウントダウンが0になると自動で解禁されます。")
else:
    st.success("🎉 解禁されました！以下のQRコードを受付でご提示ください。")
    st.markdown('<span style="color: #d9534f; font-size: 14px; font-weight: bold;">⚠️ QR コードが表示されない場合は再度メールに添付されてるURLから入り直すか、webサイトの更新ボタン「↻」を押してください</span>', unsafe_allow_html=True)

# 各チケットの描画（各QRの上に0.1秒単位の専用タイマーを配置）
for idx, row in enumerate(user_rows):
    name = row['氏名']
    student_id = row['学籍番号']
    seat = row['座席番号']
    token = row['受付キー']

    with st.container():
        st.markdown('<div class="ticket-box">', unsafe_allow_html=True)

        # 各チケット専用の0.1秒単位タイマーコンポーネント（IDをユニークにする）
        timer_id = f"ticket-timer-{idx}"

        ticket_timer_html = f"""
        <div style="font-family: sans-serif; font-size: 13px; text-align: center; padding: 10px; border-radius: 4px; margin-bottom: 12px; font-weight: bold; line-height: 1.6;" id="{timer_id}">
            ⏳ 読み込み中...
        </div>

        <script>
            const releaseTime_{idx} = Date.parse("{RELEASE_YEAR}-{RELEASE_MONTH:02d}-{RELEASE_DAY:02d}T{RELEASE_HOUR:02d}:{RELEASE_MINUTE:02d}:{RELEASE_SECOND:02d}+09:00");
            const timerEl_{idx} = document.getElementById("{timer_id}");
            const releaseLabel_{idx} = "{RELEASE_YEAR}年{RELEASE_MONTH}月{RELEASE_DAY}日 {RELEASE_HOUR:02d}:{RELEASE_MINUTE:02d}";

            function updateTimer_{idx}() {{
                const now = new Date().getTime();
                const diff = releaseTime_{idx} - now;

                if (diff <= 0) {{
                    timerEl_{idx}.style.backgroundColor = "#d4edda";
                    timerEl_{idx}.style.color = "#155724";
                    timerEl_{idx}.innerHTML = "🎉 解禁済みチケット<br><span style='font-size:12px;font-weight:normal;'>解禁日時: " + releaseLabel_{idx} + "</span>";
                    if (window.location.hash !== "#unlocked_{idx}") {{
                        window.location.hash = "#unlocked_{idx}";
                        window.location.reload();
                    }}
                }} else {{
                    const days = Math.floor(diff / (1000 * 60 * 60 * 24));
                    const hours = Math.floor((diff % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
                    const minutes = Math.floor((diff % (1000 * 60 * 60)) / (1000 * 60));
                    const seconds = Math.floor((diff % (1000 * 60)) / 1000);

                    timerEl_{idx}.style.backgroundColor = "#fff3cd";
                    timerEl_{idx}.style.color = "#856404";
                    timerEl_{idx}.innerHTML =
                        "🔒 解禁日時: " + releaseLabel_{idx} +
                        "<br>⏳ 残り: " + days + "日 " + hours + "時間 " + minutes + "分 " + seconds + "秒";
                }}
            }}
            setInterval(updateTimer_{idx}, 100);
            updateTimer_{idx}();
        </script>
        """
        components.html(ticket_timer_html, height=70)

        # 時間前はプレースホルダー、時間後はQRコードを表示（日本時間で判定）
        if datetime.now(JST) < release_dt:
            st.info(f"📌 {name} 様のQRコードは指定時間（{RELEASE_LABEL}）に表示されます。")
        else:
            col1, col2 = st.columns([1, 2])

            with col1:
                qr = qrcode.QRCode(version=1, box_size=5, border=2)
                qr.add_data(f"{token}-{seat}-{student_id}")
                qr.make(fit=True)
                img = qr.make_image(fill_color="black", back_color="white")

                buf = io.BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                st.image(buf, width=150)

            with col2:
                st.markdown(f"""
                <h4>入場者情報</h4>
                ・氏名: <b>{name} 様</b><br>
                ・学籍番号: <b>{student_id}</b><br>
                ・割当座席: <span style="color: #d9534f; font-size: 18px;"><b>{seat}</b></span><br>
                ・認証キー: <span style="font-size: 11px; color: #888;">{token}</span>
                """, unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)
