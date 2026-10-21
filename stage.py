import streamlit as st
import pandas as pd
import random
import io
import re

# --- セッションステートの初期化 ---
if "lottery_map" not in st.session_state:
    st.session_state.lottery_map = None
if 'export_df' not in st.session_state: 
    st.session_state.export_df = None

# ページの初期設定
st.set_page_config(page_title='3301教室 お笑いライブ座席抽選システム', layout='wide')

# CSSの定義（座席の初期色はすべて「白」、当選時のみブルーに変化）
css_data = '''
<style>
.stage-area { text-align: center; margin: 10px auto 20px auto; font-weight: bold; color: #333333; background-color: #f7e7d0; padding: 12px; border-radius: 4px; border: 2px solid #555; width: 60%; font-size: 16px; letter-spacing: 2px; }
.sub-area { text-align: center; font-weight: bold; color: #555; background-color: #eee; padding: 8px; border-radius: 4px; border: 1px dashed #777; margin: 10px auto; width: 60%; font-size: 14px; }
.grid-container { display: flex; flex-direction: column; align-items: center; margin-bottom: 30px; }
.row-container { display: flex; justify-content: center; align-items: center; margin-bottom: 4px; }
.row-label { width: 35px; font-weight: bold; text-align: center; color: #333; font-size: 13px; }
.block-center { display: flex; justify-content: center; }
.seat-box { width: 45px; height: 38px; margin: 1.5px; border-radius: 4px; display: inline-flex; flex-direction: column; align-items: center; justify-content: center; font-size: 11px; font-weight: bold; color: #333333; background-color: #ffffff; border: 1px solid #333333; }
.seat-assigned { background-color: #4a90e2 !important; color: white !important; border: 1px solid #357abd !important; box-shadow: 0px 2px 4px rgba(0,0,0,0.1); }
.seat-id { font-size: 9px; opacity: 0.8; }
.member-name { font-size: 10px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; max-width: 40px; }
.side-box { width: 100px; height: 260px; border: 2px solid #333; display: flex; align-items: center; justify-content: center; font-weight: bold; font-size: 14px; background-color: #f9f9f9; margin: 0 10px; text-align: center; }
</style>
'''
st.markdown(css_data, unsafe_allow_html=True)

# --- 座席マスタの定義 (A行が手前・ひな壇側、G行が奥) ---
layout = {
    'A': list(range(1, 19)),
    'B': list(range(1, 19)),
    'C': list(range(1, 19)),
    'D': list(range(1, 19)),
    'E': list(range(1, 19)),
    'F': list(range(1, 19)),
    'G': list(range(1, 19))
}

ordered_valid_seats = []
for row in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
    for s in layout[row]:
        ordered_valid_seats.append(f'{row}-{s}')

# =====================================================================
# 座席抽選メイン処理
# =====================================================================
st.title('🏫 3301教室 お笑いライブ 座席抽選システム')

uploaded_file = st.sidebar.file_uploader('『チケット抽選_抽選結果.xlsx』をアップロード', type=['xlsx', 'xls'])
st.sidebar.info(f'総座席数: {len(ordered_valid_seats)} 席 (7行 × 18列)')

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        # カラム名の自動検出
        result_col = next((c for c in df.columns if '当選' in str(c)), None)
        num_col = next((c for c in df.columns if '人数' in str(c)), None)
        mail_col = next((c for c in df.columns if 'メール' in str(c) or '@' in str(c)), None)
        
        rep_name_col = next((c for c in df.columns if '代表者氏名' in str(c) or '名前' in str(c)), None)
        rep_id_col = next((c for c in df.columns if '代表者学籍番号' in str(c)), None)
        
        c1_name_col = next((c for c in df.columns if '同行者1 氏名' in str(c)), None)
        c1_id_col = next((c for c in df.columns if '同行者1学籍番号' in str(c)), None)
        
        c2_name_col = next((c for c in df.columns if '同行者2 氏名' in str(c)), None)
        c2_id_col = next((c for c in df.columns if '同行者2学籍番号' in str(c)), None)
        
        c3_name_col = next((c for c in df.columns if '同行者3 氏名' in str(c)), None)
        c3_id_col = next((c for c in df.columns if '同行者3学籍番号' in str(c)), None)
        
        if not result_col or not rep_name_col:
            st.error('エラー: エクセルファイルに「当選」または「代表者氏名」の列が見つかりません。')
        else:
            winners_df = df[df[result_col] == '当選'].copy()
            groups = []
            total_required_seats = 0
            
            for _, row in winners_df.iterrows():
                raw_count = str(row[num_col]) if (num_col and pd.notna(row[num_col])) else "1"
                numbers_only = re.sub(r'\D', '', raw_count)
                count = int(numbers_only) if numbers_only != "" else 1
                
                email = str(row[mail_col]).strip() if (mail_col and pd.notna(row[mail_col])) else '-'
                rep_name = str(row[rep_name_col]).strip()
                token = f"TOKEN-{abs(hash(email + rep_name)) % 1000000:06d}"
                
                members = []
                if pd.notna(row[rep_name_col]):
                    members.append({'name': str(row[rep_name_col]).strip(), 'student_id': str(row[rep_id_col]).strip() if (rep_id_col and pd.notna(row[rep_id_col])) else '-'})
                if count >= 2 and c1_name_col and pd.notna(row.get(c1_name_col)):
                    members.append({'name': str(row[c1_name_col]).strip(), 'student_id': str(row[c1_id_col]).strip() if (c1_id_col and pd.notna(row[c1_id_col])) else '-'})
                if count >= 3 and c2_name_col and pd.notna(row.get(c2_name_col)):
                    members.append({'name': str(row[c2_name_col]).strip(), 'student_id': str(row[c2_id_col]).strip() if (c2_id_col and pd.notna(row[c2_id_col])) else '-'})
                if count >= 4 and c3_name_col and pd.notna(row.get(c3_name_col)):
                    members.append({'name': str(row[c3_name_col]).strip(), 'student_id': str(row[c3_id_col]).strip() if (c3_id_col and pd.notna(row[c3_id_col])) else '-'})
                
                actual_count = len(members)
                total_required_seats += actual_count
                
                groups.append({
                    'rep_name': rep_name,
                    'count': actual_count,
                    'email': email,
                    'token': token,
                    'members': members
                })
            
            st.sidebar.success(f'名簿読み込み成功！ 必要総座席数: {total_required_seats}席')
            
            if total_required_seats > len(ordered_valid_seats):
                st.error(f'エラー: 必要座席数 ({total_required_seats}席) が全席数 ({len(ordered_valid_seats)}席) を超えています！')
            elif st.sidebar.button('🎉 当選者の座席抽選を実行'):
                seat_mapping = {}
                available_seats = ordered_valid_seats.copy()
                export_data = []
                
                random.shuffle(groups)
                
                for g in groups:
                    c = g['count']
                    assigned = False
                    
                    if c > 1:
                        start_indices = list(range(len(available_seats) - c + 1))
                        random.shuffle(start_indices)
                        for start_idx in start_indices:
                            candidate_seats = available_seats[start_idx : start_idx + c]
                            rows = [s.split('-')[0] for s in candidate_seats]
                            if len(set(rows)) == 1: # 同一列内での連席
                                for i, seat in enumerate(candidate_seats):
                                    member = g['members'][i]
                                    seat_mapping[seat] = member['name']
                                    available_seats.remove(seat)
                                    export_data.append({
                                        '座席番号': seat,
                                        '氏名': member['name'],
                                        '学籍番号': member['student_id'],
                                        '代表者名': g['rep_name'],
                                        'グループ人数': c,
                                        '代表者メールアドレス': g['email'],
                                        '受付キー': g['token']
                                    })
                                assigned = True
                                break
                    
                    if not assigned:
                        for i in range(c):
                            seat = available_seats.pop(0)
                            member = g['members'][i]
                            seat_mapping[seat] = member['name']
                            export_data.append({
                                '座席番号': seat,
                                '氏名': member['name'],
                                '学籍番号': member['student_id'],
                                '代表者名': g['rep_name'],
                                'グループ人数': c,
                                '代表者メールアドレス': g['email'],
                                '受付キー': g['token']
                            })
                
                st.session_state.lottery_map = seat_mapping
                export_df = pd.DataFrame(export_data).sort_values(by='座席番号').reset_index(drop=True)
                st.session_state.export_df = export_df
                st.balloons()
    except Exception as e:
        st.error(f'ファイル読み込みエラー: {e}')

# --- 図面のビジュアルレンダリング ---
results = st.session_state.lottery_map if st.session_state.lottery_map is not None else {}

st.subheader('🗺️ 3301教室 座席レイアウト図')

grid_html = "<div class='grid-container'>"
grid_html += "<div class='row-container'>"

# 左側スペース（教卓や袖など）
grid_html += "<div class='side-box'>教卓 / 受付<br>・椅子置き場</div>"

# 座席本体エリア（上側が奥のG行、下側が手前のA行になるように並び替え）
grid_html += "<div style='display: flex; flex-direction: column;'>"
for row_label in ['G', 'F', 'E', 'D', 'C', 'B', 'A']:
    grid_html += "<div class='row-container'>"
    grid_html += f"<div class='row-label'>{row_label}</div>"
    grid_html += "<div class='block-center'>"
    for s in layout[row_label]:
        seat_code = f'{row_label}-{s}'
        if seat_code in results:
            grid_html += f"<div class='seat-box seat-assigned'><div class='seat-id'>{seat_code}</div><div class='member-name'>{results[seat_code]}</div></div>"
        else:
            grid_html += f"<div class='seat-box'><div class='seat-id'>{seat_code}</div><div class='member-name'>-</div></div>"
    grid_html += "</div>"
    grid_html += f"<div class='row-label'>{row_label}</div>"
    grid_html += "</div>"
grid_html += "</div>"

# 右側スペース（通路）
grid_html += "<div class='side-box' style='writing-mode: vertical-rl; width: 50px; font-size: 13px; color: #555;'>通路</div>"

grid_html += "</div>" # row-container 終了

# 下部（A行側）にひな壇を配置
grid_html += "<div class='sub-area'>🪑 ひな壇エリア (A行側)</div>"

grid_html += "</div>" # grid-container 終了

st.markdown(grid_html, unsafe_allow_html=True)

# データ一覧・ダウンロードエリア
if st.session_state.export_df is not None:
    with st.expander("📊 座席決定の一覧表を確認（Excelダウンロードはこちら）", expanded=True):
        st.dataframe(st.session_state.export_df, use_container_width=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.export_df.to_excel(writer, index=False, sheet_name='座席抽選結果')
        buffer.seek(0)
        st.download_button(label="📥 stage_output_seats.xlsx をダウンロード", data=buffer, file_name="stage_output_seats.xlsx")

