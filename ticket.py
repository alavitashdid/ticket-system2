import pandas as pd
import random
import copy
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# 1. Excelファイルの読み込み
file_path = "チケット抽選.xlsx"
df = pd.read_excel(file_path)

# ★【追加】代表者氏名や人数が空の行（不要な空行）を削除する
df = df.dropna(subset=[df.columns[1]]) # 2列目の「代表者氏名」などが空の行を削除
df = df.reset_index(drop=True)

# 「人数」または「枚数」に関する列を自動検出して数値化する
target_column = None
for col in df.columns:
    if "人数" in str(col) or "枚数" in str(col):
        target_column = col
        break

if target_column:
    df[target_column] = df[target_column].astype(str).str.extract(r'(\d+)')[0]
    df[target_column] = pd.to_numeric(df[target_column], errors='coerce').fillna(1).astype(int)
else:
    target_column = df.columns[2]
    df[target_column] = df[target_column].astype(str).str.extract(r'(\d+)')[0]
    df[target_column] = pd.to_numeric(df[target_column], errors='coerce').fillna(1).astype(int)

# （以降の抽選ロジックはそのまま）

# 「当選」列を初期化
df["当選"] = "落選"

# 2. 抽選ロジック（上限126人に対応）
total_seats = 126
current_seats = 0
selected_indices = []
indices = list(df.index)

# ランダムに並べ替え
random.shuffle(indices)

# 抽選処理：なるべく126に近づけるために、入る人を順次探す
remaining_indices = copy.deepcopy(indices)
for idx in remaining_indices:
    tickets = df.loc[idx, target_column]
    if current_seats + tickets <= total_seats:
        selected_indices.append(idx)
        current_seats += tickets
        
    if current_seats == total_seats:
        break

# 3. 当選処理
df.loc[selected_indices, "当選"] = "当選"
df = df.sort_index()

if "@" in df.columns:
    df = df.drop(columns=["@"])

# 4. 結果保存とデザイン調整
output_path = "チケット抽選_抽選結果.xlsx"

with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Sheet1")
    workbook = writer.book
    worksheet = writer.sheets["Sheet1"]
    
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    header_font = Font(name="MS Pゴシック", size=11, bold=True)
    data_font = Font(name="MS Pゴシック", size=11, bold=False)
    
    for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column), start=1):
        for cell in row:
            if row_idx == 1:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.font = data_font
                # 中央揃えにしたい列（必要に応じて列番号を調整してください）
                if cell.column in [1, 3, 4]:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.alignment = Alignment(horizontal="left", vertical="center")

    # 列幅の自動調整
    for col in worksheet.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                val_str = str(cell.value)
                byte_len = len(val_str.encode('utf-8'))
                max_len = max(max_len, int(byte_len * 0.9))
        worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

print(f"【抽選完了】 合計当選人数: {current_seats}人 / 結果ファイルを作成しました: {output_path}")

