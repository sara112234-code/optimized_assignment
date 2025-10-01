# 最適化された振り分け処理スクリプト
import pandas as pd
import itertools
import openpyxl
from openpyxl import Workbook

# 固定項目数
fixed_count = 13

# データフレームの作成
data = {'ID': [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20], 'Score': [80, 75, 90, 85, 70, 88, 92, 77, 83, 79, 81, 76, 84, 91, 73, 78, 82, 74, 86, 87]}
df = pd.DataFrame(data)

fixed_df = df.iloc[:fixed_count]
optimize_df = df.iloc[fixed_count:]

best_order = None
best_avg_score = -1
for perm in itertools.permutations(optimize_df['ID']):
    perm_scores = [df[df['ID'] == pid]['Score'].values[0] for pid in perm]
    avg_score = sum(perm_scores) / len(perm_scores)
    if avg_score > best_avg_score:
        best_avg_score = avg_score
        best_order = perm

final_ids = list(fixed_df['ID']) + list(best_order)
final_scores = [df[df['ID'] == pid]['Score'].values[0] for pid in final_ids]

wb = Workbook()
ws = wb.active
ws.title = "Optimized Assignment"
ws.append(["ID", "Score"])
for pid, score in zip(final_ids, final_scores):
    ws.append([pid, score])
wb.save("optimized_assignment.xlsx")
