import pandas as pd
import itertools
from openpyxl import Workbook

def run_optimization():
    data = {
        'ID': list(range(1, 21)),
        'Score': [80, 75, 90, 85, 70, 95, 88, 76, 82, 91, 77, 84, 79, 65, 68, 72, 74, 69, 73, 71]
    }
    df = pd.DataFrame(data)

    fixed_df = df[df['ID'] <= 13].copy()
    target_df = df[df['ID'] > 13].copy()

    best_order = None
    best_score = -1
    for perm in itertools.permutations(target_df['ID'].tolist()):
        perm_scores = [df[df['ID'] == pid]['Score'].values[0] for pid in perm]
        avg_score = sum(perm_scores) / len(perm_scores)
        if avg_score > best_score:
            best_score = avg_score
            best_order = perm

    optimized_target_df = pd.DataFrame({'ID': best_order})
    optimized_target_df['Score'] = optimized_target_df['ID'].apply(lambda x: df[df['ID'] == x]['Score'].values[0])

    final_df = pd.concat([fixed_df, optimized_target_df], ignore_index=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Optimized Assignment"
    ws.append(["ID", "Score"])
    for _, row in final_df.iterrows():
        ws.append([row['ID'], row['Score']])
    wb.save("optimized_assignment.xlsx")
