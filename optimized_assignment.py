import pandas as pd
import itertools

def run_optimization(uploaded_file):
    # Read the uploaded Excel file
    df = pd.read_excel(uploaded_file, engine='openpyxl')

    # 固定項目（ID 1〜13）を抽出
    fixed_df = df[df['固定'] == 1].copy()

    # 最適化対象（ID 14〜）を抽出
    target_df = df[df['固定'] != 1].copy()

    # 最適化対象の順列を生成
    permutations = list(itertools.permutations(target_df.index))

    best_score = -1
    best_order = None

    # 各順列についてスコア平均を計算
    for perm in permutations:
        ordered_df = target_df.loc[list(perm)].reset_index(drop=True)
        combined_df = pd.concat([fixed_df.reset_index(drop=True), ordered_df], ignore_index=True)
        avg_score = combined_df['スコア'].mean()
        if avg_score > best_score:
            best_score = avg_score
            best_order = combined_df

    # 結果をExcelに保存
    output_path = "optimized_assignment.xlsx"
    best_order.to_excel(output_path, index=False)
    return output_path
