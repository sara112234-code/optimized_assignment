import pandas as pd
import random

def run_optimization(input_excel_path):
    # Excelファイルを読み込む（IDとスコア列を取得）
    df = pd.read_excel(input_excel_path, engine='openpyxl')

    # 列名の正規化（ID列とスコア列を探す）
    id_col = None
    score_col = None
    for col in df.columns:
        if str(col).strip().lower() in ['id', 'no.', '番号']:
            id_col = col
        if str(col).strip().lower() in ['スコア', 'average score', 'score']:
            score_col = col
    if id_col is None:
        raise ValueError("ID列が見つかりません。列名は 'ID' にしてください。")
    if score_col is None:
        raise ValueError("スコア列が見つかりません。列名は 'スコア' または 'average score' にしてください。")

    # 固定項目（IDが1〜13）を抽出
    fixed_df = df[df[id_col].between(1, 13)].copy()

    # Must PairとMust Separateのルール（例）
    must_pair = [(3, 16), (5, 18)]
    must_separate = [(6, 17), (8, 19)]

    # 固定項目にMust PairのIDを追加（重複を避ける）
    pair_ids = set([i for pair in must_pair for i in pair])
    for pid in pair_ids:
        if pid not in fixed_df[id_col].values:
            pair_row = df[df[id_col] == pid]
            if not pair_row.empty:
                fixed_df = pd.concat([fixed_df, pair_row], ignore_index=True)

    # 最適化対象（固定以外）
    fixed_ids = fixed_df[id_col].tolist()
    remaining_df = df[~df[id_col].isin(fixed_ids)].copy()

    # 最適化：3グループに分ける（空行を挿入し、Must Separateを厳密にチェック）
    best_groups = None
    best_score = float('inf')
    remaining_ids = remaining_df[id_col].tolist()

    for _ in range(1000):
        random.shuffle(remaining_ids)
        groups = [[], [], []]
        for i, rid in enumerate(remaining_ids):
            groups[i % 3].append(rid)

        # Must Separateチェック：同じグループに含まれていないか
        valid = True
        for a, b in must_separate:
            for group in groups:
                if a in group and b in group:
                    valid = False
                    break
            if not valid:
                break
        if not valid:
            continue

        # 各グループのスコア平均の分散を計算
        variances = []
        for group in groups:
            scores = remaining_df[remaining_df[id_col].isin(group)][score_col]
            if not scores.empty:
                variances.append(scores.mean())
        score_variance = pd.Series(variances).var()

        if score_variance < best_score:
            best_score = score_variance
            best_groups = groups

    # 最終結果を構築（空行を挿入）
    result_df = fixed_df.copy()
    for group in best_groups:
        group_df = remaining_df[remaining_df[id_col].isin(group)]
        result_df = pd.concat([result_df, group_df, pd.DataFrame({id_col: [''], score_col: ['']})], ignore_index=True)

    # 結果を保存
    output_excel = "optimized_assignment.xlsx"
    result_df.to_excel(output_excel, index=False)

    return output_excel
