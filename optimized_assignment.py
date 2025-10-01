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

    # 軽量化：ランダムに1000通りの順列を試す
    best_order = None
    best_score = -1
    remaining_ids = remaining_df[id_col].tolist()

    for _ in range(1000):
        random.shuffle(remaining_ids)

        # Must Separateのチェック（隣接していないか）
        valid = True
        for a, b in must_separate:
            for i in range(len(remaining_ids) - 1):
                if (remaining_ids[i] == a and remaining_ids[i + 1] == b) or (remaining_ids[i] == b and remaining_ids[i + 1] == a):
                    valid = False
                    break
            if not valid:
                break
        if not valid:
            continue

        # スコア平均を計算
        temp_df = remaining_df.set_index(id_col).loc[remaining_ids].reset_index()
        avg_score = temp_df[score_col].mean()

        if avg_score > best_score:
            best_score = avg_score
            best_order = remaining_ids.copy()

    # 最終結果を結合
    optimized_df = pd.concat([
        fixed_df,
        remaining_df.set_index(id_col).loc[best_order].reset_index()
    ], ignore_index=True)

    # 結果を保存
    output_excel = "optimized_assignment.xlsx"
    optimized_df.to_excel(output_excel, index=False)

    return output_excel
