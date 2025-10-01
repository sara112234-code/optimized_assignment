import pandas as pd
import itertools

def run_optimization(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # Must Pair と Must Separate の定義（必要に応じて変更）
    must_pair = [(2, 14), (5, 15)]
    must_separate = [(3, 16), (6, 17)]

    # ID 1〜13 を固定とみなす
    fixed_ids = set(range(1, 14))

    # Must Pair の ID を固定に追加
    for pair in must_pair:
        fixed_ids.update(pair)

    # 固定項目と最適化対象に分割
    fixed_df = df[df['ID'].isin(fixed_ids)].copy()
    optimize_df = df[~df['ID'].isin(fixed_ids)].copy()

    # 最適化対象の ID リスト
    optimize_ids = optimize_df['ID'].tolist()

    # 最適化対象の順列を生成
    best_score = -1
    best_order = None
    for perm in itertools.permutations(optimize_ids):
        # Must Separate のチェック
        invalid = False
        for a, b in must_separate:
            for i in range(len(perm) - 1):
                if (perm[i] == a and perm[i+1] == b) or (perm[i] == b and perm[i+1] == a):
                    invalid = True
                    break
            if invalid:
                break
        if invalid:
            continue

        # スコア平均の計算
        scores = [optimize_df[optimize_df['ID'] == i]['スコア'].values[0] for i in perm]
        avg_score = sum(scores) / len(scores)
        if avg_score > best_score:
            best_score = avg_score
            best_order = perm

    # 最適順序で並べ替え
    best_optimize_df = pd.concat([optimize_df[optimize_df['ID'] == i] for i in best_order])

    # 最終結果を結合して保存
    final_df = pd.concat([fixed_df, best_optimize_df])
    final_df.to_excel("optimized_assignment.xlsx", index=False)
