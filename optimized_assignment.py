import pandas as pd
import itertools

def run_optimization(uploaded_file):
    df = pd.read_excel(uploaded_file, engine='openpyxl')
    columns = df.columns.str.lower()
    score_col = None
    for col in df.columns:
        if col.lower() in ['スコア', 'average score']:
            score_col = col
            break
    if score_col is None:
        raise ValueError("スコア列が見つかりません。'スコア' または 'average score' の列名を使用してください。")

    if 'ID' not in df.columns:
        raise ValueError("ID列が見つかりません。列名は 'ID' にしてください。")

    fixed_df = df[df['ID'].between(1, 13)].copy()

    must_pair = [(3, 16)]
    for pair in must_pair:
        for pid in pair:
            if pid not in fixed_df['ID'].values:
                pair_df = df[df['ID'] == pid]
                fixed_df = pd.concat([fixed_df, pair_df])

    fixed_ids = fixed_df['ID'].tolist()
    remaining_df = df[~df['ID'].isin(fixed_ids)].copy()

    must_separate = [(6, 17)]

    best_order = None
    best_score = -1
    for perm in itertools.permutations(remaining_df.to_dict('records')):
        ids = [row['ID'] for row in perm]
        valid = True
        for a, b in must_separate:
            if a in ids and b in ids:
                if abs(ids.index(a) - ids.index(b)) == 1:
                    valid = False
                    break
        if not valid:
            continue
        avg_score = sum([row[score_col] for row in perm]) / len(perm)
        if avg_score > best_score:
            best_score = avg_score
            best_order = perm

    optimized_df = pd.concat([fixed_df, pd.DataFrame(best_order)])
    optimized_df.reset_index(drop=True, inplace=True)

    output_path = "optimized_assignment.xlsx"
    optimized_df.to_excel(output_path, index=False)

    return output_path
