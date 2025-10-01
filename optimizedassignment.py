import pandas as pd
import random

def run_optimization(input_excel_path):
    df = pd.read_excel(input_excel_path, engine='openpyxl')

    # 列名の正規化
    id_col = next((col for col in df.columns if str(col).strip().lower() in ['id', 'no.', '番号']), None)
    score_col = next((col for col in df.columns if str(col).strip().lower() in ['スコア', 'average score', 'score']), None)
    if id_col is None or score_col is None:
        raise ValueError("ID列またはスコア列が見つかりません。")

    # 関係性列の抽出
    def extract_relations(row, colname):
        if pd.isna(row[colname]):
            return []
        return [int(x) for x in str(row[colname]).split(',') if x.strip().isdigit()]

    relations = {'must_separate': {}, 'must_pair': {}, 'prefer_separate': {}, 'prefer_pair': {}}
    for idx, row in df.iterrows():
        id_val = row[id_col]
        for key in relations:
            rel_ids = extract_relations(row, key.replace('_', ' ').title())
            if id_val not in relations[key]:
                relations[key][id_val] = set()
            relations[key][id_val].update(rel_ids)

    # グループ構成（m15f20, m15f21, m14f21）
    group_limits = [{'m': 15, 'f': 20}, {'m': 15, 'f': 21}, {'m': 14, 'f': 21}]
    best_assignment = None
    best_score_var = float('inf')

    for _ in range(1000):
        shuffled = df.sample(frac=1).reset_index(drop=True)
        groups = [[], [], []]
        gender_counts = [{'m': 0, 'f': 0} for _ in range(3)]
        valid = True

        for _, row in shuffled.iterrows():
            id_val = row[id_col]
            gender = row['gender']
            assigned = False

            for i in range(3):
                if gender_counts[i][gender] < group_limits[i][gender]:
                    conflict = False
                    for other in groups[i]:
                        for rule in ['must_separate']:
                            if id_val in relations[rule] and other in relations[rule][id_val]:
                                conflict = True
                                break
                        if conflict:
                            break
                    if not conflict:
                        groups[i].append(id_val)
                        gender_counts[i][gender] += 1
                        assigned = True
                        break
            if not assigned:
                valid = False
                break

        if not valid:
            continue

        # スコア平均の分散を計算
        variances = []
        for group in groups:
            scores = df[df[id_col].isin(group)][score_col]
            if not scores.empty:
                variances.append(scores.mean())
        score_var = pd.Series(variances).var()

        if score_var < best_score_var:
            best_score_var = score_var
            best_assignment = groups

    # 結果の構築
    result_df = pd.DataFrame()
    for i, group in enumerate(best_assignment):
        group_df = df[df[id_col].isin(group)].copy()
        group_df['Group'] = f'Group {i+1}'
        result_df = pd.concat([result_df, group_df], ignore_index=True)

    output_excel = "optimized_assignment_result.xlsx"
    result_df.to_excel(output_excel, index=False)
    return output_excel
