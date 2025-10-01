import pandas as pd
import numpy as np
from io import BytesIO

def run_optimization(uploaded_file):
    # Load Excel file
    xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
    df1 = pd.read_excel(xls, sheet_name="Sheet1", engine="openpyxl")
    df2 = pd.read_excel(xls, sheet_name="Sheet2", engine="openpyxl")
    df3 = pd.read_excel(xls, sheet_name="Sheet3", engine="openpyxl")

    # 固定メンバー（13項目）を抽出
    fixed_ids = df3.iloc[:, 0].dropna().astype(str).tolist()
    fixed_members = df1[df1["No."].astype(str).isin(fixed_ids)]

    # 残りのメンバーを抽出
    remaining = df1[~df1["No."].astype(str).isin(fixed_ids)].copy()

    # m/fごとの人数制限（パターン1）
    m_total = 15 + 15 + 14
    f_total = 20 + 21 + 21

    # m/fの性別ごとに分割
    males = remaining[remaining["gender"] == "m"].copy()
    females = remaining[remaining["gender"] == "f"].copy()

    # ランダムに抽出（試行錯誤）
    best_score = float("inf")
    best_combination = None
    np.random.seed(42)
    for _ in range(1000):
        try:
            m_sample = males.sample(n=m_total)
            f_sample = females.sample(n=f_total)
            combined = pd.concat([fixed_members, m_sample, f_sample])
            combined = combined.sample(frac=1).reset_index(drop=True)

            # 3列に分割
            group_size = len(combined) // 3
            g1 = combined.iloc[:group_size]
            g2 = combined.iloc[group_size:group_size*2]
            g3 = combined.iloc[group_size*2:]

            avg1 = g1["average score"].mean()
            avg2 = g2["average score"].mean()
            avg3 = g3["average score"].mean()
            score_diff = max(avg1, avg2, avg3) - min(avg1, avg2, avg3)

            if score_diff < best_score:
                best_score = score_diff
                best_combination = (g1, g2, g3)
        except:
            continue

    # 結果をSheet2に書き込み
    g1, g2, g3 = best_combination
    df_out = pd.DataFrame()
    df_out = pd.concat([g1.reset_index(drop=True), g2.reset_index(drop=True), g3.reset_index(drop=True)], axis=1)

    # Save to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_out.to_excel(writer, sheet_name="Sheet2", index=False)
    output.seek(0)
    with open("Book4_optimized.xlsx", "wb") as f:
        f.write(output.read())
    return "Book4_optimized.xlsx"
