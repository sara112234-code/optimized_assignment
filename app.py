import streamlit as st
import pandas as pd
from optimized_assignment import run_optimization

st.title("振り分け最適化アプリ")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください（Book4.xlsx形式）", type=["xlsx"])

if uploaded_file:
    st.success("ファイルがアップロードされました。")
    df_preview = pd.read_excel(uploaded_file, sheet_name="Sheet1", engine="openpyxl")
    st.subheader("Sheet1のプレビュー")
    st.dataframe(df_preview.head())

    if st.button("最適化を実行"):
        with st.spinner("最適化中..."):
            output_path = run_optimization(uploaded_file)
        st.success("最適化が完了しました！")
        with open(output_path, "rb") as f:
            st.download_button("結果をダウンロード", f, file_name="Book4_optimized.xlsx")
