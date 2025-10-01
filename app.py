import streamlit as st
from optimizedassignment import run_optimization

st.title("最適化振り分けアプリ（完全版）")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file:
    if st.button("最適化を実行する"):
        output_path = run_optimization(uploaded_file)
        st.success("最適化が完了しました。以下からダウンロードできます。")
        with open(output_path, "rb") as f:
            st.download_button("結果ファイルをダウンロード", f, file_name=output_path)
