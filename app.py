import streamlit as st
from optimized_assignment import run_optimization

st.title("最適化振り分けアプリ")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    if st.button("最適化を実行する"):
        run_optimization(uploaded_file)
        st.success("最適化が完了しました。結果は optimized_assignment.xlsx に保存されました。")
        with open("optimized_assignment.xlsx", "rb") as f:
            st.download_button("結果ファイルをダウンロード", f, file_name="optimized_assignment.xlsx")
