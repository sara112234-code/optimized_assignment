import streamlit as st
from optimized_assignment import run_optimization

st.title("最適化振り分けアプリ")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file is not None:
    if st.button("最適化を実行する"):
        output_path = run_optimization(uploaded_file)
        st.success("最適化が完了しました。結果ファイルを以下からダウンロードできます。")
        with open(output_path, "rb") as f:
            st.download_button("結果をダウンロード", f, file_name=output_path)
