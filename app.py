import streamlit as st
import tempfile
from optimized_assignment import run_optimization

st.title("最適化振り分けアプリ（軽量版）")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx"])

if uploaded_file:
    if st.button("最適化を実行する"):
        # 一時ファイルとして保存
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.read())
            tmp_path = tmp.name

        # 一時ファイルのパスを渡して処理
        output_path = run_optimization(tmp_path)

        st.success("最適化が完了しました。以下からダウンロードできます。")
        with open(output_path, "rb") as f:
            st.download_button("結果ファイルをダウンロード", f, file_name=output_path)
