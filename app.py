"""Run with: python -m streamlit run app.py"""
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from reporting import load_data, prepare_data, aggregate, comparison, summary_text, create_pdf

ROOT = Path(__file__).resolve().parent
st.set_page_config(page_title="AI月次報告書ツール", page_icon="📊", layout="wide")
st.title("AI月次報告書ツール")
st.caption("売上データから、月次の数字と報告書を。Excel・CSVデータはアプリが動作するサーバー上で処理します。外部AI APIは使用しません。")
st.warning("試験公開版では架空データのみ使用してください。実際の顧客データ・個人情報・機密情報はアップロードしないでください。")

with st.sidebar:
    st.header("1. データを選ぶ")
    mode = st.radio("データの種類", ["サンプルデータで試す", "ファイルを読み込む"])
    st.download_button("サンプルCSVを保存", (ROOT / "sample_data/sales.csv").read_bytes(),
                       "sample_sales.csv", "text/csv")
    st.caption("Excel (.xlsx) / CSV ・ 最大20MB / 10万行")

try:
    if mode == "サンプルデータで試す":
        content = (ROOT / "sample_data/sales.csv").read_bytes()
        filename = "sales.csv"
        st.info("架空のサンプルデータを表示しています。実際の売上ではありません。")
    else:
        uploaded = st.sidebar.file_uploader("ExcelまたはCSVを選択", type=["xlsx", "csv"])
        if uploaded is None:
            st.info("左側の「Browse files」からファイルを選択してください。")
            st.stop()
        content, filename = uploaded.getvalue(), uploaded.name
    sheet = 0
    if filename.lower().endswith(".xlsx"):
        with pd.ExcelFile(BytesIO(content), engine="openpyxl") as book:
            sheet = st.sidebar.selectbox("読み込むシート", book.sheet_names)
    frame = load_data(content, filename, sheet)
    st.sidebar.header("2. 列を確認する")
    columns = frame.columns.tolist()
    date_default = next((i for i, c in enumerate(columns) if c in ["日付", "売上日", "date"]), 0)
    amount_default = next((i for i, c in enumerate(columns) if c in ["売上金額", "売上", "金額", "amount"]), min(1, len(columns)-1))
    date_col = st.sidebar.selectbox("日付の列", columns, index=date_default)
    amount_col = st.sidebar.selectbox("売上金額の列", columns, index=amount_default)
    with st.expander("読み込んだデータを確認"):
        st.dataframe(frame.head(100), hide_index=True, width="stretch")
        st.caption(f"全{len(frame):,}行のうち先頭100行まで表示。1行を1件として集計します。")
    data = prepare_data(frame, date_col, amount_col)
    monthly = aggregate(data)
except Exception as error:
    st.error(f"データを読み込めませんでした：{error}")
    st.info("見出しのあるCSV、またはパスワードのないExcelを選択し、日付列と金額列を確認してください。")
    st.stop()

st.subheader("月次サマリー")
months = monthly.dropna().index.tolist()
month = st.selectbox("報告書の対象月", months, index=len(months)-1, format_func=lambda p: f"{p.year}年{p.month}月")
row = monthly.loc[month]
rate, change = comparison(monthly, month)
c1, c2, c3 = st.columns(3)
c1.metric("売上金額", f"¥{int(row['売上金額']):,}")
c2.metric("売上件数", f"{int(row['件数']):,} 件")
c3.metric("前月比（売上）", change if rate is not None else "比較不可")
if rate is None:
    st.caption(change)
st.subheader("月別売上の推移")
chart_frame = monthly[["売上金額"]].copy()
chart_frame.index = chart_frame.index.astype(str)
st.bar_chart(chart_frame, color="#167d8d", width="stretch")
st.caption("単位：円。データのない月は未集計です。返品などの負数も合算します。")
with st.expander("月別の集計表"):
    table = monthly.copy()
    table["前月比"] = [comparison(monthly, p)[1] if pd.notna(monthly.loc[p, "売上金額"]) else "データなし" for p in monthly.index]
    table.index = table.index.astype(str)
    st.dataframe(table, width="stretch")
st.subheader("報告書をダウンロード")
st.write(summary_text(monthly, month))
st.caption("文章は集計値から自動作成する定型文です。PDFには対象月までの最大12か月を掲載します。")
try:
    pdf = create_pdf(monthly, month)
    # Block a stale/unembedded generator from serving a misleading download.
    if b"/FontFile2" not in pdf or b"IPAexGothic" not in pdf:
        raise ValueError("日本語フォントの埋め込みを確認できません。アプリを停止して再起動してください。")
    st.download_button("月次報告書をPDFで保存", pdf, f"monthly_report_{month}_jp.pdf", "application/pdf", type="primary")
    st.caption("日本語フォント埋め込み済みPDF（ファイル名末尾：_jp.pdf）")
except Exception as error:
    st.error(f"PDFを作成できませんでした：{error}")
