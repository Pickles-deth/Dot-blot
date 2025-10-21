import streamlit as st
import numpy as np
import itertools
import pandas as pd
import os
from io import BytesIO
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.utils.dataframe import dataframe_to_rows

# ----------------------------------------
# 🌟 ページ設定
# ----------------------------------------
st.set_page_config(
    page_title="Dot Blot 最適化ツール",
    page_icon="🧪",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ----------------------------------------
# 🖼️ ロゴ表示
# ----------------------------------------
st.image("logo.png", width=150)
st.markdown("""
<h1 style='text-align:center; color:#2c3e50;'>🧪 Dot Blot 最適化ツール</h1>
<p style='text-align:center; color:gray; font-size:18px;'>
行内シャッフルによる最小SD構成を自動で探索します<br>
0は自動的に除外されます
</p>
<hr style='border:1px solid #eee;'>
""", unsafe_allow_html=True)

# ----------------------------------------
# 📥 入力エリア
# ----------------------------------------
st.sidebar.header("⚙️ 設定")
num_rows = st.sidebar.number_input("行の数（例：4）", min_value=1, max_value=10, value=4)
num_cols = st.sidebar.number_input("列の数（例：4）", min_value=1, max_value=10, value=4)
st.sidebar.markdown("---")

st.sidebar.info("各行にカンマ区切りで数値を入力してください。")

data = {}
st.markdown("### ✏️ データ入力")
cols = st.columns(2)
for i in range(num_rows):
    key = chr(65 + i)
    with cols[i % 2]:
        values = st.text_input(f"🔹 {key} 行の値", "1.0, 1.0, 1.0, 0")
        try:
            data[key] = [float(v.strip()) for v in values.split(",")]
        except:
            st.warning(f"{key}行の入力を確認してください。")

# ----------------------------------------
# 🚀 実行ボタン
# ----------------------------------------
run = st.button("🚀 計算を実行する", use_container_width=True, type="primary")

if run:
    # ----------------------------------------
    # 🧠 内部処理
    # ----------------------------------------
    def nonzero_list(row):
        return [v for v in row if v != 0.0]

    def normalize_columns(columns):
        normalized = []
        for col in columns:
            base_val = col[0]
            if base_val == 0:
                normalized.append(tuple(0.0 for _ in col))
                continue
            normalized.append(tuple((v / base_val * 100) for v in col))
        return normalized

    def calc_sum_sd_from_columns(columns):
        norm_cols = normalize_columns(columns)
        rows = list(zip(*norm_cols))
        sds = []
        means = []
        for r in rows:
            arr = np.array([x for x in r if not np.isnan(x)])
            if arr.size == 0:
                means.append(0.0); sds.append(0.0)
            elif arr.size == 1:
                means.append(float(arr[0])); sds.append(0.0)
            else:
                means.append(float(np.mean(arr)))
                sds.append(float(np.std(arr, ddof=1)))
        return sum(sds), sds, means, norm_cols

    def canonicalize_columns(cols):
        return tuple(sorted(tuple(c) for c in cols))

    nonzero_data = {k: nonzero_list(v) for k, v in data.items()}
    counts = [len(v) for v in nonzero_data.values()]
    k = min(counts)

    st.info(f"✅ 使用サンプル数（k）={k} / 非ゼロ数: {counts}")

    row_perms_list = [list(itertools.permutations(v, k)) for v in nonzero_data.values()]
    total_combinations = np.prod([len(p) for p in row_perms_list])
    st.write(f"全組み合わせ数: **{int(total_combinations):,}** 通り")

    seen = {}
    results = []

    progress = st.progress(0)
    total = int(total_combinations)
    checked = 0

    for perm_set in itertools.product(*row_perms_list):
        checked += 1
        if checked % 500 == 0:
            progress.progress(min(checked / total, 1.0))

        cols = [tuple(perm[i] for perm in perm_set) for i in range(k)]
        cols_nozero = [col for col in cols if 0.0 not in col]
        if len(cols_nozero) != k:
            continue

        key = canonicalize_columns(cols_nozero)
        if key in seen:
            continue

        sum_sd, sds, means, norm_cols = calc_sum_sd_from_columns(cols_nozero)
        seen[key] = True
        results.append({
            'sum_sd': sum_sd, 'sds': sds, 'means': means, 'columns': cols_nozero
        })

    progress.empty()
    results.sort(key=lambda x: x['sum_sd'])

    st.success("🎉 計算が完了しました！")

    # ----------------------------------------
    # 📊 表示（詳細展開＋色付き）
    # ----------------------------------------
    labels = list(nonzero_data.keys())  # ['A', 'B', 'C', 'D']
    colors = ["#1f77b4", "#ff7f0e", "#2ca02c", "#d62728", "#9467bd",
              "#8c564b", "#e377c2", "#7f7f7f", "#bcbd22", "#17becf"]

    st.markdown("### 🏆 上位ランキング（Top 10）")
    topn = min(10, len(results))

    for i in range(topn):
        r = results[i]
        with st.expander(f"🏅 Rank {i+1} — Sum_SD: {r['sum_sd']:.4f}"):
            st.write(f"**SDs:** {', '.join(f'{v:.3f}' for v in r['sds'])}")
            st.write(f"**Means:** {', '.join(f'{v:.3f}' for v in r['means'])}")

            # --- 表形式で見やすく ---
            df_cols = pd.DataFrame(r["columns"], columns=labels)
            styled_df = df_cols.style.format(precision=3)
            for j, label in enumerate(labels):
                styled_df = styled_df.set_properties(subset=[label], **{
                    "color": "white",
                    "background-color": colors[j % len(colors)],
                    "font-weight": "bold",
                    "text-align": "center"
                })
            st.dataframe(styled_df, use_container_width=True)


    # ----------------------------------------
    # 📈 Excel出力（グラフ付き） + リストを文字列化
    # ----------------------------------------
# --- Excel 出力 ---
wb = Workbook()
ws = wb.active
ws.title = "Top Results"

df = pd.DataFrame([{
    "Rank": i + 1,
    "Sum_SD": round(r["sum_sd"], 6),
    "SDs": ", ".join([f"{v:.3f}" for v in r["sds"]]),
    "Means": ", ".join([f"{v:.3f}" for v in r["means"]]),
    "Columns": str(r["columns"])
} for i, r in enumerate(results[:topn])])

for row in dataframe_to_rows(df, index=False, header=True):
    ws.append(row)

chart = BarChart()
chart.title = "Sum_SD 比較"
chart.x_axis.title = "Rank"
chart.y_axis.title = "Sum_SD"
data_ref = Reference(ws, min_col=2, min_row=1, max_row=len(df)+1)
cats_ref = Reference(ws, min_col=1, min_row=2, max_row=len(df)+1)
chart.add_data(data_ref, titles_from_data=True)
chart.set_categories(cats_ref)
ws.add_chart(chart, "H3")

output = BytesIO()
wb.save(output)
output.seek(0)

st.download_button(
    "⬇️ Excelで結果をダウンロード",
    output,
    file_name="DotBlot_Result.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)

st.markdown("<h2 style='text-align:center; color:#ff66b2;'>✨あはは、できちゃったよ✨</h2>", unsafe_allow_html=True)
