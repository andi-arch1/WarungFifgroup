import streamlit as st
import pandas as pd
import plotly.express as px
import re

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(
    page_title="Dashboard Warung FIFGROUP",
    page_icon="🛒",
    layout="wide"
)

# ==========================================
# CUSTOM CSS
# ==========================================
st.markdown("""
<style>

.main {
    background-color: #f7f9fc;
}

.block-container {
    padding-top: 2rem;
    padding-bottom: 2rem;
}

h1, h2, h3 {
    color: #1f2937;
}

[data-testid="metric-container"] {
    background-color: white;
    border: 1px solid #e5e7eb;
    padding: 18px;
    border-radius: 18px;
    box-shadow: 0px 2px 10px rgba(0,0,0,0.05);
}

</style>
""", unsafe_allow_html=True)

# ==========================================
# HEADER
# ==========================================
st.title("🛒 Dashboard Warung FIFGROUP")

st.markdown("""
Monitoring Penjualan Harian & Rekonsiliasi QRIS  
📍 Warung FIFGROUP Lantai 3 • Lantai 8 • Lantai 9
""")

st.divider()

# ==========================================
# FILE PATH
# ==========================================
stock_file = "data/stock_harian.xlsx"

qris_lt3_file = "data/qris_lt3.csv"
qris_lt8_file = "data/qris_lt8.csv"
qris_lt9_file = "data/qris_lt9.csv"

# ==========================================
# CLEAN RUPIAH FUNCTION
# ==========================================
def clean_rupiah(value):

    if pd.isna(value):
        return 0

    if isinstance(value, (int, float)):
        return int(value)

    value = str(value)

    value = re.sub(r"\D", "", value)

    if value == "":
        return 0

    return int(value)

# ==========================================
# LOAD STOCK DATA
# ==========================================
try:

    stock_df = pd.read_excel(stock_file)

except:

    st.error("❌ File stock_harian.xlsx tidak ditemukan")
    st.stop()

# ==========================================
# CLEAN STOCK DATA
# ==========================================
stock_df.columns = stock_df.columns.str.strip()

stock_df["Harga Jual"] = (
    stock_df["Harga Jual"]
    .apply(clean_rupiah)
)

stock_df["Uang Seharusnya Dibayar"] = (
    stock_df["Uang Seharusnya Dibayar"]
    .apply(clean_rupiah)
)

# ==========================================
# FORMAT LANTAI
# ==========================================
stock_df["Lantai"] = (
    "Lantai " + stock_df["Lantai"].astype(str)
)

# ==========================================
# FORMAT TANGGAL STOCK
# ==========================================
stock_df["Tanggal"] = pd.to_datetime(
    stock_df["Tanggal"]
).dt.date

# ==========================================
# LOAD QRIS CSV
# ==========================================
try:

    lt3 = pd.read_csv(qris_lt3_file)
    lt8 = pd.read_csv(qris_lt8_file)
    lt9 = pd.read_csv(qris_lt9_file)

except:

    st.error("❌ File CSV QRIS tidak ditemukan")
    st.stop()

# ==========================================
# MERCHANT MAPPING
# ==========================================
merchant_mapping = {
    "Warung FIFGROUP 1": "Lantai 3",
    "Warung FIFGROUP 2": "Lantai 8",
    "Warung FIFGROUP 3": "Lantai 9"
}

# ==========================================
# CONCAT QRIS
# ==========================================
qris_df = pd.concat(
    [lt3, lt8, lt9],
    ignore_index=True
)

# ==========================================
# CLEAN QRIS COLUMN
# ==========================================
qris_df.columns = qris_df.columns.str.strip()

# ==========================================
# FILTER STATUS BERHASIL
# ==========================================
qris_df = qris_df[
    qris_df["Status"] == "BERHASIL"
]

# ==========================================
# FORMAT TANGGAL QRIS
# ==========================================
qris_df["Tanggal"] = pd.to_datetime(
    qris_df["Tanggal Transaksi"],
    dayfirst=True
).dt.date

# ==========================================
# FORMAT NOMINAL
# ==========================================
qris_df["Total Terbayar"] = (
    qris_df["Total Terbayar"]
    .apply(clean_rupiah)
)

# ==========================================
# MAPPING LANTAI
# ==========================================
qris_df["Lantai"] = (
    qris_df["Nama Merchant"]
    .map(merchant_mapping)
)

# ==========================================
# FILTER SECTION
# ==========================================
st.subheader("🔍 Filter Dashboard")

col_filter1, col_filter2 = st.columns(2)

with col_filter1:

    lantai_filter = st.multiselect(
        "Pilih Lantai",
        options=[
            "Lantai 3",
            "Lantai 8",
            "Lantai 9"
        ],
        default=[
            "Lantai 3",
            "Lantai 8",
            "Lantai 9"
        ]
    )

with col_filter2:

    tanggal_filter = st.multiselect(
        "Pilih Tanggal",
        options=sorted(stock_df["Tanggal"].unique()),
        default=sorted(stock_df["Tanggal"].unique())
    )

# ==========================================
# FILTER DATA
# ==========================================
stock_df = stock_df[
    (stock_df["Lantai"].isin(lantai_filter)) &
    (stock_df["Tanggal"].isin(tanggal_filter))
]

qris_df = qris_df[
    (qris_df["Lantai"].isin(lantai_filter)) &
    (qris_df["Tanggal"].isin(tanggal_filter))
]

# ==========================================
# KPI
# ==========================================
total_expected = (
    stock_df["Uang Seharusnya Dibayar"]
    .sum()
)

total_qris = (
    qris_df["Total Terbayar"]
    .sum()
)

total_selisih = (
    total_qris - total_expected
)

total_terjual = (
    stock_df["Terjual"]
    .sum()
)

# ==========================================
# KPI CARDS
# ==========================================
st.subheader("📌 Ringkasan Rekonsiliasi")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "💰 Expected Revenue",
    f"Rp {total_expected:,.0f}"
)

col2.metric(
    "💳 Actual QRIS",
    f"Rp {total_qris:,.0f}"
)

col3.metric(
    "📦 Produk Terjual",
    f"{int(total_terjual)}"
)

col4.metric(
    "⚠️ Selisih",
    f"Rp {total_selisih:,.0f}"
)

st.divider()

# ==========================================
# REKONSILIASI TABLE
# ==========================================
expected_df = (
    stock_df.groupby(["Tanggal", "Lantai"])[
        "Uang Seharusnya Dibayar"
    ]
    .sum()
    .reset_index()
)

actual_df = (
    qris_df.groupby(["Tanggal", "Lantai"])[
        "Total Terbayar"
    ]
    .sum()
    .reset_index()
)

rekon_df = expected_df.merge(
    actual_df,
    on=["Tanggal", "Lantai"],
    how="left"
)

rekon_df["Total Terbayar"] = (
    rekon_df["Total Terbayar"]
    .fillna(0)
)

rekon_df["Selisih"] = (
    rekon_df["Total Terbayar"] -
    rekon_df["Uang Seharusnya Dibayar"]
)

# ==========================================
# STATUS FLAG
# ==========================================
def get_status(selisih):

    if selisih == 0:
        return "🟢 MATCH"

    elif selisih < 0:
        return "🔴 KURANG"

    else:
        return "🟡 LEBIH"

rekon_df["Status"] = (
    rekon_df["Selisih"]
    .apply(get_status)
)

# ==========================================
# CHART SECTION
# ==========================================
col_chart1, col_chart2 = st.columns(2)

# ==========================================
# EXPECTED VS QRIS
# ==========================================
with col_chart1:

    st.subheader("📈 Expected vs Actual QRIS")

    chart_df = (
        rekon_df.groupby("Lantai")[
            [
                "Uang Seharusnya Dibayar",
                "Total Terbayar"
            ]
        ]
        .sum()
        .reset_index()
    )

    fig1 = px.bar(
        chart_df,
        x="Lantai",
        y=[
            "Uang Seharusnya Dibayar",
            "Total Terbayar"
        ],
        barmode="group"
    )

    fig1.update_layout(
        height=450
    )

    st.plotly_chart(
        fig1,
        use_container_width=True
    )

# ==========================================
# TOP PRODUK
# ==========================================
with col_chart2:

    st.subheader("🔥 Top Produk Terlaris")

    top_produk = (
        stock_df.groupby("Nama Produk")[
            "Terjual"
        ]
        .sum()
        .reset_index()
        .sort_values(
            by="Terjual",
            ascending=False
        )
        .head(10)
    )

    fig2 = px.bar(
        top_produk,
        x="Nama Produk",
        y="Terjual",
        text_auto=True,
        color="Terjual"
    )

    fig2.update_layout(
        height=450
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

st.divider()

# ==========================================
# REKONSILIASI TABLE
# ==========================================
st.subheader("📋 Tabel Rekonsiliasi")

st.dataframe(
    rekon_df,
    use_container_width=True,
    hide_index=True
)

st.divider()

# ==========================================
# DETAIL STOCK
# ==========================================
st.subheader("📦 Detail Penjualan Stock")

st.dataframe(
    stock_df,
    use_container_width=True,
    hide_index=True
)

# ==========================================
# DETAIL QRIS
# ==========================================
st.subheader("💳 Detail Transaksi QRIS")

st.dataframe(
    qris_df,
    use_container_width=True,
    hide_index=True
)

# ==========================================
# DOWNLOAD CSV
# ==========================================
csv = rekon_df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="⬇️ Download Rekonsiliasi CSV",
    data=csv,
    file_name="rekonsiliasi_warung_fifgroup.csv",
    mime="text/csv"
)
