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
Monitoring Penjualan Harian  
📍 Warung FIFGROUP Lantai 3 • Lantai 8 • Lantai 9
""")

st.divider()

# ==========================================
# FILE PATH
# ==========================================
file_path = "data/stock_harian.xlsx"

# ==========================================
# CLEAN RUPIAH FUNCTION
# ==========================================
def clean_rupiah(value):

    if pd.isna(value):
        return 0

    if isinstance(value, (int, float)):
        return int(value)

    value = str(value)

    # hapus semua selain angka
    value = re.sub(r"\D", "", value)

    if value == "":
        return 0

    return int(value)

# ==========================================
# READ EXCEL
# ==========================================
try:

    df = pd.read_excel(file_path)

except:

    st.error("❌ File stock_harian.xlsx tidak ditemukan")
    st.stop()

# ==========================================
# CLEAN COLUMN
# ==========================================
df.columns = df.columns.str.strip()

# ==========================================
# CLEAN DATA
# ==========================================
df["Harga Jual"] = (
    df["Harga Jual"]
    .apply(clean_rupiah)
)

df["Uang Seharusnya Dibayar"] = (
    df["Uang Seharusnya Dibayar"]
    .apply(clean_rupiah)
)

# ==========================================
# LANTAI AS CATEGORICAL
# ==========================================
df["Lantai"] = (
    "Lantai " + df["Lantai"].astype(str)
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

    produk_filter = st.multiselect(
        "Pilih Produk",
        options=sorted(df["Nama Produk"].unique()),
        default=sorted(df["Nama Produk"].unique())
    )

# ==========================================
# FILTER DATA
# ==========================================
df = df[
    (df["Lantai"].isin(lantai_filter)) &
    (df["Nama Produk"].isin(produk_filter))
]

# ==========================================
# KPI
# ==========================================
total_revenue = (
    df["Uang Seharusnya Dibayar"]
    .sum()
)

total_terjual = (
    df["Terjual"]
    .sum()
)

total_produk = (
    df["Nama Produk"]
    .nunique()
)

warung_terlaris = (
    df.groupby("Lantai")[
        "Uang Seharusnya Dibayar"
    ]
    .sum()
    .idxmax()
)

# ==========================================
# KPI CARDS
# ==========================================
st.subheader("📌 Ringkasan Penjualan")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "💰 Total Revenue",
    f"Rp {total_revenue:,.0f}"
)

col2.metric(
    "📦 Total Produk Terjual",
    f"{int(total_terjual)}"
)

col3.metric(
    "🛍️ Total Jenis Produk",
    total_produk
)

col4.metric(
    "🏆 Warung Terlaris",
    warung_terlaris
)

st.divider()

# ==========================================
# CHART SECTION
# ==========================================
col_chart1, col_chart2 = st.columns(2)

# ==========================================
# REVENUE PER LANTAI
# ==========================================
with col_chart1:

    st.subheader("📈 Revenue per Lantai")

    lantai_chart = (
        df.groupby("Lantai")[
            "Uang Seharusnya Dibayar"
        ]
        .sum()
        .reset_index()
    )

    fig1 = px.bar(
        lantai_chart,
        x="Lantai",
        y="Uang Seharusnya Dibayar",
        text_auto=True,
        color="Lantai"
    )

    fig1.update_layout(
        height=450,
        showlegend=False,
        xaxis_title="Lantai",
        yaxis_title="Revenue"
    )

    st.plotly_chart(
        fig1,
        use_container_width=True
    )

# ==========================================
# TOP PRODUK
# ==========================================
with col_chart2:

    st.subheader("🔥 Top 10 Produk Terlaris")

    top_produk = (
        df.groupby("Nama Produk")[
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
        height=450,
        xaxis_title="Produk",
        yaxis_title="Jumlah Terjual"
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

st.divider()

# ==========================================
# REVENUE PER PRODUK
# ==========================================
st.subheader("💸 Revenue per Produk")

revenue_produk = (
    df.groupby("Nama Produk")[
        "Uang Seharusnya Dibayar"
    ]
    .sum()
    .reset_index()
    .sort_values(
        by="Uang Seharusnya Dibayar",
        ascending=False
    )
)

fig3 = px.bar(
    revenue_produk,
    x="Nama Produk",
    y="Uang Seharusnya Dibayar",
    text_auto=True,
    color="Uang Seharusnya Dibayar"
)

fig3.update_layout(
    height=500,
    xaxis_title="Produk",
    yaxis_title="Revenue"
)

st.plotly_chart(
    fig3,
    use_container_width=True
)

st.divider()

# ==========================================
# DETAIL TABLE
# ==========================================
st.subheader("📋 Detail Penjualan")

st.dataframe(
    df,
    use_container_width=True,
    hide_index=True
)

# ==========================================
# DOWNLOAD CSV
# ==========================================
csv = df.to_csv(index=False).encode("utf-8")

st.download_button(
    label="⬇️ Download CSV",
    data=csv,
    file_name="dashboard_warung_fifgroup.csv",
    mime="text/csv"
)
