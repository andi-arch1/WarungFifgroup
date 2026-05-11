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
qris_file = "data/qris_all.xlsx"

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

except Exception as e:
    st.error(f"❌ Error membaca stock file: {e}")
    st.stop()

# ==========================================
# CLEAN STOCK COLUMNS
# ==========================================
stock_df.columns = (
    stock_df.columns
    .str.strip()
    .str.replace("\ufeff", "")
)

# ==========================================
# VALIDATE STOCK COLUMNS
# ==========================================
required_stock_cols = [
    "Tanggal",
    "Lantai",
    "Nama Produk",
    "Harga Jual",
    "Terjual",
    "Uang Seharusnya Dibayar"
]

missing_stock_cols = [
    col for col in required_stock_cols
    if col not in stock_df.columns
]

if missing_stock_cols:
    st.error(f"❌ Kolom stock tidak ditemukan: {missing_stock_cols}")
    st.stop()

# ==========================================
# CLEAN STOCK DATA
# ==========================================
stock_df["Harga Jual"] = stock_df["Harga Jual"].apply(clean_rupiah)
stock_df["Uang Seharusnya Dibayar"] = stock_df["Uang Seharusnya Dibayar"].apply(clean_rupiah)
stock_df["Terjual"] = pd.to_numeric(stock_df["Terjual"], errors="coerce").fillna(0)

# Format Lantai
stock_df["Lantai"] = stock_df["Lantai"].astype(str).str.strip()

stock_df["Lantai"] = stock_df["Lantai"].replace({
    "3": "Lantai 3",
    "8": "Lantai 8",
    "9": "Lantai 9",
    "Lantai 3": "Lantai 3",
    "Lantai 8": "Lantai 8",
    "Lantai 9": "Lantai 9"
})

# Format Tanggal
stock_df["Tanggal"] = pd.to_datetime(
    stock_df["Tanggal"],
    errors="coerce"
).dt.date

# ==========================================
# LOAD QRIS EXCEL
# ==========================================
try:
    qris_df = pd.read_excel(qris_file)

except Exception as e:
    st.error(f"❌ Error membaca Excel QRIS: {e}")
    st.stop()

# ==========================================
# CLEAN QRIS COLUMNS
# ==========================================
qris_df.columns = (
    qris_df.columns
    .str.strip()
    .str.replace("\ufeff", "")
)

# ==========================================
# VALIDATE QRIS COLUMNS
# ==========================================
required_qris_cols = [
    "Nama Merchant",
    "Tanggal Transaksi",
    "Total Terbayar"
]

missing_qris_cols = [
    col for col in required_qris_cols
    if col not in qris_df.columns
]

if missing_qris_cols:
    st.error(f"❌ Kolom QRIS tidak ditemukan: {missing_qris_cols}")
    st.stop()

# ==========================================
# CLEAN QRIS DATA
# ==========================================
qris_df["Tanggal"] = pd.to_datetime(
    qris_df["Tanggal Transaksi"],
    errors="coerce"
).dt.date

qris_money_cols = [
    "Nominal Transaksi",
    "Tip",
    "Harga",
    "Total Terbayar",
    "MDR"
]

for col in qris_money_cols:
    if col in qris_df.columns:
        qris_df[col] = qris_df[col].apply(clean_rupiah)

# ==========================================
# MERCHANT MAPPING
# ==========================================
merchant_mapping = {
    "Warung FIFGROUP 1": "Lantai 3",
    "Warung FIFGROUP 2": "Lantai 8",
    "Warung FIFGROUP 3": "Lantai 9"
}

qris_df["Nama Merchant"] = qris_df["Nama Merchant"].astype(str).str.strip()

qris_df["Lantai"] = qris_df["Nama Merchant"].map(merchant_mapping)

# Kalau ada merchant yang tidak kebaca
unknown_merchants = qris_df[qris_df["Lantai"].isna()]["Nama Merchant"].dropna().unique()

if len(unknown_merchants) > 0:
    st.warning(f"⚠️ Ada Nama Merchant yang belum dimapping: {list(unknown_merchants)}")

# ==========================================
# FILTER SECTION
# ==========================================
st.subheader("🔍 Filter Dashboard")

col_filter1, col_filter2 = st.columns(2)

available_lantai = [
    "Lantai 3",
    "Lantai 8",
    "Lantai 9"
]

available_tanggal = sorted(
    stock_df["Tanggal"]
    .dropna()
    .unique()
)

with col_filter1:
    lantai_filter = st.multiselect(
        "Pilih Lantai",
        options=available_lantai,
        default=available_lantai
    )

with col_filter2:
    tanggal_filter = st.multiselect(
        "Pilih Tanggal",
        options=available_tanggal,
        default=available_tanggal
    )

# ==========================================
# FILTER DATA
# ==========================================
filtered_stock_df = stock_df[
    (stock_df["Lantai"].isin(lantai_filter)) &
    (stock_df["Tanggal"].isin(tanggal_filter))
].copy()

filtered_qris_df = qris_df[
    (qris_df["Lantai"].isin(lantai_filter)) &
    (qris_df["Tanggal"].isin(tanggal_filter))
].copy()

# ==========================================
# KPI
# ==========================================
total_expected = filtered_stock_df["Uang Seharusnya Dibayar"].sum()
total_qris = filtered_qris_df["Total Terbayar"].sum()
total_selisih = total_qris - total_expected
total_terjual = filtered_stock_df["Terjual"].sum()

# ==========================================
# KPI CARDS
# ==========================================
st.subheader("📌 Ringkasan Rekonsiliasi")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "💰 Expected Revenue",
    f"Rp {total_expected:,.0f}".replace(",", ".")
)

col2.metric(
    "💳 Actual QRIS",
    f"Rp {total_qris:,.0f}".replace(",", ".")
)

col3.metric(
    "📦 Produk Terjual",
    f"{int(total_terjual)}"
)

col4.metric(
    "⚠️ Selisih",
    f"Rp {total_selisih:,.0f}".replace(",", ".")
)

st.divider()

# ==========================================
# REKONSILIASI
# ==========================================
expected_df = (
    filtered_stock_df
    .groupby(["Tanggal", "Lantai"])["Uang Seharusnya Dibayar"]
    .sum()
    .reset_index()
)

actual_df = (
    filtered_qris_df
    .groupby(["Tanggal", "Lantai"])["Total Terbayar"]
    .sum()
    .reset_index()
)

rekon_df = expected_df.merge(
    actual_df,
    on=["Tanggal", "Lantai"],
    how="left"
)

rekon_df["Total Terbayar"] = rekon_df["Total Terbayar"].fillna(0)

rekon_df["Selisih"] = (
    rekon_df["Total Terbayar"] -
    rekon_df["Uang Seharusnya Dibayar"]
)

# ==========================================
# STATUS
# ==========================================
def get_status(selisih):
    if selisih == 0:
        return "🟢 MATCH"
    elif selisih < 0:
        return "🔴 KURANG"
    else:
        return "🟡 LEBIH"

rekon_df["Status"] = rekon_df["Selisih"].apply(get_status)

rekon_df = rekon_df.sort_values(
    by=["Tanggal", "Lantai"],
    ascending=[True, True]
)

# ==========================================
# CHART SECTION
# ==========================================
col_chart1, col_chart2 = st.columns(2)

# ==========================================
# EXPECTED VS ACTUAL
# ==========================================
with col_chart1:
    st.subheader("📈 Expected vs Actual QRIS")

    chart_df = (
        rekon_df
        .groupby("Lantai")[
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
        barmode="group",
        labels={
            "value": "Nominal",
            "variable": "Kategori"
        }
    )

    fig1.update_layout(
        height=450,
        yaxis_tickprefix="Rp "
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
        filtered_stock_df
        .groupby("Nama Produk")["Terjual"]
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
        xaxis_tickangle=-30
    )

    st.plotly_chart(
        fig2,
        use_container_width=True
    )

st.divider()

# ==========================================
# DATA SELISIH
# ==========================================
st.subheader("🚨 Data Selisih Tidak Match")

selisih_df = rekon_df[rekon_df["Selisih"] != 0].copy()

if selisih_df.empty:
    st.success("✅ Semua data sudah MATCH.")
else:
    st.dataframe(
        selisih_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Uang Seharusnya Dibayar": st.column_config.NumberColumn(
                "Uang Seharusnya Dibayar",
                format="Rp %d"
            ),
            "Total Terbayar": st.column_config.NumberColumn(
                "Total Terbayar",
                format="Rp %d"
            ),
            "Selisih": st.column_config.NumberColumn(
                "Selisih",
                format="Rp %d"
            ),
        }
    )

st.divider()

# ==========================================
# REKONSILIASI TABLE
# ==========================================
st.subheader("📋 Tabel Rekonsiliasi")

st.dataframe(
    rekon_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Uang Seharusnya Dibayar": st.column_config.NumberColumn(
            "Uang Seharusnya Dibayar",
            format="Rp %d"
        ),
        "Total Terbayar": st.column_config.NumberColumn(
            "Total Terbayar",
            format="Rp %d"
        ),
        "Selisih": st.column_config.NumberColumn(
            "Selisih",
            format="Rp %d"
        ),
    }
)

st.divider()

# ==========================================
# DETAIL STOCK
# ==========================================
st.subheader("📦 Detail Penjualan Stock")

st.dataframe(
    filtered_stock_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Harga Jual": st.column_config.NumberColumn(
            "Harga Jual",
            format="Rp %d"
        ),
        "Uang Seharusnya Dibayar": st.column_config.NumberColumn(
            "Uang Seharusnya Dibayar",
            format="Rp %d"
        ),
    }
)

st.divider()

# ==========================================
# DETAIL QRIS
# ==========================================
st.subheader("💳 Detail Transaksi QRIS")

qris_column_config = {}

for col in qris_money_cols:
    if col in filtered_qris_df.columns:
        qris_column_config[col] = st.column_config.NumberColumn(
            col,
            format="Rp %d"
        )

st.dataframe(
    filtered_qris_df,
    use_container_width=True,
    hide_index=True,
    column_config=qris_column_config
)

st.divider()

# ==========================================
# DOWNLOAD CSV
# ==========================================
csv = rekon_df.to_csv(index=False).encode("utf-8-sig")

st.download_button(
    label="⬇️ Download Rekonsiliasi CSV",
    data=csv,
    file_name="rekonsiliasi_warung_fifgroup.csv",
    mime="text/csv"
)

# ==========================================
# DOWNLOAD EXCEL
# ==========================================
output_excel = pd.ExcelWriter(
    "rekonsiliasi_warung_fifgroup.xlsx",
    engine="openpyxl"
)

rekon_df.to_excel(
    output_excel,
    sheet_name="Rekonsiliasi",
    index=False
)

selisih_df.to_excel(
    output_excel,
    sheet_name="Data Selisih",
    index=False
)

filtered_stock_df.to_excel(
    output_excel,
    sheet_name="Detail Stock",
    index=False
)

filtered_qris_df.to_excel(
    output_excel,
    sheet_name="Detail QRIS",
    index=False
)

output_excel.close()

with open("rekonsiliasi_warung_fifgroup.xlsx", "rb") as file:
    st.download_button(
        label="⬇️ Download Rekonsiliasi Excel",
        data=file,
        file_name="rekonsiliasi_warung_fifgroup.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
