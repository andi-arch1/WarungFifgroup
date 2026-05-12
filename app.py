import streamlit as st
import pandas as pd
import plotly.express as px
import re
from io import BytesIO

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
Monitoring Penjualan Harian, Rekonsiliasi QRIS, Kepatuhan Stock, dan Standar Display  
📍 Warung FIFGROUP Lantai 3 • Lantai 8 • Lantai 9
""")

st.divider()

# ==========================================
# FILE PATH
# ==========================================
master_file = "data/master_data.xlsx"
stock_file = "data/stock_harian.xlsx"
qris_file = "data/qris_all.xlsx"

# ==========================================
# HELPER FUNCTIONS
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


def format_lantai(value):
    if pd.isna(value):
        return None

    value = str(value).strip().lower()

    mapping = {
        "3": "Lantai 3",
        "8": "Lantai 8",
        "9": "Lantai 9",
        "lantai 3": "Lantai 3",
        "lantai 8": "Lantai 8",
        "lantai 9": "Lantai 9",
        "lt 3": "Lantai 3",
        "lt 8": "Lantai 8",
        "lt 9": "Lantai 9",
        "lt3": "Lantai 3",
        "lt8": "Lantai 8",
        "lt9": "Lantai 9",
        "lantai3": "Lantai 3",
        "lantai8": "Lantai 8",
        "lantai9": "Lantai 9",
    }

    return mapping.get(value, str(value).title())


def lantai_from_sheet_name(sheet_name):
    sheet_name_clean = str(sheet_name).strip().lower()

    if "3" in sheet_name_clean:
        return "Lantai 3"

    if "8" in sheet_name_clean:
        return "Lantai 8"

    if "9" in sheet_name_clean:
        return "Lantai 9"

    return sheet_name


def get_display_status(row):
    jumlah = row["Jumlah"]
    display = row["Display"]

    if pd.isna(display):
        return "🔴 Display Kosong"

    if display < 0:
        return "🔴 Display Negatif"

    if display == jumlah:
        return "🟢 Sesuai Standar"

    if display < jumlah:
        return "🟡 Perlu Restock"

    return "🔵 Lebih dari Standar"


def get_rekon_status(selisih):
    if selisih == 0:
        return "🟢 MATCH"
    elif selisih < 0:
        return "🔴 KURANG"
    else:
        return "🟡 LEBIH"


def get_stock_status(selisih):
    if selisih == 0:
        return "🟢 Patuh"
    else:
        return "🔴 Tidak Patuh"


def rupiah(value):
    return f"Rp {value:,.0f}".replace(",", ".")


# ==========================================
# LOAD MASTER DATA - MULTIPLE SHEETS
# ==========================================
try:
    master_sheets = pd.read_excel(
        master_file,
        sheet_name=None
    )

except Exception as e:
    st.error(f"❌ Error membaca master data: {e}")
    st.stop()

master_list = []

for sheet_name, df in master_sheets.items():

    if df.empty:
        continue

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\ufeff", "")
    )

    rename_cols = {}

    # Produk -> Nama Produk
    if "Produk" in df.columns and "Nama Produk" not in df.columns:
        rename_cols["Produk"] = "Nama Produk"

    # Product -> Nama Produk
    for col in df.columns:
        if col.lower() == "product" and "Nama Produk" not in df.columns:
            rename_cols[col] = "Nama Produk"

    # Standar Display -> Jumlah
    if "Standar Display" in df.columns and "Jumlah" not in df.columns:
        rename_cols["Standar Display"] = "Jumlah"

    # Qty -> Jumlah
    if "Qty" in df.columns and "Jumlah" not in df.columns:
        rename_cols["Qty"] = "Jumlah"

    # Stock Display -> Display
    if "Stock Display" in df.columns and "Display" not in df.columns:
        rename_cols["Stock Display"] = "Display"

    df = df.rename(columns=rename_cols)

    # Kalau tidak ada kolom Lantai, ambil dari nama sheet
    if "Lantai" not in df.columns and "lantai" not in df.columns:
        df["Lantai"] = lantai_from_sheet_name(sheet_name)

    if "lantai" in df.columns and "Lantai" not in df.columns:
        df = df.rename(columns={"lantai": "Lantai"})

    df["Source Sheet"] = sheet_name

    master_list.append(df)

if len(master_list) == 0:
    st.error("❌ master_data.xlsx kosong atau tidak ada data yang terbaca.")
    st.stop()

master_df = pd.concat(
    master_list,
    ignore_index=True
)

# ==========================================
# VALIDATE MASTER COLUMNS
# ==========================================
required_master_cols = [
    "Tanggal",
    "Nama Produk",
    "Jumlah",
    "Actual"
]

missing_master_cols = [
    col for col in required_master_cols
    if col not in master_df.columns
]

if missing_master_cols:
    st.error(f"❌ Kolom master data tidak ditemukan: {missing_master_cols}")
    st.info("Pastikan tiap sheet master punya kolom minimal: Tanggal, Produk/Nama Produk, Jumlah, Actual.")
    st.stop()

# ==========================================
# CLEAN MASTER DATA
# ==========================================
master_df["Tanggal"] = pd.to_datetime(
    master_df["Tanggal"],
    errors="coerce",
    dayfirst=True
).dt.date

master_df["Nama Produk"] = (
    master_df["Nama Produk"]
    .astype(str)
    .str.strip()
)

master_df["Lantai"] = master_df["Lantai"].apply(format_lantai)

master_df["Jumlah"] = pd.to_numeric(
    master_df["Jumlah"],
    errors="coerce"
).fillna(0)

master_df["Actual"] = pd.to_numeric(
    master_df["Actual"],
    errors="coerce"
).fillna(0)

# Kalau Display sudah ada di file, pakai Display dari file.
# Kalau belum ada, hitung otomatis dari Jumlah - Actual.
if "Display" in master_df.columns:
    master_df["Display"] = pd.to_numeric(
        master_df["Display"],
        errors="coerce"
    )

    master_df["Display"] = master_df["Display"].fillna(
        master_df["Jumlah"] - master_df["Actual"]
    )
else:
    master_df["Display"] = master_df["Jumlah"] - master_df["Actual"]

# Restock dibutuhkan = Jumlah - Display
master_df["Restock Dibutuhkan"] = (
    master_df["Jumlah"] -
    master_df["Display"]
)

# Kalau display lebih besar dari jumlah, tidak perlu restock
master_df.loc[
    master_df["Restock Dibutuhkan"] < 0,
    "Restock Dibutuhkan"
] = 0

master_df["Status Display"] = master_df.apply(
    get_display_status,
    axis=1
)

if "Notes" not in master_df.columns:
    master_df["Notes"] = ""

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
    .astype(str)
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
    st.info(
        "Pastikan stock_harian.xlsx punya kolom: "
        "Tanggal, Lantai, Nama Produk, Harga Jual, Terjual, Uang Seharusnya Dibayar."
    )
    st.stop()

# ==========================================
# CLEAN STOCK DATA
# ==========================================
stock_df["Tanggal"] = pd.to_datetime(
    stock_df["Tanggal"],
    errors="coerce",
    dayfirst=True
).dt.date

stock_df["Nama Produk"] = (
    stock_df["Nama Produk"]
    .astype(str)
    .str.strip()
)

stock_df["Lantai"] = stock_df["Lantai"].apply(format_lantai)

stock_df["Harga Jual"] = stock_df["Harga Jual"].apply(clean_rupiah)

stock_df["Uang Seharusnya Dibayar"] = (
    stock_df["Uang Seharusnya Dibayar"]
    .apply(clean_rupiah)
)

stock_df["Terjual"] = pd.to_numeric(
    stock_df["Terjual"],
    errors="coerce"
).fillna(0)

# ==========================================
# STOCK COMPLIANCE CHECK
# Stock Sore Olsera vs Stock Sore
# ==========================================
stock_compliance_available = (
    "Stock Sore Olsera" in stock_df.columns and
    "Stock Sore" in stock_df.columns
)

if stock_compliance_available:
    stock_df["Stock Sore Olsera"] = pd.to_numeric(
        stock_df["Stock Sore Olsera"],
        errors="coerce"
    ).fillna(0)

    stock_df["Stock Sore"] = pd.to_numeric(
        stock_df["Stock Sore"],
        errors="coerce"
    ).fillna(0)

    stock_df["Selisih Stock"] = (
        stock_df["Stock Sore"] -
        stock_df["Stock Sore Olsera"]
    )

    stock_df["Status Stock"] = stock_df["Selisih Stock"].apply(get_stock_status)

else:
    st.warning(
        "⚠️ Kolom 'Stock Sore Olsera' dan/atau 'Stock Sore' tidak ditemukan. "
        "Cek kepatuhan stock tidak ditampilkan."
    )

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
    .astype(str)
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
    st.info("Pastikan qris_all.xlsx punya kolom: Nama Merchant, Tanggal Transaksi, Total Terbayar.")
    st.stop()

# ==========================================
# CLEAN QRIS DATA
# ==========================================
qris_df["Tanggal"] = pd.to_datetime(
    qris_df["Tanggal Transaksi"],
    errors="coerce",
    dayfirst=True
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

qris_df["Nama Merchant"] = (
    qris_df["Nama Merchant"]
    .astype(str)
    .str.strip()
)

qris_df["Lantai"] = qris_df["Nama Merchant"].map(merchant_mapping)

unknown_merchants = (
    qris_df[qris_df["Lantai"].isna()]["Nama Merchant"]
    .dropna()
    .unique()
)

if len(unknown_merchants) > 0:
    st.warning(f"⚠️ Ada Nama Merchant yang belum dimapping: {list(unknown_merchants)}")

# ==========================================
# FILTER SECTION
# ==========================================
st.subheader("🔍 Filter Dashboard")

available_lantai = [
    "Lantai 3",
    "Lantai 8",
    "Lantai 9"
]

all_dates = pd.Series(
    list(master_df["Tanggal"].dropna().unique()) +
    list(stock_df["Tanggal"].dropna().unique()) +
    list(qris_df["Tanggal"].dropna().unique())
)

available_tanggal = sorted(all_dates.dropna().unique())

col_filter1, col_filter2 = st.columns(2)

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
filtered_master_df = master_df[
    (master_df["Lantai"].isin(lantai_filter)) &
    (master_df["Tanggal"].isin(tanggal_filter))
].copy()

filtered_stock_df = stock_df[
    (stock_df["Lantai"].isin(lantai_filter)) &
    (stock_df["Tanggal"].isin(tanggal_filter))
].copy()

filtered_qris_df = qris_df[
    (qris_df["Lantai"].isin(lantai_filter)) &
    (qris_df["Tanggal"].isin(tanggal_filter))
].copy()

# ==========================================
# STOCK COMPLIANCE SUMMARY PER LANTAI
# ==========================================
if stock_compliance_available and not filtered_stock_df.empty:

    stock_flag_lantai_df = (
        filtered_stock_df
        .groupby(["Tanggal", "Lantai"])
        .agg(
            Total_Item=("Nama Produk", "count"),
            Item_Tidak_Patuh=("Status Stock", lambda x: (x == "🔴 Tidak Patuh").sum()),
            Total_Selisih_Stock=("Selisih Stock", "sum")
        )
        .reset_index()
    )

    stock_flag_lantai_df["Status Kepatuhan Stock"] = stock_flag_lantai_df[
        "Item_Tidak_Patuh"
    ].apply(
        lambda x: "🟢 Patuh" if x == 0 else "🔴 Tidak Patuh"
    )

else:
    stock_flag_lantai_df = pd.DataFrame()

# ==========================================
# KPI REKONSILIASI
# ==========================================
total_expected = filtered_stock_df["Uang Seharusnya Dibayar"].sum()
total_qris = filtered_qris_df["Total Terbayar"].sum()
total_selisih = total_qris - total_expected
total_terjual = filtered_stock_df["Terjual"].sum()

st.subheader("📌 Ringkasan Rekonsiliasi")

col1, col2, col3, col4 = st.columns(4)

col1.metric(
    "💰 Expected Revenue",
    rupiah(total_expected)
)

col2.metric(
    "💳 Actual QRIS",
    rupiah(total_qris)
)

col3.metric(
    "📦 Produk Terjual",
    f"{int(total_terjual)}"
)

col4.metric(
    "⚠️ Selisih",
    rupiah(total_selisih)
)

st.divider()

# ==========================================
# KPI KEPATUHAN STOCK
# ==========================================
if stock_compliance_available:
    st.subheader("🧮 Ringkasan Kepatuhan Stock")

    total_lantai_checked = len(stock_flag_lantai_df)

    total_lantai_patuh = (
        stock_flag_lantai_df["Status Kepatuhan Stock"]
        .eq("🟢 Patuh")
        .sum()
        if not stock_flag_lantai_df.empty else 0
    )

    total_lantai_tidak_patuh = (
        stock_flag_lantai_df["Status Kepatuhan Stock"]
        .eq("🔴 Tidak Patuh")
        .sum()
        if not stock_flag_lantai_df.empty else 0
    )

    total_item_tidak_patuh = (
        stock_flag_lantai_df["Item_Tidak_Patuh"]
        .sum()
        if not stock_flag_lantai_df.empty else 0
    )

    col_s1, col_s2, col_s3, col_s4 = st.columns(4)

    col_s1.metric("🏢 Lantai Dicek", int(total_lantai_checked))
    col_s2.metric("✅ Lantai Patuh", int(total_lantai_patuh))
    col_s3.metric("🚨 Lantai Tidak Patuh", int(total_lantai_tidak_patuh))
    col_s4.metric("📦 Item Selisih Stock", int(total_item_tidak_patuh))

    st.divider()

# ==========================================
# KPI MASTER DISPLAY
# ==========================================
st.subheader("🧾 Ringkasan Standar Display")

total_item_master = len(filtered_master_df)

total_sesuai = (
    filtered_master_df["Status Display"]
    .eq("🟢 Sesuai Standar")
    .sum()
)

total_perlu_restock = (
    filtered_master_df["Status Display"]
    .eq("🟡 Perlu Restock")
    .sum()
)

total_display_negatif = (
    filtered_master_df["Status Display"]
    .eq("🔴 Display Negatif")
    .sum()
)

total_restock_qty = (
    filtered_master_df["Restock Dibutuhkan"]
    .sum()
)

col_m1, col_m2, col_m3, col_m4 = st.columns(4)

col_m1.metric("📋 Total Item Master", int(total_item_master))
col_m2.metric("✅ Sesuai Standar", int(total_sesuai))
col_m3.metric("⚠️ Perlu Restock", int(total_perlu_restock))
col_m4.metric("📦 Total Qty Restock", int(total_restock_qty))

if total_display_negatif > 0:
    st.error(f"🚨 Ada {total_display_negatif} baris dengan Display negatif. Cek input Actual/Jumlah.")

st.divider()

# ==========================================
# REKONSILIASI DATA
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

rekon_df["Status"] = rekon_df["Selisih"].apply(get_rekon_status)

rekon_df = rekon_df.sort_values(
    by=["Tanggal", "Lantai"],
    ascending=[True, True]
)

# ==========================================
# CHART EXPECTED VS ACTUAL
# ==========================================
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

if chart_df.empty:
    st.info("Belum ada data rekonsiliasi untuk filter ini.")
else:
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

st.divider()

# ==========================================
# CHART DISPLAY STATUS
# ==========================================
st.subheader("📊 Status Standar Display")

display_status_df = (
    filtered_master_df
    .groupby("Status Display")
    .size()
    .reset_index(name="Jumlah Baris")
)

if display_status_df.empty:
    st.info("Belum ada data master display untuk filter ini.")
else:
    fig_display = px.bar(
        display_status_df,
        x="Status Display",
        y="Jumlah Baris",
        text_auto=True
    )

    fig_display.update_layout(
        height=400
    )

    st.plotly_chart(
        fig_display,
        use_container_width=True
    )

st.divider()

# ==========================================
# REKAP RESTOCK PER LANTAI
# ==========================================
st.subheader("📦 Rekap Restock Dibutuhkan per Lantai")

restock_lantai_df = (
    filtered_master_df
    .groupby("Lantai")["Restock Dibutuhkan"]
    .sum()
    .reset_index()
    .sort_values("Lantai")
)

if restock_lantai_df.empty:
    st.info("Belum ada data restock untuk filter ini.")
else:
    st.dataframe(
        restock_lantai_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Restock Dibutuhkan": st.column_config.NumberColumn(
                "Restock Dibutuhkan",
                format="%d"
            )
        }
    )

st.divider()

# ==========================================
# DATA SELISIH REKONSILIASI
# ==========================================
st.subheader("🚨 Data Selisih Rekonsiliasi Tidak Match")

selisih_df = rekon_df[rekon_df["Selisih"] != 0].copy()

if selisih_df.empty:
    st.success("✅ Semua data rekonsiliasi sudah MATCH.")
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
# DATA KEPATUHAN STOCK PER LANTAI
# ==========================================
if stock_compliance_available:
    st.subheader("🚨 Flag Kepatuhan Stock per Lantai")

    stock_tidak_patuh_lantai_df = stock_flag_lantai_df[
        stock_flag_lantai_df["Status Kepatuhan Stock"] == "🔴 Tidak Patuh"
    ].copy()

    if stock_tidak_patuh_lantai_df.empty:
        st.success("✅ Semua lantai patuh. Stock Sore Olsera sudah sesuai dengan Stock Sore.")
    else:
        st.dataframe(
            stock_tidak_patuh_lantai_df,
            use_container_width=True,
            hide_index=True,
            column_config={
                "Total_Item": st.column_config.NumberColumn(
                    "Total Item",
                    format="%d"
                ),
                "Item_Tidak_Patuh": st.column_config.NumberColumn(
                    "Item Tidak Patuh",
                    format="%d"
                ),
                "Total_Selisih_Stock": st.column_config.NumberColumn(
                    "Total Selisih Stock",
                    format="%d"
                ),
            }
        )

    with st.expander("Lihat Detail Item Selisih Stock"):
        detail_stock_selisih_df = filtered_stock_df[
            filtered_stock_df["Status Stock"] == "🔴 Tidak Patuh"
        ].copy()

        if detail_stock_selisih_df.empty:
            st.success("Tidak ada item yang selisih.")
        else:
            st.dataframe(
                detail_stock_selisih_df[
                    [
                        "Tanggal",
                        "Lantai",
                        "Nama Produk",
                        "Stock Sore Olsera",
                        "Stock Sore",
                        "Selisih Stock",
                        "Status Stock"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Stock Sore Olsera": st.column_config.NumberColumn(
                        "Stock Sore Olsera",
                        format="%d"
                    ),
                    "Stock Sore": st.column_config.NumberColumn(
                        "Stock Sore",
                        format="%d"
                    ),
                    "Selisih Stock": st.column_config.NumberColumn(
                        "Selisih Stock",
                        format="%d"
                    ),
                }
            )

    st.divider()

# ==========================================
# MASTER DISPLAY CHECK TABLE
# ==========================================
st.subheader("🧾 Cek Display vs Jumlah Standar")

problem_display_df = filtered_master_df[
    filtered_master_df["Status Display"] != "🟢 Sesuai Standar"
].copy()

display_columns = [
    "Tanggal",
    "Lantai",
    "Nama Produk",
    "Jumlah",
    "Actual",
    "Display",
    "Restock Dibutuhkan",
    "Status Display",
    "Notes",
    "Source Sheet"
]

display_columns = [
    col for col in display_columns
    if col in problem_display_df.columns
]

if problem_display_df.empty:
    st.success("✅ Semua display sudah sesuai standar.")
else:
    st.dataframe(
        problem_display_df[display_columns],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Jumlah": st.column_config.NumberColumn(
                "Jumlah",
                format="%d"
            ),
            "Actual": st.column_config.NumberColumn(
                "Actual",
                format="%d"
            ),
            "Display": st.column_config.NumberColumn(
                "Display",
                format="%d"
            ),
            "Restock Dibutuhkan": st.column_config.NumberColumn(
                "Restock Dibutuhkan",
                format="%d"
            ),
        }
    )

with st.expander("Lihat Semua Data Master Display"):
    st.dataframe(
        filtered_master_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Jumlah": st.column_config.NumberColumn(
                "Jumlah",
                format="%d"
            ),
            "Actual": st.column_config.NumberColumn(
                "Actual",
                format="%d"
            ),
            "Display": st.column_config.NumberColumn(
                "Display",
                format="%d"
            ),
            "Restock Dibutuhkan": st.column_config.NumberColumn(
                "Restock Dibutuhkan",
                format="%d"
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

stock_column_config = {
    "Harga Jual": st.column_config.NumberColumn(
        "Harga Jual",
        format="Rp %d"
    ),
    "Uang Seharusnya Dibayar": st.column_config.NumberColumn(
        "Uang Seharusnya Dibayar",
        format="Rp %d"
    ),
}

if stock_compliance_available:
    stock_column_config.update({
        "Stock Sore Olsera": st.column_config.NumberColumn(
            "Stock Sore Olsera",
            format="%d"
        ),
        "Stock Sore": st.column_config.NumberColumn(
            "Stock Sore",
            format="%d"
        ),
        "Selisih Stock": st.column_config.NumberColumn(
            "Selisih Stock",
            format="%d"
        ),
    })

st.dataframe(
    filtered_stock_df,
    use_container_width=True,
    hide_index=True,
    column_config=stock_column_config
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
# TOP PRODUK TERLARIS - MOVED TO BOTTOM
# ==========================================
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

if top_produk.empty:
    st.info("Belum ada data produk untuk filter ini.")
else:
    fig_top_produk = px.bar(
        top_produk,
        x="Nama Produk",
        y="Terjual",
        text_auto=True,
        color="Terjual"
    )

    fig_top_produk.update_layout(
        height=500,
        xaxis_tickangle=-30
    )

    st.plotly_chart(
        fig_top_produk,
        use_container_width=True
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
excel_buffer = BytesIO()

with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    rekon_df.to_excel(
        writer,
        sheet_name="Rekonsiliasi",
        index=False
    )

    selisih_df.to_excel(
        writer,
        sheet_name="Data Selisih Rekon",
        index=False
    )

    filtered_master_df.to_excel(
        writer,
        sheet_name="Master Display",
        index=False
    )

    problem_display_df.to_excel(
        writer,
        sheet_name="Display Perlu Dicek",
        index=False
    )

    restock_lantai_df.to_excel(
        writer,
        sheet_name="Rekap Restock Lantai",
        index=False
    )

    if stock_compliance_available:
        stock_flag_lantai_df.to_excel(
            writer,
            sheet_name="Flag Stock Lantai",
            index=False
        )

        filtered_stock_df[
            filtered_stock_df["Status Stock"] == "🔴 Tidak Patuh"
        ].to_excel(
            writer,
            sheet_name="Detail Stock Tidak Patuh",
            index=False
        )

    filtered_stock_df.to_excel(
        writer,
        sheet_name="Detail Stock",
        index=False
    )

    filtered_qris_df.to_excel(
        writer,
        sheet_name="Detail QRIS",
        index=False
    )

excel_buffer.seek(0)

st.download_button(
    label="⬇️ Download Rekonsiliasi Excel",
    data=excel_buffer,
    file_name="rekonsiliasi_warung_fifgroup.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
