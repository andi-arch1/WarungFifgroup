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
# FILE PATH
# ==========================================
stock_file = "data/stock_harian.xlsx"
qris_file = "data/qris_all.xlsx"

# ==========================================
# SIDEBAR MENU
# ==========================================
st.sidebar.title("🛒 Warung FIFGROUP")

menu = st.sidebar.radio(
    "Pilih Menu",
    [
        "Dashboard Utama",
        "Inventory Restock"
    ]
)

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


def get_untung_rugi_status(selisih):
    if selisih == 0:
        return "🟢 Impas / Match"
    elif selisih > 0:
        return "🔵 Untung / Lebih"
    else:
        return "🔴 Rugi / Kurang"


def get_inventory_status(row):
    jumlah = row["Jumlah"]
    batas_minimum = row["Batas Minimum"]

    if jumlah <= 0:
        return "🔴 Habis"

    if jumlah <= batas_minimum:
        return "🟡 Menipis"

    return "🟢 Aman"


def rupiah(value):
    return f"Rp {value:,.0f}".replace(",", ".")


# ==========================================
# LOAD STOCK WORKBOOK
# ==========================================
try:
    workbook_sheets = pd.read_excel(
        stock_file,
        sheet_name=None
    )

except Exception as e:
    st.error(f"❌ Error membaca stock_harian.xlsx: {e}")
    st.stop()

# ==========================================
# VALIDATE REQUIRED SHEETS
# ==========================================
required_sheets = [
    "stock_harian",
    "Inventory"
]

missing_sheets = [
    sheet for sheet in required_sheets
    if sheet not in workbook_sheets.keys()
]

if missing_sheets:
    st.error(f"❌ Sheet tidak ditemukan di stock_harian.xlsx: {missing_sheets}")
    st.stop()

# ==========================================
# LOAD STOCK_HARIAN SHEET
# ==========================================
stock_df = workbook_sheets["stock_harian"].copy()

stock_df.columns = (
    stock_df.columns
    .astype(str)
    .str.strip()
    .str.replace("\ufeff", "")
)

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
    st.error(f"❌ Kolom stock_harian tidak ditemukan: {missing_stock_cols}")
    st.stop()

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
# LOAD DISPLAY MASTER SHEETS
# lantai3, lantai8, lantai9
# ==========================================
display_sheet_names = [
    sheet for sheet in workbook_sheets.keys()
    if str(sheet).strip().lower() in ["lantai3", "lantai8", "lantai9"]
]

master_list = []

for sheet_name in display_sheet_names:
    df = workbook_sheets[sheet_name].copy()

    if df.empty:
        continue

    df.columns = (
        df.columns
        .astype(str)
        .str.strip()
        .str.replace("\ufeff", "")
    )

    rename_cols = {}

    if "Produk" in df.columns and "Nama Produk" not in df.columns:
        rename_cols["Produk"] = "Nama Produk"

    if "Standar Display" in df.columns and "Jumlah" not in df.columns:
        rename_cols["Standar Display"] = "Jumlah"

    if "Qty" in df.columns and "Jumlah" not in df.columns:
        rename_cols["Qty"] = "Jumlah"

    if "Stock Display" in df.columns and "Display" not in df.columns:
        rename_cols["Stock Display"] = "Display"

    if "lantai" in df.columns and "Lantai" not in df.columns:
        rename_cols["lantai"] = "Lantai"

    df = df.rename(columns=rename_cols)

    if "Lantai" not in df.columns:
        df["Lantai"] = lantai_from_sheet_name(sheet_name)

    df["Source Sheet"] = sheet_name

    master_list.append(df)

if len(master_list) == 0:
    st.error("❌ Sheet lantai3/lantai8/lantai9 tidak ditemukan atau kosong.")
    st.stop()

master_df = pd.concat(
    master_list,
    ignore_index=True
)

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
    st.error(f"❌ Kolom master display tidak ditemukan: {missing_master_cols}")
    st.stop()

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

master_df["Restock Dibutuhkan"] = (
    master_df["Jumlah"] -
    master_df["Display"]
)

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
# LOAD INVENTORY SHEET
# ==========================================
inventory_df = workbook_sheets["Inventory"].copy()

inventory_df.columns = (
    inventory_df.columns
    .astype(str)
    .str.strip()
    .str.replace("\ufeff", "")
)

required_inventory_cols = [
    "Tanggal",
    "Produk",
    "Jumlah"
]

missing_inventory_cols = [
    col for col in required_inventory_cols
    if col not in inventory_df.columns
]

if missing_inventory_cols:
    st.error(f"❌ Kolom Inventory tidak ditemukan: {missing_inventory_cols}")
    st.stop()

inventory_df["Tanggal"] = pd.to_datetime(
    inventory_df["Tanggal"],
    errors="coerce",
    dayfirst=True
).dt.date

inventory_df["Produk"] = (
    inventory_df["Produk"]
    .astype(str)
    .str.strip()
)

inventory_df["Jumlah"] = pd.to_numeric(
    inventory_df["Jumlah"],
    errors="coerce"
).fillna(0)

if "Restock" not in inventory_df.columns:
    inventory_df["Restock"] = 0

inventory_df["Restock"] = pd.to_numeric(
    inventory_df["Restock"],
    errors="coerce"
).fillna(0)

# Ambil data inventory terakhir per produk
inventory_latest_df = (
    inventory_df
    .sort_values("Tanggal")
    .groupby("Produk", as_index=False)
    .tail(1)
    .copy()
)

# Kalau ada kolom Stock Awal, pakai itu.
# Kalau belum ada, stock awal dihitung dari jumlah maksimum historis per produk.
if "Stock Awal" in inventory_df.columns:
    stock_awal_df = (
        inventory_df[["Produk", "Stock Awal"]]
        .copy()
    )

    stock_awal_df["Stock Awal"] = pd.to_numeric(
        stock_awal_df["Stock Awal"],
        errors="coerce"
    )

    stock_awal_df = (
        stock_awal_df
        .dropna(subset=["Stock Awal"])
        .groupby("Produk", as_index=False)["Stock Awal"]
        .max()
    )

else:
    stock_awal_df = (
        inventory_df
        .groupby("Produk", as_index=False)["Jumlah"]
        .max()
        .rename(columns={"Jumlah": "Stock Awal"})
    )

inventory_latest_df = inventory_latest_df.merge(
    stock_awal_df,
    on="Produk",
    how="left"
)

inventory_latest_df["Stock Awal"] = (
    inventory_latest_df["Stock Awal"]
    .fillna(inventory_latest_df["Jumlah"])
)

inventory_latest_df["Batas Minimum"] = (
    inventory_latest_df["Stock Awal"] * 0.10
)

inventory_latest_df["Status Inventory"] = inventory_latest_df.apply(
    get_inventory_status,
    axis=1
)

inventory_latest_df["Saran Restock"] = (
    inventory_latest_df["Stock Awal"] -
    inventory_latest_df["Jumlah"]
)

inventory_latest_df.loc[
    inventory_latest_df["Saran Restock"] < 0,
    "Saran Restock"
] = 0

inventory_latest_df = inventory_latest_df.sort_values(
    by=["Status Inventory", "Produk"],
    ascending=[True, True]
)

# ==========================================
# LOAD QRIS EXCEL
# ==========================================
try:
    qris_df = pd.read_excel(qris_file)

except Exception as e:
    st.error(f"❌ Error membaca qris_all.xlsx: {e}")
    st.stop()

qris_df.columns = (
    qris_df.columns
    .astype(str)
    .str.strip()
    .str.replace("\ufeff", "")
)

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
# INVENTORY RESTOCK PAGE
# ==========================================
if menu == "Inventory Restock":

    st.title("📦 Inventory Restock")

    st.markdown("""
    Monitoring inventory untuk menentukan produk yang **habis** atau **menipis**.  
    Produk dianggap menipis jika jumlah inventory sudah berada di bawah atau sama dengan **10% dari stock awal**.
    """)

    st.divider()

    total_produk_inventory = len(inventory_latest_df)

    total_habis = (
        inventory_latest_df["Status Inventory"]
        .eq("🔴 Habis")
        .sum()
    )

    total_menipis = (
        inventory_latest_df["Status Inventory"]
        .eq("🟡 Menipis")
        .sum()
    )

    total_aman = (
        inventory_latest_df["Status Inventory"]
        .eq("🟢 Aman")
        .sum()
    )

    total_saran_restock = inventory_latest_df[
        inventory_latest_df["Status Inventory"].isin(["🔴 Habis", "🟡 Menipis"])
    ]["Saran Restock"].sum()

    col_i1, col_i2, col_i3, col_i4 = st.columns(4)

    col_i1.metric("📋 Total Produk", int(total_produk_inventory))
    col_i2.metric("🔴 Habis", int(total_habis))
    col_i3.metric("🟡 Menipis", int(total_menipis))
    col_i4.metric("📦 Saran Qty Restock", int(total_saran_restock))

    st.divider()

    st.subheader("🚨 Produk Perlu Restock ke Supplier")

    restock_supplier_df = inventory_latest_df[
        inventory_latest_df["Status Inventory"].isin(["🔴 Habis", "🟡 Menipis"])
    ].copy()

    if restock_supplier_df.empty:
        st.success("✅ Semua inventory masih aman. Belum perlu restock supplier.")
    else:
        st.dataframe(
            restock_supplier_df[
                [
                    "Tanggal",
                    "Produk",
                    "Jumlah",
                    "Stock Awal",
                    "Batas Minimum",
                    "Saran Restock",
                    "Status Inventory"
                ]
            ],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Jumlah": st.column_config.NumberColumn(
                    "Jumlah Sekarang",
                    format="%d"
                ),
                "Stock Awal": st.column_config.NumberColumn(
                    "Stock Awal",
                    format="%d"
                ),
                "Batas Minimum": st.column_config.NumberColumn(
                    "Batas Minimum 10%",
                    format="%.1f"
                ),
                "Saran Restock": st.column_config.NumberColumn(
                    "Saran Restock",
                    format="%d"
                ),
            }
        )

    st.divider()

    st.subheader("📊 Status Inventory")

    inventory_status_df = (
        inventory_latest_df
        .groupby("Status Inventory")
        .size()
        .reset_index(name="Jumlah Produk")
    )

    fig_inventory = px.bar(
        inventory_status_df,
        x="Status Inventory",
        y="Jumlah Produk",
        text_auto=True
    )

    fig_inventory.update_layout(
        height=400
    )

    st.plotly_chart(
        fig_inventory,
        use_container_width=True
    )

    st.divider()

    st.subheader("📋 Semua Data Inventory Terakhir")

    st.dataframe(
        inventory_latest_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Jumlah": st.column_config.NumberColumn(
                "Jumlah Sekarang",
                format="%d"
            ),
            "Stock Awal": st.column_config.NumberColumn(
                "Stock Awal",
                format="%d"
            ),
            "Batas Minimum": st.column_config.NumberColumn(
                "Batas Minimum 10%",
                format="%.1f"
            ),
            "Saran Restock": st.column_config.NumberColumn(
                "Saran Restock",
                format="%d"
            ),
        }
    )

    st.divider()

    inventory_excel_buffer = BytesIO()

    with pd.ExcelWriter(inventory_excel_buffer, engine="openpyxl") as writer:
        restock_supplier_df.to_excel(
            writer,
            sheet_name="Perlu Restock",
            index=False
        )

        inventory_latest_df.to_excel(
            writer,
            sheet_name="Inventory Terakhir",
            index=False
        )

        inventory_df.to_excel(
            writer,
            sheet_name="Raw Inventory",
            index=False
        )

    inventory_excel_buffer.seek(0)

    st.download_button(
        label="⬇️ Download Data Inventory Restock",
        data=inventory_excel_buffer,
        file_name="inventory_restock_supplier.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.stop()

# ==========================================
# DASHBOARD UTAMA PAGE
# ==========================================
st.title("🛒 Dashboard Warung FIFGROUP")

st.markdown("""
Monitoring Penjualan Harian, Rekonsiliasi QRIS, Kepatuhan Stock, dan Standar Display  
📍 Warung FIFGROUP Lantai 3 • Lantai 8 • Lantai 9
""")

st.divider()

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
# UNTUNG RUGI PER LANTAI
# ==========================================
expected_lantai_df = (
    filtered_stock_df
    .groupby("Lantai")["Uang Seharusnya Dibayar"]
    .sum()
    .reset_index()
)

actual_lantai_df = (
    filtered_qris_df
    .groupby("Lantai")["Total Terbayar"]
    .sum()
    .reset_index()
)

untung_rugi_lantai_df = expected_lantai_df.merge(
    actual_lantai_df,
    on="Lantai",
    how="outer"
)

untung_rugi_lantai_df["Uang Seharusnya Dibayar"] = (
    untung_rugi_lantai_df["Uang Seharusnya Dibayar"]
    .fillna(0)
)

untung_rugi_lantai_df["Total Terbayar"] = (
    untung_rugi_lantai_df["Total Terbayar"]
    .fillna(0)
)

untung_rugi_lantai_df["Selisih"] = (
    untung_rugi_lantai_df["Total Terbayar"] -
    untung_rugi_lantai_df["Uang Seharusnya Dibayar"]
)

untung_rugi_lantai_df["Status"] = (
    untung_rugi_lantai_df["Selisih"]
    .apply(get_untung_rugi_status)
)

lantai_order = {
    "Lantai 3": 1,
    "Lantai 8": 2,
    "Lantai 9": 3
}

untung_rugi_lantai_df["Urutan"] = (
    untung_rugi_lantai_df["Lantai"]
    .map(lantai_order)
)

untung_rugi_lantai_df = (
    untung_rugi_lantai_df
    .sort_values("Urutan")
    .drop(columns=["Urutan"])
)

# ==========================================
# STOCK SUMMARY PER LANTAI
# ==========================================
if stock_compliance_available and not filtered_stock_df.empty:

    stock_summary_lantai_df = (
        filtered_stock_df
        .groupby("Lantai")
        .agg(
            Total_Item=("Nama Produk", "count"),
            Item_Selisih_Stock=("Status Stock", lambda x: (x == "🔴 Tidak Patuh").sum()),
            Total_Selisih_Stock=("Selisih Stock", lambda x: x.abs().sum())
        )
        .reset_index()
    )

    stock_summary_lantai_df["Status Kepatuhan Stock"] = stock_summary_lantai_df[
        "Item_Selisih_Stock"
    ].apply(
        lambda x: "🟢 Patuh" if x == 0 else "🔴 Tidak Patuh"
    )

    stock_summary_lantai_df["Urutan"] = stock_summary_lantai_df["Lantai"].map(lantai_order)

    stock_summary_lantai_df = (
        stock_summary_lantai_df
        .sort_values("Urutan")
        .drop(columns=["Urutan"])
    )

else:
    stock_summary_lantai_df = pd.DataFrame()

# ==========================================
# DISPLAY SUMMARY PER LANTAI
# ==========================================
if not filtered_master_df.empty:

    display_summary_lantai_df = (
        filtered_master_df
        .groupby("Lantai")
        .agg(
            Total_Item_Master=("Nama Produk", "count"),
            Sesuai_Standar=("Status Display", lambda x: (x == "🟢 Sesuai Standar").sum()),
            Perlu_Restock=("Status Display", lambda x: (x == "🟡 Perlu Restock").sum()),
            Display_Negatif=("Status Display", lambda x: (x == "🔴 Display Negatif").sum()),
            Total_Qty_Restock=("Restock Dibutuhkan", "sum")
        )
        .reset_index()
    )

    display_summary_lantai_df["Status Standar Display"] = display_summary_lantai_df.apply(
        lambda row: (
            "🔴 Ada Display Negatif"
            if row["Display_Negatif"] > 0
            else "🟡 Perlu Restock"
            if row["Perlu_Restock"] > 0
            else "🟢 Sesuai Standar"
        ),
        axis=1
    )

    display_summary_lantai_df["Urutan"] = display_summary_lantai_df["Lantai"].map(lantai_order)

    display_summary_lantai_df = (
        display_summary_lantai_df
        .sort_values("Urutan")
        .drop(columns=["Urutan"])
    )

else:
    display_summary_lantai_df = pd.DataFrame()

# ==========================================
# KPI REKONSILIASI
# ==========================================
total_expected = filtered_stock_df["Uang Seharusnya Dibayar"].sum()
total_qris = filtered_qris_df["Total Terbayar"].sum()
total_selisih = total_qris - total_expected
total_terjual = filtered_stock_df["Terjual"].sum()

st.subheader("📌 Ringkasan Rekonsiliasi")

col1, col2, col3, col4 = st.columns(4)

col1.metric("💰 Expected Revenue", rupiah(total_expected))
col2.metric("💳 Actual QRIS", rupiah(total_qris))
col3.metric("📦 Produk Terjual", f"{int(total_terjual)}")
col4.metric("⚠️ Selisih", rupiah(total_selisih))

st.divider()

# ==========================================
# RINGKASAN UNTUNG RUGI PER LANTAI
# ==========================================
st.subheader("💸 Ringkasan Untung / Rugi per Lantai")

col_u3, col_u8, col_u9 = st.columns(3)

for col, lantai in zip(
    [col_u3, col_u8, col_u9],
    ["Lantai 3", "Lantai 8", "Lantai 9"]
):
    data_lantai = untung_rugi_lantai_df[
        untung_rugi_lantai_df["Lantai"] == lantai
    ]

    if data_lantai.empty:
        col.metric(f"🏢 {lantai}", "Tidak Ada Data")
    else:
        selisih = data_lantai["Selisih"].iloc[0]
        status = data_lantai["Status"].iloc[0]

        col.metric(
            f"🏢 {lantai}",
            rupiah(selisih),
            status.replace("🔵 ", "").replace("🔴 ", "").replace("🟢 ", "")
        )

st.dataframe(
    untung_rugi_lantai_df,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Uang Seharusnya Dibayar": st.column_config.NumberColumn(
            "Expected Revenue",
            format="Rp %d"
        ),
        "Total Terbayar": st.column_config.NumberColumn(
            "Actual QRIS",
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
# RINGKASAN KEPATUHAN STOCK PER LANTAI
# ==========================================
if stock_compliance_available:
    st.subheader("🧮 Ringkasan Kepatuhan Stock per Lantai")

    col_l3, col_l8, col_l9 = st.columns(3)

    for col, lantai in zip(
        [col_l3, col_l8, col_l9],
        ["Lantai 3", "Lantai 8", "Lantai 9"]
    ):
        data_lantai = stock_summary_lantai_df[
            stock_summary_lantai_df["Lantai"] == lantai
        ]

        if data_lantai.empty:
            col.metric(f"🏢 {lantai}", "Tidak Ada Data")
        else:
            item_selisih = int(data_lantai["Item_Selisih_Stock"].iloc[0])
            total_selisih = int(data_lantai["Total_Selisih_Stock"].iloc[0])
            status = data_lantai["Status Kepatuhan Stock"].iloc[0]

            col.metric(
                f"🏢 {lantai}",
                status.replace("🔴 ", "").replace("🟢 ", ""),
                f"{item_selisih} item selisih | Total selisih: {total_selisih}"
            )

    st.dataframe(
        stock_summary_lantai_df,
        use_container_width=True,
        hide_index=True
    )

    st.divider()

# ==========================================
# DETAIL STOCK TIDAK PATUH PER LANTAI
# ==========================================
if stock_compliance_available:
    st.subheader("🔎 Detail Stock Tidak Patuh per Lantai")

    tab_stock_l3, tab_stock_l8, tab_stock_l9 = st.tabs([
        "🏢 Lantai 3",
        "🏢 Lantai 8",
        "🏢 Lantai 9"
    ])

    stock_tabs = {
        "Lantai 3": tab_stock_l3,
        "Lantai 8": tab_stock_l8,
        "Lantai 9": tab_stock_l9
    }

    for lantai, tab in stock_tabs.items():
        with tab:
            detail_lantai_df = filtered_stock_df[
                (filtered_stock_df["Lantai"] == lantai) &
                (filtered_stock_df["Status Stock"] == "🔴 Tidak Patuh")
            ].copy()

            st.markdown(f"### {lantai}")

            if detail_lantai_df.empty:
                st.success("✅ Tidak ada item selisih stock di lantai ini.")
            else:
                total_item_selisih = len(detail_lantai_df)
                total_selisih_abs = detail_lantai_df["Selisih Stock"].abs().sum()
                total_selisih_net = detail_lantai_df["Selisih Stock"].sum()

                col_a, col_b, col_c = st.columns(3)

                col_a.metric("📦 Item Tidak Patuh", int(total_item_selisih))
                col_b.metric("📊 Total Selisih Abs", int(total_selisih_abs))
                col_c.metric("🧾 Total Selisih Net", int(total_selisih_net))

                detail_stock_cols = [
                    "Tanggal",
                    "Lantai",
                    "Nama Produk",
                    "Stock Sore Olsera",
                    "Stock Sore",
                    "Selisih Stock",
                    "Status Stock"
                ]

                detail_stock_cols = [
                    col for col in detail_stock_cols
                    if col in detail_lantai_df.columns
                ]

                st.dataframe(
                    detail_lantai_df[detail_stock_cols],
                    use_container_width=True,
                    hide_index=True
                )

    st.divider()

# ==========================================
# RINGKASAN STANDAR DISPLAY PER LANTAI
# ==========================================
st.subheader("🧾 Ringkasan Standar Display per Lantai")

col_d3, col_d8, col_d9 = st.columns(3)

for col, lantai in zip(
    [col_d3, col_d8, col_d9],
    ["Lantai 3", "Lantai 8", "Lantai 9"]
):
    data_lantai = display_summary_lantai_df[
        display_summary_lantai_df["Lantai"] == lantai
    ]

    if data_lantai.empty:
        col.metric(f"🏢 {lantai}", "Tidak Ada Data")
    else:
        total_item = int(data_lantai["Total_Item_Master"].iloc[0])
        sesuai = int(data_lantai["Sesuai_Standar"].iloc[0])
        perlu_restock = int(data_lantai["Perlu_Restock"].iloc[0])
        total_qty_restock = int(data_lantai["Total_Qty_Restock"].iloc[0])
        status = data_lantai["Status Standar Display"].iloc[0]

        col.metric(
            f"🏢 {lantai}",
            status.replace("🔴 ", "").replace("🟡 ", "").replace("🟢 ", ""),
            f"{sesuai}/{total_item} sesuai | {perlu_restock} restock | Qty {total_qty_restock}"
        )

st.dataframe(
    display_summary_lantai_df,
    use_container_width=True,
    hide_index=True
)

st.divider()

# ==========================================
# DETAIL DISPLAY RESTOCK PER LANTAI
# ==========================================
st.subheader("🔎 Detail Display / Restock per Lantai")

tab_display_l3, tab_display_l8, tab_display_l9 = st.tabs([
    "🏢 Lantai 3",
    "🏢 Lantai 8",
    "🏢 Lantai 9"
])

display_tabs = {
    "Lantai 3": tab_display_l3,
    "Lantai 8": tab_display_l8,
    "Lantai 9": tab_display_l9
}

for lantai, tab in display_tabs.items():
    with tab:
        detail_display_lantai_df = filtered_master_df[
            (filtered_master_df["Lantai"] == lantai) &
            (filtered_master_df["Status Display"] != "🟢 Sesuai Standar")
        ].copy()

        st.markdown(f"### {lantai}")

        if detail_display_lantai_df.empty:
            st.success("✅ Semua display sudah sesuai standar di lantai ini.")
        else:
            total_perlu_restock = (
                detail_display_lantai_df["Status Display"]
                .eq("🟡 Perlu Restock")
                .sum()
            )

            total_display_negatif = (
                detail_display_lantai_df["Status Display"]
                .eq("🔴 Display Negatif")
                .sum()
            )

            total_qty_restock = detail_display_lantai_df["Restock Dibutuhkan"].sum()

            col_a, col_b, col_c = st.columns(3)

            col_a.metric("⚠️ Item Perlu Restock", int(total_perlu_restock))
            col_b.metric("🚨 Display Negatif", int(total_display_negatif))
            col_c.metric("📦 Qty Restock", int(total_qty_restock))

            display_detail_columns = [
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

            display_detail_columns = [
                col for col in display_detail_columns
                if col in detail_display_lantai_df.columns
            ]

            st.dataframe(
                detail_display_lantai_df[display_detail_columns],
                use_container_width=True,
                hide_index=True
            )

st.divider()

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
# DETAIL STOCK
# ==========================================
st.subheader("📦 Detail Penjualan Stock")

st.dataframe(
    filtered_stock_df,
    use_container_width=True,
    hide_index=True
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
# TOP PRODUK TERLARIS
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
# DOWNLOAD EXCEL
# ==========================================
excel_buffer = BytesIO()

with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    rekon_df.to_excel(
        writer,
        sheet_name="Rekonsiliasi",
        index=False
    )

    untung_rugi_lantai_df.to_excel(
        writer,
        sheet_name="Untung Rugi Lantai",
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

    display_summary_lantai_df.to_excel(
        writer,
        sheet_name="Summary Display Lantai",
        index=False
    )

    if stock_compliance_available:
        stock_summary_lantai_df.to_excel(
            writer,
            sheet_name="Summary Stock Lantai",
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

    inventory_latest_df.to_excel(
        writer,
        sheet_name="Inventory Terakhir",
        index=False
    )

excel_buffer.seek(0)

st.download_button(
    label="⬇️ Download Rekonsiliasi Excel",
    data=excel_buffer,
    file_name="rekonsiliasi_warung_fifgroup.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
