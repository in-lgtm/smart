import io
import os
import re
import sqlite3
from datetime import datetime
import pandas as pd
import streamlit as st

# -----------------------------------------------------------------------------
# APP CONFIG & SAFE DIRECTORY SETUP
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Material Inventory System",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_FILE = os.path.join(BASE_DIR, "inventory.db")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
DATASHEETS_DIR = os.path.join(BASE_DIR, "datasheets")

for folder_path in [IMAGES_DIR, DATASHEETS_DIR]:
    if os.path.exists(folder_path) and not os.path.isdir(folder_path):
        os.remove(folder_path)
    os.makedirs(folder_path, exist_ok=True)


# -----------------------------------------------------------------------------
# DATABASE CONNECTION & MUTATION HELPERS
# -----------------------------------------------------------------------------
def get_db_connection():
    conn = sqlite3.connect(DB_FILE, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL;")
    conn.execute("PRAGMA foreign_keys=ON;")
    return conn


def execute_db_query(query, params=()):
    conn = get_db_connection()
    try:
        cursor = conn.cursor()
        cursor.execute(query, params)
        conn.commit()
        last_id = cursor.lastrowid
        return last_id
    finally:
        conn.close()


def sanitize_filename(name):
    clean = re.sub(r"[^\w\-_]", "_", str(name))
    return clean[:30]


def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='products'")
        table_exists = cursor.fetchone()

        if table_exists:
            cursor.execute("PRAGMA table_info(products)")
            columns = [column[1] for column in cursor.fetchall()]
            required_cols = [
                "sds_hazard_class", "source_origin", "batch_lot", "where_used",
                "unit", "delivery_time_unit",
            ]
            if not all(col in columns for col in required_cols):
                cursor.execute("DROP TABLE IF EXISTS products")
                cursor.execute("DROP TABLE IF EXISTS stock_entries")
                cursor.execute("DROP TABLE IF EXISTS contacts")

        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='stock_entries'")
        stock_table_exists = cursor.fetchone()

        if stock_table_exists:
            cursor.execute("PRAGMA table_info(stock_entries)")
            se_columns = [column[1] for column in cursor.fetchall()]
            if "unit" not in se_columns or "movement_type" not in se_columns:
                cursor.execute("DROP TABLE IF EXISTS stock_entries")

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS products (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                icon TEXT DEFAULT '📦',
                sku TEXT UNIQUE NOT NULL,
                product TEXT NOT NULL,
                characteristics TEXT,
                suppliers TEXT,
                price REAL DEFAULT 0.0,
                discount REAL DEFAULT 0.0,
                transport_price REAL DEFAULT 0.0,
                expiring_date TEXT,
                delivery_time INTEGER DEFAULT 5,
                delivery_time_unit TEXT DEFAULT 'days',
                ubication TEXT,
                monthly_usage INTEGER DEFAULT 0,
                min_stock INTEGER DEFAULT 0,
                quantity REAL DEFAULT 0.0,
                unit TEXT DEFAULT 'pieces',
                description TEXT,
                where_used TEXT,
                source_origin TEXT,
                batch_lot TEXT,
                sds_hazard_class TEXT,
                photo_path TEXT,
                datasheet_path TEXT
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS stock_entries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                entry_date TEXT NOT NULL,
                quantity REAL NOT NULL,
                unit TEXT DEFAULT 'pieces',
                movement_type TEXT DEFAULT 'IN',
                price REAL NOT NULL,
                note TEXT,
                FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
            )
        """)

        cursor.execute("""
            CREATE TABLE IF NOT EXISTS contacts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                product_id INTEGER NOT NULL,
                name TEXT,
                surname TEXT,
                phone TEXT,
                email TEXT,
                country TEXT,
                FOREIGN KEY (product_id) REFERENCES products (id) ON DELETE CASCADE
            )
        """)

        cursor.execute("SELECT COUNT(*) FROM products")
        if cursor.fetchone()[0] == 0:
            seed_data = [
                (
                    "🏷️",
                    "MAT-VEL-401",
                    "Velcro 401 Black",
                    "1 Box = 14 U | Breakdown: 1 Box x 350 mts + 8 U x 25 mts",
                    "Velcro Industrial",
                    120.00,
                    5.0,
                    10.00,
                    "2030-12-31",
                    5,
                    "Zone A - Rack 01",
                    200,
                    250,
                    550.0,
                    "1 Box x 350 mts (350m) + 8 U x 25 mts (200m). Total: 550 mts.",
                    "Vertical/horizontal textile fastening and modular panels.",
                    "Barcelona, Spain",
                    "LOT-VEL401-26",
                    "Non-hazardous",
                ),
                (
                    "🏷️",
                    "MAT-VEL-758",
                    "Velcro 758 Black",
                    "1 Box = 11 U | Breakdown: 20 Box x 495 mts + 7 U x 45 mts",
                    "Velcro Industrial",
                    180.00,
                    8.0,
                    15.00,
                    "2030-12-31",
                    5,
                    "Zone A - Rack 02",
                    1500,
                    2000,
                    10215.0,
                    "20 Box x 495 mts (9,900m) + 7 U x 45 mts (315m). Total: 10,215 mts.",
                    "Large roof closure and insulating curtains.",
                    "Barcelona, Spain",
                    "LOT-VEL758-26",
                    "Non-hazardous",
                ),
                (
                    "🧪",
                    "MAT-GLU-TUN400",
                    "Glue Tunsan 400 ml",
                    "1 Box = 28 U | Breakdown: 23 Box x 28 U + 10 U loose",
                    "Tunsan Chemical Supplies",
                    8.50,
                    5.0,
                    0.80,
                    "2027-06-30",
                    4,
                    "Zone B - Shelf 01",
                    150,
                    100,
                    654.0,
                    "23 Box x 28 U (644 U) + 10 U loose. Total: 654 U.",
                    "Fast housing sealing and light adhesion.",
                    "Valencia, Spain",
                    "LOT-TUN-400-A",
                    "Class 3 Flammable",
                ),
                (
                    "🧪",
                    "MAT-GLU-HUI600",
                    "Glue HUITIAN 600 ml",
                    "1 Box = 20 U | Total Stock: 15 U loose",
                    "Huitian Adhesives",
                    12.00,
                    0.0,
                    1.20,
                    "2027-04-15",
                    7,
                    "Zone B - Shelf 02",
                    30,
                    20,
                    15.0,
                    "15 U loose. Total: 15 U.",
                    "Elastic and industrial partition sealing.",
                    "Hubei, China",
                    "LOT-HUI-600X",
                    "Irritante",
                ),
                (
                    "🧪",
                    "MAT-GLU-DOW600",
                    "Glue DOW 600 ml",
                    "1 Box = 20 U | Total Stock: 13 U loose",
                    "Dow Chemical Europe",
                    16.50,
                    10.0,
                    1.50,
                    "2027-09-30",
                    3,
                    "Zone B - Shelf 03",
                    40,
                    25,
                    13.0,
                    "13 U loose. Total: 13 U.",
                    "Glass and structural joint sealing.",
                    "Wiesbaden, Germany",
                    "LOT-DOW-600D",
                    "Low VOC",
                ),
                (
                    "🧪",
                    "MAT-GLU-SEA600",
                    "Glue SEAL 600 ml",
                    "1 Box = 12 U | Breakdown: 3 Box x 12 U",
                    "Seal Industrial Solutions",
                    11.00,
                    0.0,
                    1.00,
                    "2027-08-10",
                    5,
                    "Zone B - Shelf 04",
                    25,
                    15,
                    36.0,
                    "3 Box x 12 U. Total: 36 U.",
                    "Waterproof sealing of frames and moldings.",
                    "Milan, Italy",
                    "LOT-SEA-36X",
                    "Non-hazardous",
                ),
                (
                    "⚡",
                    "MAT-PV-TRAD",
                    "PV TRADICIONAL",
                    "Standard Photovoltaic Module | Total Stock: 23 U",
                    "PV Solar Tech",
                    140.00,
                    12.0,
                    12.00,
                    "2035-12-31",
                    10,
                    "Zone C - Rack PV1",
                    15,
                    10,
                    23.0,
                    "23 U. Total: 23 U.",
                    "Traditional solar roof installation.",
                    "Madrid, Spain",
                    "LOT-PV-TRAD-01",
                    "Electrical",
                ),
                (
                    "⚡",
                    "MAT-PV-560W",
                    "PV 560W",
                    "High-Efficiency PV Panel 560W | Total Stock: 5 U",
                    "PV Solar Tech",
                    210.00,
                    15.0,
                    18.00,
                    "2035-12-31",
                    10,
                    "Zone C - Rack PV2",
                    8,
                    10,
                    5.0,
                    "5 U. Total: 5 U.",
                    "High-density solar power generation.",
                    "Jiangsu, China",
                    "LOT-PV-560W-26",
                    "Electrical",
                ),
                (
                    "📦",
                    "MAT-PV-WGV",
                    "PV White Glue Velcro Vertical",
                    "White PV Panel with Integrated Vertical Velcro | Total Stock: 127 U",
                    "Custom Solar Flex",
                    165.00,
                    10.0,
                    14.00,
                    "2032-12-31",
                    7,
                    "Zone C - Rack PV3",
                    50,
                    30,
                    127.0,
                    "127 U. Total: 127 U.",
                    "Fast vertical photovoltaic mounting on canvas.",
                    "Porto, Portugal",
                    "LOT-PV-WGV-127",
                    "Non-hazardous",
                ),
                (
                    "📦",
                    "MAT-PV-WGH",
                    "PV White Glue Velcro Horizontal",
                    "White PV Panel with Integrated Horizontal Velcro | Total Stock: 2 U",
                    "Custom Solar Flex",
                    165.00,
                    10.0,
                    14.00,
                    "2032-12-31",
                    7,
                    "Zone C - Rack PV3",
                    20,
                    15,
                    2.0,
                    "2 U. Total: 2 U.",
                    "Fast horizontal photovoltaic mounting.",
                    "Porto, Portugal",
                    "LOT-PV-WGH-02",
                    "Non-hazardous",
                ),
                (
                    "📦",
                    "MAT-PV-WHITE",
                    "PV White",
                    "Standard White Flex PV Module | Total Stock: 110 U",
                    "Custom Solar Flex",
                    150.00,
                    8.0,
                    12.00,
                    "2032-12-31",
                    6,
                    "Zone C - Rack PV4",
                    40,
                    25,
                    110.0,
                    "110 U. Total: 110 U.",
                    "Architectural white photovoltaic integration.",
                    "Porto, Portugal",
                    "LOT-PV-W-110",
                    "Non-hazardous",
                ),
                (
                    "📦",
                    "MAT-PV-BLACK",
                    "PV Black",
                    "Full Black Flex PV Module | Total Stock: 32 U",
                    "Custom Solar Flex",
                    155.00,
                    8.0,
                    12.00,
                    "2032-12-31",
                    6,
                    "Zone C - Rack PV4",
                    30,
                    20,
                    32.0,
                    "32 U. Total: 32 U.",
                    "Aesthetic Full Black installation on dark surfaces.",
                    "Porto, Portugal",
                    "LOT-PV-B-32",
                    "Non-hazardous",
                ),
            ]

            for item in seed_data:
                cursor.execute(
                    """
                    INSERT INTO products (
                        icon, sku, product, characteristics, suppliers, price, discount, transport_price,
                        expiring_date, delivery_time, ubication, monthly_usage, min_stock, quantity,
                        description, where_used, source_origin, batch_lot, sds_hazard_class
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                    item,
                )

                prod_id = cursor.lastrowid
                cursor.execute(
                    """
                    INSERT INTO stock_entries (product_id, entry_date, quantity, unit, movement_type, price, note)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                """,
                    (
                        prod_id,
                        datetime.now().strftime("%Y-%m-%d"),
                        item[13],
                        "pieces",
                        "IN",
                        item[5],
                        "Initial Batch",
                    ),
                )

        conn.commit()


init_db()


def calc_landed_cost(price, discount, transport):
    price = price or 0.0
    discount = discount or 0.0
    transport = transport or 0.0
    return round((price * (1 - discount / 100.0)) + transport, 2)


def safe_path_exists(path):
    return path is not None and bool(path) and os.path.exists(str(path)) and not os.path.isdir(str(path))


def safe_float(value, default=0.0):
    try:
        result = float(value)
        return default if pd.isna(result) else result
    except (TypeError, ValueError):
        return default


def safe_int(value, default=0):
    return int(safe_float(value, default))


# -----------------------------------------------------------------------------
# ADVANCED MULTI-SHEET EXCEL EXPORT & IMPORT
# -----------------------------------------------------------------------------
def export_database_to_multi_sheet_excel():
    conn = get_db_connection()
    try:
        products_df = pd.read_sql_query("SELECT * FROM products", conn)
        entries_df = pd.read_sql_query("SELECT * FROM stock_entries", conn)
    finally:
        conn.close()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        products_df.to_excel(
            writer, sheet_name="Master_Inventory", index=False
        )

        for _, prod in products_df.iterrows():
            prod_entries = entries_df[entries_df["product_id"] == prod["id"]]
            sheet_title = sanitize_filename(f"{prod['product']}")

            if prod_entries.empty:
                prod_entries = pd.DataFrame(
                    columns=["entry_date", "quantity", "price", "note"]
                )
            else:
                prod_entries = prod_entries[
                    ["entry_date", "quantity", "price", "note"]
                ]

            prod_entries.to_excel(
                writer, sheet_name=sheet_title, index=False
            )

    return output.getvalue()


def import_excel_to_database(uploaded_file):
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        if "Master_Inventory" in excel_file.sheet_names:
            p_df = pd.read_excel(excel_file, sheet_name="Master_Inventory")
            conn = get_db_connection()
            try:
                conn.execute("DELETE FROM products")
                conn.execute("DELETE FROM stock_entries")

                p_df.to_sql("products", conn, if_exists="append", index=False)

                db_prods = pd.read_sql_query(
                    "SELECT id, product FROM products", conn
                )

                for _, prod_row in db_prods.iterrows():
                    sheet_title = sanitize_filename(f"{prod_row['product']}")
                    if sheet_title in excel_file.sheet_names:
                        e_df = pd.read_excel(
                            excel_file, sheet_name=sheet_title
                        )
                        if not e_df.empty:
                            e_df["product_id"] = prod_row["id"]
                            e_df.to_sql(
                                "stock_entries",
                                conn,
                                if_exists="append",
                                index=False,
                            )
                conn.commit()
            finally:
                conn.close()

        return True, "Multi-sheet database successfully imported!"
    except Exception as err:
        return False, f"Import error: {str(err)}"


# -----------------------------------------------------------------------------
# TOP HEADER & LANGUAGE SWITCHER
# -----------------------------------------------------------------------------
if "current_lang" not in st.session_state:
    st.session_state["current_lang"] = "en"

col_header, col_en, col_es = st.columns([4, 1, 1])

with col_en:
    btn_type_en = (
        "primary" if st.session_state["current_lang"] == "en" else "secondary"
    )
    if st.button("🇬🇧 English", key="lang_btn_en", type=btn_type_en):
        st.session_state["current_lang"] = "en"
        if "selected_product_id" in st.session_state:
            del st.session_state["selected_product_id"]
        st.rerun()

with col_es:
    btn_type_es = (
        "primary" if st.session_state["current_lang"] == "es" else "secondary"
    )
    if st.button("🇪🇸 Español", key="lang_btn_es", type=btn_type_es):
        st.session_state["current_lang"] = "es"
        if "selected_product_id" in st.session_state:
            del st.session_state["selected_product_id"]
        st.rerun()

es = st.session_state["current_lang"] == "es"

txt = {
    "title": (
        "📦 INVENTARIO DE MATERIALES" if es else "📦 MATERIAL INVENTORY SYSTEM"
    ),
    "search_lbl": "🔍 Búsqueda Universal" if es else "🔍 Universal Search",
    "search_ph": (
        "Buscar por Velcro, Tunsan, Dow, PV, Zona A..."
        if es
        else "Search by Velcro, Tunsan, Dow, PV, Zone A..."
    ),
    "sort_lbl": "Orden / Ranking" if es else "Sort / Ranking",
    "tab1": "🎴 Tarjetas / Vista Modal" if es else "🎴 Cards / Modal View",
    "tab2": "📊 Tabla Master Excel" if es else "📊 Master Excel Table",
    "add_btn": "➕ Añadir Nuevo Producto" if es else "➕ Add New Product",
    "save_btn": "💾 Guardar Cambios" if es else "💾 Save Changes",
    "delete_btn": "🗑️ Eliminar Producto" if es else "🗑️ Delete Product",
    "landed": "Coste Final Net" if es else "Landed Cost",
    "low_stock": "🚨 STOCK BAJO" if es else "🚨 LOW STOCK",
    "ok_stock": "🟢 STOCK OK" if es else "🟢 STOCK OK",
    "entries_sec": (
        "📅 Historial de Entradas (Múltiples Registros)"
        if es
        else "📅 Stock Entry History (Multiple Records)"
    ),
    "export_btn": (
        "📥 Descargar Excel (Con Hojas por Producto)"
        if es
        else "📥 Download Excel (With Sheet Per Product)"
    ),
    "import_btn": (
        "📤 Cargar Excel Completo" if es else "📤 Upload Multi-Sheet Excel"
    ),
}

with col_header:
    st.title(txt["title"])


# -----------------------------------------------------------------------------
# DATA LOAD & FILTERING
# -----------------------------------------------------------------------------
def load_products():
    conn = get_db_connection()
    try:
        df = pd.read_sql_query("SELECT * FROM products", conn)
        contacts_df = pd.read_sql_query(
            "SELECT product_id, name, surname FROM contacts", conn
        )
    finally:
        conn.close()

    df["landed_cost"] = df.apply(
        lambda r: calc_landed_cost(r["price"], r["discount"], r["transport_price"]),
        axis=1,
    )

    contacts_summary = (
        contacts_df.assign(
            full_name=lambda d: (d["name"].fillna("") + " " + d["surname"].fillna("")).str.strip()
        )
        .groupby("product_id")["full_name"]
        .apply(lambda names: ", ".join(n for n in names if n))
    )
    df["contacts"] = df["id"].map(contacts_summary).fillna("")

    return df


df_products = load_products()

col_search, col_sort = st.columns([3, 1])

with col_search:
    search_query = st.text_input(
        txt["search_lbl"], placeholder=txt["search_ph"]
    ).lower()

with col_sort:
    sort_option = st.selectbox(
        txt["sort_lbl"],
        options=[
            "Default",
            "Low Stock First",
            "High Stock First",
            "Landed Cost: Low to High",
            "Landed Cost: High to Low",
            "Name: A-Z",
        ],
    )

filtered_df = df_products.copy()

if search_query:
    filtered_df = filtered_df[
        filtered_df.apply(
            lambda row: search_query in row.astype(str).str.lower().str.cat(sep=" "),
            axis=1,
        )
    ]

if sort_option == "Low Stock First":
    filtered_df["stock_diff"] = filtered_df["quantity"] - filtered_df["min_stock"]
    filtered_df = filtered_df.sort_values("stock_diff", ascending=True)
elif sort_option == "High Stock First":
    filtered_df = filtered_df.sort_values("quantity", ascending=False)
elif sort_option == "Landed Cost: Low to High":
    filtered_df = filtered_df.sort_values("landed_cost", ascending=True)
elif sort_option == "Landed Cost: High to Low":
    filtered_df = filtered_df.sort_values("landed_cost", ascending=False)
elif sort_option == "Name: A-Z":
    filtered_df = filtered_df.sort_values("product", ascending=True)

# -----------------------------------------------------------------------------
# MAIN APP TABS
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs([txt["tab1"], txt["tab2"]])

# TAB 1: CARDS VIEW
with tab1:
    # INSTANT ADD PRODUCT WITH AUTO-MODAL OPENING
    if st.button(txt["add_btn"], type="primary"):
        new_timestamp = int(datetime.now().timestamp())
        new_sku = f"MAT-NEW-{new_timestamp}"
        new_prod_name = f"New Material {new_timestamp % 10000}" if not es else f"Nuevo Material {new_timestamp % 10000}"
        
        new_id = execute_db_query(
            """
            INSERT INTO products (sku, product, quantity, min_stock, price, ubication)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (new_sku, new_prod_name, 10.0, 5.0, 10.00, "Zone A"),
        )
        
        execute_db_query(
            """
            INSERT INTO stock_entries (product_id, entry_date, quantity, price, note)
            VALUES (?, ?, ?, ?, ?)
            """,
            (new_id, datetime.now().strftime("%Y-%m-%d"), 10.0, 10.00, "Initial Record"),
        )
        
        # Set session state to immediately open modal for new item
        st.session_state["selected_product_id"] = new_id
        st.toast(f"Added product #{new_id}!", icon="✅")
        st.rerun()

    st.write("---")

    cols = st.columns(3)
    for idx, row in filtered_df.reset_index().iterrows():
        col = cols[idx % 3]
        with col:
            is_low = row["quantity"] <= row["min_stock"]
            badge = txt["low_stock"] if is_low else txt["ok_stock"]

            with st.container(border=True):
                if safe_path_exists(row["photo_path"]):
                    st.image(row["photo_path"], width="stretch")
                else:
                    st.markdown(
                        f"<div style='height:120px; background:#1e293b; color:#94a3b8; display:flex; align-items:center; justify-content:center; border-radius:8px; font-weight:bold; font-size:24px;'>{row['icon']} {row['product'][:15]}</div>",
                        unsafe_allow_html=True,
                    )

                st.subheader(f"{row['icon']} {row['product']}")
                st.caption(f"SKU: {row['sku']} | {badge}")

                st.write(
                    f"📍 **{'Ubicación' if es else 'Location'}:** {row['ubication']}"
                )
                st.write(
                    f"📦 **{'Cantidad Total' if es else 'Total Qty'}:** `{row['quantity']}`"
                )
                st.write(f"💶 **{txt['landed']}:** `€{row['landed_cost']}`")

                if st.button(
                    f"📄 {'Ficha & Entradas' if es else 'Details & History'}",
                    key=f"card_btn_{row['id']}",
                ):
                    st.session_state["selected_product_id"] = row["id"]

# TAB 2: MASTER GRID VIEW & EXCEL EXPANDER
with tab2:
    with st.expander("📂 Import / Export Multi-Sheet Excel Workbook"):
        col_ex_b, col_im_b = st.columns(2)

        with col_ex_b:
            excel_bytes = export_database_to_multi_sheet_excel()
            st.download_button(
                label=txt["export_btn"],
                data=excel_bytes,
                file_name=f"full_inventory_with_product_sheets_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

        with col_im_b:
            up_file = st.file_uploader(
                txt["import_btn"], type=["xlsx"], key="excel_uploader_tab2"
            )
            if up_file is not None:
                if st.button("Apply Excel Import", type="primary"):
                    ok, msg = import_excel_to_database(up_file)
                    if ok:
                        st.toast(msg, icon="✅")
                        st.rerun()
                    else:
                        st.error(msg)

    st.markdown("### 📊 Master Editable Grid")

    edited_df = st.data_editor(
        filtered_df[
            [
                "id",
                "icon",
                "sku",
                "product",
                "characteristics",
                "suppliers",
                "delivery_time",
                "delivery_time_unit",
                "price",
                "discount",
                "transport_price",
                "landed_cost",
                "quantity",
                "unit",
                "min_stock",
                "ubication",
                "contacts",
            ]
        ],
        disabled=["id", "sku", "landed_cost", "contacts"],
        column_config={
            "delivery_time_unit": st.column_config.SelectboxColumn(
                "Delivery Unit", options=["days", "weeks"]
            ),
            "unit": st.column_config.SelectboxColumn(
                "Stock Unit", options=["m", "kg", "rolls", "pieces", "boxes"]
            ),
            "contacts": st.column_config.TextColumn(
                "Contacts", help="Edit contacts from the product's Details & History modal"
            ),
        },
        width="stretch",
        hide_index=True,
    )

    if st.button("💾 Apply Grid Edits to Database"):
        for idx, r in edited_df.iterrows():
            execute_db_query(
                """
                UPDATE products SET
                    icon = ?, product = ?, characteristics = ?, suppliers = ?,
                    delivery_time = ?, delivery_time_unit = ?,
                    price = ?, discount = ?, transport_price = ?,
                    quantity = ?, unit = ?, min_stock = ?, ubication = ?
                WHERE id = ?
                """,
                (
                    r["icon"],
                    r["product"],
                    r["characteristics"],
                    r["suppliers"],
                    safe_int(r["delivery_time"]),
                    r["delivery_time_unit"],
                    safe_float(r["price"]),
                    safe_float(r["discount"]),
                    safe_float(r["transport_price"]),
                    safe_float(r["quantity"]),
                    r["unit"],
                    safe_int(r["min_stock"]),
                    r["ubication"],
                    r["id"],
                ),
            )
        st.toast("Database inventory.db updated successfully!", icon="✅")
        st.rerun()

# -----------------------------------------------------------------------------
# DETAILED PRODUCT MODAL
# -----------------------------------------------------------------------------
if "selected_product_id" in st.session_state:
    p_id = st.session_state["selected_product_id"]

    conn = get_db_connection()
    try:
        product = conn.execute(
            "SELECT * FROM products WHERE id = ?", (p_id,)
        ).fetchone()
    finally:
        conn.close()

    if product:

        @st.dialog(
            f"{product['icon']} {product['product']}",
            width="large",
        )
        def product_modal():
            st.caption(
                f"SKU: {product['sku']} | Ubicación: {product['ubication']}"
            )

            st.markdown("### 🖼️ Photo & Technical Datasheet Upload")
            col_img, col_pdf = st.columns(2)

            with col_img:
                if safe_path_exists(product["photo_path"]):
                    st.image(
                        product["photo_path"],
                        caption="Current Photo",
                        width="stretch",
                    )

                uploaded_img = st.file_uploader(
                    "Select Photo (PNG/JPG)", type=["png", "jpg", "jpeg"]
                )

            with col_pdf:
                if safe_path_exists(product["datasheet_path"]):
                    st.success("🟢 Technical Datasheet Attached")
                    with open(product["datasheet_path"], "rb") as pdf_file:
                        st.download_button(
                            "📥 Download Datasheet PDF",
                            pdf_file,
                            file_name=f"{sanitize_filename(product['product'])}_{product['sku']}_datasheet.pdf",
                        )

                uploaded_pdf = st.file_uploader(
                    "Select Datasheet (PDF)", type=["pdf"]
                )

            st.write("---")

            # Stock Entries History
            st.markdown(f"### {txt['entries_sec']}")

            conn = get_db_connection()
            try:
                entries = conn.execute(
                    "SELECT * FROM stock_entries WHERE product_id = ? ORDER BY entry_date DESC",
                    (p_id,),
                ).fetchall()
            finally:
                conn.close()

            def signed_qty(entry):
                magnitude = safe_float(entry["quantity"])
                return -magnitude if entry["movement_type"] == "OUT" else magnitude

            if entries:
                for entry in entries:
                    e_col1, e_col2, e_col3, e_col4, e_col5 = st.columns(
                        [2, 2, 2, 3, 1]
                    )
                    move_icon = "➖" if entry["movement_type"] == "OUT" else "➕"
                    e_col1.write(f"📅 {entry['entry_date']}")
                    e_col2.write(
                        f"{move_icon} `{entry['quantity']} {entry['unit'] or 'pieces'}`"
                    )
                    e_col3.write(f"💶 `€{entry['price']:.2f}`")
                    e_col4.caption(f"{entry['note'] or '-'}")

                    if e_col5.button("🗑️", key=f"del_entry_{entry['id']}"):
                        execute_db_query(
                            "DELETE FROM stock_entries WHERE id = ?", (entry["id"],)
                        )

                        conn = get_db_connection()
                        try:
                            rem_entries = conn.execute(
                                "SELECT * FROM stock_entries WHERE product_id = ? ORDER BY entry_date DESC",
                                (p_id,),
                            ).fetchall()
                        finally:
                            conn.close()

                        new_total = sum(signed_qty(r) for r in rem_entries)
                        if rem_entries:
                            latest = rem_entries[0]
                            execute_db_query(
                                "UPDATE products SET quantity = ?, unit = ?, price = ? WHERE id = ?",
                                (new_total, latest["unit"], latest["price"], p_id),
                            )
                        else:
                            execute_db_query(
                                "UPDATE products SET quantity = ? WHERE id = ?",
                                (0.0, p_id),
                            )

                        st.toast("Entry removed!", icon="✅")
                        st.rerun()

            with st.expander("➕➖ Register Stock Movement"):
                with st.form(key=f"add_entry_form_{p_id}"):
                    c1, c2, c3, c4 = st.columns(4)
                    e_date = c1.date_input("Date", value=datetime.now())
                    e_move = c2.selectbox(
                        "Movement", options=["➕ Add", "➖ Remove"]
                    )
                    e_qty = c3.number_input(
                        "Quantity", value=0.0, min_value=0.0
                    )
                    unit_options = ["m", "kg", "rolls", "pieces", "boxes"]
                    current_unit = product["unit"] if product["unit"] in unit_options else "pieces"
                    e_unit = c4.selectbox(
                        "Unit", options=unit_options, index=unit_options.index(current_unit)
                    )
                    e_price = st.number_input(
                        "Price (€)", value=safe_float(product["price"])
                    )
                    e_note = st.text_input(
                        "Note / Comment",
                        value="Restock" if e_move == "➕ Add" else "Stock removal",
                    )

                    if st.form_submit_button("Save Movement"):
                        movement_type = "IN" if e_move == "➕ Add" else "OUT"
                        current_qty = safe_float(product["quantity"])

                        if movement_type == "OUT" and e_qty > current_qty:
                            st.error(
                                f"Cannot remove {e_qty} {e_unit} — only {current_qty} {product['unit'] or 'pieces'} in stock."
                            )
                        else:
                            execute_db_query(
                                """
                                INSERT INTO stock_entries (product_id, entry_date, quantity, unit, movement_type, price, note)
                                VALUES (?, ?, ?, ?, ?, ?, ?)
                                """,
                                (
                                    p_id,
                                    e_date.strftime("%Y-%m-%d"),
                                    e_qty,
                                    e_unit,
                                    movement_type,
                                    e_price,
                                    e_note,
                                ),
                            )

                            delta = e_qty if movement_type == "IN" else -e_qty
                            new_total_qty = current_qty + delta
                            execute_db_query(
                                """
                                UPDATE products SET quantity = ?, unit = ?, price = ? WHERE id = ?
                                """,
                                (new_total_qty, e_unit, e_price, p_id),
                            )
                            st.toast("Stock movement recorded!", icon="✅")
                            st.rerun()

            st.write("---")

            st.markdown("### 👤 Contacts")

            conn = get_db_connection()
            try:
                contacts = conn.execute(
                    "SELECT * FROM contacts WHERE product_id = ? ORDER BY id", (p_id,)
                ).fetchall()
            finally:
                conn.close()

            if contacts:
                contacts_df = pd.DataFrame(
                    [dict(c) for c in contacts]
                )[["name", "surname", "phone", "email", "country"]]
                st.dataframe(contacts_df, width="stretch", hide_index=True)

                del_options = {
                    f"{c['name']} {c['surname']} ({c['email'] or 'no email'})": c["id"]
                    for c in contacts
                }
                col_del_sel, col_del_btn = st.columns([3, 1])
                sel_contact_label = col_del_sel.selectbox(
                    "Remove a contact", options=list(del_options.keys()), key=f"contact_del_sel_{p_id}"
                )
                if col_del_btn.button("🗑️ Remove", key=f"contact_del_btn_{p_id}"):
                    execute_db_query(
                        "DELETE FROM contacts WHERE id = ?", (del_options[sel_contact_label],)
                    )
                    st.toast("Contact removed!", icon="✅")
                    st.rerun()
            else:
                st.caption("No contacts registered for this product yet.")

            with st.expander("➕ Add New Contact"):
                with st.form(key=f"add_contact_form_{p_id}"):
                    ct1, ct2 = st.columns(2)
                    c_name = ct1.text_input("Name")
                    c_surname = ct2.text_input("Surname")
                    ct3, ct4 = st.columns(2)
                    c_phone = ct3.text_input("Phone")
                    c_email = ct4.text_input("Email")
                    c_country = st.text_input("Country")

                    if st.form_submit_button("Add Contact"):
                        execute_db_query(
                            """
                            INSERT INTO contacts (product_id, name, surname, phone, email, country)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """,
                            (p_id, c_name, c_surname, c_phone, c_email, c_country),
                        )
                        st.toast("Contact added!", icon="✅")
                        st.rerun()

            st.write("---")

            st.markdown("### 📘 Master Details & Pricing")
            with st.form(key=f"edit_prod_form_{p_id}"):
                col_a, col_b = st.columns(2)
                f_name = col_a.text_input("Product Name", value=product["product"])
                f_icon = col_b.text_input("Icon (Emoji)", value=product["icon"])

                f_desc = st.text_area(
                    "Description / Breakdown", value=product["description"]
                )
                f_char = st.text_area(
                    "Characteristics / Packaging",
                    value=product["characteristics"],
                )
                f_where = st.text_area(
                    "Where Used (Process)", value=product["where_used"]
                )

                col_sup, col_del1, col_del2 = st.columns(3)
                f_supplier = col_sup.text_input(
                    "Provider / Supplier", value=product["suppliers"] or ""
                )
                f_delivery_time = col_del1.number_input(
                    "Delivery Time", value=safe_int(product["delivery_time"]), min_value=0
                )
                delivery_unit_options = ["days", "weeks"]
                current_delivery_unit = (
                    product["delivery_time_unit"]
                    if product["delivery_time_unit"] in delivery_unit_options
                    else "days"
                )
                f_delivery_unit = col_del2.selectbox(
                    "Delivery Time Unit",
                    options=delivery_unit_options,
                    index=delivery_unit_options.index(current_delivery_unit),
                )

                c1, c2, c3 = st.columns(3)
                f_price = c1.number_input(
                    "Base Price (€)", value=safe_float(product["price"])
                )
                f_disc = c2.number_input(
                    "Discount (%)", value=safe_float(product["discount"])
                )
                f_trans = c3.number_input(
                    "Transport Fee (€)", value=safe_float(product["transport_price"])
                )

                c4, c5, c6 = st.columns(3)
                f_qty = c4.number_input(
                    "Total Quantity", value=safe_float(product["quantity"])
                )
                f_min = c5.number_input(
                    "Min Stock Level", value=safe_int(product["min_stock"])
                )
                f_ubic = c6.text_input("Ubication", value=product["ubication"] or "")

                btn_save = st.form_submit_button(
                    txt["save_btn"], type="primary"
                )

                if btn_save:
                    clean_name = sanitize_filename(f_name)
                    clean_sku = sanitize_filename(product["sku"])

                    new_photo_path = product["photo_path"]
                    if uploaded_img:
                        ext = uploaded_img.name.split(".")[-1]
                        filename = f"{clean_name}_{clean_sku}_photo.{ext}"
                        
                        if os.path.exists(IMAGES_DIR) and not os.path.isdir(IMAGES_DIR):
                            os.remove(IMAGES_DIR)
                        os.makedirs(IMAGES_DIR, exist_ok=True)
                        
                        new_photo_path = os.path.join(IMAGES_DIR, filename)
                        with open(new_photo_path, "wb") as f:
                            f.write(uploaded_img.getbuffer())

                    new_pdf_path = product["datasheet_path"]
                    if uploaded_pdf:
                        filename = f"{clean_name}_{clean_sku}_datasheet.pdf"
                        
                        if os.path.exists(DATASHEETS_DIR) and not os.path.isdir(DATASHEETS_DIR):
                            os.remove(DATASHEETS_DIR)
                        os.makedirs(DATASHEETS_DIR, exist_ok=True)
                        
                        new_pdf_path = os.path.join(DATASHEETS_DIR, filename)
                        with open(new_pdf_path, "wb") as f:
                            f.write(uploaded_pdf.getbuffer())

                    execute_db_query(
                        """
                        UPDATE products SET
                            product = ?, icon = ?, description = ?, characteristics = ?,
                            where_used = ?, suppliers = ?, delivery_time = ?, delivery_time_unit = ?,
                            price = ?, discount = ?, transport_price = ?,
                            quantity = ?, min_stock = ?, ubication = ?,
                            photo_path = ?, datasheet_path = ?
                        WHERE id = ?
                        """,
                        (
                            f_name,
                            f_icon,
                            f_desc,
                            f_char,
                            f_where,
                            f_supplier,
                            f_delivery_time,
                            f_delivery_unit,
                            f_price,
                            f_disc,
                            f_trans,
                            f_qty,
                            f_min,
                            f_ubic,
                            new_photo_path,
                            new_pdf_path,
                            p_id,
                        ),
                    )

                    st.toast("Product details saved successfully!", icon="✅")
                    del st.session_state["selected_product_id"]
                    st.rerun()

            confirm_key = f"confirm_delete_{p_id}"
            if not st.session_state.get(confirm_key, False):
                if st.button(txt["delete_btn"], type="secondary"):
                    st.session_state[confirm_key] = True
                    st.rerun()
            else:
                st.warning(
                    "¿Seguro que quieres eliminar este producto? Esto también borrará su historial de entradas."
                    if es
                    else "Are you sure you want to delete this product? This will also remove its stock entry history."
                )
                col_confirm, col_cancel = st.columns(2)
                if col_confirm.button(
                    "✅ " + ("Sí, eliminar" if es else "Yes, delete"),
                    type="primary",
                    key=f"confirm_yes_{p_id}",
                ):
                    execute_db_query("DELETE FROM products WHERE id = ?", (p_id,))
                    del st.session_state[confirm_key]
                    del st.session_state["selected_product_id"]
                    st.toast("Product removed!", icon="✅")
                    st.rerun()
                if col_cancel.button(
                    "Cancelar" if es else "Cancel", key=f"confirm_no_{p_id}"
                ):
                    del st.session_state[confirm_key]
                    st.rerun()

        product_modal()