import io
import os
import re
import sqlite3
from datetime import datetime
import altair as alt
import pandas as pd
import streamlit as st
from openpyxl.chart import BarChart, LineChart, Reference

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


def ensure_column(cursor, table, column, column_def):
    cursor.execute(f"PRAGMA table_info({table})")
    existing_cols = [row[1] for row in cursor.fetchall()]
    if column not in existing_cols:
        cursor.execute(f"ALTER TABLE {table} ADD COLUMN {column} {column_def}")


def init_db():
    with get_db_connection() as conn:
        cursor = conn.cursor()

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

        for col, col_def in [
            ("delivery_time_unit", "TEXT DEFAULT 'days'"),
            ("unit", "TEXT DEFAULT 'pieces'"),
            ("source_origin", "TEXT"),
            ("batch_lot", "TEXT"),
            ("sds_hazard_class", "TEXT"),
            ("where_used", "TEXT"),
        ]:
            ensure_column(cursor, "products", col, col_def)

        for col, col_def in [
            ("unit", "TEXT DEFAULT 'pieces'"),
            ("movement_type", "TEXT DEFAULT 'IN'"),
        ]:
            ensure_column(cursor, "stock_entries", col, col_def)

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
        contacts_df = pd.read_sql_query("SELECT * FROM contacts", conn)
    finally:
        conn.close()

    entry_cols = ["entry_date", "movement_type", "quantity", "unit", "price", "note"]
    contact_cols = ["name", "surname", "phone", "email", "country"]

    if not entries_df.empty:
        entries_df = entries_df.copy()
        entries_df["quantity"] = entries_df.apply(
            lambda r: -abs(r["quantity"]) if r["movement_type"] == "OUT" else abs(r["quantity"]),
            axis=1,
        )

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        products_df.to_excel(
            writer, sheet_name="Master_Inventory", index=False
        )

        contacts_export = contacts_df.merge(
            products_df[["id", "product", "sku"]],
            left_on="product_id",
            right_on="id",
            how="left",
            suffixes=("", "_product"),
        )
        if contacts_export.empty:
            contacts_export = pd.DataFrame(columns=["product", "sku"] + contact_cols)
        else:
            contacts_export = contacts_export[["product", "sku"] + contact_cols]
        contacts_export.to_excel(writer, sheet_name="Contacts", index=False)

        for _, prod in products_df.iterrows():
            prod_entries = entries_df[entries_df["product_id"] == prod["id"]]
            sheet_title = sanitize_filename(f"{prod['product']}")

            if prod_entries.empty:
                prod_entries = pd.DataFrame(columns=entry_cols)
            else:
                prod_entries = prod_entries[entry_cols].sort_values("entry_date", ascending=False)

            prod_entries.to_excel(
                writer, sheet_name=sheet_title, index=False
            )

    return output.getvalue()


def import_excel_to_database(uploaded_file, es=False):
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        if "Master_Inventory" in excel_file.sheet_names:
            p_df = pd.read_excel(excel_file, sheet_name="Master_Inventory")
            conn = get_db_connection()
            try:
                conn.execute("DELETE FROM products")
                conn.execute("DELETE FROM stock_entries")
                conn.execute("DELETE FROM contacts")

                p_df.to_sql("products", conn, if_exists="append", index=False)

                db_prods = pd.read_sql_query(
                    "SELECT id, product, sku FROM products", conn
                )

                for _, prod_row in db_prods.iterrows():
                    sheet_title = sanitize_filename(f"{prod_row['product']}")
                    if sheet_title in excel_file.sheet_names:
                        e_df = pd.read_excel(
                            excel_file, sheet_name=sheet_title
                        )
                        if not e_df.empty:
                            e_df["product_id"] = prod_row["id"]
                            if "quantity" in e_df.columns:
                                e_df["quantity"] = e_df["quantity"].abs()
                            e_df.to_sql(
                                "stock_entries",
                                conn,
                                if_exists="append",
                                index=False,
                            )

                if "Contacts" in excel_file.sheet_names:
                    c_df = pd.read_excel(excel_file, sheet_name="Contacts")
                    if not c_df.empty and "sku" in c_df.columns:
                        c_df = c_df.merge(
                            db_prods[["id", "sku"]], on="sku", how="left"
                        )
                        c_df = c_df.rename(columns={"id": "product_id"})
                        c_df = c_df.dropna(subset=["product_id"])
                        contact_cols = [
                            "product_id", "name", "surname", "phone", "email", "country"
                        ]
                        c_df = c_df[[col for col in contact_cols if col in c_df.columns]]
                        c_df.to_sql(
                            "contacts", conn, if_exists="append", index=False
                        )

                conn.commit()
            finally:
                conn.close()

        return True, (
            "¡Base de datos multi-hoja importada correctamente!"
            if es
            else "Multi-sheet database successfully imported!"
        )
    except Exception as err:
        return False, (
            f"Error de importación: {str(err)}" if es else f"Import error: {str(err)}"
        )


# -----------------------------------------------------------------------------
# PER-PRODUCT STOCK HISTORY & CHARTING
# -----------------------------------------------------------------------------
FREQ_MAP = {"Daily": "D", "Weekly": "W", "Monthly": "ME", "Yearly": "YE"}

ICON_OPTIONS = [
    "📦", "🏷️", "🧪", "⚡", "🔩", "🛠️", "🔧", "🪛", "🧵", "🎨",
    "🧱", "🪵", "🔌", "🧯", "🧰", "🧲", "🛢️", "🧻", "📎", "🪝",
]


def get_stock_history_df(product_id):
    conn = get_db_connection()
    try:
        df = pd.read_sql_query(
            """
            SELECT entry_date, movement_type, quantity, unit, price, note
            FROM stock_entries WHERE product_id = ? ORDER BY entry_date, id
            """,
            conn,
            params=(product_id,),
        )
    finally:
        conn.close()

    if df.empty:
        return df

    df["entry_date"] = pd.to_datetime(df["entry_date"])
    df["signed_qty"] = df.apply(
        lambda r: -abs(r["quantity"]) if r["movement_type"] == "OUT" else abs(r["quantity"]),
        axis=1,
    )
    df["running_stock"] = df["signed_qty"].cumsum()
    return df


def aggregate_stock_history(history_df, freq_label):
    freq = FREQ_MAP[freq_label]
    if history_df.empty:
        return pd.DataFrame(
            columns=["period", "stock_in", "stock_out", "net_change", "running_stock"]
        )

    indexed = history_df.set_index("entry_date")
    stock_in = indexed[indexed["movement_type"] == "IN"]["quantity"].resample(freq).sum()
    stock_out = indexed[indexed["movement_type"] == "OUT"]["quantity"].resample(freq).sum()
    net_change = indexed["signed_qty"].resample(freq).sum()
    running_stock = indexed["running_stock"].resample(freq).last().ffill().fillna(0)

    agg = pd.DataFrame(
        {
            "stock_in": stock_in,
            "stock_out": stock_out,
            "net_change": net_change,
            "running_stock": running_stock,
        }
    ).fillna(0)
    agg.index.name = "period"
    agg = agg.reset_index()

    date_fmt = {"D": "%Y-%m-%d", "W": "%Y-%m-%d", "ME": "%Y-%m", "YE": "%Y"}[freq]
    agg["period"] = agg["period"].dt.strftime(date_fmt)
    return agg


def export_stock_chart_excel(product_name, agg_df, freq_label, es=False):
    sheet_title = "Datos de Stock" if es else "Stock Data"
    if es:
        agg_df = agg_df.rename(
            columns={
                "period": "Período",
                "stock_in": "Entradas de Stock",
                "stock_out": "Salidas de Stock",
                "net_change": "Cambio Neto",
                "running_stock": "Stock Acumulado",
            }
        )
    else:
        agg_df = agg_df.rename(
            columns={
                "period": "Period",
                "stock_in": "Stock In",
                "stock_out": "Stock Out",
                "net_change": "Net Change",
                "running_stock": "Running Stock",
            }
        )
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        agg_df.to_excel(writer, sheet_name=sheet_title, index=False)
        worksheet = writer.sheets[sheet_title]

        n_rows = len(agg_df) + 1

        bar = BarChart()
        movement_word = "Movimiento de Stock" if es else "Stock Movement"
        freq_word = (
            {"Daily": "Diario", "Weekly": "Semanal", "Monthly": "Mensual", "Yearly": "Anual"}.get(
                freq_label, freq_label
            )
            if es
            else freq_label
        )
        bar.title = f"{product_name} — {movement_word} ({freq_word})"
        bar.y_axis.title = "Unidades Movidas" if es else "Units Moved"
        bar.x_axis.title = "Período" if es else "Period"
        in_ref = Reference(worksheet, min_col=2, min_row=1, max_row=n_rows)
        out_ref = Reference(worksheet, min_col=3, min_row=1, max_row=n_rows)
        cats = Reference(worksheet, min_col=1, min_row=2, max_row=n_rows)
        bar.add_data(in_ref, titles_from_data=True)
        bar.add_data(out_ref, titles_from_data=True)
        bar.set_categories(cats)

        line = LineChart()
        run_ref = Reference(worksheet, min_col=5, min_row=1, max_row=n_rows)
        line.add_data(run_ref, titles_from_data=True)
        line.y_axis.axId = 200
        line.y_axis.title = "Stock Acumulado" if es else "Running Stock"
        line.y_axis.crosses = "max"

        bar += line
        bar.width = 26
        bar.height = 12
        worksheet.add_chart(bar, "H2")

    return output.getvalue()


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


def t(en_text, es_text):
    return es_text if es else en_text


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
    "tab3": "📈 Gráficos de Stock" if es else "📈 Stock Graphs",
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
            "SELECT product_id, name, surname, phone, email, country FROM contacts", conn
        )
    finally:
        conn.close()

    df["landed_cost"] = df.apply(
        lambda r: calc_landed_cost(r["price"], r["discount"], r["transport_price"]),
        axis=1,
    )

    def format_contact(row):
        parts = [
            f"{(row['name'] or '').strip()} {(row['surname'] or '').strip()}".strip(),
            row["phone"] or "",
            row["email"] or "",
            row["country"] or "",
        ]
        return " | ".join(p for p in parts if p)

    contacts_summary = (
        contacts_df.assign(contact_line=contacts_df.apply(format_contact, axis=1))
        .groupby("product_id")["contact_line"]
        .apply(lambda lines: "; ".join(l for l in lines if l))
    )
    df["contacts"] = df["id"].map(contacts_summary).fillna("")

    return df


df_products = load_products()

col_search, col_sort = st.columns([3, 1])

with col_search:
    search_query = st.text_input(
        txt["search_lbl"], placeholder=txt["search_ph"]
    ).lower()

SORT_OPTIONS = {
    "default": t("Default", "Por Defecto"),
    "low_stock": t("Low Stock First", "Menor Stock Primero"),
    "high_stock": t("High Stock First", "Mayor Stock Primero"),
    "cost_low_high": t("Landed Cost: Low to High", "Coste Final: Menor a Mayor"),
    "cost_high_low": t("Landed Cost: High to Low", "Coste Final: Mayor a Menor"),
    "name_az": t("Name: A-Z", "Nombre: A-Z"),
    "expiry_soonest": t("Expiry: Soonest First", "Caducidad: Más Próxima Primero"),
    "expiry_latest": t("Expiry: Latest First", "Caducidad: Más Lejana Primero"),
}

with col_sort:
    sort_option = st.selectbox(
        txt["sort_lbl"],
        options=list(SORT_OPTIONS.keys()),
        format_func=lambda k: SORT_OPTIONS[k],
    )

filtered_df = df_products.copy()

if search_query:
    filtered_df = filtered_df[
        filtered_df.apply(
            lambda row: search_query in row.astype(str).str.lower().str.cat(sep=" "),
            axis=1,
        )
    ]

if sort_option == "low_stock":
    filtered_df["stock_diff"] = filtered_df["quantity"] - filtered_df["min_stock"]
    filtered_df = filtered_df.sort_values("stock_diff", ascending=True)
elif sort_option == "high_stock":
    filtered_df = filtered_df.sort_values("quantity", ascending=False)
elif sort_option == "cost_low_high":
    filtered_df = filtered_df.sort_values("landed_cost", ascending=True)
elif sort_option == "cost_high_low":
    filtered_df = filtered_df.sort_values("landed_cost", ascending=False)
elif sort_option == "name_az":
    filtered_df = filtered_df.sort_values("product", ascending=True)
elif sort_option in ("expiry_soonest", "expiry_latest"):
    filtered_df["_expiry_parsed"] = pd.to_datetime(filtered_df["expiring_date"], errors="coerce")
    filtered_df = filtered_df.sort_values(
        "_expiry_parsed", ascending=(sort_option == "expiry_soonest"), na_position="last"
    ).drop(columns="_expiry_parsed")

# -----------------------------------------------------------------------------
# MAIN APP TABS
# -----------------------------------------------------------------------------
tab1, tab2, tab3 = st.tabs([txt["tab1"], txt["tab2"], txt["tab3"]])

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
        st.toast(t(f"Added product #{new_id}!", f"¡Producto #{new_id} añadido!"), icon="✅")
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
    with st.expander(
        t("📂 Import / Export Multi-Sheet Excel Workbook", "📂 Importar / Exportar Libro Excel Multi-Hoja")
    ):
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
                if st.button(t("Apply Excel Import", "Aplicar Importación de Excel"), type="primary"):
                    ok, msg = import_excel_to_database(up_file, es)
                    if ok:
                        st.toast(msg, icon="✅")
                        st.rerun()
                    else:
                        st.error(msg)

    st.markdown(f"### {t('📊 Master Editable Grid', '📊 Cuadrícula Editable Maestra')}")

    edited_df = st.data_editor(
        filtered_df[
            [
                "id",
                "icon",
                "sku",
                "product",
                "description",
                "characteristics",
                "where_used",
                "suppliers",
                "source_origin",
                "batch_lot",
                "sds_hazard_class",
                "delivery_time",
                "delivery_time_unit",
                "expiring_date",
                "price",
                "discount",
                "transport_price",
                "landed_cost",
                "quantity",
                "unit",
                "monthly_usage",
                "min_stock",
                "ubication",
                "contacts",
            ]
        ],
        disabled=["id", "sku", "landed_cost", "contacts"],
        column_config={
            "icon": st.column_config.Column(t("Icon", "Icono")),
            "sku": st.column_config.Column("SKU"),
            "product": st.column_config.Column(t("Product", "Producto")),
            "description": st.column_config.Column(t("Description", "Descripción")),
            "characteristics": st.column_config.Column(t("Characteristics", "Características")),
            "where_used": st.column_config.Column(t("Where Used", "Dónde se Usa")),
            "suppliers": st.column_config.Column(t("Supplier", "Proveedor")),
            "source_origin": st.column_config.Column(t("Source / Origin", "Origen")),
            "batch_lot": st.column_config.Column(t("Batch / Lot", "Lote")),
            "sds_hazard_class": st.column_config.Column(t("Hazard Class", "Clase de Riesgo")),
            "delivery_time": st.column_config.Column(t("Delivery Time", "Tiempo de Entrega")),
            "delivery_time_unit": st.column_config.SelectboxColumn(
                t("Delivery Unit", "Unidad de Entrega"), options=["days", "weeks"]
            ),
            "expiring_date": st.column_config.Column(t("Expiry Date", "Fecha de Caducidad")),
            "price": st.column_config.Column(t("Price (€)", "Precio (€)")),
            "discount": st.column_config.Column(t("Discount (%)", "Descuento (%)")),
            "transport_price": st.column_config.Column(t("Transport Fee (€)", "Transporte (€)")),
            "landed_cost": st.column_config.Column(t("Landed Cost", "Coste Final")),
            "quantity": st.column_config.Column(t("Quantity", "Cantidad")),
            "unit": st.column_config.SelectboxColumn(
                t("Stock Unit", "Unidad de Stock"), options=["m", "kg", "rolls", "pieces", "boxes"]
            ),
            "monthly_usage": st.column_config.Column(t("Monthly Usage", "Uso Mensual")),
            "min_stock": st.column_config.Column(t("Min Stock", "Stock Mínimo")),
            "ubication": st.column_config.Column(t("Location", "Ubicación")),
            "contacts": st.column_config.TextColumn(
                t(
                    "Contacts (Name | Phone | Email | Country)",
                    "Contactos (Nombre | Teléfono | Email | País)",
                ),
                help=t(
                    "Edit contacts from the product's Details & History modal",
                    "Edita los contactos desde el modal Ficha y Entradas del producto",
                ),
                width="large",
            ),
        },
        width="stretch",
        hide_index=True,
    )

    if st.button(t("💾 Apply Grid Edits to Database", "💾 Aplicar Cambios a la Base de Datos")):
        for idx, r in edited_df.iterrows():
            execute_db_query(
                """
                UPDATE products SET
                    icon = ?, product = ?, description = ?, characteristics = ?, where_used = ?,
                    suppliers = ?, source_origin = ?, batch_lot = ?, sds_hazard_class = ?,
                    delivery_time = ?, delivery_time_unit = ?, expiring_date = ?,
                    price = ?, discount = ?, transport_price = ?,
                    quantity = ?, unit = ?, monthly_usage = ?, min_stock = ?, ubication = ?
                WHERE id = ?
                """,
                (
                    r["icon"],
                    r["product"],
                    r["description"],
                    r["characteristics"],
                    r["where_used"],
                    r["suppliers"],
                    r["source_origin"],
                    r["batch_lot"],
                    r["sds_hazard_class"],
                    safe_int(r["delivery_time"]),
                    r["delivery_time_unit"],
                    r["expiring_date"],
                    safe_float(r["price"]),
                    safe_float(r["discount"]),
                    safe_float(r["transport_price"]),
                    safe_float(r["quantity"]),
                    r["unit"],
                    safe_int(r["monthly_usage"]),
                    safe_int(r["min_stock"]),
                    r["ubication"],
                    r["id"],
                ),
            )
        st.toast(
            t("Database inventory.db updated successfully!", "¡Base de datos inventory.db actualizada correctamente!"),
            icon="✅",
        )
        st.rerun()

# TAB 3: PER-PRODUCT STOCK GRAPHS
with tab3:
    st.markdown(f"### {txt['tab3']}")

    if df_products.empty:
        st.info(t("No products available.", "No hay productos disponibles."))
    else:
        product_options = {
            f"{row['icon']} {row['product']} ({row['sku']})": row["id"]
            for _, row in df_products.sort_values("product").iterrows()
        }

        FREQ_LABELS = {
            "Daily": t("Daily", "Diario"),
            "Weekly": t("Weekly", "Semanal"),
            "Monthly": t("Monthly", "Mensual"),
            "Yearly": t("Yearly", "Anual"),
        }

        col_prod, col_freq = st.columns([2, 1])
        selected_label = col_prod.selectbox(
            t("Select Product", "Seleccionar Producto"),
            options=list(product_options.keys()),
            key="graph_product_select",
        )
        freq_label = col_freq.selectbox(
            t("Granularity", "Granularidad"),
            options=list(FREQ_LABELS.keys()),
            format_func=lambda k: FREQ_LABELS[k],
            index=2,
            key="graph_freq_select",
        )

        selected_pid = product_options[selected_label]
        history_df = get_stock_history_df(selected_pid)

        if history_df.empty:
            st.info(
                t(
                    "No stock entries recorded for this product yet.",
                    "Este producto aún no tiene entradas de stock registradas.",
                )
            )
        else:
            agg_df = aggregate_stock_history(history_df, freq_label)

            movement_labels = {"stock_in": t("In", "Entrada"), "stock_out": t("Out", "Salida")}

            chart_df = agg_df.melt(
                id_vars=["period", "running_stock"],
                value_vars=["stock_in", "stock_out"],
                var_name="movement",
                value_name="units",
            )
            chart_df["movement"] = chart_df["movement"].map(movement_labels)

            bar_layer = (
                alt.Chart(chart_df)
                .mark_bar()
                .encode(
                    x=alt.X("period:N", title=t("Period", "Período"), sort=None),
                    y=alt.Y("units:Q", title=t("Units Moved", "Unidades Movidas")),
                    color=alt.Color(
                        "movement:N",
                        scale=alt.Scale(
                            domain=list(movement_labels.values()), range=["#22c55e", "#ef4444"]
                        ),
                        legend=alt.Legend(title=t("Movement", "Movimiento")),
                    ),
                    xOffset="movement:N",
                    tooltip=["period", "movement", "units"],
                )
            )

            line_layer = (
                alt.Chart(agg_df)
                .mark_line(point=True, color="#3b82f6")
                .encode(
                    x=alt.X("period:N", sort=None),
                    y=alt.Y("running_stock:Q", title=t("Running Stock", "Stock Acumulado")),
                    tooltip=["period", "running_stock"],
                )
            )

            combo_chart = (
                alt.layer(bar_layer, line_layer)
                .resolve_scale(y="independent")
                .properties(height=420)
            )
            st.altair_chart(combo_chart, width="stretch")

            st.markdown("#### 📋 " + t("Data Table", "Tabla de Datos"))
            st.dataframe(
                agg_df.rename(
                    columns={
                        "period": t("Period", "Período"),
                        "stock_in": t("Stock In", "Entradas de Stock"),
                        "stock_out": t("Stock Out", "Salidas de Stock"),
                        "net_change": t("Net Change", "Cambio Neto"),
                        "running_stock": t("Running Stock", "Stock Acumulado"),
                    }
                ),
                width="stretch",
                hide_index=True,
            )

            chart_excel_bytes = export_stock_chart_excel(selected_label, agg_df, freq_label, es)
            st.download_button(
                label="📥 Download Chart + Table (Excel)"
                if not es
                else "📥 Descargar Gráfico + Tabla (Excel)",
                data=chart_excel_bytes,
                file_name=f"stock_chart_{sanitize_filename(selected_label)}_{freq_label.lower()}_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                width="stretch",
            )

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
                f"SKU: {product['sku']} | {t('Location', 'Ubicación')}: {product['ubication']}"
            )

            st.markdown(f"### 🖼️ {t('Photo & Technical Datasheet Upload', 'Foto y Ficha Técnica (Datasheet)')}")
            col_img, col_pdf = st.columns(2)

            with col_img:
                if safe_path_exists(product["photo_path"]):
                    st.image(
                        product["photo_path"],
                        caption=t("Current Photo", "Foto Actual"),
                        width="stretch",
                    )

                uploaded_img = st.file_uploader(
                    t("Select Photo (PNG/JPG)", "Seleccionar Foto (PNG/JPG)"),
                    type=["png", "jpg", "jpeg"],
                )

            with col_pdf:
                if safe_path_exists(product["datasheet_path"]):
                    st.success("🟢 " + t("Technical Datasheet Attached", "Ficha Técnica Adjunta"))
                    with open(product["datasheet_path"], "rb") as pdf_file:
                        st.download_button(
                            "📥 " + t("Download Datasheet PDF", "Descargar Ficha Técnica PDF"),
                            pdf_file,
                            file_name=f"{sanitize_filename(product['product'])}_{product['sku']}_datasheet.pdf",
                        )

                uploaded_pdf = st.file_uploader(
                    t("Select Datasheet (PDF)", "Seleccionar Ficha Técnica (PDF)"),
                    type=["pdf"],
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

                        st.toast(t("Entry removed!", "¡Entrada eliminada!"), icon="✅")
                        st.rerun()

            with st.expander("➕➖ " + t("Register Stock Movement", "Registrar Movimiento de Stock")):
                with st.form(key=f"add_entry_form_{p_id}"):
                    c1, c2, c3, c4 = st.columns(4)
                    e_date = c1.date_input(t("Date", "Fecha"), value=datetime.now())
                    e_move = c2.selectbox(
                        t("Movement", "Movimiento"),
                        options=["IN", "OUT"],
                        format_func=lambda v: ("➕ " + t("Add", "Añadir")) if v == "IN" else ("➖ " + t("Remove", "Quitar")),
                    )
                    e_qty = c3.number_input(
                        t("Quantity", "Cantidad"), value=0.0, min_value=0.0
                    )
                    unit_options = ["m", "kg", "rolls", "pieces", "boxes"]
                    current_unit = product["unit"] if product["unit"] in unit_options else "pieces"
                    e_unit = c4.selectbox(
                        t("Unit", "Unidad"), options=unit_options, index=unit_options.index(current_unit)
                    )
                    e_price = st.number_input(
                        t("Price (€)", "Precio (€)"), value=safe_float(product["price"])
                    )
                    e_note = st.text_input(
                        t("Note / Comment", "Nota / Comentario"),
                        value=t("Restock", "Reposición") if e_move == "IN" else t("Stock removal", "Retiro de stock"),
                    )

                    if st.form_submit_button(t("Save Movement", "Guardar Movimiento")):
                        movement_type = e_move
                        current_qty = safe_float(product["quantity"])

                        if movement_type == "OUT" and e_qty > current_qty:
                            st.error(
                                t(
                                    f"Cannot remove {e_qty} {e_unit} — only {current_qty} {product['unit'] or 'pieces'} in stock.",
                                    f"No se puede quitar {e_qty} {e_unit} — solo hay {current_qty} {product['unit'] or 'pieces'} en stock.",
                                )
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
                            st.toast(t("Stock movement recorded!", "¡Movimiento de stock registrado!"), icon="✅")
                            st.rerun()

            st.write("---")

            st.markdown("### 👤 " + t("Contacts", "Contactos"))

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
                )[["name", "surname", "phone", "email", "country"]].rename(
                    columns={
                        "name": t("Name", "Nombre"),
                        "surname": t("Surname", "Apellido"),
                        "phone": t("Phone", "Teléfono"),
                        "email": t("Email", "Correo"),
                        "country": t("Country", "País"),
                    }
                )
                st.dataframe(contacts_df, width="stretch", hide_index=True)

                del_options = {
                    f"{c['name']} {c['surname']} ({c['email'] or t('no email', 'sin correo')})": c["id"]
                    for c in contacts
                }
                col_del_sel, col_del_btn = st.columns([3, 1])
                sel_contact_label = col_del_sel.selectbox(
                    t("Remove a contact", "Eliminar un contacto"),
                    options=list(del_options.keys()),
                    key=f"contact_del_sel_{p_id}",
                )
                if col_del_btn.button("🗑️ " + t("Remove", "Eliminar"), key=f"contact_del_btn_{p_id}"):
                    execute_db_query(
                        "DELETE FROM contacts WHERE id = ?", (del_options[sel_contact_label],)
                    )
                    st.toast(t("Contact removed!", "¡Contacto eliminado!"), icon="✅")
                    st.rerun()
            else:
                st.caption(
                    t(
                        "No contacts registered for this product yet.",
                        "Este producto aún no tiene contactos registrados.",
                    )
                )

            with st.expander("➕ " + t("Add New Contact", "Añadir Nuevo Contacto")):
                with st.form(key=f"add_contact_form_{p_id}"):
                    ct1, ct2 = st.columns(2)
                    c_name = ct1.text_input(t("Name", "Nombre"))
                    c_surname = ct2.text_input(t("Surname", "Apellido"))
                    ct3, ct4 = st.columns(2)
                    c_phone = ct3.text_input(t("Phone", "Teléfono"))
                    c_email = ct4.text_input(t("Email", "Correo"))
                    c_country = st.text_input(t("Country", "País"))

                    if st.form_submit_button(t("Add Contact", "Añadir Contacto")):
                        execute_db_query(
                            """
                            INSERT INTO contacts (product_id, name, surname, phone, email, country)
                            VALUES (?, ?, ?, ?, ?, ?)
                            """,
                            (p_id, c_name, c_surname, c_phone, c_email, c_country),
                        )
                        st.toast(t("Contact added!", "¡Contacto añadido!"), icon="✅")
                        st.rerun()

            st.write("---")

            st.markdown("### 📘 " + t("Master Details & Pricing", "Ficha Maestra y Precios"))
            with st.form(key=f"edit_prod_form_{p_id}"):
                col_a, col_b = st.columns(2)
                f_name = col_a.text_input(t("Product Name", "Nombre del Producto"), value=product["product"])

                icon_options = ICON_OPTIONS if product["icon"] in ICON_OPTIONS else [product["icon"]] + ICON_OPTIONS
                f_icon = col_b.selectbox(
                    t("Icon", "Icono"),
                    options=icon_options,
                    index=icon_options.index(product["icon"]),
                )

                f_desc = st.text_area(
                    t("Description / Breakdown", "Descripción / Desglose"), value=product["description"]
                )
                f_char = st.text_area(
                    t("Characteristics / Packaging", "Características / Empaque"),
                    value=product["characteristics"],
                )
                f_where = st.text_area(
                    t("Where Used (Process)", "Dónde se Usa (Proceso)"), value=product["where_used"]
                )

                col_src, col_batch, col_haz = st.columns(3)
                f_source = col_src.text_input(
                    t("Source / Origin", "Origen"), value=product["source_origin"] or ""
                )
                f_batch = col_batch.text_input(
                    t("Batch / Lot", "Lote"), value=product["batch_lot"] or ""
                )
                f_hazard = col_haz.text_input(
                    t("Hazard Class", "Clase de Riesgo"), value=product["sds_hazard_class"] or ""
                )

                col_sup, col_del1, col_del2, col_exp = st.columns(4)
                f_supplier = col_sup.text_input(
                    t("Provider / Supplier", "Proveedor"), value=product["suppliers"] or ""
                )
                f_delivery_time = col_del1.number_input(
                    t("Delivery Time", "Tiempo de Entrega"), value=safe_int(product["delivery_time"]), min_value=0
                )
                delivery_unit_options = ["days", "weeks"]
                delivery_unit_labels = {"days": t("days", "días"), "weeks": t("weeks", "semanas")}
                current_delivery_unit = (
                    product["delivery_time_unit"]
                    if product["delivery_time_unit"] in delivery_unit_options
                    else "days"
                )
                f_delivery_unit = col_del2.selectbox(
                    t("Delivery Time Unit", "Unidad de Entrega"),
                    options=delivery_unit_options,
                    format_func=lambda k: delivery_unit_labels[k],
                    index=delivery_unit_options.index(current_delivery_unit),
                )
                try:
                    current_expiry = datetime.strptime(product["expiring_date"], "%Y-%m-%d")
                except (TypeError, ValueError):
                    current_expiry = datetime.now()
                f_expiry = col_exp.date_input(
                    t("Expiry Date", "Fecha de Caducidad"), value=current_expiry
                )

                c1, c2, c3 = st.columns(3)
                f_price = c1.number_input(
                    t("Base Price (€)", "Precio Base (€)"), value=safe_float(product["price"])
                )
                f_disc = c2.number_input(
                    t("Discount (%)", "Descuento (%)"), value=safe_float(product["discount"])
                )
                f_trans = c3.number_input(
                    t("Transport Fee (€)", "Transporte (€)"), value=safe_float(product["transport_price"])
                )

                c4, c5, c6 = st.columns(3)
                f_qty = c4.number_input(
                    t("Total Quantity", "Cantidad Total"), value=safe_float(product["quantity"])
                )
                f_min = c5.number_input(
                    t("Min Stock Level", "Nivel Mínimo de Stock"), value=safe_int(product["min_stock"])
                )
                f_ubic = c6.text_input(t("Location", "Ubicación"), value=product["ubication"] or "")

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
                            where_used = ?, source_origin = ?, batch_lot = ?, sds_hazard_class = ?,
                            suppliers = ?, delivery_time = ?, delivery_time_unit = ?, expiring_date = ?,
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
                            f_source,
                            f_batch,
                            f_hazard,
                            f_supplier,
                            f_delivery_time,
                            f_delivery_unit,
                            f_expiry.strftime("%Y-%m-%d"),
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

                    st.toast(t("Product details saved successfully!", "¡Detalles del producto guardados correctamente!"), icon="✅")
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
                    st.toast(t("Product removed!", "¡Producto eliminado!"), icon="✅")
                    st.rerun()
                if col_cancel.button(
                    "Cancelar" if es else "Cancel", key=f"confirm_no_{p_id}"
                ):
                    del st.session_state[confirm_key]
                    st.rerun()

        product_modal()