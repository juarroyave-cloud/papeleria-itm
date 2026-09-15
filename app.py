# ============================================================
# PAPELERÍA ITM
# Sistema de Inventario y Gestión de Papelería
# ============================================================
#
# Ejecutar en PowerShell:
# python -m streamlit run papeleria.app.py
#
# Seguridad:
# El administrador se configura mediante Streamlit Secrets:
# ADMIN_USER = "admin"
# ADMIN_PASSWORD = "una_clave_segura"
# Nunca se muestran credenciales en la pantalla de acceso.
#
# Estructura del Excel:
# A = ÍTEM
# B = DESCRIPCIÓN / ESPECIFICACIONES TÉCNICAS
# C = CANTIDAD
# D = VALOR UNITARIO SIN IVA
#
# El IVA se calcula automáticamente al 19%.
# ============================================================

import sqlite3
import hashlib
import textwrap
import unicodedata
from pathlib import Path
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px


# ============================================================
# CONFIGURACIÓN
# ============================================================

st.set_page_config(
    page_title="Papelería ITM",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = Path(__file__).resolve().parent
DB_FILE = BASE_DIR / "inventario_papeleria.db"
IVA = 0.19


# ============================================================
# COLORES
# ============================================================

AZUL_ITM = "#00529B"
AZUL_OSCURO = "#003B73"
AZUL_MEDIO = "#0072BC"
FONDO = "#F4F7FA"
TEXTO = "#1F2937"
GRIS = "#64748B"
VERDE = "#198754"
ROJO = "#DC3545"
AMARILLO = "#FFC107"


# ============================================================
# FUNCIONES DE PRESENTACIÓN
# ============================================================

def html(contenido):
    """Renderiza HTML sin que Streamlit lo muestre como código."""
    st.markdown(
        textwrap.dedent(contenido),
        unsafe_allow_html=True
    )


def fecha_actual():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def hash_password(password):
    return hashlib.sha256(
        str(password).encode("utf-8")
    ).hexdigest()


def credenciales_admin():
    """Obtiene credenciales administrativas desde Streamlit Secrets."""
    try:
        usuario = str(st.secrets.get("ADMIN_USER", "admin")).strip()
        password = str(st.secrets.get("ADMIN_PASSWORD", "")).strip()
    except Exception:
        usuario = "admin"
        password = ""
    return usuario or "admin", password


def credenciales_consulta():
    """Obtiene credenciales del usuario de consulta desde Streamlit Secrets."""
    try:
        usuario = str(st.secrets.get("CONSULTA_USER", "consulta")).strip()
        password = str(st.secrets.get("CONSULTA_PASSWORD", "")).strip()
    except Exception:
        usuario = "consulta"
        password = ""
    return usuario or "consulta", password


def dinero(valor):
    try:
        valor = float(valor)
    except Exception:
        valor = 0

    return "$ " + f"{valor:,.0f}".replace(",", ".")


def numero(valor):
    """Convierte números de Excel a float."""
    try:
        if pd.isna(valor):
            return 0.0

        if isinstance(valor, (int, float)):
            return float(valor)

        texto = str(valor).strip()
        texto = texto.replace("$", "").replace(" ", "")

        # Formato colombiano: 1.234.567,89
        if "." in texto and "," in texto:
            texto = texto.replace(".", "").replace(",", ".")
        elif "," in texto:
            texto = texto.replace(",", ".")

        return float(texto)
    except Exception:
        return 0.0


def normalizar_texto(valor):
    texto = "" if valor is None else str(valor).strip().upper()
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    texto = " ".join(texto.split())
    return texto


def limpiar_item(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip()

    try:
        n = float(texto)
        if n.is_integer():
            return str(int(n))
    except Exception:
        pass

    return texto


# ============================================================
# ESTILOS
# ============================================================

html(f"""
<style>
.stApp {{
    background-color: {FONDO};
}}

.main {{
    background-color: {FONDO};
}}

h1, h2, h3 {{
    color: {AZUL_OSCURO} !important;
}}

label {{
    color: {TEXTO} !important;
    font-weight: 600 !important;
}}

.header-itm {{
    background: linear-gradient(
        110deg,
        {AZUL_OSCURO} 0%,
        {AZUL_ITM} 55%,
        {AZUL_MEDIO} 100%
    );
    padding: 28px 35px;
    border-radius: 14px;
    margin-bottom: 25px;
    box-shadow: 0 5px 18px rgba(0,0,0,.12);
}}

.header-itm h1 {{
    color: white !important;
    margin: 0;
    font-size: 34px;
    font-weight: 800;
}}

.header-itm p {{
    color: white !important;
    margin: 8px 0 0 0;
    font-size: 16px;
}}

.page-title {{
    color: {AZUL_OSCURO} !important;
    font-size: 30px;
    font-weight: 800;
    margin-bottom: 4px;
}}

.page-subtitle {{
    color: {GRIS} !important;
    font-size: 15px;
    margin-bottom: 22px;
}}

.metric-card {{
    background: white;
    padding: 18px;
    border-radius: 12px;
    border-left: 5px solid {AZUL_ITM};
    box-shadow: 0 3px 12px rgba(0,0,0,.07);
    min-height: 110px;
}}

.metric-title {{
    color: {GRIS};
    font-size: 14px;
    font-weight: 600;
}}

.metric-value {{
    color: {AZUL_ITM};
    font-size: 24px;
    font-weight: 800;
    margin-top: 8px;
}}

.info-box {{
    background-color: #EAF3FB;
    border-left: 5px solid {AZUL_ITM};
    padding: 14px;
    border-radius: 8px;
    margin: 10px 0;
}}

.warning-box {{
    background-color: #FFF8DD;
    border-left: 5px solid {AMARILLO};
    padding: 14px;
    border-radius: 8px;
    margin: 10px 0;
}}

.success-box {{
    background-color: #E9F7EF;
    border-left: 5px solid {VERDE};
    padding: 14px;
    border-radius: 8px;
    margin: 10px 0;
}}

section[data-testid="stSidebar"] {{
    background: linear-gradient(
        180deg,
        {AZUL_OSCURO} 0%,
        {AZUL_ITM} 100%
    );
}}

section[data-testid="stSidebar"] * {{
    color: white !important;
}}

.sidebar-logo {{
    text-align: center;
    padding: 10px 0 15px 0;
}}

.sidebar-logo .logo {{
    font-size: 46px;
}}

.sidebar-logo h2 {{
    color: white !important;
    margin: 0;
    font-size: 27px;
}}

.sidebar-logo p {{
    color: #EAF3FB !important;
    font-size: 13px;
}}

.user-card {{
    background: rgba(255,255,255,.12);
    border-radius: 10px;
    padding: 13px;
    margin-top: 10px;
}}

.stButton > button {{
    border-radius: 8px;
    font-weight: 700;
    min-height: 42px;
}}
</style>
""")


# ============================================================
# BASE DE DATOS
# ============================================================

def conectar():
    conn = sqlite3.connect(
        str(DB_FILE),
        timeout=30
    )
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def migrar_tabla_inventario(conn):
    """
    Corrige automáticamente bases creadas con versiones anteriores.
    Evita errores como:
    no such column: stock_minimo
    """

    columnas = conn.execute(
        "PRAGMA table_info(inventario)"
    ).fetchall()

    existentes = {fila[1] for fila in columnas}

    columnas_faltantes = {
        "descripcion": "TEXT DEFAULT ''",
        "cantidad": "REAL DEFAULT 0",
        "valor_sin_iva": "REAL DEFAULT 0",
        "iva": "REAL DEFAULT 0.19",
        "valor_con_iva": "REAL DEFAULT 0",
        "stock_minimo": "REAL DEFAULT 10",
        "fecha_actualizacion": "TEXT"
    }

    for nombre, tipo in columnas_faltantes.items():
        if nombre not in existentes:
            conn.execute(
                f"ALTER TABLE inventario ADD COLUMN {nombre} {tipo}"
            )

    conn.commit()


def crear_base_datos():
    conn = conectar()

    conn.execute("""
        CREATE TABLE IF NOT EXISTS inventario (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            item TEXT UNIQUE NOT NULL,
            descripcion TEXT DEFAULT '',
            cantidad REAL DEFAULT 0,
            valor_sin_iva REAL DEFAULT 0,
            iva REAL DEFAULT 0.19,
            valor_con_iva REAL DEFAULT 0,
            stock_minimo REAL DEFAULT 10,
            fecha_actualizacion TEXT
        )
    """)

    migrar_tabla_inventario(conn)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS pedidos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_pedido TEXT UNIQUE NOT NULL,
            fecha TEXT,
            usuario TEXT,
            area TEXT,
            valor_total REAL DEFAULT 0,
            estado TEXT DEFAULT 'APROBADO'
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS detalle_pedidos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            pedido_id INTEGER,
            item TEXT,
            descripcion TEXT,
            cantidad REAL,
            valor_unitario REAL,
            valor_total REAL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            fecha TEXT,
            tipo TEXT,
            item TEXT,
            cantidad REAL,
            valor_unitario REAL,
            valor_total REAL,
            area TEXT,
            usuario TEXT,
            observacion TEXT
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario TEXT UNIQUE NOT NULL,
            password TEXT NOT NULL,
            rol TEXT NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS areas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL
        )
    """)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS importaciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            clave TEXT UNIQUE NOT NULL,
            fecha TEXT,
            usuario TEXT,
            registros INTEGER DEFAULT 0
        )
    """)

    # Administrador: credenciales configuradas fuera del código.
    admin_usuario, admin_password = credenciales_admin()
    existe_admin = conn.execute(
        "SELECT id FROM usuarios WHERE usuario = ?",
        (admin_usuario,)
    ).fetchone()

    if existe_admin is None and admin_password:
        conn.execute("""
            INSERT INTO usuarios (usuario, password, rol)
            VALUES (?, ?, ?)
        """, (
            admin_usuario,
            hash_password(admin_password),
            "Administrador"
        ))
    elif existe_admin is not None and admin_password:
        # Mantiene sincronizada la contraseña administrativa con el Secret.
        conn.execute(
            "UPDATE usuarios SET password = ?, rol = ? WHERE usuario = ?",
            (hash_password(admin_password), "Administrador", admin_usuario)
        )

    # Usuario de consulta: puede consultar inventario, pedidos y reportes,
    # pero no tiene acceso a compras ni administración.
    consulta_usuario, consulta_password = credenciales_consulta()
    existe_consulta = conn.execute(
        "SELECT id FROM usuarios WHERE usuario = ?",
        (consulta_usuario,)
    ).fetchone()
    if existe_consulta is None and consulta_password:
        conn.execute("""
            INSERT INTO usuarios (usuario, password, rol)
            VALUES (?, ?, ?)
        """, (consulta_usuario, hash_password(consulta_password), "Consulta"))
    elif existe_consulta is not None and consulta_password:
        conn.execute(
            "UPDATE usuarios SET password = ?, rol = ? WHERE usuario = ?",
            (hash_password(consulta_password), "Consulta", consulta_usuario)
        )

    # Áreas iniciales
    areas = [
        "Administración",
        "Académica",
        "Biblioteca",
        "Bienestar Institucional",
        "Compras",
        "Contabilidad",
        "Financiera",
        "Gestión Humana",
        "Mantenimiento",
        "Planeación",
        "Rectoría",
        "Tecnología",
        "Otra"
    ]

    for area in areas:
        conn.execute(
            "INSERT OR IGNORE INTO areas (nombre) VALUES (?)",
            (area,)
        )

    conn.commit()
    conn.close()


# ============================================================
# EXCEL
# ============================================================

def localizar_excel():
    """
    Busca primero 'papeleria listado.xlsx'.
    Si no existe, busca cualquier .xlsx en la carpeta.
    """
    preferido = BASE_DIR / "papeleria listado.xlsx"

    if preferido.exists():
        return preferido

    archivos = sorted(BASE_DIR.glob("*.xlsx"))

    if archivos:
        return archivos[0]

    return None


def importar_excel(reemplazar=False):
    archivo = localizar_excel()

    if archivo is None:
        return False, (
            "No se encontró ningún archivo Excel (.xlsx) "
            "en la carpeta de la aplicación."
        )

    try:
        libro = pd.ExcelFile(archivo, engine="openpyxl")

        if "papeleria" in libro.sheet_names:
            hoja = "papeleria"
        else:
            hoja = libro.sheet_names[0]

        df = pd.read_excel(
            archivo,
            sheet_name=hoja,
            header=0,
            engine="openpyxl"
        )

        if df.empty:
            return False, "El Excel está vacío."

        if len(df.columns) < 4:
            return False, (
                "El Excel debe tener al menos cuatro columnas: "
                "Ítem, Descripción, Cantidad y Valor sin IVA."
            )

        # A, B, C, D
        col_item = df.columns[0]
        col_desc = df.columns[1]
        col_cantidad = df.columns[2]
        col_valor = df.columns[3]

        conn = conectar()

        nuevos = 0
        actualizados = 0
        procesados = 0

        for _, fila in df.iterrows():

            item = limpiar_item(fila[col_item])

            if not item:
                continue

            if item.lower() in ["nan", "item", "ítem"]:
                continue

            descripcion = str(fila[col_desc]).strip()

            if descripcion.lower() == "nan":
                descripcion = ""

            cantidad = numero(fila[col_cantidad])
            valor_sin_iva = numero(fila[col_valor])

            valor_iva = valor_sin_iva * IVA
            valor_con_iva = valor_sin_iva + valor_iva

            existe = conn.execute(
                "SELECT id FROM inventario WHERE item = ?",
                (item,)
            ).fetchone()

            if existe is None:
                conn.execute("""
                    INSERT INTO inventario (
                        item,
                        descripcion,
                        cantidad,
                        valor_sin_iva,
                        iva,
                        valor_con_iva,
                        stock_minimo,
                        fecha_actualizacion
                    )
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    item,
                    descripcion,
                    cantidad,
                    valor_sin_iva,
                    IVA,
                    valor_con_iva,
                    10,
                    fecha_actual()
                ))
                nuevos += 1

            elif reemplazar:
                conn.execute("""
                    UPDATE inventario
                    SET descripcion = ?,
                        cantidad = ?,
                        valor_sin_iva = ?,
                        iva = ?,
                        valor_con_iva = ?,
                        fecha_actualizacion = ?
                    WHERE item = ?
                """, (
                    descripcion,
                    cantidad,
                    valor_sin_iva,
                    IVA,
                    valor_con_iva,
                    fecha_actual(),
                    item
                ))
                actualizados += 1

            procesados += 1

        conn.commit()
        conn.close()

        return True, {
            "archivo": archivo.name,
            "hoja": hoja,
            "procesados": procesados,
            "nuevos": nuevos,
            "actualizados": actualizados
        }

    except Exception as e:
        return False, f"Error leyendo el Excel: {e}"


# ============================================================
# CONSULTAS
# ============================================================

def obtener_inventario():
    conn = conectar()

    try:
        return pd.read_sql_query("""
            SELECT
                id,
                item AS "Ítem",
                descripcion AS "Descripción",
                cantidad AS "Cantidad disponible",
                valor_sin_iva AS "Valor sin IVA",
                iva AS "IVA",
                valor_con_iva AS "Valor con IVA",
                cantidad * valor_con_iva AS "Valor total",
                stock_minimo AS "Stock mínimo"
            FROM inventario
            ORDER BY
                CASE
                    WHEN item GLOB '[0-9]*'
                    THEN CAST(item AS INTEGER)
                    ELSE 999999999
                END,
                item
        """, conn)
    finally:
        conn.close()


def obtener_areas():
    conn = conectar()

    try:
        rows = conn.execute(
            "SELECT nombre FROM areas ORDER BY nombre"
        ).fetchall()

        return [r[0] for r in rows]
    finally:
        conn.close()


def obtener_items_disponibles():
    df = obtener_inventario()

    if df.empty:
        return df

    return df[df["Cantidad disponible"] > 0].copy()


def etiquetas_articulos(df):
    """Crea etiquetas legibles para los selectores: código - descripción."""
    etiquetas = {}
    if df is None or df.empty:
        return etiquetas
    for _, fila in df.iterrows():
        item = str(fila["Ítem"])
        descripcion = str(fila["Descripción"]).strip()
        etiquetas[item] = f"{item} - {descripcion}" if descripcion else item
    return etiquetas


# ============================================================
# PEDIDOS
# ============================================================

def generar_numero_pedido():
    conn = conectar()

    try:
        fila = conn.execute("""
            SELECT id
            FROM pedidos
            ORDER BY id DESC
            LIMIT 1
        """).fetchone()
    finally:
        conn.close()

    consecutivo = 1 if fila is None else int(fila[0]) + 1

    return f"PED-{consecutivo:06d}"


def crear_pedido(usuario, area, carrito):
    if not carrito:
        return False, "El pedido está vacío."

    conn = conectar()

    try:
        # Validación antes de modificar inventario
        for registro in carrito:

            item = str(registro["item"])
            solicitado = float(registro["cantidad"])

            fila = conn.execute("""
                SELECT cantidad
                FROM inventario
                WHERE item = ?
            """, (item,)).fetchone()

            if fila is None:
                raise Exception(
                    f"El artículo {item} no existe."
                )

            disponible = float(fila[0])

            if solicitado <= 0:
                raise Exception(
                    f"La cantidad para {item} debe ser mayor que cero."
                )

            if solicitado > disponible:
                raise Exception(
                    f"Stock insuficiente para {item}. "
                    f"Disponible: {disponible:,.0f}"
                    .replace(",", ".")
                )

        numero_pedido = generar_numero_pedido()

        cursor = conn.execute("""
            INSERT INTO pedidos (
                numero_pedido,
                fecha,
                usuario,
                area,
                valor_total,
                estado
            )
            VALUES (?, ?, ?, ?, ?, ?)
        """, (
            numero_pedido,
            fecha_actual(),
            usuario,
            area,
            0,
            "APROBADO"
        ))

        pedido_id = cursor.lastrowid
        total = 0

        for registro in carrito:

            item = str(registro["item"])
            cantidad = float(registro["cantidad"])

            fila = conn.execute("""
                SELECT descripcion, valor_con_iva, cantidad
                FROM inventario
                WHERE item = ?
            """, (item,)).fetchone()

            if fila is None:
                raise Exception(
                    f"No se encontró el artículo {item}."
                )

            descripcion = fila[0]
            valor_unitario = float(fila[1])
            stock_actual = float(fila[2])

            if cantidad > stock_actual:
                raise Exception(
                    f"Stock insuficiente para {item}."
                )

            valor_total = cantidad * valor_unitario
            total += valor_total

            conn.execute("""
                INSERT INTO detalle_pedidos (
                    pedido_id,
                    item,
                    descripcion,
                    cantidad,
                    valor_unitario,
                    valor_total
                )
                VALUES (?, ?, ?, ?, ?, ?)
            """, (
                pedido_id,
                item,
                descripcion,
                cantidad,
                valor_unitario,
                valor_total
            ))

            # Descontar del inventario
            conn.execute("""
                UPDATE inventario
                SET cantidad = cantidad - ?,
                    fecha_actualizacion = ?
                WHERE item = ?
            """, (
                cantidad,
                fecha_actual(),
                item
            ))

            # Kardex
            conn.execute("""
                INSERT INTO movimientos (
                    fecha,
                    tipo,
                    item,
                    cantidad,
                    valor_unitario,
                    valor_total,
                    area,
                    usuario,
                    observacion
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                fecha_actual(),
                "SALIDA",
                item,
                cantidad,
                valor_unitario,
                valor_total,
                area,
                usuario,
                f"Pedido {numero_pedido}"
            ))

        conn.execute("""
            UPDATE pedidos
            SET valor_total = ?
            WHERE id = ?
        """, (total, pedido_id))

        conn.commit()

        return True, numero_pedido

    except Exception as e:
        conn.rollback()
        return False, str(e)

    finally:
        conn.close()


# ============================================================
# COMPRAS
# ============================================================

def registrar_compra(item, cantidad, valor_sin_iva, usuario):
    cantidad = float(cantidad)
    valor_sin_iva = float(valor_sin_iva)

    if cantidad <= 0:
        return False, "La cantidad debe ser mayor que cero."

    if valor_sin_iva < 0:
        return False, "El valor no puede ser negativo."

    conn = conectar()

    try:
        fila = conn.execute("""
            SELECT cantidad
            FROM inventario
            WHERE item = ?
        """, (item,)).fetchone()

        if fila is None:
            return False, "El artículo no existe."

        cantidad_actual = float(fila[0])
        nueva_cantidad = cantidad_actual + cantidad

        valor_con_iva = valor_sin_iva * (1 + IVA)
        valor_total = cantidad * valor_con_iva

        conn.execute("""
            UPDATE inventario
            SET cantidad = ?,
                valor_sin_iva = ?,
                iva = ?,
                valor_con_iva = ?,
                fecha_actualizacion = ?
            WHERE item = ?
        """, (
            nueva_cantidad,
            valor_sin_iva,
            IVA,
            valor_con_iva,
            fecha_actual(),
            item
        ))

        conn.execute("""
            INSERT INTO movimientos (
                fecha,
                tipo,
                item,
                cantidad,
                valor_unitario,
                valor_total,
                area,
                usuario,
                observacion
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            fecha_actual(),
            "ENTRADA",
            item,
            cantidad,
            valor_con_iva,
            valor_total,
            "",
            usuario,
            "Nueva compra"
        ))

        conn.commit()

        return True, nueva_cantidad

    except Exception as e:
        conn.rollback()
        return False, str(e)

    finally:
        conn.close()


# ============================================================
# AJUSTAR INVENTARIO
# ============================================================

def actualizar_inventario(
    item,
    nueva_cantidad,
    valor_sin_iva,
    stock_minimo,
    usuario
):
    conn = conectar()

    try:
        fila = conn.execute("""
            SELECT cantidad
            FROM inventario
            WHERE item = ?
        """, (item,)).fetchone()

        if fila is None:
            return False, "El artículo no existe."

        anterior = float(fila[0])
        nueva = float(nueva_cantidad)
        valor_sin_iva = float(valor_sin_iva)
        stock_minimo = float(stock_minimo)

        if nueva < 0:
            return False, "La cantidad no puede ser negativa."

        if valor_sin_iva < 0:
            return False, "El valor no puede ser negativo."

        if stock_minimo < 0:
            return False, "El stock mínimo no puede ser negativo."

        diferencia = nueva - anterior
        valor_con_iva = valor_sin_iva * (1 + IVA)

        conn.execute("""
            UPDATE inventario
            SET cantidad = ?,
                valor_sin_iva = ?,
                iva = ?,
                valor_con_iva = ?,
                stock_minimo = ?,
                fecha_actualizacion = ?
            WHERE item = ?
        """, (
            nueva,
            valor_sin_iva,
            IVA,
            valor_con_iva,
            stock_minimo,
            fecha_actual(),
            item
        ))

        if diferencia != 0:

            tipo = (
                "AJUSTE_ENTRADA"
                if diferencia > 0
                else "AJUSTE_SALIDA"
            )

            conn.execute("""
                INSERT INTO movimientos (
                    fecha,
                    tipo,
                    item,
                    cantidad,
                    valor_unitario,
                    valor_total,
                    area,
                    usuario,
                    observacion
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                fecha_actual(),
                tipo,
                item,
                diferencia,
                valor_con_iva,
                diferencia * valor_con_iva,
                "",
                usuario,
                "Ajuste manual de inventario"
            ))

        conn.commit()

        return True, nueva

    except Exception as e:
        conn.rollback()
        return False, str(e)

    finally:
        conn.close()


# ============================================================
# LOGIN
# ============================================================

def header():
    html("""
    <div class="header-itm">
        <h1>📦 Papelería ITM</h1>
        <p>Sistema de Inventario y Gestión de Papelería</p>
    </div>
    """)


def login():
    header()

    _, centro, _ = st.columns([1, 2, 1])

    with centro:

        st.markdown(
            f"""
            <h2 style="
                text-align:center;
                color:{AZUL_OSCURO} !important;
            ">
                🔐 Ingreso al sistema
            </h2>
            """,
            unsafe_allow_html=True
        )

        usuario = st.text_input(
            "Usuario",
            key="login_usuario"
        )

        password = st.text_input(
            "Contraseña",
            type="password",
            key="login_password"
        )

        admin_usuario_config, admin_password_config = credenciales_admin()

        if not admin_password_config:
            st.warning(
                "El administrador aún no está configurado. En Streamlit Cloud, agregue ADMIN_PASSWORD en Settings → Secrets."
            )

        if st.button(
            "Ingresar",
            use_container_width=True
        ):

            if not usuario or not password:
                st.error(
                    "Ingrese usuario y contraseña."
                )
                return

            conn = conectar()

            fila = conn.execute("""
                SELECT usuario, rol
                FROM usuarios
                WHERE usuario = ?
                AND password = ?
            """, (
                usuario.strip(),
                hash_password(password)
            )).fetchone()

            conn.close()

            if fila:

                st.session_state["logueado"] = True
                st.session_state["usuario"] = fila[0]
                st.session_state["rol"] = fila[1]
                st.session_state["carrito"] = []

                st.rerun()

            else:
                st.error(
                    "Usuario o contraseña incorrectos."
                )

        st.info(
            "Las credenciales se configuran de forma segura mediante Streamlit Secrets. "
            "El usuario de consulta no tiene acceso a compras ni administración."
        )


# ============================================================
# INICIO
# ============================================================

def pagina_inicio():
    header()

    html("""
    <div class="page-title">🏠 Inicio</div>
    <div class="page-subtitle">
        Resumen general del inventario de Papelería ITM.
    </div>
    """)

    df = obtener_inventario()

    if df.empty:
        st.warning(
            "No hay artículos cargados. "
            "Verifique el archivo Excel."
        )
        return

    articulos = len(df)
    unidades = df["Cantidad disponible"].sum()

    total_sin_iva = (
        df["Cantidad disponible"] *
        df["Valor sin IVA"]
    ).sum()

    total_con_iva = (
        df["Cantidad disponible"] *
        df["Valor con IVA"]
    ).sum()

    stock_bajo = len(
        df[
            df["Cantidad disponible"]
            <=
            df["Stock mínimo"]
        ]
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    datos = [
        ("📦 Artículos", f"{articulos}"),
        ("🔢 Unidades", f"{unidades:,.0f}".replace(",", ".")),
        ("💰 Valor sin IVA", dinero(total_sin_iva)),
        ("💵 Valor con IVA", dinero(total_con_iva)),
        ("⚠️ Stock bajo", f"{stock_bajo}")
    ]

    for columna, (titulo, valor) in zip(
        [c1, c2, c3, c4, c5],
        datos
    ):
        with columna:
            html(f"""
            <div class="metric-card">
                <div class="metric-title">{titulo}</div>
                <div class="metric-value">{valor}</div>
            </div>
            """)

    st.divider()

    st.subheader("📊 Consumo por área")

    conn = conectar()

    try:
        consumo = pd.read_sql_query("""
            SELECT
                area AS "Área",
                SUM(cantidad) AS "Cantidad",
                SUM(valor_total) AS "Valor"
            FROM movimientos
            WHERE tipo = 'SALIDA'
            GROUP BY area
            ORDER BY SUM(cantidad) DESC
        """, conn)
    finally:
        conn.close()

    if consumo.empty:
        st.info(
            "Todavía no existen pedidos registrados."
        )
    else:
        a, b = st.columns(2)

        with a:
            fig = px.bar(
                consumo,
                x="Área",
                y="Cantidad",
                title="Cantidad consumida por área"
            )
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(
                fig,
                use_container_width=True
            )

        with b:
            fig = px.bar(
                consumo,
                x="Área",
                y="Valor",
                title="Valor consumido por área"
            )
            fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(
                fig,
                use_container_width=True
            )


# ============================================================
# INVENTARIO
# ============================================================

def pagina_inventario():
    header()

    html("""
    <div class="page-title">📦 Inventario</div>
    <div class="page-subtitle">
        Consulta de existencias actuales de papelería.
    </div>
    """)

    df = obtener_inventario()

    if df.empty:
        st.warning("No hay artículos registrados.")
        return

    busqueda = st.text_input(
        "🔎 Buscar artículo",
        placeholder="Número de ítem o descripción..."
    )

    if busqueda.strip():

        texto = busqueda.strip().lower()

        df = df[
            df["Ítem"].astype(str).str.lower().str.contains(
                texto, na=False
            )
            |
            df["Descripción"].astype(str).str.lower().str.contains(
                texto, na=False
            )
        ]

    filtro = st.selectbox(
        "Filtro de inventario",
        [
            "Todos",
            "Disponibles",
            "Stock bajo",
            "Agotados"
        ]
    )

    if filtro == "Disponibles":
        df = df[
            df["Cantidad disponible"]
            >
            df["Stock mínimo"]
        ]

    elif filtro == "Stock bajo":
        df = df[
            (df["Cantidad disponible"] <= df["Stock mínimo"])
            &
            (df["Cantidad disponible"] > 0)
        ]

    elif filtro == "Agotados":
        df = df[
            df["Cantidad disponible"] <= 0
        ]

    st.write(
        f"Artículos encontrados: **{len(df)}**"
    )

    tabla = df.copy()

    tabla["IVA"] = "19%"
    tabla["Valor sin IVA"] = tabla["Valor sin IVA"].apply(dinero)
    tabla["Valor con IVA"] = tabla["Valor con IVA"].apply(dinero)
    tabla["Valor total"] = tabla["Valor total"].apply(dinero)

    st.dataframe(
        tabla[
            [
                "Ítem",
                "Descripción",
                "Cantidad disponible",
                "Valor sin IVA",
                "IVA",
                "Valor con IVA",
                "Valor total",
                "Stock mínimo"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

    csv = df.to_csv(
        index=False
    ).encode("utf-8-sig")

    st.download_button(
        "⬇️ Descargar inventario",
        data=csv,
        file_name="inventario_papeleria_itm.csv",
        mime="text/csv"
    )


# ============================================================
# VALORES E IVA
# ============================================================

def pagina_iva():
    header()

    html("""
    <div class="page-title">💰 Valores e IVA</div>
    <div class="page-subtitle">
        Valores unitarios y cálculo automático del IVA del 19%.
    </div>
    """)

    df = obtener_inventario()

    if df.empty:
        st.warning("No hay artículos registrados.")
        return

    tabla = df.copy()

    tabla["Valor IVA"] = (
        tabla["Valor sin IVA"] * IVA
    )

    tabla["IVA"] = "19%"
    tabla["Valor sin IVA"] = tabla["Valor sin IVA"].apply(dinero)
    tabla["Valor IVA"] = tabla["Valor IVA"].apply(dinero)
    tabla["Valor con IVA"] = tabla["Valor con IVA"].apply(dinero)
    tabla["Valor total"] = tabla["Valor total"].apply(dinero)

    st.dataframe(
        tabla[
            [
                "Ítem",
                "Descripción",
                "Valor sin IVA",
                "IVA",
                "Valor IVA",
                "Valor con IVA",
                "Cantidad disponible",
                "Valor total"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# PEDIDOS
# ============================================================

def pagina_pedidos():
    header()

    html("""
    <div class="page-title">🛒 Realizar pedido</div>
    <div class="page-subtitle">
        Seleccione el área, los artículos y las cantidades.
    </div>
    """)

    df = obtener_items_disponibles()

    if df.empty:
        st.warning(
            "No hay inventario disponible para realizar pedidos."
        )
        return

    if "carrito" not in st.session_state:
        st.session_state["carrito"] = []

    areas = obtener_areas()
    OPCION_NUEVA_AREA = "➕ Agregar otra área..."

    opciones_area = areas + [OPCION_NUEVA_AREA]

    area = st.selectbox(
        "🏢 Área solicitante",
        opciones_area
    )

    if area == OPCION_NUEVA_AREA:
        area_nueva = st.text_input(
            "Nombre de la nueva área",
            placeholder="Ej. Facultad de Ingenierías",
            key="nueva_area_pedido"
        ).strip()

        guardar_area = st.button(
            "💾 Guardar nueva área",
            use_container_width=True
        )

        if guardar_area:
            if not area_nueva:
                st.error("Ingrese el nombre de la nueva área.")
            else:
                conn = conectar()
                try:
                    existente = conn.execute(
                        "SELECT nombre FROM areas WHERE LOWER(TRIM(nombre)) = LOWER(TRIM(?))",
                        (area_nueva,)
                    ).fetchone()

                    if existente:
                        st.warning(f"El área '{existente[0]}' ya existe.")
                        st.session_state["area_pedido_seleccionada"] = existente[0]
                    else:
                        conn.execute(
                            "INSERT INTO areas (nombre) VALUES (?)",
                            (area_nueva,)
                        )
                        conn.commit()
                        st.success(
                            f"Área '{area_nueva}' creada correctamente. Ya está disponible para esta solicitud."
                        )
                        st.session_state["area_pedido_seleccionada"] = area_nueva
                        st.session_state["nueva_area_pedido"] = ""
                except sqlite3.IntegrityError:
                    st.error("No fue posible crear el área porque ya existe.")
                finally:
                    conn.close()

                st.rerun()

        # Si acabamos de crear una nueva área, úsala automáticamente.
        area = st.session_state.get("area_pedido_seleccionada", area_nueva)

    st.divider()

    items = df["Ítem"].astype(str).tolist()

    col1, col2, col3 = st.columns([5, 2, 2])

    with col1:
        etiquetas = etiquetas_articulos(df)
        item = st.selectbox(
            "📦 Ítem",
            items,
            format_func=lambda x: etiquetas.get(str(x), str(x))
        )

    seleccionado = df[
        df["Ítem"].astype(str) == str(item)
    ]

    if seleccionado.empty:
        st.error("No se encontró el artículo.")
        return

    articulo = seleccionado.iloc[0]

    disponible = float(
        articulo["Cantidad disponible"]
    )

    maximo = max(1, int(disponible))

    with col2:
        cantidad = st.number_input(
            "Cantidad",
            min_value=1,
            max_value=maximo,
            value=1,
            step=1
        )

    with col3:
        st.write("")
        st.write("")
        agregar = st.button(
            "➕ Agregar",
            use_container_width=True
        )

    html(f"""
    <div class="info-box">
        <strong>Ítem:</strong> {articulo["Ítem"]}<br>
        <strong>Descripción:</strong> {articulo["Descripción"]}<br>
        <strong>Disponible:</strong> {disponible:,.0f}<br>
        <strong>Valor unitario con IVA:</strong>
        {dinero(articulo["Valor con IVA"])}
    </div>
    """)

    if agregar:

        ya_existe = False

        for registro in st.session_state["carrito"]:

            if str(registro["item"]) == str(item):

                nueva = (
                    registro["cantidad"]
                    +
                    int(cantidad)
                )

                if nueva > disponible:
                    st.error(
                        "La cantidad acumulada supera "
                        "el inventario disponible."
                    )
                else:
                    registro["cantidad"] = nueva

                ya_existe = True
                break

        if not ya_existe:
            st.session_state["carrito"].append({
                "item": str(item),
                "cantidad": int(cantidad)
            })

        st.rerun()

    st.divider()
    st.subheader("📋 Detalle del pedido")

    carrito = st.session_state["carrito"]

    if not carrito:
        st.info("No hay artículos agregados al pedido.")
        return

    filas = []

    for registro in carrito:

        encontrado = df[
            df["Ítem"].astype(str)
            ==
            str(registro["item"])
        ]

        if encontrado.empty:
            continue

        art = encontrado.iloc[0]

        cant = float(registro["cantidad"])
        unitario = float(art["Valor con IVA"])

        filas.append({
            "Ítem": registro["item"],
            "Descripción": art["Descripción"],
            "Cantidad": cant,
            "Disponible": art["Cantidad disponible"],
            "Valor unitario": unitario,
            "Valor total": cant * unitario
        })

    pedido_df = pd.DataFrame(filas)

    if pedido_df.empty:
        st.warning("El pedido no contiene artículos válidos.")
        return

    mostrar = pedido_df.copy()
    mostrar["Valor unitario"] = mostrar["Valor unitario"].apply(dinero)
    mostrar["Valor total"] = mostrar["Valor total"].apply(dinero)

    st.dataframe(
        mostrar,
        use_container_width=True,
        hide_index=True
    )

    total = pedido_df["Valor total"].sum()

    st.metric(
        "💰 Total del pedido",
        dinero(total)
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button(
            "🗑️ Vaciar pedido",
            use_container_width=True
        ):
            st.session_state["carrito"] = []
            st.rerun()

    with c2:
        if st.button(
            "✅ Confirmar pedido",
            use_container_width=True
        ):

            correcto, resultado = crear_pedido(
                st.session_state["usuario"],
                area,
                carrito
            )

            if correcto:
                st.session_state["carrito"] = []

                st.success(
                    f"Pedido {resultado} registrado correctamente."
                )

                st.rerun()
            else:
                st.error(resultado)


# ============================================================
# HISTORIAL
# ============================================================

def pagina_historial():
    header()

    html("""
    <div class="page-title">📋 Historial de pedidos</div>
    <div class="page-subtitle">
        Consulta de pedidos registrados.
    </div>
    """)

    conn = conectar()

    try:

        if st.session_state["rol"] == "Administrador":

            df = pd.read_sql_query("""
                SELECT
                    numero_pedido AS "Pedido",
                    fecha AS "Fecha",
                    usuario AS "Usuario",
                    area AS "Área",
                    valor_total AS "Valor total",
                    estado AS "Estado"
                FROM pedidos
                ORDER BY id DESC
            """, conn)

        else:

            df = pd.read_sql_query("""
                SELECT
                    numero_pedido AS "Pedido",
                    fecha AS "Fecha",
                    usuario AS "Usuario",
                    area AS "Área",
                    valor_total AS "Valor total",
                    estado AS "Estado"
                FROM pedidos
                WHERE usuario = ?
                ORDER BY id DESC
            """, conn, params=(
                st.session_state["usuario"],
            ))

    finally:
        conn.close()

    if df.empty:
        st.info("No existen pedidos registrados.")
        return

    df["Valor total"] = df["Valor total"].apply(dinero)

    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# REPORTES
# ============================================================

def pagina_reportes():
    header()

    html("""
    <div class="page-title">📊 Reportes</div>
    <div class="page-subtitle">
        Análisis del consumo de papelería por área y artículo.
    </div>
    """)

    conn = conectar()

    try:
        movimientos = pd.read_sql_query("""
            SELECT *
            FROM movimientos
            WHERE tipo = 'SALIDA'
        """, conn)
    finally:
        conn.close()

    if movimientos.empty:
        st.info("No existen consumos registrados.")
        return

    consumo_area = (
        movimientos
        .groupby("area", as_index=False)
        .agg({
            "cantidad": "sum",
            "valor_total": "sum"
        })
        .sort_values("cantidad", ascending=False)
    )

    st.subheader("🏢 Consumo por área")

    c1, c2 = st.columns(2)

    with c1:
        fig = px.bar(
            consumo_area,
            x="area",
            y="cantidad",
            title="Cantidad de artículos consumidos"
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(
            fig,
            use_container_width=True
        )

    with c2:
        fig = px.bar(
            consumo_area,
            x="area",
            y="valor_total",
            title="Valor consumido por área"
        )
        fig.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(
            fig,
            use_container_width=True
        )

    st.subheader("📦 Artículos más consumidos")

    consumo_item = (
        movimientos
        .groupby("item", as_index=False)
        .agg({
            "cantidad": "sum",
            "valor_total": "sum"
        })
        .sort_values("cantidad", ascending=False)
        .head(15)
    )

    fig = px.bar(
        consumo_item,
        x="cantidad",
        y="item",
        orientation="h",
        title="Top 15 artículos por cantidad"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    st.subheader("💰 Artículos con mayor valor de consumo")

    valor_item = (
        movimientos
        .groupby("item", as_index=False)["valor_total"]
        .sum()
        .sort_values("valor_total", ascending=False)
        .head(15)
    )

    fig = px.bar(
        valor_item,
        x="valor_total",
        y="item",
        orientation="h",
        title="Top 15 artículos por valor"
    )

    st.plotly_chart(
        fig,
        use_container_width=True
    )

    st.subheader("📋 Resumen por área")

    resumen = consumo_area.copy()
    resumen["valor_total"] = resumen["valor_total"].apply(dinero)

    resumen.columns = [
        "Área",
        "Cantidad consumida",
        "Valor consumido"
    ]

    st.dataframe(
        resumen,
        use_container_width=True,
        hide_index=True
    )


# ============================================================
# COMPRAS
# ============================================================

def pagina_compras():
    header()

    html("""
    <div class="page-title">🛒 Registrar compra</div>
    <div class="page-subtitle">
        Las nuevas compras se suman al inventario existente.
    </div>
    """)

    if st.session_state["rol"] != "Administrador":
        st.error(
            "Esta sección es exclusiva del administrador."
        )
        return

    df = obtener_inventario()

    if df.empty:
        st.warning("No existen artículos.")
        return

    etiquetas = etiquetas_articulos(df)

    item = st.selectbox(
        "Artículo",
        df["Ítem"].astype(str).tolist(),
        format_func=lambda x: etiquetas.get(str(x), str(x)),
        key="compra_item"
    )

    encontrado = df[
        df["Ítem"].astype(str) == str(item)
    ]

    if encontrado.empty:
        st.error("No se encontró el artículo.")
        return

    articulo = encontrado.iloc[0]

    html(f"""
    <div class="info-box">
        <strong>Descripción:</strong>
        {articulo["Descripción"]}<br>
        <strong>Inventario actual:</strong>
        {articulo["Cantidad disponible"]:,.0f}<br>
        <strong>Valor actual sin IVA:</strong>
        {dinero(articulo["Valor sin IVA"])}
    </div>
    """)

    cantidad = st.number_input(
        "Cantidad comprada",
        min_value=1,
        value=1,
        step=1
    )

    valor_sin_iva = st.number_input(
        "Nuevo valor unitario SIN IVA",
        min_value=0.0,
        value=float(articulo["Valor sin IVA"]),
        step=100.0
    )

    st.info(
        "Valor unitario con IVA: "
        +
        dinero(valor_sin_iva * (1 + IVA))
    )

    if st.button(
        "💾 Registrar compra",
        use_container_width=True
    ):

        correcto, resultado = registrar_compra(
            item,
            cantidad,
            valor_sin_iva,
            st.session_state["usuario"]
        )

        if correcto:
            st.success(
                "Compra registrada correctamente."
            )
            st.info(
                "Nuevo inventario: "
                +
                f"{resultado:,.0f}".replace(",", ".")
            )
            st.rerun()
        else:
            st.error(resultado)



# ============================================================
# EDICIÓN ADMINISTRATIVA DE SALIDAS
# ============================================================

def obtener_salidas():
    conn = conectar()
    try:
        return pd.read_sql_query("""
            SELECT
                m.id AS "ID",
                m.fecha AS "Fecha",
                m.tipo AS "Tipo",
                m.item AS "Código",
                COALESCE(i.descripcion, m.item) AS "Ítem",
                m.cantidad AS "Cantidad",
                m.valor_unitario AS "Valor unitario",
                m.valor_total AS "Valor total",
                m.area AS "Área",
                m.usuario AS "Usuario",
                m.observacion AS "Observación"
            FROM movimientos m
            LEFT JOIN inventario i ON CAST(i.item AS TEXT) = CAST(m.item AS TEXT)
            WHERE m.tipo IN ('SALIDA', 'AJUSTE_SALIDA')
            ORDER BY m.id DESC
        """, conn)
    finally:
        conn.close()


def actualizar_salida(movimiento_id, nueva_cantidad, nueva_area, usuario):
    conn = conectar()
    try:
        fila = conn.execute("""
            SELECT item, cantidad, valor_unitario, area, tipo, observacion
            FROM movimientos
            WHERE id = ?
        """, (movimiento_id,)).fetchone()

        if fila is None:
            return False, "No se encontró la salida."

        item, cantidad_anterior, valor_unitario, area_anterior, tipo, observacion = fila
        nueva_cantidad = float(nueva_cantidad)

        if nueva_cantidad <= 0:
            return False, "La cantidad debe ser mayor que cero."

        # Las salidas históricas importadas desde GASTO POR AREA representan
        # consumos ya ocurridos. No se modifica el inventario actual al editarlas.
        valor_total = nueva_cantidad * float(valor_unitario)

        conn.execute("""
            UPDATE movimientos
            SET cantidad = ?, valor_total = ?, area = ?, observacion = ?
            WHERE id = ?
        """, (
            nueva_cantidad,
            valor_total,
            nueva_area.strip(),
            f"{observacion or ''} | Modificada por {usuario}",
            movimiento_id
        ))

        conn.commit()
        return True, nueva_cantidad

    except Exception as e:
        conn.rollback()
        return False, str(e)
    finally:
        conn.close()


def importar_salidas_historicas(usuario):
    archivo = localizar_excel()
    if archivo is None:
        return False, "No se encontró el archivo Excel."

    clave = "GASTO_POR_AREA_HISTORICO_V1"
    conn = conectar()
    try:
        ya = conn.execute(
            "SELECT id FROM importaciones WHERE clave = ?", (clave,)
        ).fetchone()
        if ya:
            return False, "Las salidas históricas ya fueron importadas. No se volverán a duplicar."

        libro = pd.ExcelFile(archivo, engine="openpyxl")
        if "GASTO POR AREA" not in libro.sheet_names:
            return False, "El Excel no contiene la hoja 'GASTO POR AREA'."

        gasto = pd.read_excel(
            archivo, sheet_name="GASTO POR AREA", header=0, engine="openpyxl"
        )
        if gasto.empty or len(gasto.columns) < 19:
            return False, "La hoja GASTO POR AREA no tiene la estructura esperada."

        inventario = pd.read_sql_query(
            "SELECT item, descripcion, valor_con_iva FROM inventario", conn
        )
        if inventario.empty:
            return False, "No existe inventario cargado para relacionar las salidas."

        exactos = {}
        for _, r in inventario.iterrows():
            clave_desc = normalizar_texto(r["descripcion"])
            if clave_desc:
                exactos.setdefault(clave_desc, r)

        # A = descripción; C:R = áreas. S = Total general y se excluye.
        col_desc = gasto.columns[0]
        columnas_areas = list(gasto.columns[2:18])
        registros = 0
        no_encontrados = []

        for _, fila in gasto.iterrows():
            descripcion = normalizar_texto(fila[col_desc])
            if not descripcion or descripcion == "NAN":
                continue

            articulo = exactos.get(descripcion)
            if articulo is None:
                no_encontrados.append(str(fila[col_desc]))
                continue

            item = str(articulo["item"])
            valor_unitario = float(articulo["valor_con_iva"] or 0)

            for col_area in columnas_areas:
                cantidad = numero(fila[col_area])
                if cantidad <= 0:
                    continue

                area = str(col_area).strip()
                if not area or area.lower() == "nan":
                    continue

                conn.execute(
                    "INSERT OR IGNORE INTO areas (nombre) VALUES (?)", (area,)
                )
                conn.execute("""
                    INSERT INTO movimientos (
                        fecha, tipo, item, cantidad, valor_unitario,
                        valor_total, area, usuario, observacion
                    ) VALUES (?, 'SALIDA', ?, ?, ?, ?, ?, ?, ?)
                """, (
                    fecha_actual(), item, cantidad, valor_unitario,
                    cantidad * valor_unitario, area, usuario,
                    "Importación histórica desde GASTO POR AREA"
                ))
                registros += 1

        conn.execute("""
            INSERT INTO importaciones (clave, fecha, usuario, registros)
            VALUES (?, ?, ?, ?)
        """, (clave, fecha_actual(), usuario, registros))
        conn.commit()

        return True, {
            "registros": registros,
            "productos_no_encontrados": sorted(set(no_encontrados))
        }
    except Exception as e:
        conn.rollback()
        return False, f"Error importando salidas históricas: {e}"
    finally:
        conn.close()


def pagina_editar_salidas():
    header()

    html("""
    <div class="page-title">✏️ Editar salidas</div>
    <div class="page-subtitle">
        Modificación de salidas históricas. Esta opción es exclusiva del administrador. La edición de una salida histórica no modifica el inventario actual.
    </div>
    """)

    if st.session_state.get("rol") != "Administrador":
        st.error("No tiene permisos de administrador.")
        return

    salidas = obtener_salidas()

    if salidas.empty:
        st.info("No existen salidas registradas.")
        return

    salidas_mostrar = salidas.copy()
    salidas_mostrar["Valor unitario"] = salidas_mostrar["Valor unitario"].apply(dinero)
    salidas_mostrar["Valor total"] = salidas_mostrar["Valor total"].apply(dinero)

    st.dataframe(
        salidas_mostrar,
        use_container_width=True,
        hide_index=True
    )

    ids = salidas["ID"].astype(int).tolist()
    etiquetas_salidas = {
        int(r["ID"]): (
            f"#{int(r['ID'])} | {r['Ítem']} | {r['Área']} | "
            f"{float(r['Cantidad']):,.0f}".replace(",", ".") + " unidades"
        )
        for _, r in salidas.iterrows()
    }
    movimiento_id = st.selectbox(
        "Seleccione la salida que desea modificar",
        ids,
        format_func=lambda x: etiquetas_salidas.get(int(x), str(x))
    )

    registro = salidas[salidas["ID"] == movimiento_id].iloc[0]

    st.info(
        f"Ítem: {registro['Ítem']} | "
        f"Área actual: {registro['Área']} | "
        f"Cantidad actual: {registro['Cantidad']:,.0f}".replace(",", ".")
    )

    nueva_cantidad = st.number_input(
        "Nueva cantidad",
        min_value=1.0,
        value=float(registro["Cantidad"]),
        step=1.0
    )

    areas = obtener_areas()
    area_actual = str(registro["Área"])
    opciones = areas.copy()
    if area_actual and area_actual not in opciones:
        opciones.insert(0, area_actual)

    nueva_area = st.selectbox(
        "Nueva área",
        opciones,
        index=opciones.index(area_actual) if area_actual in opciones else 0
    )

    if st.button(
        "💾 Guardar modificación de salida",
        type="primary",
        use_container_width=True
    ):
        correcto, resultado = actualizar_salida(
            int(movimiento_id),
            nueva_cantidad,
            nueva_area,
            st.session_state["usuario"]
        )

        if correcto:
            st.session_state["flash_salida"] = {
                "tipo": "success",
                "mensaje": "La salida fue actualizada correctamente.",
                "detalle": "Nuevo inventario del artículo: " + f"{resultado:,.0f}".replace(",", ".")
            }
            st.rerun()
        else:
            st.error(resultado)


# ============================================================
# ADMINISTRACIÓN
# ============================================================

def pagina_administracion():
    header()

    html("""
    <div class="page-title">⚙️ Administración</div>
    <div class="page-subtitle">
        Administración del inventario, usuarios, áreas y kardex.
    </div>
    """)

    if st.session_state["rol"] != "Administrador":
        st.error(
            "No tiene permisos de administrador."
        )
        return

    tabs = st.tabs([
        "📦 Inventario",
        "✏️ Salidas",
        "📥 Excel",
        "➕ Artículo",
        "🏢 Áreas",
        "👥 Usuarios",
        "📋 Kardex"
    ])

    # --------------------------------------------------------
    # INVENTARIO
    # --------------------------------------------------------

    with tabs[0]:

        st.subheader("📦 Modificar inventario")

        df = obtener_inventario()

        if df.empty:
            st.warning("No existen artículos.")
        else:

            etiquetas = etiquetas_articulos(df)
            items_admin = df["Ítem"].astype(str).tolist()

            item = st.selectbox(
                "Artículo",
                items_admin,
                format_func=lambda x: etiquetas.get(str(x), str(x)),
                key="admin_item"
            )

            encontrado = df[
                df["Ítem"].astype(str) == str(item)
            ]

            if encontrado.empty:
                st.error("No se encontró el artículo.")
            else:

                articulo = encontrado.iloc[0]

                cantidad = st.number_input(
                    "Nueva cantidad disponible",
                    min_value=0.0,
                    value=float(
                        articulo["Cantidad disponible"]
                    ),
                    step=1.0,
                    key="admin_cantidad"
                )

                valor = st.number_input(
                    "Valor unitario SIN IVA",
                    min_value=0.0,
                    value=float(
                        articulo["Valor sin IVA"]
                    ),
                    step=100.0,
                    key="admin_valor"
                )

                stock = st.number_input(
                    "Stock mínimo",
                    min_value=0.0,
                    value=float(
                        articulo["Stock mínimo"]
                    ),
                    step=1.0,
                    key="admin_stock"
                )

                st.info(
                    "Valor con IVA: "
                    +
                    dinero(valor * (1 + IVA))
                )

                if st.button(
                    "💾 Guardar cambios",
                    use_container_width=True
                ):

                    correcto, resultado = actualizar_inventario(
                        item,
                        cantidad,
                        valor,
                        stock,
                        st.session_state["usuario"]
                    )

                    if correcto:
                        st.success(
                            "Inventario actualizado correctamente."
                        )
                        st.rerun()
                    else:
                        st.error(resultado)

    # --------------------------------------------------------
    # EXCEL
    # --------------------------------------------------------

    with tabs[2]:

        st.subheader("📥 Cargar inventario desde Excel")

        archivo = localizar_excel()

        if archivo:
            st.success(
                f"Excel encontrado: {archivo.name}"
            )
            st.code(str(archivo))
        else:
            st.error(
                "No se encontró ningún archivo .xlsx "
                "en la carpeta del programa."
            )

        html("""
        <div class="info-box">
            <strong>Estructura esperada:</strong><br>
            A → ÍTEM<br>
            B → DESCRIPCIÓN / ESPECIFICACIONES TÉCNICAS<br>
            C → CANTIDAD<br>
            D → VALOR UNITARIO SIN IVA
        </div>
        """)

        st.write(
            "La importación inicial agrega los artículos "
            "que no existan."
        )

        if st.button(
            "📥 Cargar inventario inicial",
            use_container_width=True
        ):

            correcto, resultado = importar_excel(
                reemplazar=False
            )

            if correcto:

                st.success(
                    "Importación realizada correctamente."
                )

                st.write(
                    f"Archivo: **{resultado['archivo']}**"
                )

                st.write(
                    f"Registros procesados: **{resultado['procesados']}**"
                )

                st.write(
                    f"Artículos nuevos: **{resultado['nuevos']}**"
                )

                st.rerun()

            else:
                st.error(resultado)

        st.divider()

        st.subheader("📊 Salidas históricas por área")
        st.write(
            "Importa una sola vez las salidas de la hoja GASTO POR AREA. "
            "La columna A se toma como producto, C:R como áreas y S (Total general) se excluye. "
            "Estas salidas se registran como histórico y no descuentan el inventario actual."
        )

        if st.button(
            "📊 Importar salidas históricas",
            use_container_width=True
        ):
            correcto, resultado = importar_salidas_historicas(
                st.session_state["usuario"]
            )
            if correcto:
                st.success(
                    f"Se importaron {resultado['registros']} movimientos históricos."
                )
                if resultado["productos_no_encontrados"]:
                    st.warning(
                        "Productos de GASTO POR AREA que no pudieron relacionarse: "
                        + "; ".join(resultado["productos_no_encontrados"][:20])
                    )
                st.rerun()
            else:
                st.warning(resultado)

        st.divider()

        st.subheader("🔄 Sincronizar Excel")

        st.warning(
            "Esta opción reemplaza las cantidades y valores "
            "de los artículos existentes por los datos del Excel."
        )

        if st.button(
            "🔄 Sincronizar Excel",
            use_container_width=True
        ):

            correcto, resultado = importar_excel(
                reemplazar=True
            )

            if correcto:

                st.success(
                    "Sincronización terminada."
                )

                st.write(
                    f"Registros procesados: **{resultado['procesados']}**"
                )

                st.write(
                    f"Artículos actualizados: **{resultado['actualizados']}**"
                )

                st.rerun()

            else:
                st.error(resultado)

    # --------------------------------------------------------
    # NUEVO ARTÍCULO
    # --------------------------------------------------------

    # --------------------------------------------------------
    # SALIDAS
    # --------------------------------------------------------
    with tabs[1]:
        pagina_editar_salidas()

    with tabs[3]:

        st.subheader("➕ Crear nuevo artículo")

        item = st.text_input(
            "Número / código del ítem"
        )

        descripcion = st.text_area(
            "Descripción"
        )

        cantidad = st.number_input(
            "Cantidad inicial",
            min_value=0.0,
            value=0.0,
            step=1.0
        )

        valor = st.number_input(
            "Valor unitario SIN IVA",
            min_value=0.0,
            value=0.0,
            step=100.0
        )

        stock = st.number_input(
            "Stock mínimo",
            min_value=0.0,
            value=10.0,
            step=1.0
        )

        if st.button(
            "➕ Crear artículo",
            use_container_width=True
        ):

            if not item.strip():
                st.error("Ingrese el código del ítem.")
            elif not descripcion.strip():
                st.error("Ingrese la descripción.")
            else:

                conn = conectar()

                try:

                    conn.execute("""
                        INSERT INTO inventario (
                            item,
                            descripcion,
                            cantidad,
                            valor_sin_iva,
                            iva,
                            valor_con_iva,
                            stock_minimo,
                            fecha_actualizacion
                        )
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    """, (
                        item.strip(),
                        descripcion.strip(),
                        cantidad,
                        valor,
                        IVA,
                        valor * (1 + IVA),
                        stock,
                        fecha_actual()
                    ))

                    conn.commit()

                    st.success(
                        "Artículo creado correctamente."
                    )

                    st.rerun()

                except sqlite3.IntegrityError:
                    st.error(
                        "Ese artículo ya existe."
                    )

                finally:
                    conn.close()

    # --------------------------------------------------------
    # ÁREAS
    # --------------------------------------------------------

    with tabs[4]:

        st.subheader("🏢 Administración de áreas")

        nueva_area = st.text_input(
            "Nueva área"
        )

        if st.button(
            "➕ Agregar área"
        ):

            if not nueva_area.strip():
                st.error("Ingrese el nombre del área.")
            else:

                conn = conectar()

                try:

                    conn.execute(
                        "INSERT INTO areas (nombre) VALUES (?)",
                        (nueva_area.strip(),)
                    )

                    conn.commit()

                    st.success(
                        "Área agregada correctamente."
                    )

                    st.rerun()

                except sqlite3.IntegrityError:
                    st.error("El área ya existe.")

                finally:
                    conn.close()

        conn = conectar()

        try:
            areas = pd.read_sql_query("""
                SELECT
                    id AS "ID",
                    nombre AS "Área"
                FROM areas
                ORDER BY nombre
            """, conn)
        finally:
            conn.close()

        st.dataframe(
            areas,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # USUARIOS
    # --------------------------------------------------------

    with tabs[5]:

        st.subheader("👥 Crear usuario")

        usuario = st.text_input(
            "Usuario nuevo"
        )

        password = st.text_input(
            "Contraseña",
            type="password"
        )

        rol = st.selectbox(
            "Rol",
            ["Usuario", "Administrador"]
        )

        if st.button(
            "➕ Crear usuario"
        ):

            if not usuario.strip():
                st.error("Ingrese el usuario.")
            elif not password:
                st.error("Ingrese la contraseña.")
            else:

                conn = conectar()

                try:

                    conn.execute("""
                        INSERT INTO usuarios (
                            usuario,
                            password,
                            rol
                        )
                        VALUES (?, ?, ?)
                    """, (
                        usuario.strip(),
                        hash_password(password),
                        rol
                    ))

                    conn.commit()

                    st.success(
                        "Usuario creado correctamente."
                    )

                    st.rerun()

                except sqlite3.IntegrityError:
                    st.error(
                        "El usuario ya existe."
                    )

                finally:
                    conn.close()

        conn = conectar()

        try:
            usuarios = pd.read_sql_query("""
                SELECT
                    id AS "ID",
                    usuario AS "Usuario",
                    rol AS "Rol"
                FROM usuarios
                ORDER BY usuario
            """, conn)
        finally:
            conn.close()

        st.dataframe(
            usuarios,
            use_container_width=True,
            hide_index=True
        )

    # --------------------------------------------------------
    # KARDEX
    # --------------------------------------------------------

    with tabs[6]:

        st.subheader("📋 Kardex / movimientos")

        conn = conectar()

        try:
            movimientos = pd.read_sql_query("""
                SELECT
                    fecha AS "Fecha",
                    tipo AS "Tipo",
                    item AS "Ítem",
                    cantidad AS "Cantidad",
                    valor_unitario AS "Valor unitario",
                    valor_total AS "Valor total",
                    area AS "Área",
                    usuario AS "Usuario",
                    observacion AS "Observación"
                FROM movimientos
                ORDER BY id DESC
            """, conn)
        finally:
            conn.close()

        if movimientos.empty:
            st.info("No existen movimientos.")
        else:

            movimientos["Valor unitario"] = (
                movimientos["Valor unitario"].apply(dinero)
            )

            movimientos["Valor total"] = (
                movimientos["Valor total"].apply(dinero)
            )

            st.dataframe(
                movimientos,
                use_container_width=True,
                hide_index=True
            )


# ============================================================
# SIDEBAR
# ============================================================

def sidebar():

    with st.sidebar:

        html("""
        <div class="sidebar-logo">
            <div class="logo">🏛️</div>
            <h2>ITM</h2>
            <p>Institución Universitaria</p>
        </div>
        """)

        st.divider()

        html("""
        <div style="
            text-align:center;
            font-size:19px;
            font-weight:800;
            margin-bottom:15px;
        ">
            📦 PAPELERÍA ITM
        </div>
        """)

        opciones = [
            "🏠 Inicio",
            "📦 Inventario",
            "💰 Valores e IVA",
            "🛒 Realizar pedido",
            "📋 Historial de pedidos",
            "📊 Reportes"
        ]

        if st.session_state["rol"] == "Administrador":
            opciones += [
                "🛒 Registrar compra",
                "⚙️ Administración"
            ]

        pagina = st.radio(
            "MENÚ PRINCIPAL",
            opciones
        )

        st.divider()

        html(f"""
        <div class="user-card">
            <strong>👤 {st.session_state["usuario"]}</strong><br>
            <small>🔐 {st.session_state["rol"]}</small>
        </div>
        """)

        st.write("")

        if st.button(
            "🚪 Cerrar sesión",
            use_container_width=True
        ):

            st.session_state["logueado"] = False
            st.session_state.pop("usuario", None)
            st.session_state.pop("rol", None)
            st.session_state["carrito"] = []

            st.rerun()

        return pagina


# ============================================================
# MAIN
# ============================================================

def main():

    # Crear base y migrar automáticamente
    crear_base_datos()

    # Importación automática si la base está vacía
    try:
        df = obtener_inventario()

        if df.empty:
            importar_excel(reemplazar=False)

    except Exception as e:

        st.error(
            "Error inicializando el inventario."
        )

        st.exception(e)

        return

    if "logueado" not in st.session_state:
        st.session_state["logueado"] = False

    if "carrito" not in st.session_state:
        st.session_state["carrito"] = []

    if "flash_salida" in st.session_state:
        flash = st.session_state.pop("flash_salida")
        st.success(f"{flash['mensaje']} {flash['detalle']}")

    if not st.session_state["logueado"]:
        login()
        return

    pagina = sidebar()

    if pagina == "🏠 Inicio":
        pagina_inicio()

    elif pagina == "📦 Inventario":
        pagina_inventario()

    elif pagina == "💰 Valores e IVA":
        pagina_iva()

    elif pagina == "🛒 Realizar pedido":
        pagina_pedidos()

    elif pagina == "📋 Historial de pedidos":
        pagina_historial()

    elif pagina == "📊 Reportes":
        pagina_reportes()

    elif pagina == "🛒 Registrar compra":
        pagina_compras()

    elif pagina == "⚙️ Administración":
        pagina_administracion()


# ============================================================
# ARRANQUE
# ============================================================

if __name__ == "__main__":
    main()
