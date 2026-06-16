import streamlit as st
import pandas as pd
import re
import os
import glob
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ─────────────────────────────────────────────
# CONFIGURACIÓN INICIAL
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Control de Proyectos | Techos Rentables",
    page_icon="☀️",
    layout="centered"
)

# Ocultar elementos de Streamlit
st.markdown("""
    <style>
    [data-testid="stHeader"]       {display: none !important;}
    [data-testid="stDecoration"]   {display: none !important;}
    [data-testid="stToolbar"]      {display: none !important;}
    [data-testid="stFooter"]       {display: none !important;}
    [data-testid="stViewerBadge"] {display: none !important;}
    footer                         {display: none !important;}
    [data-testid="manage-app-button"] {display: none !important;}
    #MainMenu                      {visibility: hidden !important;}
    .stDeployButton                {display: none !important;}
    </style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────
PREFIJO_GEEST     = "Proceso Comercial _ Operativo"  # Nombre exacto que arroja Geest
MAX_INTENTOS_CEDULA = 3

# ─────────────────────────────────────────────
# HELPERS: LOCALIZACIÓN DEL ARCHIVO DE DATOS
# ─────────────────────────────────────────────

def encontrar_archivo_datos() -> str | None:
    """
    Busca todos los archivos de Excel y retorna el más reciente.
    Como Geest le pone fecha al nombre, siempre tomará el último exportado.
    """
    candidatos = glob.glob("*.xlsx")
    validos = [f for f in candidatos if f.startswith(PREFIJO_GEEST) or f.startswith("base_datos")]
    
    if validos:
        # El más reciente por fecha de modificación del sistema de archivos
        return max(validos, key=os.path.getmtime)

    return None

# ─────────────────────────────────────────────
# MOTOR DE CÁLCULO PONDERADO
# ─────────────────────────────────────────────

def calcular_avance_ponderado(valor, peso: float, es_porcentaje: bool) -> float:
    if valor == "vacio" or pd.isna(valor) or str(valor).strip() == "":
        return 0.0

    if not es_porcentaje:
        return float(peso)

    val_str = str(valor)
    porcentajes = [int(x) for x in re.findall(r'(\d+)%', val_str)]

    if porcentajes:
        return (min(sum(porcentajes), 100) / 100.0) * float(peso)

    val_lower = val_str.lower()
    palabras_ok = ["si", "sí", "liquidado", "pagado", "entregado",
                   "finalizado", "aprobado", "ok"]
    if any(p in val_lower for p in palabras_ok):
        return float(peso)

    return float(peso) / 2.0

# ─────────────────────────────────────────────
# CARGA DE BASE DE DATOS
# ─────────────────────────────────────────────

@st.cache_data
def cargar_datos():
    ruta = encontrar_archivo_datos()
    if ruta is None:
        return None, None

    df = pd.read_excel(ruta)
    df['Folio'] = df['Folio'].astype(str).str.strip()
    df = df.fillna("vacio")
    return df, ruta   # retornamos también la ruta para mostrársela al admin

# ─────────────────────────────────────────────
# ETAPAS DEL PROYECTO
# ─────────────────────────────────────────────

ETAPAS = [
    {"nombre": "Firma de Contrato y Anticipo",    "col": "Fecha Contrato (RV)",                                  "peso": 1,  "es_porcentaje": False},
    {"nombre": "Cronograma de Instalación",        "col": "Cronograma Instalación",                               "peso": 2,  "es_porcentaje": False},
    {"nombre": "Diseños Eléctricos",               "col": "Plano Diseños Electricos (i)",                         "peso": 5,  "es_porcentaje": False},
    {"nombre": "Órdenes de Compras Odoo",          "col": "Ordenes de Compras Odoo",                              "peso": 10, "es_porcentaje": False},
    {"nombre": "Pagos Órdenes",                    "col": "% Ordenes de Compras",                                 "peso": 10, "es_porcentaje": True},
    {"nombre": "Instalación: Puesta de Materiales","col": "Observación (Puesta de materiales en sitio)",          "peso": 10, "es_porcentaje": False},
    {"nombre": "Instalación: Sistema de Captación","col": "% Instalación Sistema de Captación",                   "peso": 20, "es_porcentaje": True},
    {"nombre": "Instalación: Cableado DC",         "col": "% Instalación Cableado DC",                            "peso": 10, "es_porcentaje": True},
    {"nombre": "Instalación: Micros e Inversores", "col": "% Instalación Micros / Inversores",                    "peso": 10, "es_porcentaje": True},
    {"nombre": "Instalación: Cableado AC",         "col": "% Instalación Cableado AC",                            "peso": 10, "es_porcentaje": True},
    {"nombre": "Adecuación de Equipo de Medida",   "col": "% Adecuación Equipo de Medida",                        "peso": 3,  "es_porcentaje": True},
    {"nombre": "Certificación RETIE",              "col": "Solicitud RETIE",                                      "peso": 3,  "es_porcentaje": False},
    {"nombre": "Cierre y Pruebas de Sistemas",     "col": "Comprobante RETIE",                                    "peso": 3,  "es_porcentaje": False},
    {"nombre": "Interconexión Operador de Red",    "col": "Fecha entrega Sistema",                                 "peso": 3,  "es_porcentaje": False},
]

FOTOS_POR_ETAPA = {
    "Instalación: Puesta de Materiales":    "Registro Fotografico - Instalación Puesta de Materiales en Sitio",
    "Instalación: Sistema de Captación":    "Registro fotografico - Instalación Sistema de Captación",
    "Instalación: Cableado DC":             "Registro Fotografico instalación - Instalación cableado DC",
    "Instalación: Micros e Inversores":     "Registro fotografico - Instalación Micros / Inversores",
    "Instalación: Cableado AC":             "Registro Fotografico - Instalación AC",
}

# ─────────────────────────────────────────────
# SIDEBAR – MÓDULO ADMIN
# ─────────────────────────────────────────────

def _secrets_ok() -> bool:
    """Verifica secrets de admin. EMAIL_PASS se valida solo al enviar PQR."""
    try:
        _ = st.secrets["ADMIN_USER"]
        _ = st.secrets["ADMIN_PASS"]
        return True
    except KeyError:
        return False

with st.sidebar:
    st.write("### 🔒 Centro de Gestión PMO")
    st.write("Solo para actualización de base de datos.")

    admin_user  = st.text_input("Usuario", key="sb_user")
    admin_pass  = st.text_input("Contraseña", type="password", key="sb_pass")

    acceso_admin = False
    if admin_user or admin_pass:
        if not _secrets_ok():
            st.error("⚠️ Secrets no configurados en el servidor.")
        else:
            try:
                if (admin_user == st.secrets["ADMIN_USER"] and
                        admin_pass == st.secrets["ADMIN_PASS"]):
                    acceso_admin = True
                    st.success("✅ Acceso concedido")
                else:
                    st.error("❌ Credenciales incorrectas")
            except Exception:
                st.error("⚠️ Error al validar credenciales.")

    if acceso_admin:
        st.divider()

        # Mostrar archivo actualmente en uso
        _, ruta_actual = cargar_datos()
        if ruta_actual:
            st.info(f"📂 Archivo activo: `{os.path.basename(ruta_actual)}`")
        else:
            st.warning("⚠️ No hay archivo de datos activo.")

        st.write("**Actualizar Base de Datos (Geest)**")
        st.caption("Sube el archivo Excel original exportado de Geest.")

        archivo_subido = st.file_uploader("Archivo Excel (.xlsx)", type=["xlsx"])

        if archivo_subido is not None:
            try:
                # ESTRATEGIA ANTIBLOQUEO DEFINITIVA: Guardar con el nombre original de Geest
                nombre_original = archivo_subido.name
                
                with open(nombre_original, "wb") as f:
                    f.write(archivo_subido.getbuffer())
                
                st.cache_data.clear()
                st.success(f"🔄 Archivo '{nombre_original}' cargado con éxito. Los clientes ya ven la nueva información.")
            except Exception as e:
                st.error(f"⚠️ Error al guardar el archivo: {e}")

# ─────────────────────────────────────────────
# INTERFAZ PRINCIPAL
# ─────────────────────────────────────────────

if os.path.exists("EDITABLES-TRN.png"):
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.image("EDITABLES-TRN.png", use_container_width=True)

st.markdown("<h1 style='text-align:center;'>Control de Proyectos</h1>",       unsafe_allow_html=True)
st.markdown("<p style='text-align:center;'>Ingresa tu número de folio para conocer el estado de tu sistema solar.</p>", unsafe_allow_html=True)
st.divider()

# ── Inicializar estado de intentos ──────────────────────────────────────────
if "intentos_cedula" not in st.session_state or st.session_state.intentos_cedula is None:
    st.session_state.intentos_cedula = 0
if "folio_actual" not in st.session_state:
    st.session_state.folio_actual = None
if "folios_bloqueados" not in st.session_state:
    st.session_state.folios_bloqueados = []

# ── Cargar datos ─────────────────────────────────────────────────────────────
df, _ = cargar_datos()

if df is None:
    st.warning("⏳ Estamos actualizando nuestra base de datos. Por favor, intenta de nuevo en unos minutos.")
    st.stop()

# ── Búsqueda por folio ───────────────────────────────────────────────────────
folio_input = st.text_input("Número de Folio", placeholder="Ej: 10293")

if not folio_input:
    st.stop()

# Limpiar el input para la búsqueda
folio_busqueda = folio_input.strip()

with st.spinner("Buscando tu proyecto..."):
    proyecto = df[df['Folio'] == folio_busqueda]

if proyecto.empty:
    st.error("❌ No se encontró ningún proyecto asociado a ese Folio.")
    st.stop()

datos = proyecto.iloc[0]

# ── Bloqueo de Seguridad Permanente en la Sesión ─────────────────────────────
if folio_busqueda in st.session_state.folios_bloqueados:
    st.error(
        "🔒 Has superado el número de intentos permitidos para este folio. "
        "Por favor comunícate con nuestra línea de atención al cliente."
    )
    st.stop()

# ── Reseteo de Intentos si el usuario cambia de folio ────────────────────────
if st.session_state.folio_actual != folio_busqueda:
    st.session_state.intentos_cedula = 0
    st.session_state.folio_actual = folio_busqueda

# ── Validación de identidad (cédula/NIT) ─────────────────────────────────────
intentos_restantes = MAX_INTENTOS_CEDULA - st.session_state.intentos_cedula
st.info(f"🔒 Por seguridad, necesitamos validar tu identidad. ({intentos_restantes} intento(s) restante(s))")

cedula_db       = str(datos.get('Cedula Cliente (RV)', 'vacio'))
cedula_db_limpia = re.sub(r'[\.\,\s]', '', cedula_db)

cedula_input = st.text_input("Ingresa tu número de Cédula o NIT (sin puntos)", type="password")

if not cedula_input:
    st.stop()

cedula_input_limpia = re.sub(r'[\.\,\s]', '', cedula_input.strip())

if cedula_input_limpia != cedula_db_limpia or cedula_db_limpia == "vacio":
    st.session_state.intentos_cedula += 1
    intentos_restantes_ahora = MAX_INTENTOS_CEDULA - st.session_state.intentos_cedula
    
    if intentos_restantes_ahora > 0:
        st.error(f"❌ El documento ingresado no coincide. Te quedan {intentos_restantes_ahora} intento(s).")
    else:
        st.session_state.folios_bloqueados.append(folio_busqueda)
        st.error(
            "🔒 Has superado el número de intentos permitidos. "
            "Por favor comunícate con nuestra línea de atención para actualizar tus datos."
        )
    st.stop()

# ─────────────────────────────────────────────
# IDENTIDAD VERIFICADA — MOSTRAR PROYECTO
# ─────────────────────────────────────────────
st.session_state.intentos_cedula = 0

st.success("¡Identidad verificada exitosamente!")
st.subheader(f"👤 Cliente: {datos['Nombre Cliente']}")
st.divider()

# ── Calcular avance ───────────────────────────────────────────────────────────
puntaje_total = 0.0
etapa_actual  = "Inicio de Proyecto"
desglose      = []

for etapa in ETAPAS:
    avance_fase = calcular_avance_ponderado(
        datos.get(etapa["col"], "vacio"),
        etapa["peso"],
        etapa["es_porcentaje"]
    )
    puntaje_total += avance_fase
    completada = avance_fase >= etapa["peso"]
    en_curso   = 0 < avance_fase < etapa["peso"]

    if avance_fase > 0:
        etapa_actual = etapa["nombre"]

    icono = "✅" if completada else ("🔄" if en_curso else "⏳")
    desglose.append((icono, etapa["nombre"], etapa["peso"], round(avance_fase, 1)))

porcentaje_global = min(int(puntaje_total), 100)

# ── Barra de progreso global ──────────────────────────────────────────────────
st.write("### 📊 Estado General del Proyecto")
st.progress(porcentaje_global / 100.0)
st.info(f"**Progreso Total:** {porcentaje_global}% completado\n\n**Etapa Actual:** {etapa_actual}")

# ── Desglose visual de etapas ─────────────────────────────────────────────────
with st.expander("📋 Ver detalle de todas las etapas"):
    for icono, nombre, peso, avance in desglose:
        pct_etapa = round((avance / peso) * 100) if peso else 0
        st.markdown(f"{icono} **{nombre}** — {pct_etapa}%")

st.divider()

# ── Galería fotográfica ───────────────────────────────────────────────────────
if etapa_actual in FOTOS_POR_ETAPA:
    col_fotos = FOTOS_POR_ETAPA[etapa_actual]
    fotos_data = datos.get(col_fotos, "vacio")

    if pd.notna(fotos_data) and fotos_data != "vacio" and str(fotos_data).strip():
        st.write(f"### 📸 Evidencia en campo: {etapa_actual}")
        links_fotos = [l for l in re.split(r'[\s,;\n]+', str(fotos_data)) if l.startswith("http")]

        if links_fotos:
            cols_img = st.columns(3)
            for idx, link in enumerate(links_fotos):
                with cols_img[idx % 3]:
                    st.image(link, use_container_width=True)
        else:
            st.info(f"Las fotografías de la etapa actual ({etapa_actual}) se encuentran en proceso de carga.")
    else:
        st.info(f"Las fotografías de la etapa actual ({etapa_actual}) se encuentran en proceso de carga.")

st.divider()

# ── Formulario PQR ────────────────────────────────────────────────────────────
st.write("### 📩 ¿Tienes alguna duda o solicitud?")

with st.form(key="form_pqr", clear_on_submit=True):
    tipo_solicitud = st.selectbox(
        "Tipo de Solicitud",
        ["Petición (Duda o Consulta)", "Queja", "Reclamo", "Sugerencia"]
    )
    asunto  = st.text_input("Asunto", placeholder="Ej: Duda con las fechas de montaje")
    mensaje = st.text_area("Detalle de tu solicitud")
    boton_enviar = st.form_submit_button("Enviar Solicitud")

    if boton_enviar:
        if not asunto or not mensaje:
            st.error("⚠️ Completa todos los campos.")
        else:
            try:
                correo_remitente = "atencionalcliente@techosrentables.com"
                password_remitente = st.secrets["EMAIL_PASS"]
                correo_destino = "atencionalcliente@techosrentables.com"

                timestamp = datetime.now().strftime("%Y-%m-%d %H:%M")

                msg = MIMEMultipart()
                msg['From']    = correo_remitente
                msg['To']      = correo_destino
                msg['Subject'] = (
                    f"Nueva PQR ({tipo_solicitud}) | "
                    f"Folio: {folio_input} | "
                    f"{datos['Nombre Cliente']} | {timestamp}"
                )

                cuerpo_correo = f"""
Se ha recibido una nueva solicitud de atención al cliente.

📅 Fecha y hora : {timestamp}
👤 Cliente      : {datos['Nombre Cliente']}
📄 Folio        : {folio_input}
📋 Tipo         : {tipo_solicitud}
📌 Asunto       : {asunto}

📝 Mensaje:
{mensaje}
                """.strip()

                msg.attach(MIMEText(cuerpo_correo, 'plain'))

                server = smtplib.SMTP_SSL('mail.techosrentables.com', 465)
                server.login(correo_remitente, password_remitente)
                server.send_message(msg)
                server.quit()

                st.success("✅ Tu solicitud ha sido enviada con éxito. Nos comunicaremos contigo pronto.")

            except KeyError:
                st.error("⚠️ Configuración de correo no disponible. Contáctanos directamente.")
            except smtplib.SMTPAuthenticationError:
                st.error("⚠️ Error de autenticación en el servidor de correo.")
            except smtplib.SMTPException as e:
                st.error(f"⚠️ Error al enviar el correo. Detalle técnico: {e}")
            except Exception as e:
                st.error(f"⚠️ Error inesperado: {e}")
