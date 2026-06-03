import streamlit as st
import pandas as pd
import re
import os
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Configuración inicial
st.set_page_config(
    page_title="Control de Proyectos | Techos Rentables", 
    page_icon="☀️", 
    layout="centered"
)

ARCHIVO_BD = "base_datos.xlsx"

# --- MOTOR DE CÁLCULO PONDERADO (PMO) ---
def calcular_avance_ponderado(valor, peso, es_porcentaje):
    if valor == "vacio" or pd.isna(valor) or str(valor).strip() == "":
        return 0.0
    if not es_porcentaje:
        return float(peso)
    val_str = str(valor)
    porcentajes = [int(x) for x in re.findall(r'(\d+)%', val_str)]
    if porcentajes:
        promedio = sum(porcentajes) / len(porcentajes)
        return (promedio / 100.0) * peso
    val_lower = val_str.lower()
    if any(palabra in val_lower for palabra in ["si", "sí", "liquidado", "pagado", "entregado", "finalizado", "aprobado", "ok"]):
        return float(peso)
    return float(peso) / 2.0

@st.cache_data
def cargar_datos():
    if os.path.exists(ARCHIVO_BD):
        df = pd.read_excel(ARCHIVO_BD)
    elif os.path.exists("Proceso Comercial _ Operativo [2026-Jun-01 10.21].xlsx"):
        df = pd.read_excel("Proceso Comercial _ Operativo [2026-Jun-01 10.21].xlsx")
    else:
        return None
        
    df['Folio'] = df['Folio'].astype(str).str.strip()
    df = df.fillna("vacio")
    return df

# --- BARRA LATERAL OCULTA: MÓDULO PMO / ADMINISTRADOR ---
with st.sidebar:
    st.write("### 🔒 Acceso Centro de Gestión")
    st.write("Solo para actualización de base de datos.")
    
    admin_user = st.text_input("Usuario")
    admin_pass = st.text_input("Contraseña", type="password")
    
    # AQUÍ LLAMAMOS A LA CONTRASEÑA SECRETA DE ADMINISTRADOR
    if admin_user == "admin" and admin_pass == st.secrets["ADMIN_PASS"]:
        st.success("✅ Acceso concedido")
        st.divider()
        st.write("**Actualizar Base de Datos (Geest)**")
        
        archivo_subido = st.file_uploader("Sube el archivo Excel más reciente", type=["xlsx"])
        
        if archivo_subido is not None:
            with open(ARCHIVO_BD, "wb") as f:
                f.write(archivo_subido.getbuffer())
            
            st.cache_data.clear()
            st.success("🔄 Base de datos actualizada correctamente. Los clientes ya ven la nueva información.")
    elif admin_user or admin_pass:
        st.error("❌ Credenciales incorrectas")

# --- INTERFAZ PRINCIPAL (CLIENTES) ---
st.title("☀️ Control de Proyectos")
st.write("Ingresa tu número de folio para conocer el estado de tu sistema solar.")

df = cargar_datos()

if df is None:
    st.warning("⏳ Estamos actualizando nuestra base de datos. Por favor, intenta de nuevo en unos minutos.")
else:
    folio_input = st.text_input("Número de Folio", placeholder="Ej: 10293")

    if folio_input:
        proyecto = df[df['Folio'] == folio_input.strip()]
        
        if not proyecto.empty:
            datos = proyecto.iloc[0]
            st.success("¡Proyecto localizado!")
            st.subheader(f"👤 Cliente: {datos['Nombre Cliente']}")
            st.divider()
            
            etapas_mapeo = [
                {"nombre": "Firma de Contrato y Anticipo", "col": "Fecha Contrato (RV)", "peso": 1, "es_porcentaje": False},
                {"nombre": "Cronograma de Instalación", "col": "Cronograma Instalación", "peso": 2, "es_porcentaje": False},
                {"nombre": "Diseños Eléctricos", "col": "Plano Diseños Electricos (i)", "peso": 5, "es_porcentaje": False},
                {"nombre": "Órdenes de Compras Odoo", "col": "Ordenes de Compras Odoo", "peso": 10, "es_porcentaje": False},
                {"nombre": "Pagos Órdenes", "col": "% Ordenes de Compras", "peso": 10, "es_porcentaje": True},
                {"nombre": "Instalación: Puesta de Materiales", "col": "Observación (Puesta de materiales en sitio)", "peso": 10, "es_porcentaje": False},
                {"nombre": "Instalación: Sistema de Captación", "col": "% Instalación Sistema de Captación", "peso": 20, "es_porcentaje": True},
                {"nombre": "Instalación: Cableado DC", "col": "% Instalación Cableado DC", "peso": 10, "es_porcentaje": True},
                {"nombre": "Instalación: Micros e Inversores", "col": "% Instalación Micros / Inversores", "peso": 10, "es_porcentaje": True},
                {"nombre": "Instalación: Cableado AC", "col": "% Instalación Cableado AC", "peso": 10, "es_porcentaje": True},
                {"nombre": "Adecuación de Equipo de Medida", "col": "% Adecuación Equipo de Medida", "peso": 3, "es_porcentaje": True},
                {"nombre": "Certificación RETIE", "col": "Solicitud RETIE", "peso": 3, "es_porcentaje": False},
                {"nombre": "Cierre y Pruebas de Sistemas", "col": "Comprobante RETIE", "peso": 3, "es_porcentaje": False},
                {"nombre": "Interconexión Operador de Red", "col": "Fecha entrega Sistema", "peso": 3, "es_porcentaje": False}
            ]
            
            puntaje_total = 0.0
            etapa_actual = "Inicio de Proyecto"
            
            for etapa in etapas_mapeo:
                avance_fase = calcular_avance_ponderado(datos.get(etapa["col"], "vacio"), etapa["peso"], etapa["es_porcentaje"])
                puntaje_total += avance_fase
                if avance_fase > 0:
                    etapa_actual = etapa["nombre"]
            
            porcentaje_global = min(int(puntaje_total), 100)
            
            st.write("### 📊 Estado General del Proyecto")
            st.progress(porcentaje_global / 100.0)
            st.info(f"**Progreso Total:** {porcentaje_global}% completado\n\n**Etapa Actual:** {etapa_actual}")
            
            st.divider()
            
            # --- FORMULARIO DE PQR ---
            st.write("### 📩 ¿Tienes alguna duda o solicitud?")
            with st.form(key="form_pqr", clear_on_submit=True):
                tipo_solicitud = st.selectbox("Tipo de Solicitud", ["Petición (Duda o Consulta)", "Queja", "Reclamo", "Sugerencia"])
                asunto = st.text_input("Asunto", placeholder="Ej: Duda con las fechas de montaje")
                mensaje = st.text_area("Detalle de tu solicitud")
                boton_enviar = st.form_submit_button("Enviar Solicitud")
                
                if boton_enviar and asunto and mensaje:
                    correo_remitente = "atencionalcliente@techosrentables.com" 
                    
                    # AQUÍ LLAMAMOS A LA CONTRASEÑA SECRETA DEL CORREO
                    password_remitente = st.secrets["EMAIL_PASS"] 
                    
                    correo_destino = "atencionalcliente@techosrentables.com"
                    
                    msg = MIMEMultipart()
                    msg['From'] = correo_remitente
                    msg['To'] = correo_destino
                    msg['Subject'] = f"Nueva PQR ({tipo_solicitud}) - Folio: {folio_input} - {datos['Nombre Cliente']}"
                    
                    cuerpo_correo = f"""
                    Se ha recibido una nueva solicitud:
                    👤 Cliente: {datos['Nombre Cliente']}
                    📄 Folio: {folio_input}
                    📋 Tipo: {tipo_solicitud}
                    📌 Asunto: {asunto}
                    📝 Mensaje: {mensaje}
                    """
                    msg.attach(MIMEText(cuerpo_correo, 'plain'))
                    
                    try:
                        server = smtplib.SMTP_SSL('mail.techosrentables.com', 465)
                        server.login(correo_remitente, password_remitente)
                        server.send_message(msg)
                        server.quit()
                        st.success("✅ Tu solicitud ha sido enviada con éxito.")
                    except Exception as e:
                        st.error(f"⚠️ Hubo un error al enviar el correo. Detalle: {e}")
                elif boton_enviar:
                    st.error("⚠️ Completa todos los campos.")
        else:
            st.error("❌ No se encontró ningún proyecto asociado a ese Folio.")
