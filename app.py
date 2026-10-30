# ============================================================
# SISTEMA DE COTIZACIONES - Streamlit (PROGAS SAS)
# ============================================================
import streamlit as st
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from fpdf import FPDF
from datetime import datetime
import gspread
from google.oauth2 import service_account
import json
import os

# ============================================================
# CONFIGURACIÓN
# ============================================================
URL_SERVICIOS = "https://docs.google.com/spreadsheets/d/e/2PACX-1vRyvT2p0fJ5gawd-itBIG5jbOZXy-Bz4JQvb7t62CeOhTrfzEyUX6d-g7vT5Wm-Xs1acHRx2mnje7dE/pub?gid=60935932&single=true&output=csv"
EMAIL_REMITENTE = "progas.esp@gmail.com"
CONTRASENA_APP = "oyjy rvqa fkdn tvds"

OPERARIOS = [
    "Luis Villalba - Subgerente",
    "Osaris Cruz - Gerente",
    "Guillermo Solo - Técnico",
    "Gustavo Vidal - Analista"
]

st.set_page_config(page_title="Sistema de Cotizaciones PROGAS SAS ESP", page_icon="📋", layout="centered")

# ============================================================
# FUNCIONES
# ============================================================
def limpiar_precio(valor):
    if pd.isna(valor):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace('$', '').replace(' ', '')
    if texto == '' or texto == '0':
        return 0.0
    if ',' in texto and '.' in texto:
        texto = texto.replace('.', '').replace(',', '.')
    elif ',' in texto:
        texto = texto.replace(',', '.')
    elif '.' in texto:
        partes = texto.split('.')
        if len(partes) == 2 and len(partes[1]) <= 2:
            pass
        else:
            texto = texto.replace('.', '')
    try:
        return float(texto)
    except:
        return 0.0

@st.cache_data(ttl=300)
def cargar_servicios():
    try:
        df = pd.read_csv(URL_SERVICIOS, header=0, encoding='utf-8')
        df = df.dropna(how='all')
        df.columns = [col.strip().upper() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Error al cargar servicios: {e}")
        return pd.DataFrame()

def generar_pdf(cotizacion, lista_items, total):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 18)
    pdf.cell(0, 10, "PROGAS SAS ESP", ln=True, align="C")
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 6, f"NIT. 9018087574", ln=True, align="C")
    pdf.ln(3)
    pdf.set_font("Arial", "B", 14)
    pdf.cell(0, 8, "COTIZACION", ln=True, align="C")
    pdf.ln(5)
    pdf.set_font("Arial", "", 11)
    pdf.cell(0, 6, f"Numero: {cotizacion['numero']}", ln=True)
    pdf.cell(0, 6, f"Fecha: {cotizacion['fecha']}", ln=True)
    pdf.cell(0, 6, f"Cliente: {cotizacion['cliente']}", ln=True)
    pdf.cell(0, 6, f"Correo: {cotizacion['correo']}", ln=True)
    pdf.cell(0, 6, f"Elaborado por: {cotizacion['operario']}", ln=True)
    pdf.ln(10)
    pdf.set_font("Arial", "B", 10)
    pdf.set_fill_color(220, 220, 220)
    pdf.cell(20, 7, "Cod", border=1, fill=True, align="C")
    pdf.cell(90, 7, "Producto / Servicio", border=1, fill=True)
    pdf.cell(20, 7, "Cant", border=1, fill=True, align="C")
    pdf.cell(30, 7, "P. Unit", border=1, fill=True, align="R")
    pdf.cell(30, 7, "Subtotal", border=1, fill=True, align="R", ln=True)
    pdf.set_font("Arial", "", 10)
    for item in lista_items:
        pdf.cell(20, 6, str(item['codigo']), border=1, align="C")
        pdf.cell(90, 6, str(item['producto'])[:40], border=1)
        pdf.cell(20, 6, str(item['cantidad']), border=1, align="C")
        pdf.cell(30, 6, f"${item['precio']:,.0f}", border=1, align="R")
        pdf.cell(30, 6, f"${item['subtotal']:,.0f}", border=1, align="R", ln=True)
    pdf.ln(5)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(130, 8, "TOTAL:", align="R")
    pdf.cell(30, 8, f"${total:,.0f}", border=1, align="R", ln=True)
    nombre_archivo = f"Cotizacion_{cotizacion['numero'].replace(' ', '_')}.pdf"
    pdf.output(nombre_archivo)
    return nombre_archivo

def enviar_correo(destinatario, asunto, cuerpo, archivo_pdf):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMITENTE
        msg['To'] = destinatario
        msg['Subject'] = asunto
        msg.attach(MIMEText(cuerpo, 'plain'))
        with open(archivo_pdf, "rb") as archivo:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(archivo.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f"attachment; filename={archivo_pdf}")
            msg.attach(part)
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(EMAIL_REMITENTE, CONTRASENA_APP)
        servidor.send_message(msg)
        servidor.quit()
        return True
    except Exception as e:
        st.error(f"Error al enviar correo: {e}")
        return False

def guardar_en_sheets(cotizacion, lista_items, total):
    """Guarda la cotización en Google Sheets - versión simple"""
    try:
        import os
        import json
        import gspread
        from google.oauth2 import service_account
        
        # Leer credenciales del archivo JSON (SIN MODIFICAR NADA)
        ruta_credenciales = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'credenciales.json')
        
        with open(ruta_credenciales, 'r') as f:
            creds_data = json.load(f)
        
        # Autenticar DIRECTAMENTE con el diccionario
        creds = service_account.Credentials.from_service_account_info(
            creds_data,
            scopes=['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        )
        
        client = gspread.authorize(creds)
        spreadsheet = client.open('Sistema de Cotizaciones')
        
        # Hoja Cotizaciones
        try:
            hoja_cotizaciones = spreadsheet.worksheet('Cotizaciones')
        except:
            hoja_cotizaciones = spreadsheet.add_worksheet(title='Cotizaciones', rows=100, cols=10)
            hoja_cotizaciones.append_row(['Número', 'Fecha', 'Cliente', 'Correo', 'Operario', 'Total', 'Estado'])
        
        hoja_cotizaciones.append_row([
            cotizacion['numero'],
            cotizacion['fecha'],
            cotizacion['cliente'],
            cotizacion['correo'],
            cotizacion['operario'],
            f"${total:,.0f}",
            'Guardada'
        ])
        
        # Hoja Detalle_Cotizaciones
        try:
            hoja_detalle = spreadsheet.worksheet('Detalle_Cotizaciones')
        except:
            hoja_detalle = spreadsheet.add_worksheet(title='Detalle_Cotizaciones', rows=500, cols=10)
            hoja_detalle.append_row(['Número_Cotizacion', 'Código', 'Producto', 'Cantidad', 'Precio', 'Subtotal'])
        
        for item in lista_items:
            hoja_detalle.append_row([
                cotizacion['numero'],
                item['codigo'],
                item['producto'],
                item['cantidad'],
                f"${item['precio']:,.0f}",
                f"${item['subtotal']:,.0f}"
            ])
        
        return True
    except Exception as e:
        st.error(f"❌ Error al guardar en Google Sheets: {e}")
        return False
        
# ============================================================
# INICIALIZAR ESTADO
# ============================================================
if 'lista_items' not in st.session_state:
    st.session_state.lista_items = []
if 'contador_recarga' not in st.session_state:
    st.session_state.contador_recarga = 0

# ============================================================
# INTERFAZ PRINCIPAL
# ============================================================
st.title(" Sistema de Cotizaciones")
st.markdown("---")

df_servicios = cargar_servicios()
if df_servicios.empty:
    st.error(" No se pudo cargar la lista de servicios.")
    st.stop()

col_codigo = None
col_producto = None
col_precio = None
for col in df_servicios.columns:
    if 'COD' in col or col == 'ID':
        col_codigo = col
    if 'PROD' in col or 'SERV' in col or 'DESC' in col:
        col_producto = col
    if 'COST' in col or 'PREC' in col or 'VALOR' in col:
        col_precio = col

if not col_codigo or not col_producto or not col_precio:
    st.error(f"❌ No se encontraron las columnas necesarias.")
    st.info(f"**Columnas detectadas:** {list(df_servicios.columns)}")
    st.stop()

# ============================================================
# SECCIÓN 1: DATOS DEL CLIENTE
# ============================================================
st.subheader("1. Datos de la Cotización")
col1, col2 = st.columns(2)

with col1:
    if "numero_cot" not in st.session_state:
        st.session_state.numero_cot = f"COT-{datetime.now().strftime('%Y%m%d-%H%M')}"
    numero_cot = st.text_input("Número de cotización", value=st.session_state.numero_cot, key=f"num_{st.session_state.contador_recarga}")
    fecha_cot = st.date_input("Fecha", value=datetime.now(), key=f"fecha_{st.session_state.contador_recarga}")
    cliente = st.text_input("Nombre del cliente / Empresa", key=f"cliente_{st.session_state.contador_recarga}")

with col2:
    correo = st.text_input("Correo electrónico del cliente", key=f"correo_{st.session_state.contador_recarga}")
    operario = st.selectbox("Elaborado por (Operario)", OPERARIOS, key=f"operario_{st.session_state.contador_recarga}")

# ============================================================
# SECCIÓN 2: AGREGAR ÍTEMS
# ============================================================
st.markdown("---")
st.subheader("2. Agregar Productos / Servicios")

opciones_servicios = ["-- Selecciona un producto/servicio --"]
for _, row in df_servicios.iterrows():
    codigo = str(row[col_codigo])
    producto = str(row[col_producto])
    precio_num = limpiar_precio(row[col_precio])
    opciones_servicios.append(f"{codigo} - {producto} - ${precio_num:,.0f}")

with st.form("form_agregar_item", clear_on_submit=True):
    st.write("**Selecciona un producto/servicio:**")
    seleccion = st.selectbox("Código - Producto - Precio", opciones_servicios, index=0)
    cantidad = st.number_input("Cantidad", min_value=0, value=0, step=1)
    
    col_boton1, col_boton2 = st.columns(2)
    with col_boton1:
        btn_agregar = st.form_submit_button("➕ Agregar ítem", use_container_width=True, type="primary")
    with col_boton2:
        btn_finalizar = st.form_submit_button("✅ Finalizar cotización", use_container_width=True, type="secondary")

    if btn_agregar:
        if seleccion == "-- Selecciona un producto/servicio --":
            st.warning("⚠️ Debes seleccionar un producto/servicio primero.")
        elif cantidad == 0:
            st.warning("⚠️ Debes ingresar una cantidad mayor a 0.")
        else:
            partes = seleccion.split(" - ")
            codigo_sel, producto_sel = partes[0], partes[1]
            precio_str = partes[2].replace('$', '').replace(',', '')
            precio_exacto = float(precio_str)
            
            st.session_state.lista_items.append({
                'codigo': codigo_sel, 
                'producto': producto_sel, 
                'cantidad': cantidad,
                'precio': precio_exacto, 
                'subtotal': precio_exacto * cantidad
            })
            st.success(f"✅ Agregado: {producto_sel} x{cantidad} = ${precio_exacto * cantidad:,.0f}")
            st.rerun()

    if btn_finalizar:
        if seleccion != "-- Selecciona un producto/servicio --" and cantidad > 0:
            partes = seleccion.split(" - ")
            codigo_sel, producto_sel = partes[0], partes[1]
            precio_str = partes[2].replace('$', '').replace(',', '')
            precio_exacto = float(precio_str)
            
            st.session_state.lista_items.append({
                'codigo': codigo_sel, 
                'producto': producto_sel, 
                'cantidad': cantidad,
                'precio': precio_exacto, 
                'subtotal': precio_exacto * cantidad
            })
            st.success(f"✅ Agregado: {producto_sel} x{cantidad} = ${precio_exacto * cantidad:,.0f}")
            st.rerun()

# ============================================================
# SECCIÓN 3: RESUMEN TIPO FACTURA
# ============================================================
st.markdown("---")

if len(st.session_state.lista_items) == 0:
    st.info("Aún no has agregado ítems a la cotización.")
else:
    st.markdown("""
    <div style='background-color: #1a5276; color: white; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;'>
        <h2 style='margin: 0; font-size: 24px;'>PROGAS SAS ESP</h2>
        <p style='margin: 5px 0; font-size: 14px;'>Empresa de Servicios Públicos</p>
        <p style='margin: 5px 0; font-size: 14px;'>NIT: 9018087574</p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown(f"""
    <div style='background-color: #f8f9fa; padding: 15px; border-radius: 8px; margin-bottom: 20px; border-left: 4px solid #1a5276;'>
        <p style='margin: 5px 0;'><strong>Número:</strong> {numero_cot}</p>
        <p style='margin: 5px 0;'><strong>Fecha:</strong> {fecha_cot.strftime('%d/%m/%Y')}</p>
        <p style='margin: 5px 0;'><strong>Cliente:</strong> {cliente}</p>
        <p style='margin: 5px 0;'><strong>Correo:</strong> {correo}</p>
        <p style='margin: 5px 0;'><strong>Elaborado por:</strong> {operario}</p>
    </div>
    """, unsafe_allow_html=True)
    
    df_tabla = pd.DataFrame(st.session_state.lista_items)
    df_tabla['precio_fmt'] = df_tabla['precio'].apply(lambda x: f"${x:,.0f}")
    df_tabla['subtotal_fmt'] = df_tabla['subtotal'].apply(lambda x: f"${x:,.0f}")
    df_tabla = df_tabla[['codigo', 'producto', 'cantidad', 'precio_fmt', 'subtotal_fmt']]
    df_tabla.columns = ['Código', 'Producto / Servicio', 'Cant.', 'Precio Unit.', 'Subtotal']
    
    st.dataframe(df_tabla, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.markdown("**Eliminar ítems:**")
    for idx, item in enumerate(st.session_state.lista_items):
        if st.button(f"🗑️ Eliminar: {item['codigo']} - {item['producto']}", key=f"del_{idx}", use_container_width=True):
            st.session_state.lista_items.pop(idx)
            st.rerun()
    
    total = sum(item.get('subtotal', 0) for item in st.session_state.lista_items)
    
    st.markdown(
        f"""
        <div style='background-color: #27ae60; color: white; padding: 20px; border-radius: 10px; margin: 20px 0; text-align: right;'>
            <h2 style='margin: 0; font-size: 28px;'>💰 TOTAL: ${total:,.0f}</h2>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    col_accion1, col_accion2, col_accion3 = st.columns(3)
    with col_accion1:
        if st.button("️ Limpiar todo", use_container_width=True):
            st.session_state.lista_items = []
            if "numero_cot" in st.session_state:
                del st.session_state["numero_cot"]
            st.session_state.contador_recarga += 1
            st.rerun()
    
    with col_accion2:
        if st.button("📄 Generar PDF", use_container_width=True):
            st.session_state.accion = "generar_pdf"
            st.rerun()
    
    with col_accion3:
        if st.button("📧 Enviar por correo", use_container_width=True, disabled=correo == ""):
            st.session_state.accion = "enviar_correo"
            st.rerun()

# ============================================================
# SECCIÓN 4: GENERAR PDF Y ENVIAR
# ============================================================
if st.session_state.get('accion') in ['generar_pdf', 'enviar_correo']:
    
    if not cliente or not correo:
        st.error("❌ Debes llenar el nombre del cliente y el correo.")
        if st.button("🔙 Volver"):
            st.session_state.accion = None
            st.rerun()
    else:
        with st.spinner("📄 Generando PDF..."):
            cotizacion = {
                'numero': numero_cot,
                'fecha': fecha_cot.strftime('%d/%m/%Y'),
                'cliente': cliente,
                'correo': correo,
                'operario': operario
            }
            
            archivo_pdf = generar_pdf(cotizacion, st.session_state.lista_items, total)
            st.success(f"✅ PDF generado: `{archivo_pdf}`")
            
            with open(archivo_pdf, "rb") as f:
                st.download_button(label="⬇️ Descargar PDF", data=f, file_name=archivo_pdf, mime="application/pdf")
            
            if st.session_state.accion == 'enviar_correo':
                with st.spinner("📧 Enviando correo y guardando en Google Sheets..."):
                    asunto = f"Cotización {numero_cot} - {cliente}"
                    cuerpo = f"""Hola {cliente},

Adjunto encontrará la cotización solicitada de PROGAS SAS ESP.

Número: {numero_cot}
Fecha: {fecha_cot.strftime('%d/%m/%Y')}
Total: ${total:,.0f}

Cordialmente,
{operario}
PROGAS SAS
"""
                    resultado_correo = enviar_correo(correo, asunto, cuerpo, archivo_pdf)
                    resultado_sheets = guardar_en_sheets(cotizacion, st.session_state.lista_items, total)
                    
                    if resultado_correo and resultado_sheets:
                        st.success(f"✅ Correo enviado exitosamente a {correo}")
                        st.success(f"✅ Cotización guardada en Google Sheets")
                    elif resultado_correo:
                        st.success(f"✅ Correo enviado a {correo}")
                        st.warning(f"⚠️ No se pudo guardar en Google Sheets")
                    elif resultado_sheets:
                        st.error(f"❌ No se pudo enviar el correo")
                        st.success(f"✅ Cotización guardada en Google Sheets")
                    else:
                        st.error(f"❌ No se pudo enviar el correo ni guardar en Google Sheets")
            
            st.markdown("---")
            if st.button("🔙 Volver al resumen", use_container_width=True):
                st.session_state.accion = None
                st.rerun()
