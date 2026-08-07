import streamlit as st
import pandas as pd
from supabase import create_client, Client
import io
import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fpdf import FPDF
from PIL import Image

# --- 1. CONEXIÓN A TU BASE DE DATOS EN LA NUBE ---
URL_SUPABASE = "https://ahmjxfmfnbuhmirtfjzu.supabase.co"
CLAVE_SUPABASE = "sb_publishable_yRq3_3K7MK2d3Zxt6DmQBg_fBht0_Bv" 

@st.cache_resource
def init_connection():
    return create_client(URL_SUPABASE, CLAVE_SUPABASE)

try:
    supabase: Client = init_connection()
except Exception as e:
    st.error("Error conectando a la base de datos.")

# --- FUNCIONES PARA LEER DATOS ---
def obtener_clientes():
    try:
        return supabase.table("clientes").select("*").execute().data
    except:
        return []

def obtener_productos():
    try:
        return supabase.table("productos").select("*").execute().data
    except:
        return []

def obtener_historial():
    try:
        return supabase.table("cotizaciones").select("*").order("id", desc=True).execute().data
    except:
        return []

def obtener_siguiente_folio():
    try:
        res = supabase.table("cotizaciones").select("id").order("id", desc=True).limit(1).execute()
        if res.data:
            return res.data[0]["id"] + 1
        return 1
    except:
        return 1

# --- 2. EXCEL GENERAL ---
def generar_excel(datos_cotizacion, cliente, subtotal, iva, total):
    wb = Workbook()
    ws = wb.active
    ws.title = "Cotización"

    header_fill = PatternFill(start_color="2B3A42", end_color="2B3A42", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    title_font = Font(size=16, bold=True, color="1F4E79")
    bold_font = Font(bold=True)
    border_thin = Border(left=Side(style='thin', color='E0E0E0'), right=Side(style='thin', color='E0E0E0'), top=Side(style='thin', color='E0E0E0'), bottom=Side(style='thin', color='E0E0E0'))

    ws['A1'] = "COTIZACIÓN QUIMSAGI"
    ws['A1'].font = title_font
    ws['A2'] = f"Cliente: {cliente}"
    ws['A2'].font = bold_font

    if datos_cotizacion:
        headers = list(datos_cotizacion[0].keys())
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_num)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for row_num, row_data in enumerate(datos_cotizacion, 5):
            for col_num, key in enumerate(headers, 1):
                cell = ws.cell(row=row_num, column=col_num)
                cell.value = row_data[key]
                cell.border = border_thin
                if key in ["Precio Unit.", "Subtotal"]:
                    cell.number_format = '"$"#,##0.00'
                elif key == "Cant.":
                    cell.alignment = Alignment(horizontal="center")

        final_row = len(datos_cotizacion) + 5
        ws.cell(row=final_row+1, column=6).value = "Subtotal:"
        ws.cell(row=final_row+1, column=6).font = bold_font
        ws.cell(row=final_row+1, column=7).value = subtotal
        ws.cell(row=final_row+1, column=7).number_format = '"$"#,##0.00'
        ws.cell(row=final_row+2, column=6).value = "IVA (16%):"
        ws.cell(row=final_row+2, column=6).font = bold_font
        ws.cell(row=final_row+2, column=7).value = iva
        ws.cell(row=final_row+2, column=7).number_format = '"$"#,##0.00'
        ws.cell(row=final_row+3, column=6).value = "Total:"
        ws.cell(row=final_row+3, column=6).font = bold_font
        ws.cell(row=final_row+3, column=7).value = total
        ws.cell(row=final_row+3, column=7).number_format = '"$"#,##0.00'

    ws.column_dimensions['A'].width = 15
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 12
    ws.column_dimensions['G'].width = 15

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()

# --- FILTRO LIMPIADOR ---
def limpiar_texto(texto):
    if not isinstance(texto, str):
        return str(texto)
    return texto.replace('', '').replace('\ufffd', '')

def armar_direccion(c):
    partes = [
        f"{c.get('CALLE', '')} {c.get('NO_EXTERIOR', '')}".strip(),
        f"Col. {c.get('COLONIA', '')}" if c.get('COLONIA') else "",
        c.get('MUNICIPIO', ''),
        c.get('ESTADO', ''),
        f"CP {c.get('CP', '')}" if c.get('CP') else ""
    ]
    validas = [p for p in partes if p and p.lower() != "none" and p.lower() != "nan"]
    return ", ".join(validas)

# --- 3. CLASE ESPECIAL PDF (Cotización) ---
class PDFQuimsagi(FPDF):
    def footer(self):
        self.set_y(-25)
        self.set_font("Arial", 'B', 10)
        self.set_text_color(26, 82, 118) 
        self.cell(0, 5, "Favor de confirmar la cotizacion con su vendedor", ln=True, align='C')
        self.set_font("Arial", 'B', 9)
        self.cell(0, 5, "Ventas: ventas1quimsagi@gmail.com  |  Tel: 998 459 2513", ln=True, align='C')
        self.cell(0, 5, "Administracion: direccionquimsagi@gmail.com", ln=True, align='C')

def generar_pdf(datos_cotizacion, cliente_info, subtotal, iva, total, vendedor, folio):
    pdf = PDFQuimsagi()
    pdf.add_page()
    
    logo_path = None
    if os.path.exists("logo.png"): logo_path = "logo.png"
    elif os.path.exists("logo.jpg"): logo_path = "logo.jpg"
    elif os.path.exists("QUIMSAGI LOGO FINAL.jpeg"): logo_path = "QUIMSAGI LOGO FINAL.jpeg"

    if logo_path:
        try:
            img = Image.open(logo_path).convert("RGBA")
            alpha = img.split()[3]
            alpha = alpha.point(lambda p: p * 0.15) 
            img.putalpha(alpha)
            wm_io = io.BytesIO()
            img.save(wm_io, format='PNG')
            wm_io.seek(0)
            pdf.image(wm_io, x=35, y=90, w=140)
        except Exception as e:
            pass 
        pdf.image(logo_path, 10, 10, 42)
    else:
        pdf.set_font("Arial", 'B', 22)
        pdf.set_text_color(26, 82, 118)
        pdf.cell(50, 10, "Quimsagi", ln=False)

    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f"COTIZACION #{folio}", ln=True, align='R')
    
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, "PRODUCTOS DE LIMPIEZA A TU MEDIDA", ln=True, align='R')
    pdf.ln(10)
    
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("Arial", 'B', 11)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 8, "  Datos del Cliente", border=0, fill=True, ln=True)
    pdf.ln(2)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "Cliente:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.cell(110, 5, limpiar_texto(cliente_info.get("RAZON_SOCIAL", "")), border=0)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "Atendio:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.cell(40, 5, limpiar_texto(vendedor), border=0, ln=True)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "RFC:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.cell(110, 5, limpiar_texto(cliente_info.get("RFC", "")), border=0)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "Fecha:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.cell(40, 5, datetime.now().strftime("%d/%m/%Y"), border=0, ln=True)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "Direccion:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.multi_cell(0, 5, limpiar_texto(armar_direccion(cliente_info)), border=0)
    pdf.ln(5)
    
    pdf.set_fill_color(26, 82, 118)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 9)
    
    col_w = {"Clave": 20, "Producto": 85, "Cant.": 15, "Unidad": 20, "Precio Unit.": 25, "Subtotal": 25}
    
    pdf.cell(col_w["Clave"], 8, "Clave", border=1, align='C', fill=True)
    pdf.cell(col_w["Producto"], 8, "Descripcion", border=1, align='C', fill=True)
    pdf.cell(col_w["Cant."], 8, "Cant.", border=1, align='C', fill=True)
    pdf.cell(col_w["Unidad"], 8, "Unidad", border=1, align='C', fill=True)
    pdf.cell(col_w["Precio Unit."], 8, "Precio", border=1, align='C', fill=True)
    pdf.cell(col_w["Subtotal"], 8, "Total", border=1, align='C', fill=True)
    pdf.ln()
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", '', 9)
    
    for row in datos_cotizacion:
        pdf.cell(col_w["Clave"], 8, limpiar_texto(row["Clave"]), border=1)
        pdf.cell(col_w["Producto"], 8, limpiar_texto(row["Producto"])[:45], border=1)
        pdf.cell(col_w["Cant."], 8, str(row["Cant."]), border=1, align='C')
        pdf.cell(col_w["Unidad"], 8, limpiar_texto(row["Unidad"]), border=1, align='C')
        pdf.cell(col_w["Precio Unit."], 8, f"${row['Precio Unit.']:,.2f}", border=1, align='R')
        pdf.cell(col_w["Subtotal"], 8, f"${row['Subtotal']:,.2f}", border=1, align='R')
        pdf.ln()
        
    pdf.ln(5)
    margen_izq = col_w["Clave"] + col_w["Producto"] + col_w["Cant."] + col_w["Unidad"]
    
    pdf.set_font("Arial", '', 9)
    pdf.cell(margen_izq, 7, "", border=0)
    pdf.cell(col_w["Precio Unit."], 7, "Subtotal:", border=0, align='R')
    pdf.cell(col_w["Subtotal"], 7, f"${subtotal:,.2f}", border=0, align='R')
    pdf.ln()
    
    pdf.cell(margen_izq, 7, "", border=0)
    pdf.cell(col_w["Precio Unit."], 7, "IVA (16%):", border=0, align='R')
    pdf.cell(col_w["Subtotal"], 7, f"${iva:,.2f}", border=0, align='R')
    pdf.ln()
    
    pdf.set_font("Arial", 'B', 10)
    pdf.set_fill_color(240, 240, 240)
    pdf.cell(margen_izq, 8, "", border=0)
    pdf.cell(col_w["Precio Unit."], 8, "Total:", border=1, align='R', fill=True)
    pdf.cell(col_w["Subtotal"], 8, f"${total:,.2f}", border=1, align='R', fill=True)
    pdf.ln(15)
    
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(26, 82, 118) 
    pdf.cell(0, 6, "DATOS PARA TRANSFERENCIA BANCARIA", ln=True)
    
    pdf.set_font("Arial", 'B', 9) 
    pdf.cell(15, 5, "Banco:", border=0)
    pdf.cell(80, 5, "Mercado Pago", border=0, ln=True)
    
    pdf.cell(15, 5, "Titular:", border=0)
    pdf.cell(80, 5, "Quimsagi", border=0, ln=True)
    
    pdf.cell(15, 5, "CLABE:", border=0)
    pdf.cell(80, 5, "722969013491074912", border=0, ln=True)

    pdf_bytes = pdf.output(dest='S')
    return bytes(pdf_bytes) if not isinstance(pdf_bytes, str) else pdf_bytes.encode('latin-1')

# --- 4. GENERADOR PDF ESTADO DE CUENTA CLIENTE ---
def generar_estado_cuenta_pdf(nombre_cliente, cliente_info, lista_facturas_cliente):
    pdf = FPDF()
    pdf.add_page()
    
    logo_path = None
    if os.path.exists("logo.png"): logo_path = "logo.png"
    elif os.path.exists("logo.jpg"): logo_path = "logo.jpg"
    elif os.path.exists("QUIMSAGI LOGO FINAL.jpeg"): logo_path = "QUIMSAGI LOGO FINAL.jpeg"

    if logo_path:
        try: pdf.image(logo_path, 10, 10, 35)
        except: pass

    pdf.set_font("Arial", 'B', 16)
    pdf.set_text_color(26, 82, 118)
    pdf.cell(0, 10, "ESTADO DE CUENTA", ln=True, align='R')
    pdf.set_font("Arial", 'I', 9)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 4, "QUIMSAGI - PRODUCTOS DE LIMPIEZA", ln=True, align='R')
    pdf.ln(10)
    
    # Datos del cliente
    pdf.set_fill_color(240, 240, 240)
    pdf.set_font("Arial", 'B', 10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 7, "  Informacion del Cliente", border=0, fill=True, ln=True)
    pdf.ln(2)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "Cliente:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.cell(100, 5, limpiar_texto(nombre_cliente), border=0, ln=True)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "RFC:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.cell(100, 5, limpiar_texto(cliente_info.get("RFC", "N/D")), border=0, ln=True)
    
    pdf.set_font("Arial", 'B', 9)
    pdf.cell(20, 5, "Direccion:", border=0)
    pdf.set_font("Arial", '', 9)
    pdf.multi_cell(0, 5, limpiar_texto(armar_direccion(cliente_info)), border=0)
    pdf.ln(10)
    
    # Tabla de facturas/cotizaciones del cliente
    pdf.set_fill_color(26, 82, 118)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 9)
    
    pdf.cell(15, 8, "Folio", border=1, align='C', fill=True)
    pdf.cell(25, 8, "Fecha", border=1, align='C', fill=True)
    pdf.cell(35, 8, "Folio Fiscal", border=1, align='C', fill=True)
    pdf.cell(30, 8, "Operativo", border=1, align='C', fill=True)
    pdf.cell(30, 8, "Financiero", border=1, align='C', fill=True)
    pdf.cell(20, 8, "Atraso", border=1, align='C', fill=True)
    pdf.cell(25, 8, "Total", border=1, align='C', fill=True)
    pdf.ln()
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", '', 8)
    
    suma_total_deuda = 0
    for fac in lista_facturas_cliente:
        # Calcular dias transcurridos
        f_creacion = pd.to_datetime(fac.get("fecha"))
        dias = (datetime.now() - f_creacion.tz_localize(None) if f_creacion.tzinfo else datetime.now() - f_creacion).days
        if fac.get("estatus_financiero") == "Pagada":
            dias_str = "Pagado"
        else:
            dias_str = f"{dias} dias"
            suma_total_deuda += float(fac.get("total", 0))
            
        pdf.cell(15, 7, str(fac.get("id")), border=1, align='C')
        pdf.cell(25, 7, f_creacion.strftime('%d/%m/%Y'), border=1, align='C')
        pdf.cell(35, 7, limpiar_texto(fac.get("folio_fiscal") or "S/F"), border=1, align='C')
        pdf.cell(30, 7, limpiar_texto(fac.get("estatus_operativo", "Pendiente")), border=1, align='C')
        pdf.cell(30, 7, limpiar_texto(fac.get("estatus_financiero", "Pendiente")), border=1, align='C')
        pdf.cell(20, 7, dias_str, border=1, align='C')
        pdf.cell(25, 7, f"${float(fac.get('total', 0)):,.2f}", border=1, align='R')
        pdf.ln()
        
    pdf.ln(5)
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(135, 8, "SALDO PENDIENTE TOTAL:", border=0, align='R')
    pdf.cell(25, 8, f"${suma_total_deuda:,.2f}", border=1, align='R', fill=True)

    pdf_bytes = pdf.output(dest='S')
    return bytes(pdf_bytes) if not isinstance(pdf_bytes, str) else pdf_bytes.encode('latin-1')

# --- PÁGINA PRINCIPAL ---
st.set_page_config(page_title="Cotizador Quimsagi", page_icon="📋", layout="wide")

if 'cotizacion_actual' not in st.session_state: st.session_state.cotizacion_actual = []
if 'autenticado' not in st.session_state: st.session_state.autenticado = False

if not st.session_state.autenticado:
    st.title("🔒 Acceso al Sistema")
    usuario = st.text_input("Usuario")
    contrasena = st.text_input("Contraseña", type="password")
    
    if st.button("Entrar"):
        if usuario == "admin" and contrasena == "12345":
            st.session_state.autenticado = True
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos")
else:
    st.sidebar.title("Menú Principal")
    menu = st.sidebar.radio("Ir a:", ["📝 Cotizador", "📂 Historial y Cobranza", "📊 Métricas", "⚙️ Panel de Administrador"])
    
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

    # 1. COTIZADOR
    if menu == "📝 Cotizador":
        st.title("📝 Generar Nueva Cotización")
        clientes = obtener_clientes()
        productos = obtener_productos()
        
        if len(clientes) == 0 or len(productos) == 0:
            st.warning("⚠️ Faltan clientes o productos en la base de datos.")
        else:
            folio_actual = obtener_siguiente_folio()
            
            st.subheader("1. Datos Generales")
            col_vendedor, col_folio = st.columns([3, 1])
            with col_vendedor:
                lista_vendedores = ["Gerente Gilber Carbajal", "Vendedora Laisha", "Vendedor Omar", "Vendedora Grisy"]
                vendedor_seleccionado = st.selectbox("Atendido por:", lista_vendedores)
            with col_folio:
                st.info(f"Folio actual: **#{folio_actual}**")

            lista_razones_sociales = [c.get("RAZON_SOCIAL", "Sin Razón Social") for c in clientes if c.get("RAZON_SOCIAL")]
            cliente_seleccionado = st.selectbox("Selecciona un cliente:", lista_razones_sociales)
            
            st.subheader("2. Agregar Productos y Precios")
            mapa_productos = {f"{p.get('CLAVE') or 'S/C'} - {p.get('PRODUCTO', 'Sin Nombre')}": p for p in productos if p.get("PRODUCTO")}
            lista_opciones_productos = list(mapa_productos.keys())
            
            col1, col2, col3 = st.columns([2, 1, 1])
            with col1: 
                producto_seleccionado_display = st.selectbox("Selecciona un producto:", lista_opciones_productos)
            
            prod_info = mapa_productos.get(producto_seleccionado_display)
            producto_seleccionado = prod_info.get("PRODUCTO") if prod_info else ""
            precio_catalogo = float(prod_info.get("PRECIO", 0) or 0) if prod_info else 0.0
            
            with col2: 
                cantidad = st.number_input("Cantidad", min_value=1, value=1)
            with col3:
                precio_personalizado = st.number_input("Precio Unitario ($)", min_value=0.0, value=precio_catalogo, format="%.2f")
                
            if st.button("Agregar a la cotización"):
                if prod_info:
                    descuento_pct = float(prod_info.get("DESCUENTO", 0) or 0)
                    
                    if precio_personalizado != precio_catalogo:
                        precio_final_unitario = precio_personalizado
                    else:
                        descuento_dinero = precio_catalogo * (descuento_pct / 100)
                        precio_final_unitario = precio_catalogo - descuento_dinero
                        
                    subtotal_linea = precio_final_unitario * cantidad
                    
                    st.session_state.cotizacion_actual.append({
                        "Clave": prod_info.get("Clave", prod_info.get("CLAVE", "")),
                        "Producto": producto_seleccionado,
                        "Cant.": cantidad,
                        "Unidad": prod_info.get("UNIDAD", "Pza"),
                        "Precio Unit.": precio_final_unitario, 
                        "Desc. %": f"{descuento_pct}%",
                        "Subtotal": subtotal_linea
                    })
                    st.success(f"Se agregó {cantidad}x {producto_seleccionado} a ${precio_final_unitario:,.2f}")

            if len(st.session_state.cotizacion_actual) > 0:
                st.markdown("---")
                st.subheader("🛒 Resumen de la Cotización")
                df = pd.DataFrame(st.session_state.cotizacion_actual)
                df_vista = df.copy()
                df_vista["Precio Unit."] = df_vista["Precio Unit."].apply(lambda x: f"${x:,.2f}")
                df_vista["Subtotal"] = df_vista["Subtotal"].apply(lambda x: f"${x:,.2f}")
                st.dataframe(df_vista, use_container_width=True, hide_index=True)
                
                suma_subtotal = df["Subtotal"].sum()
                iva = suma_subtotal * 0.16
                total_final = suma_subtotal + iva
                
                col_blank, col_totales = st.columns([3, 1])
                with col_totales:
                    st.write(f"**Subtotal:** ${suma_subtotal:,.2f}")
                    st.write(f"**IVA (16%):** ${iva:,.2f}")
                    st.write(f"### **Total:** ${total_final:,.2f}")
                
                st.markdown("---")
                col_btn1, col_btn2, col_btn3, col_btn4 = st.columns(4)
                
                cliente_completo = next((c for c in clientes if c.get("RAZON_SOCIAL") == cliente_seleccionado), {})
                
                with col_btn1:
                    if st.button("💾 Guardar en Historial", use_container_width=True):
                        cotizacion_data = {
                            "cliente": cliente_seleccionado, 
                            "vendedor": vendedor_seleccionado,
                            "total": total_final,
                            "detalles": st.session_state.cotizacion_actual,
                            "estatus_operativo": "Pendiente de autorización",
                            "estatus_financiero": "Pendiente de cobro"
                        }
                        try:
                            supabase.table("cotizaciones").insert(cotizacion_data).execute()
                            st.success("¡Guardada exitosamente!")
                            st.session_state.cotizacion_actual = []
                            st.rerun()
                        except Exception as e: 
                            st.error(f"Error al guardar: {e}")
                with col_btn2:
                    st.download_button("📥 Descargar Excel", data=generar_excel(st.session_state.cotizacion_actual, cliente_seleccionado, suma_subtotal, iva, total_final), file_name=f"Cotizacion_Folio{folio_actual}_{cliente_seleccionado}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with col_btn3:
                    st.download_button("📄 Descargar PDF", data=generar_pdf(st.session_state.cotizacion_actual, cliente_completo, suma_subtotal, iva, total_final, vendedor_seleccionado, folio_actual), file_name=f"Cotizacion_Folio{folio_actual}_{cliente_seleccionado}.pdf", mime="application/pdf", use_container_width=True)
                with col_btn4:
                    if st.button("🗑️ Limpiar Todo", use_container_width=True):
                        st.session_state.cotizacion_actual = []
                        st.rerun()

    # 2. HISTORIAL Y COBRANZA
    elif menu == "📂 Historial y Cobranza":
        st.title("📂 Historial, Operación y Cobranza (CxC)")
        historial = obtener_historial()
        if len(historial) == 0:
            st.info("Aún no hay cotizaciones guardadas.")
        else:
            df_hist = pd.DataFrame(historial)
            
            # Limpieza y defaults por si hay columnas vacías
            for col, val in [("vendedor", "S/D"), ("estatus_operativo", "Pendiente de autorización"), ("estatus_financiero", "Pendiente de cobro"), ("folio_fiscal", "-"), ("forma_pago", "-")]:
                if col not in df_hist.columns: df_hist[col] = val
                else: df_hist[col] = df_hist[col].fillna(val)

            # Cálculo de días de atraso para la tabla
            dias_atraso_lista = []
            for idx, row in df_hist.iterrows():
                f_creacion = pd.to_datetime(row.get("fecha"))
                dias = (datetime.now() - f_creacion.tz_localize(None) if f_creacion.tzinfo else datetime.now() - f_creacion).days
                if row.get("estatus_financiero") == "Pagada":
                    dias_atraso_lista.text = "Pagado"
                    dias_atraso_lista.append("Pagado")
                else:
                    dias_atraso_lista.append(f"{dias} días")
            df_hist["Atraso"] = dias_atraso_lista

            df_vista = df_hist[["id", "cliente", "vendedor", "estatus_operativo", "estatus_financiero", "Atraso", "total", "fecha"]].copy()
            df_vista["total"] = df_vista["total"].apply(lambda x: f"${float(x):,.2f}")
            df_vista["fecha"] = pd.to_datetime(df_vista["fecha"]).dt.strftime('%d/%m/%Y %H:%M')
            
            st.dataframe(df_vista, column_config={"id": "Folio", "cliente": "Cliente", "vendedor": "Atendió", "estatus_operativo": "Operativo", "estatus_financiero": "Financiero", "Atraso": "Antigüedad", "total": "Total", "fecha": "Fecha"}, use_container_width=True, hide_index=True)

            # PANEL DE ACTUALIZACIÓN DE ESTATUS Y COBRANZA
            st.markdown("---")
            st.subheader("⚡ Administrar Estatus y Cobranza de un Folio")
            
            opciones_folios = [f"Folio #{r['id']} - {r['cliente']}" for r in historial]
            folio_sel_str = st.selectbox("Selecciona el folio a modificar:", opciones_folios)
            id_sel = int(folio_sel_str.split("#")[1].split(" -")[0])
            reg_sel = next((r for r in historial if r["id"] == id_sel), {})

            col_c1, col_c2 = st.columns(2)
            with col_c1:
                st.markdown("##### 📌 Carril Operativo")
                nuevo_op = st.selectbox("Estatus Operativo:", ["Pendiente de autorización", "Autorizado", "Facturado"], index=["Pendiente de autorización", "Autorizado", "Facturado"].index(reg_sel.get("estatus_operativo", "Pendiente de autorización")))
                
                nuevo_folio_fiscal = reg_sel.get("folio_fiscal", "")
                if nuevo_op == "Facturado":
                    nuevo_folio_fiscal = st.text_input("Folio Fiscal de la Factura:", value=reg_sel.get("folio_fiscal", ""))

with col_c2:
                st.markdown("##### 💰 Carril Financiero (CxC)")
                opciones_fin = ["Pendiente de cobro", "Pagada"]
                fin_actual = reg_sel.get("estatus_financiero", "Pendiente de cobro")
                idx_fin = opciones_fin.index(fin_actual) if fin_actual in opciones_fin else 0
                nuevo_fin = st.selectbox("Estatus Financiero:", opciones_fin, index=idx_fin)
                
                nueva_forma_pago = "Transferencia"
                nueva_fecha_pago = str(date.today())
                
                if nuevo_fin == "Pagada":
                    opciones_pago = ["Transferencia", "Efectivo", "Tarjeta", "Cheque"]
                    fp_actual = reg_sel.get("forma_pago")
                    idx_fp = opciones_pago.index(fp_actual) if fp_actual in opciones_pago else 0
                    nueva_forma_pago = st.selectbox("Forma de Pago:", opciones_pago, index=idx_fp)
                    
                    fe_actual = reg_sel.get("fecha_pago")
                    try:
                        dt_val = datetime.strptime(str(fe_actual), "%Y-%m-%d").date() if fe_actual else date.today()
                    except:
                        dt_val = date.today()
                    nueva_fecha_pago = str(st.date_input("Fecha en que se pagó:", value=dt_val))
            if st.button("💾 Guardar Cambios de Estatus", use_container_width=True):
                datos_actualizados = {
                    "estatus_operativo": nuevo_op,
                    "folio_fiscal": nuevo_folio_fiscal if nuevo_op == "Facturado" else None,
                    "estatus_financiero": nuevo_fin,
                    "forma_pago": nueva_forma_pago if nuevo_fin == "Pagada" else None,
                    "fecha_pago": nueva_fecha_pago if nuevo_fin == "Pagada" else None
                }
                try:
                    supabase.table("cotizaciones").update(datos_actualizados).eq("id", id_sel).execute()
                    st.success("¡Estatus y cobranza actualizados correctamente!")
                    st.rerun()
                except Exception as e:
                    st.error(f"Error al actualizar: {e}")

            # RECUPERAR PDF Y ESTADO DE CUENTA
            st.markdown("---")
            col_desc1, col_desc2 = st.columns(2)
            
            with col_desc1:
                st.subheader("📄 Recuperar PDF de Cotización")
                folios_detalles = [f"Folio #{r['id']} - {r['cliente']}" for r in historial if r.get("detalles")]
                if folios_detalles:
                    f_dl = st.selectbox("Elige la cotización:", folios_detalles, key="dl_cot")
                    id_dl = int(f_dl.split("#")[1].split(" -")[0])
                    reg_dl = next((r for r in historial if r["id"] == id_dl), None)
                    if reg_dl and reg_dl.get("detalles"):
                        cli_comp = next((c for c in obtener_clientes() if c.get("RAZON_SOCIAL") == reg_dl["cliente"]), {"RAZON_SOCIAL": reg_dl["cliente"]})
                        dt_cot = reg_dl["detalles"]
                        s_sub = sum([float(i["Subtotal"]) for i in dt_cot])
                        iv = s_sub * 0.16
                        pdf_rec = generar_pdf(dt_cot, cli_comp, s_sub, iv, float(reg_dl["total"]), reg_dl.get("vendedor", "S/D"), id_dl)
                        st.download_button("📥 Descargar PDF de Cotización", data=pdf_rec, file_name=f"Cotizacion_Folio{id_dl}_{reg_dl['cliente']}.pdf", mime="application/pdf", use_container_width=True)

            with col_desc2:
                st.subheader("📑 Estado de Cuenta por Cliente")
                lista_nombres_clientes = list(set([r["cliente"] for r in historial]))
                if lista_nombres_clientes:
                    cli_edo_cta = st.selectbox("Selecciona cliente para estado de cuenta:", lista_nombres_clientes)
                    facturas_cliente = [r for r in historial if r["cliente"] == cli_edo_cta]
                    cli_info_completo = next((c for c in obtener_clientes() if c.get("RAZON_SOCIAL") == cli_edo_cta), {"RAZON_SOCIAL": cli_edo_cta})
                    
                    pdf_edo_cta = generar_estado_cuenta_pdf(cli_edo_cta, cli_info_completo, facturas_cliente)
                    st.download_button("📥 Descargar Estado de Cuenta (PDF)", data=pdf_edo_cta, file_name=f"EstadoDeCuenta_{cli_edo_cta}.pdf", mime="application/pdf", use_container_width=True)

    # 3. MÉTRICAS E INTELIGENCIA DE NEGOCIOS
    elif menu == "📊 Métricas":
        st.title("📊 Panel de Inteligencia y Cartera Vencida")
        historial = obtener_historial()
        
        if len(historial) > 0:
            df = pd.DataFrame(historial)
            df['total'] = df['total'].astype(float)
            
            for col, val in [("estatus_operativo", "Pendiente de autorización"), ("estatus_financiero", "Pendiente de cobro"), ("vendedor", "S/D")]:
                if col not in df.columns: df[col] = val
                else: df[col] = df[col].fillna(val)

            # KPIs Financieros y Operativos
            col1, col2, col3, col4 = st.columns(4)
            t_historico = df['total'].sum()
            t_cobrado = df[df['estatus_financiero'] == 'Pagada']['total'].sum()
            t_por_cobrar = df[df['estatus_financiero'] == 'Pendiente de cobro']['total'].sum()
            t_autorizado = df[df['estatus_operativo'] == 'Autorizado']['total'].sum()
            
            col1.metric("Total Cotizado", f"${t_historico:,.2f}")
            col2.metric("Cobrado en Banco", f"${t_cobrado:,.2f}")
            col3.metric("Por Cobrar (CxC)", f"${t_por_cobrar:,.2f}", delta_color="inverse")
            col4.metric("Operativo Autorizado", f"${t_autorizado:,.2f}")
            
            st.markdown("---")
            
            # Gráficas
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                st.subheader("Ventas por Vendedor")
                st.bar_chart(df.groupby("vendedor")["total"].sum())
            with col_g2:
                st.subheader("Estado Financiero (Cobranza)")
                st.bar_chart(df.groupby("estatus_financiero")["total"].sum())
                
            st.markdown("---")
            st.subheader("Top Clientes con mayor deuda (Pendiente de cobro)")
            df_deuda = df[df['estatus_financiero'] == 'Pendiente de cobro']
            if len(df_deuda) > 0:
                top_deuda = df_deuda.groupby("cliente")["total"].sum().sort_values(ascending=False).head(5).reset_index()
                top_deuda["total"] = top_deuda["total"].apply(lambda x: f"${x:,.2f}")
                st.dataframe(top_deuda, column_config={"cliente": "Cliente", "total": "Deuda Pendiente"}, use_container_width=True, hide_index=True)
            else:
                st.success("¡Excelente! No hay saldos pendientes de cobro en este momento.")
        else:
            st.info("Aún no hay suficientes datos para mostrar métricas.")

    # 4. ADMINISTRACIÓN
    elif menu == "⚙️ Panel de Administrador":
        st.title("⚙️ Panel de Administración")
        tab1, tab2 = st.tabs(["📦 Nuevo Producto", "🏢 Nuevo Cliente"])
        with tab1:
            col1, col2 = st.columns(2)
            with col1:
                clave = st.text_input("CLAVE")
                producto = st.text_input("PRODUCTO")
                unidad = st.selectbox("UNIDAD", ["Pza", "Kg", "Litro", "Galón", "Servicio"])
            with col2:
                precio = st.number_input("PRECIO", min_value=0.0, format="%.2f")
                descuento = st.number_input("DESCUENTO (%)", min_value=0.0, format="%.2f")
            if st.button("Guardar Producto"):
                nuevo_producto = {"CLAVE": clave, "PRODUCTO": producto, "PRECIO": precio, "DESCUENTO": descuento, "UNIDAD": unidad}
                try:
                    supabase.table("productos").insert(nuevo_producto).execute()
                    st.success("¡Guardado!")
                except: st.error("Error al guardar.")
        with tab2:
            rfc = st.text_input("RFC")
            razon_social = st.text_input("RAZÓN SOCIAL")
            forma_pago = st.selectbox("FORMA DE PAGO", ["Transferencia", "Efectivo", "Tarjeta", "Crédito"])
            if st.button("Guardar Cliente"):
                nuevo_cliente = {"RFC": rfc, "RAZON_SOCIAL": razon_social, "FORMA_PAGO": forma_pago}
                try:
                    supabase.table("clientes").insert(nuevo_cliente).execute()
                    st.success("¡Guardado!")
                except: st.error("Error al guardar.")
