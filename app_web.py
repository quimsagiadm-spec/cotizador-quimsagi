import streamlit as st
import pandas as pd
from supabase import create_client, Client
import io
import os
from datetime import datetime
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

# --- 2. EXCEL ---
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

# --- 3. CLASE ESPECIAL PDF ---
class PDFQuimsagi(FPDF):
    def footer(self):
        self.set_y(-25)
        self.set_font("Arial", 'B', 10)
        self.set_text_color(26, 82, 118) 
        self.cell(0, 5, "Favor de confirmar la cotizacion con su vendedor", ln=True, align='C')
        self.set_font("Arial", 'B', 9)
        self.cell(0, 5, "Ventas: ventas1quimsagi@gmail.com  |  Tel: 998 459 2513", ln=True, align='C')
        self.cell(0, 5, "Administracion: direccionquimsagi@gmail.com", ln=True, align='C')

# --- 4. GENERADOR DEL PDF PREMIUM ---
def generar_pdf(datos_cotizacion, cliente_info, subtotal, iva, total, vendedor, folio):
    pdf = PDFQuimsagi()
    pdf.add_page()
    
    # Marca de Agua y Logo
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

    # Títulos y Folio
    pdf.set_font("Arial", 'B', 18)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f"COTIZACION #{folio}", ln=True, align='R')
    
    pdf.set_font("Arial", 'I', 10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(0, 5, "PRODUCTOS DE LIMPIEZA A TU MEDIDA", ln=True, align='R')
    pdf.ln(10)
    
    # Datos del cliente y Vendedor
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
    
    # Tabla de productos
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
    
    # Datos Bancarios
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

# --- PÁGINA ---
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
    menu = st.sidebar.radio("Ir a:", ["📝 Cotizador", "📂 Historial", "⚙️ Panel de Administrador"])
    
    if st.sidebar.button("Cerrar Sesión"):
        st.session_state.autenticado = False
        st.rerun()

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
                        "Clave": prod_info.get("CLAVE", ""),
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
                            "detalles": st.session_state.cotizacion_actual
                        }
                        try:
                            supabase.table("cotizaciones").insert(cotizacion_data).execute()
                            st.success("¡Guardada exitosamente!")
                            st.session_state.cotizacion_actual = []
                            st.rerun()
                        except Exception as e: 
                            st.error(f"Error al guardar. Detalles: {e}")
                with col_btn2:
                    st.download_button("📥 Descargar Excel", data=generar_excel(st.session_state.cotizacion_actual, cliente_seleccionado, suma_subtotal, iva, total_final), file_name=f"Cotizacion_Folio{folio_actual}_{cliente_seleccionado}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with col_btn3:
                    st.download_button("📄 Descargar PDF", data=generar_pdf(st.session_state.cotizacion_actual, cliente_completo, suma_subtotal, iva, total_final, vendedor_seleccionado, folio_actual), file_name=f"Cotizacion_Folio{folio_actual}_{cliente_seleccionado}.pdf", mime="application/pdf", use_container_width=True)
                with col_btn4:
                    if st.button("🗑️ Limpiar Todo", use_container_width=True):
                        st.session_state.cotizacion_actual = []
                        st.rerun()

    elif menu == "📂 Historial":
        st.title("📂 Historial de Cotizaciones")
        historial = obtener_historial()
        if len(historial) == 0:
            st.info("Aún no hay cotizaciones guardadas.")
        else:
            df_historial = pd.DataFrame(historial)
            df_historial["total"] = df_historial["total"].apply(lambda x: f"${float(x):,.2f}")
            df_historial["fecha"] = pd.to_datetime(df_historial["fecha"]).dt.strftime('%d/%m/%Y %H:%M')
            
            # Verificación de la columna vendedor (por las dudas)
            if "vendedor" not in df_historial.columns: 
                df_historial["vendedor"] = "S/D"
            else:
                df_historial["vendedor"] = df_historial["vendedor"].fillna("S/D")
                
            st.dataframe(df_historial[["id", "cliente", "vendedor", "total", "fecha"]], column_config={"id": "Folio", "cliente": "Cliente", "vendedor": "Atendió", "total": "Total", "fecha": "Fecha"}, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.subheader("📄 Recuperar PDF de cotización anterior")
            
            opciones_folios = [f"Folio #{r['id']} - {r['cliente']}" for r in historial if r.get("detalles")]
            
            if len(opciones_folios) > 0:
                folio_descarga = st.selectbox("Selecciona la cotización a descargar:", opciones_folios)
                
                id_seleccionado = int(folio_descarga.split("#")[1].split(" -")[0])
                registro = next((r for r in historial if r["id"] == id_seleccionado), None)
                
                if registro and registro.get("detalles"):
                    cliente_completo = next((c for c in obtener_clientes() if c.get("RAZON_SOCIAL") == registro["cliente"]), {"RAZON_SOCIAL": registro["cliente"]})
                    datos_cot = registro["detalles"]
                    suma_subtotal = sum([float(item["Subtotal"]) for item in datos_cot])
                    iva = suma_subtotal * 0.16
                    vendedor_registro = registro.get("vendedor", "S/D")
                    
                    pdf_recuperado = generar_pdf(datos_cot, cliente_completo, suma_subtotal, iva, float(registro["total"]), vendedor_registro, id_seleccionado)
                    
                    st.download_button("📥 Descargar PDF Recuperado", data=pdf_recuperado, file_name=f"Cotizacion_Folio{id_seleccionado}_{registro['cliente']}.pdf", mime="application/pdf", use_container_width=True)
            else:
                st.info("Solo las cotizaciones nuevas que guardes a partir de hoy aparecerán aquí para descargar su PDF.")

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