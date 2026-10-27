import streamlit as st
import pypdf
import pandas as pd
import re
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
import io
import datetime
import zoneinfo

# Zona horaria oficial de Ecuador (Guayaquil/Cuenca UTC-5)
TZ_ECUADOR = zoneinfo.ZoneInfo("America/Guayaquil")

def obtener_hora_ecuador():
    return datetime.datetime.now(TZ_ECUADOR).strftime("%d/%m/%Y %H:%M")

# Configuración de página ancha
st.set_page_config(
    page_title="Cruce Hospitalario FEFO - Hospital Vicente Corral Moscoso",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS con la imagen permanente del hospital de fondo y diseño gerencial impecable
st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    .stApp {
        background-image: linear-gradient(rgba(15, 23, 42, 0.82), rgba(15, 23, 42, 0.88)), 
                          url('https://upload.wikimedia.org/wikipedia/commons/4/4b/Hospital_Vicente_Corral_Moscoso%2C_Cuenca.jpg');
        background-size: cover;
        background-position: center;
        background-attachment: fixed;
    }

    /* Estilo de Cabecera Institucional Tipo Portal Oficial */
    .portal-header {
        background: linear-gradient(135deg, #0b3c5d 0%, #1d2731 100%);
        padding: 25px 20px;
        border-radius: 8px;
        color: white;
        text-align: center;
        box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        margin-bottom: 20px;
        border-bottom: 4px solid #328cc1;
    }

    .portal-title {
        font-size: 26px;
        font-weight: 800;
        letter-spacing: 0.5px;
        text-transform: uppercase;
        margin: 0;
        color: #ffffff;
    }

    .portal-subtitle {
        font-size: 13px;
        font-weight: 500;
        letter-spacing: 2px;
        text-transform: uppercase;
        color: #66b3ff;
        margin-top: 6px;
    }

    /* Tarjetas de Métricas Profesionales con Fondo Sólido para Alta Visibilidad */
    div[data-testid="stMetric"] {
        background-color: #ffffff !important;
        border: 1px solid #d0d7de !important;
        padding: 15px !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 10px rgba(0,0,0,0.15) !important;
    }

    div[data-testid="stMetric"] label {
        color: #1f4e78 !important;
        font-weight: 700 !important;
        font-size: 13px !important;
    }

    div[data-testid="stMetric"] div[data-testid="stMetricValue"] {
        color: #0b3c5d !important;
        font-weight: 800 !important;
        font-size: 24px !important;
    }

    .acta-document-box {
        background-color: rgba(255, 255, 255, 0.97);
        border: 1px solid #d0d7de;
        border-radius: 8px;
        padding: 30px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        margin-top: 15px;
    }

    .acta-header-title {
        color: #1f4e78;
        font-weight: 800;
        font-size: 18px;
        text-align: center;
        margin-bottom: 2px;
        text-transform: uppercase;
    }

    .acta-sub-title {
        color: #1f4e78;
        font-weight: 700;
        font-size: 14px;
        text-align: center;
        margin-bottom: 15px;
    }
    </style>
""", unsafe_allow_html=True)

# Cabecera Principal Institucional
st.markdown("""
    <div class="portal-header">
        <div style="font-size: 30px; margin-bottom: 3px;">🏥</div>
        <h1 class="portal-title">Hospital Provincial General Docente Vicente Corral Moscoso</h1>
        <div class="portal-subtitle">Sistema Automatizado de Cruce Hospitalario FEFO • Bodega & Farmacia</div>
    </div>
""", unsafe_allow_html=True)

# Pestañas Superiores de Navegación
tab_matriz, tab_actas, tab_lotes, tab_ayuda = st.tabs([
    "Matriz Consolidada", 
    "Actas Oficiales", 
    "Lotes y Caducidades", 
    "Plantillas y Ayuda"
])

# Módulo de Carga en Lote (Limpio, solo con los 3 archivos necesarios)
with st.expander("📂 Módulo de Carga en Lote (Requisiciones & Inventarios)", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        uploaded_pdfs = st.file_uploader("1. Requisiciones PDF (Lote)", type=["pdf"], accept_multiple_files=True, key="pdf_up")
    with col2:
        uploaded_lotes = st.file_uploader("2. Excel de Lotes y F. Cad", type=["xlsx", "xls"], key="lotes_up")
    with col3:
        uploaded_unificado = st.file_uploader("3. Excel Unificado (Stock General)", type=["xlsx", "xls"], key="unif_up")

# Funciones de procesamiento seguro y lógica FEFO
def cargar_excel_seguro(file_bytes):
    try:
        return pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
    except Exception:
        try:
            return pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')
        except Exception:
            return pd.read_excel(io.BytesIO(file_bytes))

def clean_exact(s):
    if not isinstance(s, str):
        return ""
    s = unicodedata.normalize('NFD', s).encode('ascii', 'ignore').decode("utf-8")
    s = s.upper().strip()
    s = re.sub(r'[^A-Z0-9\s]', '', s)
    s = re.sub(r'\s+', ' ', s)
    return s

def generar_pdf_actas(df_final, metadata_actas):
    output_pdf = io.BytesIO()
    doc = SimpleDocTemplate(output_pdf, pagesize=letter, rightMargin=18, leftMargin=18, topMargin=20, bottomMargin=20)
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=10, alignment=1, textColor=colors.HexColor('#1F4E78'))
    subtitle_style = ParagraphStyle('SubStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7.5, alignment=1, textColor=colors.darkgray)
    meta_style = ParagraphStyle('MetaStyle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=8, alignment=0, textColor=colors.black)
    cell_style = ParagraphStyle('CellStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5)
    cell_center = ParagraphStyle('CellCenter', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, alignment=1)
    cell_red = ParagraphStyle('CellRed', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, textColor=colors.HexColor('#C00000'))
    cell_red_center = ParagraphStyle('CellRedCenter', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=6.5, textColor=colors.HexColor('#C00000'), alignment=1)
    sig_style = ParagraphStyle('SigStyle', parent=styles['Normal'], fontName='Helvetica', fontSize=6.5, alignment=1)

    fecha_emision_ec = obtener_hora_ecuador()

    for i, meta in enumerate(metadata_actas):
        p_name_actual = meta['nombre_pdf']
        df_acta_actual = df_final[df_final['PDF Origen'] == p_name_actual]
        if df_acta_actual.empty:
            continue

        elements.append(Paragraph("HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO", title_style))
        elements.append(Paragraph(f"ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO) - Archivo: {p_name_actual}", subtitle_style))
        elements.append(Spacer(1, 4))
        elements.append(Paragraph(f"<b>SERVICIO:</b> {meta['servicio']} &nbsp;&nbsp;|&nbsp;&nbsp; <b>EMISIÓN:</b> {fecha_emision_ec}", meta_style))
        elements.append(Spacer(1, 4))

        table_data = [['Nro.', 'Código', 'SKU', 'Dispositivo Médico', 'Cant. Ap.', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad', 'Stock']]
        for _, row in df_acta_actual.iterrows():
            is_red = row['Alerta_Rojo'] or (row['Stock Bodega'] <= 0)
            c_style = cell_red if is_red else cell_style
            c_style_ctr = cell_red_center if is_red else cell_center
            table_data.append([
                str(row['Nro.']),
                Paragraph(str(row['Código']), c_style_ctr),
                Paragraph(str(row['SKU']), c_style_ctr),
                Paragraph(str(row['Dispositivo Médico']), c_style),
                str(int(row['Cant. Ap.'])),
                Paragraph(str(row['Lotes y Cantidades a Despachar']), c_style),
                Paragraph(str(row['Fechas Caducidad']), c_style),
                Paragraph(f"{row['Stock Bodega']:,.0f}", c_style_ctr)
            ])

        t = Table(table_data, colWidths=[20, 45, 55, 135, 30, 135, 65, 36])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4E78')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 7),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,0), 3),
            ('TOPPADDING', (0,0), (-1,0), 3),
            ('GRID', (0,0), (-1,-1), 0.5, colors.lightgrey),
        ]))
        elements.append(t)
        elements.append(Spacer(1, 10))

        firma_data = [[
            Paragraph(f"<b>ELABORÓ / RECIBIDO:</b><br/><br/><br/>_________________________________<br/><b>{meta['elaborado_por']}</b>", sig_style),
            Paragraph(f"<b>ENTREGADO BODEGA:</b><br/><br/><br/>_________________________________<br/><b>{meta['entregado_por']}</b>", sig_style),
            Paragraph(f"<b>VALIDÓ / AUTORIZÓ:</b><br/><br/><br/>_________________________________<br/><b>{meta['validado_por']}</b>", sig_style)
        ]]
        t_firmas = Table(firma_data, colWidths=[185, 185, 185])
        elements.append(t_firmas)
        if i < len(metadata_actas) - 1:
            elements.append(PageBreak())

    doc.build(elements)
    output_pdf.seek(0)
    return output_pdf.getvalue()

def procesar_cruce_hospitalario_lote(lista_pdfs, excel_lotes_bytes, excel_unificado_bytes):
    df_lotes = cargar_excel_seguro(excel_lotes_bytes)
    df_unificado = cargar_excel_seguro(excel_unificado_bytes)

    df_lotes['clean_name'] = df_lotes['Producto'].apply(clean_exact)
    df_unificado['clean_name'] = df_unificado['Producto'].apply(clean_exact)

    df_lotes_grouped = df_lotes.groupby(['clean_name', 'Lote', 'F. Cad'], as_index=False)['Existen'].sum()
    df_lotes_grouped['F. Cad DT'] = pd.to_datetime(df_lotes_grouped['F. Cad'], errors='coerce')
    df_lotes_grouped = df_lotes_grouped.sort_values(by=['clean_name', 'F. Cad DT'], ascending=[True, True])

    todos_los_resultados = []
    metadata_actas = [] 

    for idx_pdf, (nombre_pdf, pdf_bytes) in enumerate(lista_pdfs):
        try:
            reader = pypdf.PdfReader(io.BytesIO(pdf_bytes))
            pdf_text = ""
            for page in reader.pages:
                text_page = page.extract_text()
                if text_page:
                    pdf_text += text_page + "\n"
        except Exception as e:
            continue

        servicio_nombre = "GENERAL"
        match_serv = re.search(r'Servicio\s*[:]\s*([^\r\n]+)', pdf_text, re.IGNORECASE)
        if match_serv:
            raw_serv = match_serv.group(1).strip()
            for stop_word in ['FECHA', 'NRO', 'VERSIÓN', 'VERSION', 'CÓDIGO', 'CODIGO']:
                if stop_word in raw_serv.upper():
                    raw_serv = raw_serv.upper().split(stop_word)[0].strip()
            if len(raw_serv) > 1:
                servicio_nombre = raw_serv

        firmantes_encontrados = []
        texto_plano_continuo = re.sub(r'[\r\n]+', ' ', pdf_text)
        matches_firma = re.finditer(r'(?:Firmado\s+electr[oó]nicamente\s+por|por\s*[:])\s*([A-ZÁÉÍÓÚÑ\s]{5,50})', texto_plano_continuo, re.IGNORECASE)
        for m in matches_firma:
            candidato = m.group(1).strip()
            palabras = [w for w in candidato.split() if w.isalpha() and len(w) > 1 and w.upper() not in ['VALIDAR', 'UNICAMENTE', 'FIRMAEC', 'TIME', 'STAMPING', 'SECURITY', 'DATA', 'QR', 'POR', 'DE', 'DEL', 'LA', 'LAS', 'LOS', 'EL', 'Y']]
            if len(palabras) >= 2:
                nombre_valido = " ".join(palabras[:4]).upper()
                if nombre_valido not in firmantes_encontrados:
                    firmantes_encontrados.append(nombre_valido)

        elaborado_por = firmantes_encontrados[0] if len(firmantes_encontrados) > 0 else "RESPONSABLE DE SERVICIO"
        validado_por = firmantes_encontrados[1] if len(firmantes_encontrados) > 1 else ("DANIEL GUILLERMO DELEG SARI" if "EMERGENCIA" in servicio_nombre.upper() else "XIMENA KARINA CABRERA ZENTENO")
        entregado_por = "GUARDALMACÉN / DESPACHADOR"

        idx_tabla = pdf_text.find("Nro.")
        texto_items = pdf_text[idx_tabla:] if idx_tabla != -1 else pdf_text
        clean_blob = re.sub(r'\s+', ' ', texto_items)
        items_extraidos = []
        
        for i in range(1, 100):
            pattern_item = rf'(?:^|\s){i}\s+(.+?)\s*\(\s*AUTORIZADA[^\)]+\)\s*(\d+)\s+(\d+)'
            m = re.search(pattern_item, clean_blob)
            if m:
                desc = m.group(1).strip()
                sol = int(m.group(2))
                aprob = int(m.group(3))
                items_extraidos.append({
                    'Nro.': i,
                    'Dispositivo Médico': desc,
                    'Cantidades Solicitadas': sol,
                    'Cantidades Aprobadas': aprob
                })
                clean_blob = clean_blob.replace(m.group(0), "", 1)
            else:
                if i > 3:
                    break

        df_pdf = pd.DataFrame(items_extraidos)
        if not df_pdf.empty:
            df_pdf['Nro.'] = range(1, len(df_pdf) + 1)
            df_pdf['clean_name'] = df_pdf['Dispositivo Médico'].apply(clean_exact)

            for _, pdf_row in df_pdf.iterrows():
                p_name = pdf_row['clean_name']
                cant_pendiente = pdf_row['Cantidades Aprobadas']
                
                match_unif = df_unificado[df_unificado['clean_name'] == p_name]
                if not match_unif.empty:
                    codigo = match_unif.iloc[0].get('Código', match_unif.iloc[0].get('Codigo', 'SIN-CODIGO'))
                    col_sku = [c for c in match_unif.columns if str(c).strip().upper() == 'SKU']
                    sku_val = match_unif.iloc[0][col_sku[0]] if col_sku else match_unif.iloc[0].get('SKU', 'SIN-SKU')
                    sku = str(sku_val).strip() if pd.notna(sku_val) else 'SIN-SKU'

                    col_bodega = [c for c in match_unif.columns if 'BOD' in str(c).upper() or 'FARM' in str(c).upper()][0] if [c for c in match_unif.columns if 'BOD' in str(c).upper()] else match_unif.columns[-1]
                    stock_bodega = match_unif.iloc[0][col_bodega]
                    if pd.isna(stock_bodega):
                        stock_bodega = 0.0
                else:
                    codigo = 'SIN-CODIGO'
                    sku = 'SIN-SKU'
                    stock_bodega = 0.0

                lotes_prod = df_lotes_grouped[df_lotes_grouped['clean_name'] == p_name]

                if lotes_prod.empty or float(stock_bodega) <= 0:
                    item_data = {
                        'Nro.': pdf_row['Nro.'],
                        'PDF Origen': nombre_pdf,
                        'Servicio': servicio_nombre,
                        'Código': str(codigo),
                        'SKU': sku,
                        'Dispositivo Médico': pdf_row['Dispositivo Médico'],
                        'Cant. Sol.': pdf_row['Cantidades Solicitadas'],
                        'Cant. Ap.': pdf_row['Cantidades Aprobadas'],
                        'Lotes y Cantidades a Despachar': 'SIN LOTE DISPONIBLE (0 unids)',
                        'Fechas Caducidad': 'S/F',
                        'Stock Bodega': float(stock_bodega),
                        'Alerta_Rojo': True
                    }
                else:
                    lotes_info = []
                    fechas_info = []
                    for _, l_row in lotes_prod.iterrows():
                        if cant_pendiente <= 0:
                            break
                        disp_lote = l_row['Existen']
                        if pd.isna(disp_lote):
                            continue
                        if disp_lote >= cant_pendiente:
                            cant_a_usar = cant_pendiente
                            cant_pendiente = 0
                        else:
                            cant_a_usar = disp_lote
                            cant_pendiente -= disp_lote
                        
                        lotes_info.append(f"{l_row['Lote']} ({int(cant_a_usar)} unids)")
                        fecha_fmt = l_row['F. Cad DT'].strftime('%d/%m/%Y') if pd.notna(l_row['F. Cad DT']) else str(l_row['F. Cad'])
                        fechas_info.append(f"{l_row['Lote']}: {fecha_fmt}")
                    
                    if cant_pendiente > 0:
                        lotes_info.append(f"FALTAN ({int(cant_pendiente)} unids)")
                        fechas_info.append("STOCK INSUFICIENTE")
                        
                    item_data = {
                        'Nro.': pdf_row['Nro.'],
                        'PDF Origen': nombre_pdf,
                        'Servicio': servicio_nombre,
                        'Código': str(codigo),
                        'SKU': sku,
                        'Dispositivo Médico': pdf_row['Dispositivo Médico'],
                        'Cant. Sol.': pdf_row['Cantidades Solicitadas'],
                        'Cant. Ap.': pdf_row['Cantidades Aprobadas'],
                        'Lotes y Cantidades a Despachar': " / ".join(lotes_info),
                        'Fechas Caducidad': " / ".join(fechas_info),
                        'Stock Bodega': float(stock_bodega),
                        'Alerta_Rojo': False
                    }
                todos_los_resultados.append(item_data)

        metadata_actas.append({
            'nombre_pdf': nombre_pdf,
            'servicio': servicio_nombre,
            'elaborado_por': elaborado_por,
            'validado_por': validado_por,
            'entregado_por': entregado_por
        })

    df_final = pd.DataFrame(todos_los_resultados)
    if df_final.empty:
        return None, None, 0, 0, 0, 0, None, []

    output_excel = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Lote FEFO"
    headers = ['Nro.', 'PDF Origen', 'Servicio', 'Código', 'SKU', 'Dispositivo Médico', 'Cant. Sol.', 'Cant. Ap.', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad', 'Stock Bodega']
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for r_idx, row in df_final.iterrows():
        ws.append([
            row['Nro.'], row['PDF Origen'], row['Servicio'], row['Código'], row['SKU'], 
            row['Dispositivo Médico'], row['Cant. Sol.'], row['Cant. Ap.'], 
            row['Lotes y Cantidades a Despachar'], row['Fechas Caducidad'], row['Stock Bodega']
        ])
    wb.save(output_excel)
    output_excel.seek(0)

    res_pdf_bytes = generar_pdf_actas(df_final, metadata_actas)

    total_req = len(lista_pdfs)
    total_items = len(df_final)
    cant_aprobada = int(df_final['Cant. Ap.'].sum())
    alertas = int(df_final['Alerta_Rojo'].sum())

    return output_excel.getvalue(), res_pdf_bytes, total_req, total_items, cant_aprobada, alertas, df_final, metadata_actas

# Botón de ejecución en lote
if uploaded_pdfs and uploaded_lotes and uploaded_unificado:
    if st.button("🚀 Ejecutar Cruce FEFO en Lote", type="primary", use_container_width=True):
        with st.spinner("Procesando datos institucionales..."):
            lista_tuplas_pdfs = [(p.name, p.read()) for p in uploaded_pdfs]
            excel_bytes = uploaded_lotes.read()
            unif_bytes = uploaded_unificado.read()
            
            res_excel, res_pdf, req_n, items_n, cant_ap_n, alertas_n, df_resultado, meta_actas = procesar_cruce_hospitalario_lote(lista_tuplas_pdfs, excel_bytes, unif_bytes)
            
            if res_excel:
                st.session_state['res_excel'] = res_excel
                st.session_state['res_pdf'] = res_pdf
                st.session_state['df_resultado'] = df_resultado
                st.session_state['req_n'] = req_n
                st.session_state['items_n'] = items_n
                st.session_state['cant_ap_n'] = cant_ap_n
                st.session_state['alertas_n'] = alertas_n
                st.session_state['meta_actas'] = meta_actas

# PESTAÑA 1: MATRIZ CONSOLIDADA
with tab_matriz:
    if 'df_resultado' in st.session_state:
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        with m1:
            st.metric("Requisiciones", st.session_state['req_n'], "Servicios")
        with m2:
            st.metric("Ítems Solicitados", st.session_state['items_n'], "Unids pedidas")
        with m3:
            st.metric("Cant. Aprobada", st.session_state['cant_ap_n'], "Para dispensación")
        with m4:
            st.metric("Despacho con Lote", st.session_state['items_n'] - st.session_state['alertas_n'], "Cubiertos")
        with m5:
            st.metric("Alertas / Faltante", st.session_state['alertas_n'], "Sin stock")
        with m6:
            cob = ((st.session_state['items_n'] - st.session_state['alertas_n']) / st.session_state['items_n'] * 100) if st.session_state['items_n'] > 0 else 0
            st.metric("Cobertura FEFO", f"{cob:.1f}%")

        st.markdown("---")

        f_col1, f_col2, f_col3, f_col4 = st.columns([2, 1, 1, 1])
        with f_col1:
            busqueda = st.text_input("🔍 Buscar por dispositivo, código, SKU, lote o PDF...", "")
        with f_col2:
            filtro_estado = st.selectbox("Filtro", ["Todos", "Con Alerta / Faltante", "100% Cubiertos"])
        with f_col3:
            st.write("")
            st.download_button("📥 Descargar XLSX", data=st.session_state['res_excel'], file_name="Consolidado_FEFO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        with f_col4:
            st.write("")
            st.download_button("📥 Descargar PDF", data=st.session_state['res_pdf'], file_name="Actas_Oficiales.pdf", mime="application/pdf", use_container_width=True)

        df_mostrar = st.session_state['df_resultado']
        if busqueda:
            mask = df_mostrar.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)
            df_mostrar = df_mostrar[mask]

        if "Con Alerta" in filtro_estado:
            df_mostrar = df_mostrar[df_mostrar['Alerta_Rojo'] == True]
        elif "100%" in filtro_estado:
            df_mostrar = df_mostrar[df_mostrar['Alerta_Rojo'] == False]

        def color_fefo_rows(row):
            if row.get('Alerta_Rojo', False):
                return ['background-color: #fde8e4; color: #c00000; font-weight: bold'] * len(row)
            return [''] * len(row)

        st.dataframe(df_mostrar.style.apply(color_fefo_rows, axis=1), use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Cargue los archivos en el módulo superior y haga clic en **Ejecutar Cruce FEFO en Lote**.")

# PESTAÑA 2: ACTAS OFICIALES
with tab_actas:
    if 'df_resultado' in st.session_state and 'meta_actas' in st.session_state:
        
        modo_edicion = st.toggle("✏️ Activar modo de edición de firmas individuales", value=False)

        meta_list = st.session_state['meta_actas']
        servicios_disponibles = [m['servicio'] for m in meta_list]
        servicios_disponibles_unicos = list(dict.fromkeys(servicios_disponibles))

        st.markdown("**Seleccionar Requisición / Acta:**")
        
        c_sel1, c_sel2 = st.columns([3, 1])
        with c_sel1:
            servicio_seleccionado = st.radio(
                "Servicio",
                options=servicios_disponibles_unicos + ["Ver Todas"],
                horizontal=True,
                label_visibility="collapsed"
            )
        with c_sel2:
            st.download_button(
                "📥 Descargar Todas en PDF", 
                data=st.session_state['res_pdf'], 
                file_name="Actas_Oficiales_Consolidadas.pdf", 
                mime="application/pdf",
                use_container_width=True
            )

        hora_emision_ec = obtener_hora_ecuador()

        for idx, meta in enumerate(meta_list):
            if servicio_seleccionado != "Ver Todas" and meta['servicio'] != servicio_seleccionado:
                continue

            p_name = meta['nombre_pdf']
            df_acta = st.session_state['df_resultado'][st.session_state['df_resultado']['PDF Origen'] == p_name]
            
            if df_acta.empty:
                continue

            num_alertas = df_acta['Alerta_Rojo'].sum()

            st.markdown('<div class="acta-document-box">', unsafe_allow_html=True)
            st.markdown('<div class="acta-header-title">HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="acta-sub-title">ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO)<br><span style="font-weight:normal; font-size:12px; color:gray;">Archivo de Origen: {p_name}</span></div>', unsafe_allow_html=True)

            col_meta1, col_meta2 = st.columns(2)
            with col_meta1:
                st.markdown(f"**SERVICIO SOLICITANTE:** <span style='color:#1f4e78;'>{meta['servicio']}</span>", unsafe_allow_html=True)
            with col_meta2:
                st.markdown(f"<div style='text-align: right;'><b>EMISIÓN:</b> {hora_emision_ec}</div>", unsafe_allow_html=True)

            st.markdown("---")

            if num_alertas > 0:
                st.markdown(f"""
                    <div style="background-color: #fde8e4; border: 1px solid #f5c2c7; color: #842029; padding: 10px 15px; border-radius: 6px; font-size: 13px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: center;">
                        <span>⚠️ Esta acta contiene <b>{num_alertas} ítem(s) con alerta roja</b> (stock insuficiente o sin lote disponible).</span>
                        <span style="font-weight:bold; text-decoration:underline;">Requiere validación de bodega</span>
                    </div>
                """, unsafe_allow_html=True)

            def color_acta_rows(row):
                if row.get('Alerta_Rojo', False):
                    return ['background-color: #fde8e4; color: #c00000; font-weight: bold'] * len(row)
                return [''] * len(row)

            cols_acta = ['Nro.', 'Código', 'Dispositivo Médico', 'Cant. Ap.', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad', 'Stock Bodega']
            st.dataframe(df_acta[cols_acta].style.apply(color_acta_rows, axis=1), use_container_width=True, hide_index=True)

            st.markdown("""
                <div style="background-color: #1f4e78; color: white; text-align: center; font-weight: bold; font-size: 12px; padding: 6px; border-radius: 4px; margin-top: 25px; margin-bottom: 15px;">
                    FIRMAS DE CONFORMIDAD Y RESPONSABILIDAD OFICIAL
                </div>
            """, unsafe_allow_html=True)

            f_col1, f_col2, f_col3 = st.columns(3)
            
            if modo_edicion:
                with f_col1:
                    st.markdown('<div style="border: 1px solid #d0d7de; border-radius: 6px; padding: 15px; background-color: #fafbfc;">', unsafe_allow_html=True)
                    st.markdown('<div style="background-color: #1f4e78; color: white; font-weight: bold; font-size: 11px; padding: 4px; border-radius: 4px; text-align: center; margin-bottom: 10px;">ELABORÓ / RECIBIDO:</div>', unsafe_allow_html=True)
                    nuevo_elab = st.text_input("Elaboró", value=meta['elaborado_por'], key=f"elab_{idx}", label_visibility="collapsed")
                    st.markdown('<div style="font-size: 10px; color: gray; text-align: center; margin-top: 5px;">Servicio Solicitante</div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                with f_col2:
                    st.markdown('<div style="border: 1px solid #d0d7de; border-radius: 6px; padding: 15px; background-color: #fafbfc;">', unsafe_allow_html=True)
                    st.markdown('<div style="background-color: #1f4e78; color: white; font-weight: bold; font-size: 11px; padding: 4px; border-radius: 4px; text-align: center; margin-bottom: 10px;">ENTREGADO BODEGA:</div>', unsafe_allow_html=True)
                    nuevo_entre = st.text_input("Entregado", value=meta['entregado_por'], key=f"entre_{idx}", label_visibility="collapsed")
                    st.markdown('<div style="font-size: 10px; color: gray; text-align: center; margin-top: 5px;">Bodega Central / Farmacia</div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                with f_col3:
                    st.markdown('<div style="border: 1px solid #d0d7de; border-radius: 6px; padding: 15px; background-color: #fafbfc;">', unsafe_allow_html=True)
                    st.markdown('<div style="background-color: #1f4e78; color: white; font-weight: bold; font-size: 11px; padding: 4px; border-radius: 4px; text-align: center; margin-bottom: 10px;">VALIDÓ / AUTORIZÓ:</div>', unsafe_allow_html=True)
                    nuevo_val = st.text_input("Validó", value=meta['validado_por'], key=f"val_{idx}", label_visibility="collapsed")
                    st.markdown('<div style="font-size: 10px; color: gray; text-align: center; margin-top: 5px;">Jefatura / Responsable</div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)

                meta_list[idx]['elaborado_por'] = nuevo_elab
                meta_list[idx]['entregado_por'] = nuevo_entre
                meta_list[idx]['validado_por'] = nuevo_val

                if st.button(f"💾 Guardar Cambios de esta Acta ({p_name})", key=f"btn_save_{idx}", type="primary"):
                    st.session_state['res_pdf'] = generar_pdf_actas(st.session_state['df_resultado'], meta_list)
                    st.success(f"¡Acta de {meta['servicio']} actualizada y guardada con éxito!")
            else:
                with f_col1:
                    st.markdown(f"""
                        <div style="border: 1px solid #d0d7de; border-radius: 6px; padding: 15px; text-align: center; background-color: #fafbfc;">
                            <div style="background-color: #1f4e78; color: white; font-weight: bold; font-size: 11px; padding: 4px; border-radius: 4px; margin-bottom: 10px;">ELABORÓ / RECIBIDO:</div>
                            <div style="height: 30px;"></div>
                            <div style="border-top: 1px solid #333; margin-top: 10px; padding-top: 5px; font-weight: bold; font-size: 11px; color: #1f4e78;">{meta['elaborado_por']}</div>
                            <div style="font-size: 10px; color: gray;">Servicio Solicitante</div>
                        </div>
                    """, unsafe_allow_html=True)
                with f_col2:
                    st.markdown(f"""
                        <div style="border: 1px solid #d0d7de; border-radius: 6px; padding: 15px; text-align: center; background-color: #fafbfc;">
                            <div style="background-color: #1f4e78; color: white; font-weight: bold; font-size: 11px; padding: 4px; border-radius: 4px; margin-bottom: 10px;">ENTREGADO BODEGA:</div>
                            <div style="height: 30px;"></div>
                            <div style="border-top: 1px solid #333; margin-top: 10px; padding-top: 5px; font-weight: bold; font-size: 11px; color: #1f4e78;">{meta['entregado_por']}</div>
                            <div style="font-size: 10px; color: gray;">Bodega Central / Farmacia</div>
                        </div>
                    """, unsafe_allow_html=True)
                with f_col3:
                    st.markdown(f"""
                        <div style="border: 1px solid #d0d7de; border-radius: 6px; padding: 15px; text-align: center; background-color: #fafbfc;">
                            <div style="background-color: #1f4e78; color: white; font-weight: bold; font-size: 11px; padding: 4px; border-radius: 4px; margin-bottom: 10px;">VALIDÓ / AUTORIZÓ:</div>
                            <div style="height: 30px;"></div>
                            <div style="border-top: 1px solid #333; margin-top: 10px; padding-top: 5px; font-weight: bold; font-size: 11px; color: #1f4e78;">{meta['validado_por']}</div>
                            <div style="font-size: 10px; color: gray;">Jefatura / Responsable</div>
                        </div>
                    """, unsafe_allow_html=True)

            st.markdown("</div><br>", unsafe_allow_html=True)

    else:
        st.info("ℹ️ Ejecute el cruce FEFO para visualizar el formato interactivo de las Actas Oficiales.")

# PESTAÑA 3: LOTES Y CADUCIDADES
with tab_lotes:
    st.subheader("📊 Lotes y Caducidades Registradas")
    if 'df_resultado' in st.session_state:
        st.dataframe(st.session_state['df_resultado'][['Nro.', 'PDF Origen', 'Servicio', 'Dispositivo Médico', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad']], use_container_width=True, hide_index=True)
    else:
        st.info("Cargue los archivos para ver la distribución de lotes y fechas de caducidad.")

# PESTAÑA 4: PLANTILLAS Y AYUDA
with tab_ayuda:
    st.subheader("💡 Plantillas y Ayuda del Sistema")
    st.markdown("""
    * **Paso 1:** Suba sus archivos PDF con los pedidos de requisición de los diferentes servicios hospitalarios.
    * **Paso 2:** Cargue el Excel de existencias por lote y fecha de caducidad (`F. Cad`).
    * **Paso 3:** Cargue el archivo Excel unificado de stock general de bodega o farmacia.
    * **Paso 4:** Presione el botón **Ejecutar Cruce FEFO en Lote** para consolidar automáticamente todo el sistema con el formato oficial.
    """)
