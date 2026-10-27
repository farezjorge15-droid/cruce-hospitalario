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

# Configuración inicial de la página
st.set_page_config(
    page_title="Cruce Hospitalario FEFO - Hospital Vicente Corral Moscoso",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .stApp { background-color: #f8f9fa; }
    </style>
""", unsafe_allow_html=True)

# Encabezado superior institucional
st.markdown("### 🏥 Cruce Hospitalario FEFO - Hospital Vicente Corral Moscoso")

# PESTAÑAS SUPERIORES IDÉNTICAS A LA INTERFAZ DE LA IMAGEN
tab_matriz, tab_actas, tab_lotes, tab_ayuda = st.tabs([
    "Matriz Consolidada", 
    "Actas Oficiales", 
    "Lotes y Caducidades", 
    "Plantillas y Ayuda"
])

# MÓDULO DE CARGA EN LOTE (Siempre disponible)
with st.expander("📂 Módulo de Carga en Lote (Requisiciones & Inventarios)", expanded=True):
    col1, col2, col3 = st.columns(3)
    with col1:
        uploaded_pdfs = st.file_uploader("1. Requisiciones PDF (Lote)", type=["pdf"], accept_multiple_files=True, key="pdf_up")
    with col2:
        uploaded_lotes = st.file_uploader("2. Excel de Lotes y F. Cad", type=["xlsx", "xls"], key="lotes_up")
    with col3:
        uploaded_unificado = st.file_uploader("3. Excel Unificado (Stock General)", type=["xlsx", "xls"], key="unif_up")

# Funciones de procesamiento seguro
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
                        'PDF Origen': nombre_pdf,
                        'Servicio': servicio_nombre,
                        'Nro.': pdf_row['Nro.'],
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
                        'PDF Origen': nombre_pdf,
                        'Servicio': servicio_nombre,
                        'Nro.': pdf_row['Nro.'],
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
        return None, None, 0, 0, 0, 0, None

    # Excel output
    output_excel = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Consolidado Lote FEFO"
    headers = ['PDF Origen', 'Servicio', 'Nro.', 'Código', 'SKU', 'Dispositivo Médico', 'Cant. Sol.', 'Cant. Ap.', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad', 'Stock Bodega']
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
            row['PDF Origen'], row['Servicio'], row['Nro.'], row['Código'], row['SKU'], 
            row['Dispositivo Médico'], row['Cant. Sol.'], row['Cant. Ap.'], 
            row['Lotes y Cantidades a Despachar'], row['Fechas Caducidad'], row['Stock Bodega']
        ])
    wb.save(output_excel)
    output_excel.seek(0)

    # PDF output
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

    for i, meta in enumerate(metadata_actas):
        p_name_actual = meta['nombre_pdf']
        df_acta_actual = df_final[df_final['PDF Origen'] == p_name_actual]
        if df_acta_actual.empty:
            continue

        elements.append(Paragraph("HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO", title_style))
        elements.append(Paragraph(f"ACTA OFICIAL DE DISPENSACIÓN Y CRUCE DE STOCK (FEFO) - Archivo: {p_name_actual}", subtitle_style))
        elements.append(Spacer(1, 4))
        elements.append(Paragraph(f"<b>SERVICIO:</b> {meta['servicio']} &nbsp;&nbsp;|&nbsp;&nbsp; <b>EMISIÓN:</b> {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", meta_style))
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

    total_req = len(lista_pdfs)
    total_items = len(df_final)
    cant_aprobada = int(df_final['Cant. Ap.'].sum())
    alertas = int(df_final['Alerta_Rojo'].sum())

    return output_excel.getvalue(), output_pdf.getvalue(), total_req, total_items, cant_aprobada, alertas, df_final

# Botón para ejecutar cruce en lote
if uploaded_pdfs and uploaded_lotes and uploaded_unificado:
    if st.button("🚀 Ejecutar Cruce FEFO en Lote", type="primary", use_container_width=True):
        with st.spinner("Procesando datos institucionales..."):
            lista_tuplas_pdfs = [(p.name, p.read()) for p in uploaded_pdfs]
            excel_bytes = uploaded_lotes.read()
            unif_bytes = uploaded_unificado.read()
            
            res_excel, res_pdf, req_n, items_n, cant_ap_n, alertas_n, df_resultado = procesar_cruce_hospitalario_lote(lista_tuplas_pdfs, excel_bytes, unif_bytes)
            
            if res_excel:
                st.session_state['res_excel'] = res_excel
                st.session_state['res_pdf'] = res_pdf
                st.session_state['df_resultado'] = df_resultado
                st.session_state['req_n'] = req_n
                st.session_state['items_n'] = items_n
                st.session_state['cant_ap_n'] = cant_ap_n
                st.session_state['alertas_n'] = alertas_n

# CONTENIDO DE LA PESTAÑA: MATRIZ CONSOLIDADA
with tab_matriz:
    if 'df_resultado' in st.session_state:
        # Tarjetas de métricas idénticas a la imagen de AI Studio
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

        # Barra de búsqueda, filtros de estado y botón de descarga XLSX
        f_col1, f_col2, f_col3 = st.columns([2, 1, 1])
        with f_col1:
            busqueda = st.text_input("🔍 Buscar por dispositivo, código, SKU, lote o PDF...", "")
        with f_col2:
            filtro_estado = st.selectbox("Filtro Estado", ["Todos", "Con Alerta / Faltante", "100% Cubiertos"])
        with f_col3:
            st.write("")
            st.download_button("📥 Descargar XLSX", data=st.session_state['res_excel'], file_name="Consolidado_FEFO.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        # Filtrado dinámico
        df_mostrar = st.session_state['df_resultado']
        if busqueda:
            mask = df_mostrar.astype(str).apply(lambda x: x.str.contains(busqueda, case=False)).any(axis=1)
            df_mostrar = df_mostrar[mask]

        if "Con Alerta" in filtro_estado:
            df_mostrar = df_mostrar[df_mostrar['Alerta_Rojo'] == True]
        elif "100%" in filtro_estado:
            df_mostrar = df_mostrar[df_mostrar['Alerta_Rojo'] == False]

        st.dataframe(df_mostrar, use_container_width=True, hide_index=True)
    else:
        st.info("ℹ️ Cargue los archivos en el módulo superior y haga clic en **Ejecutar Cruce FEFO en Lote** para habilitar la matriz y las métricas.")

# CONTENIDO DE LA PESTAÑA: ACTAS OFICIALES
with tab_actas:
    st.subheader("📑 Actas Oficiales de Dispensación (PDF)")
    if 'res_pdf' in st.session_state:
        st.success("Actas oficiales generadas correctamente y listas para descarga.")
        st.download_button("📥 Descargar Actas en PDF", data=st.session_state['res_pdf'], file_name="Actas_Oficiales.pdf", mime="application/pdf")
    else:
        st.info("Ejecute el cruce FEFO para generar las actas oficiales.")

# CONTENIDO DE LA PESTAÑA: LOTES Y CADUCIDADES
with tab_lotes:
    st.subheader("📊 Lotes y Caducidades Registradas")
    if 'df_resultado' in st.session_state:
        st.dataframe(st.session_state['df_resultado'][['PDF Origen', 'Servicio', 'Dispositivo Médico', 'Lotes y Cantidades a Despachar', 'Fechas Caducidad']], use_container_width=True, hide_index=True)
    else:
        st.info("Cargue los archivos para ver la distribución de lotes y fechas de caducidad.")

# CONTENIDO DE LA PESTAÑA: PLANTILLAS Y AYUDA
with tab_ayuda:
    st.subheader("💡 Plantillas y Ayuda")
    st.markdown("""
    * **Paso 1:** Suba sus archivos PDF con los pedidos de requisición de los diferentes servicios hospitalarios.
    * **Paso 2:** Cargue el Excel de existencias por lote y fecha de caducidad (`F. Cad`).
    * **Paso 3:** Cargue el archivo Excel unificado de stock general de bodega o farmacia.
    * **Paso 4:** Presione el botón **Ejecutar Cruce FEFO en Lote** para consolidar automáticamente todo el sistema.
    """)
