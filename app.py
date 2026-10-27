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

# Configuración inicial de Streamlit
st.set_page_config(
    page_title="Cruce Hospitalario FEFO - Hospital Vicente Corral Moscoso",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Estilos CSS avanzados para la interfaz institucional
st.markdown("""
    <style>
    /* Ocultar elementos predeterminados molestos */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
    /* Fondo limpio general */
    .stApp {
        background-color: #f4f7f6;
    }

    /* Encabezado Principal */
    .header-box {
        background: linear-gradient(135deg, #0f4c81 0%, #1b365d 100%);
        padding: 25px 35px;
        border-radius: 14px;
        color: white;
        margin-bottom: 25px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.12);
        display: flex;
        align-items: center;
        justify-content: space-between;
    }
    
    .header-text-container {
        flex-grow: 1;
    }
    
    .hospital-title {
        font-size: 26px;
        font-weight: 800;
        margin: 0;
        color: #ffffff;
        letter-spacing: 0.3px;
        text-transform: uppercase;
    }
    
    .sub-title {
        font-size: 14px;
        color: #b0bec5;
        margin-top: 6px;
        font-weight: 500;
    }
    
    .msp-logo {
        background-color: white;
        padding: 8px 15px;
        border-radius: 10px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.15);
        display: flex;
        align-items: center;
        justify-content: center;
    }

    /* Tarjetas de carga de archivos de alta definición */
    .upload-card-header {
        background-color: #ffffff;
        padding: 15px 20px;
        border-radius: 12px 12px 0 0;
        border-bottom: 2px solid #0f4c81;
        font-weight: bold;
        font-size: 16px;
        color: #0f4c81;
        margin-bottom: 10px;
    }
    </style>
""", unsafe_allow_html=True)

# ENCABEZADO INSTITUCIONAL CON LOGO MSP, TÍTULO DEL HOSPITAL Y SUBTÍTULO
st.markdown("""
    <div class="header-box">
        <div class="header-text-container">
            <h1 class="hospital-title">HOSPITAL PROVINCIAL GENERAL DOCENTE VICENTE CORRAL MOSCOSO</h1>
            <p class="sub-title">Sistema de Cruce FEFO y Lotes — Bodega y Farmacia</p>
        </div>
        <div class="msp-logo">
            <img src="https://upload.wikimedia.org/wikipedia/commons/8/8c/Escudo_de_Ecuador_%28oficial%29.svg" height="55" alt="MSP Ecuador">
        </div>
    </div>
""", unsafe_allow_html=True)

st.write("")

# SECCIÓN DE CARGA EN 3 COLUMNAS LIMPIAS
col1, col2, col3 = st.columns(3)

with col1:
    st.markdown('<div class="upload-card-header">📄 1. Requisiciones PDF</div>', unsafe_allow_html=True)
    uploaded_pdfs = st.file_uploader("Seleccione o arrastre los PDFs de pedidos", type=["pdf"], accept_multiple_files=True, key="pdf_up")

with col2:
    st.markdown('<div class="upload-card-header">📊 2. Excel de Lotes</div>', unsafe_allow_html=True)
    uploaded_lotes = st.file_uploader("Subir reporte de lotes y caducidades", type=["xlsx", "xls"], key="lotes_up")

with col3:
    st.markdown('<div class="upload-card-header">📦 3. Stock General</div>', unsafe_allow_html=True)
    uploaded_unificado = st.file_uploader("Subir stock unificado de bodega", type=["xlsx", "xls"], key="unif_up")

# LÓGICA DE PROCESAMIENTO
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
        matches_firma = re.finditer(r'(?:Firmado\s+electr[oó]nic
