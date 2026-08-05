import os
import re
import time
import datetime
import pandas as pd
from pypdf import PdfReader
from docx import Document
import streamlit as st

# Configuración de página con la barra lateral abierta por defecto
st.set_page_config(
    page_title="Generador de Resoluciones SENA",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- GESTIÓN DE ARCHIVOS ---
def obtener_archivo_existente(extensiones, prefijo=""):
    for f in os.listdir(BASE_DIR):
        if f.lower().endswith(extensiones) and not f.startswith('~$'):
            if prefijo:
                if prefijo.lower() in f.lower():
                    return os.path.join(BASE_DIR, f)
            else:
                return os.path.join(BASE_DIR, f)
    return None

# Cargar archivos por defecto si existen
EXCEL_HISTORIAL = os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx") if os.path.exists(os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx")) else obtener_archivo_existente(('.xlsx', '.xls'))
PLANTILLA_WORD = obtener_archivo_existente(('.docx',))
MAESTRO_CARGOS = os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf") if os.path.exists(os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf")) else obtener_archivo_existente(('.pdf',), prefijo="MAESTRO")

# Días festivos oficiales
FESTIVOS_COLOMBIA = [
    datetime.date(2026, 1, 1),   datetime.date(2026, 1, 12),  datetime.date(2026, 3, 23),
    datetime.date(2026, 4, 2),   datetime.date(2026, 4, 3),   datetime.date(2026, 5, 1),
    datetime.date(2026, 5, 18),  datetime.date(2026, 6, 8),   datetime.date(2026, 6, 15),
    datetime.date(2026, 6, 29),  datetime.date(2026, 7, 20),  datetime.date(2026, 8, 7),
    datetime.date(2026, 10, 12), datetime.date(2026, 11, 2),  datetime.date(2026, 11, 16),
    datetime.date(2026, 12, 8),  datetime.date(2026, 12, 25),
]

def calcular_fecha_fin(fecha_inicio, dias_habiles=15):
    fecha_actual = fecha_inicio
    dias_contados = 0
    while dias_contados < dias_habiles:
        if fecha_actual.weekday() < 5 and fecha_actual not in FESTIVOS_COLOMBIA:
            dias_contados += 1
        if dias_contados < dias_habiles:
            fecha_actual += datetime.timedelta(days=1)
    return fecha_actual

def extraer_datos_pdf(file_bytes, filename_pdf=""):
    lector = PdfReader(file_bytes)
    texto = ""
    for pag in lector.pages:
        txt = pag.extract_text()
        if txt:
            texto += txt + "\n"
            
    meses_dict = {1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio", 7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre"}
    meses_nom = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}

    texto_unificado = " ".join(texto.split())

    rad_match = re.search(r"(\d{2}\-\d{1,2}\-\d{4}\-\d{4,8})", texto_unificado)
    if not rad_match and filename_pdf:
        rad_match = re.search(r"(\d{2}\-\d{1,2}\-\d{4}\-\d{4,8})", filename_pdf)
    radicado = rad_match.group(1).strip() if rad_match else ""

    fecha_rad_str = ""
    match_fecha_txt = re.search(r"(\d{1,2})\s+de\s+([a-zA-Z]+)\s+de\s+(\d{4})", texto_unificado, re.IGNORECASE)
    if match_fecha_txt:
        fecha_rad_str = f"{int(match_fecha_txt.group(1)):02d} de {match_fecha_txt.group(2).lower()} de {match_fecha_txt.group(3)}"
    else:
        sticker_match = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})", texto_unificado)
        if sticker_match:
            fecha_rad_str = f"{int(sticker_match.group(1)):02d} de {meses_dict.get(int(sticker_match.group(2)), 'junio')} de {sticker_match.group(3)}"

    todas_cedulas = re.findall(r"(\d{7,10})", texto_unificado)

    p_inicio, p_fin = "", ""
    periodo_match = re.search(r"(?:comprendido\s+entre\s+el|periodo\s+del|periodo\s+causado\s+del)\s+(\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4})\s+(?:al|hasta\s+el|y\s+el)\s+(\d{1,2}\s+de\s+[a-zA-Z]+\s+de\s+\d{4})", texto_unificado, re.IGNORECASE)
    if periodo_match:
        p_inicio = periodo_match.group(1).strip()
        p_fin = periodo_match.group(2).strip()

    disfrute_match = re.search(r"a\s+partir\s+del\s+(\d{1,2}\s+de\s+[a-zA-Z]+(?:\s+de\s+\d{4})?)", texto_unificado, re.IGNORECASE)
    fecha_disfrute_obj = datetime.date.today()
    if disfrute_match:
        partes = disfrute_match.group(1).lower().replace(".", "").strip().split()
        if len(partes) >= 3:
            fecha_disfrute_obj = datetime.date(int(partes[4]) if len(partes)>=5 else 2026, meses_nom.get(partes[2], 7), int(partes[0]))

    return {
        "radicado": radicado,
        "fecha_radicado": fecha_rad_str,
        "periodo_inicio": p_inicio,
        "periodo_fin": p_fin,
        "fecha_inicio_obj": fecha_disfrute_obj,
        "cedulas_extraidas": todas_cedulas,
        "texto_completo": texto_unificado.upper()
    }

def reemplazar_respetando_formato(doc, dic_reemplazos):
    def procesar_p(p):
        for k, v in dic_reemplazos.items():
            if k in p.text:
                full_text = p.text.replace(k, str(v))
                runs_no_img = [r for r in p.runs if not any(tag in r._element.xml for tag in ['w:drawing', 'w:pict', 'a:blip', 'v:shape'])]
                if runs_no_img:
                    runs_no_img[0].text = full_text
                    for r in runs_no_img[1:]:
                        r.text = ""

    for p in doc.paragraphs:
        procesar_p(p)
    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    procesar_p(p)

# --- BARRA LATERAL DE ADMINISTRACIÓN ---
with st.sidebar:
    st.header("⚙️ Configuración y Bases de Datos")
    
    st.subheader("1. Excel de KactuS / Vacaciones")
    uploaded_excel = st.file_uploader("Actualizar Excel Kactus (.xlsx)", type=["xlsx", "xls"], key="excel_uploader")
    if uploaded_excel:
        path_tmp = os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx")
        with open(path_tmp, "wb") as f:
            f.write(uploaded_excel.getbuffer())
        EXCEL_HISTORIAL = path_tmp
        st.success("✅ Base KactuS actualizada.")

    st.subheader("2. Maestro de Cargos (PDF)")
    uploaded_maestro = st.file_uploader("Actualizar PDF Maestro (.pdf)", type=["pdf"], key="maestro_uploader")
    if uploaded_maestro:
        path_tmp_m = os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf")
        with open(path_tmp_m, "wb") as f:
            f.write(uploaded_maestro.getbuffer())
        MAESTRO_CARGOS = path_tmp_m
        st.success("✅ Maestro de Cargos actualizado.")

    st.markdown("---")
    st.info(f"**Plantilla Word:** {'Detectada' if PLANTILLA_WORD else 'No encontrada'}\n\n**Base Kactus:** {'Cargada' if EXCEL_HISTORIAL else 'No encontrada'}")

# --- CONTENIDO PRINCIPAL ---
st.title("🏛️ Sistema Automático de Resoluciones de Vacaciones")
st.markdown("Carga la carta de solicitud enviada por el funcionario (PDF) para generar la resolución oficial en Word.")

if not EXCEL_HISTORIAL or not PLANTILLA_WORD:
    st.warning("⚠️ Asegúrate de tener cargada la base Excel de KactuS y la plantilla de Word en la barra lateral.")
else:
    archivo_pdf = st.file_uploader("Arrastra aquí la carta de solicitud recibida (.pdf)", type=["pdf"], key="pdf_uploader")

    if archivo_pdf is not None:
        datos_carta = extraer_datos_pdf(archivo_pdf, archivo_pdf.name)
        
        xls = pd.ExcelFile(EXCEL_HISTORIAL)
        nombre_hoja = 'KactuS - KNmVacac' if 'KactuS - KNmVacac' in xls.sheet_names else xls.sheet_names[0]
        df_kactus = pd.read_excel(EXCEL_HISTORIAL, sheet_name=nombre_hoja)

        nombres_limpios = df_kactus['Nombre del Empleado'].fillna('').astype(str)
        apellidos_limpios = df_kactus['Apellidos Empleado'].fillna('').astype(str)
        df_kactus['Nombre_Completo'] = (nombres_limpios + " " + apellidos_limpios).str.strip().str.upper()
        
        lista_funcionarios = sorted([n for n in df_kactus['Nombre_Completo'].unique() if len(n) > 2])

        indice_sugerido = 0
        for c_ext in datos_carta['cedulas_extraidas']:
            filas_c = df_kactus[df_kactus['Identificación'].astype(str).str.contains(c_ext)]
            if not filas_c.empty:
                nom_c = filas_c.iloc[0]['Nombre_Completo']
                if nom_c in lista_funcionarios:
                    indice_sugerido = lista_funcionarios.index(nom_c)
                    break

        st.success("📄 Carta analizada correctamente.")
        
        col_func, col_datos = st.columns([1, 1])
        with col_func:
            st.markdown("### 👤 Confirmación del Solicitante:")
            solicitante_elegido = st.selectbox(
                "Verifica o selecciona el funcionario:",
                options=lista_funcionarios,
                index=indice_sugerido
            )

        fila_encontrada = df_kactus[df_kactus['Nombre_Completo'] == solicitante_elegido].iloc[0]
        cedula_num = int(fila_encontrada['Identificación'])
        cedula_puntos = f"{cedula_num:,}".replace(",", ".")
        nombre_completo = solicitante_elegido

        st.markdown("---")
        st.markdown("### 📅 Ajusta o confirma las fechas e información extraída:")
        col1, col2 = st.columns(2)
        with col1:
            radicado_final = st.text_input("Número de Radicado:", value=datos_carta['radicado'])
            fecha_rad_final = st.text_input("Fecha del Radicado:", value=datos_carta['fecha_radicado'])
            p_ini_final = st.text_input("Período Causado (Inicio):", value=datos_carta['periodo_inicio'])
        with col2:
            p_fin_final = st.text_input("Período Causado (Fin):", value=datos_carta['periodo_fin'])
            f_ini_obj_final = st.date_input("Fecha Inicio Disfrute:", value=datos_carta['fecha_inicio_obj'])

        st.markdown("---")

        if st.button("⚡ Generar Resolución en Word", type="primary"):
            genero = str(fila_encontrada.get('Sexo', '')).upper()
            if 'F' in genero or nombre_completo.startswith(('BLANCA', 'MARIA', 'ANGELA', 'NEILA', 'NIDIA', 'YADIRA', 'KATHERINE', 'SANDRA', 'PATRICIA', 'LILIANA', 'CLAUDIA', 'SONIA', 'ROSA', 'ANA', 'CONSUELO', 'NORA', 'IRMA')):
                texto_funcionario = "la funcionaria"
                texto_funcionario_a = "a la funcionaria"
            else:
                texto_funcionario = "el funcionario"
                texto_funcionario_a = "al funcionario"

            meses_esp = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
            fecha_fin_obj = calcular_fecha_fin(f_ini_obj_final, 15)
            
            dia_fin_str = f"{fecha_fin_obj.day:02d}" if fecha_fin_obj.day < 10 else f"{fecha_fin_obj.day}"
            fecha_fin_str = f"{dia_fin_str} de {meses_esp[fecha_fin_obj.month - 1]} de {fecha_fin_obj.year}"
            
            dia_ini_str = f"{f_ini_obj_final.day:02d}" if f_ini_obj_final.day < 10 else f"{f_ini_obj_final.day}"
            fecha_inicio_formateada = f"{dia_ini_str} de {meses_esp[f_ini_obj_final.month - 1]} de {f_ini_obj_final.year}"

            hoy = datetime.date.today()
            fecha_hoy_str = f"{hoy.day:02d} de {meses_esp[hoy.month - 1]} de {hoy.year}"

            doc = Document(PLANTILLA_WORD)
            
            reemplazos = {
                "[TITULO_DIRECTOR_COMPLETO]": "DIRECTOR REGIONAL DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\" REGIONAL BOYACÁ",
                "[TEXTO_FUNCIONARIO]": texto_funcionario,
                "[TEXTO_FUNCIONARIO_A]": texto_funcionario_a,
                "[NOMBRE_EMPLEADO]": nombre_completo,
                "[CEDULA]": cedula_puntos,
                "[CARGO]": str(fila_encontrada.get('Cargo', 'Profesional G04')),
                "[CENTRO_FORMACION]": "Centro Industrial de Mantenimiento y Manufactura de la regional Boyacá",
                "[RADICADO]": radicado_final,
                "[FECHA_RADICADO]": fecha_rad_final,
                "[FECHA_INICIO]": fecha_inicio_formateada,
                "[FECHA_FIN]": fecha_fin_str,
                "[PERIODO_INICIO]": p_ini_final,
                "[PERIODO_FIN]": p_fin_final,
                "[CIUDAD_CENTRO]": "Sogamoso",
                "[FECHA_HOY]": fecha_hoy_str,
                "[NOMBRE_JEFE_FIRMA]": "Director Regional",
                "[CARGO_JEFE_FIRMA]": "Director Regional Boyacá"
            }
            
            reemplazar_respetando_formato(doc, reemplazos)

            timestamp_unico = int(time.time())
            nombre_archivo_salida = f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}_{timestamp_unico}.docx"
            salida_path = os.path.join(BASE_DIR, nombre_archivo_salida)
            doc.save(salida_path)

            st.balloons()
            st.success(f"✅ ¡Resolución generada con éxito!")

            with open(salida_path, "rb") as file_docx:
                st.download_button(
                    label=f"📥 DESCARGAR RESOLUCIÓN DE {nombre_completo}",
                    data=file_docx,
                    file_name=f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
