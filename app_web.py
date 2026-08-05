import os
import re
import time
import datetime
import pandas as pd
from pypdf import PdfReader
from docx import Document
import streamlit as st

st.set_page_config(
    page_title="Generador de Resoluciones SENA",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def obtener_archivo_existente(extensiones, prefijo=""):
    for f in os.listdir(BASE_DIR):
        if f.lower().endswith(extensiones) and not f.startswith('~$'):
            if prefijo:
                if prefijo.lower() in f.lower():
                    return os.path.join(BASE_DIR, f)
            else:
                return os.path.join(BASE_DIR, f)
    return None

file_excel_local = os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx") if os.path.exists(os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx")) else obtener_archivo_existente(('.xlsx', '.xls'))
file_maestro_local = os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf") if os.path.exists(os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf")) else obtener_archivo_existente(('.pdf',), prefijo="MAESTRO")
file_word_local = obtener_archivo_existente(('.docx',))

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

    # 1. RADICADO
    rad_match = re.search(r"(\d{2}\-\d{1,2}\-\d{4}\-\d{4,8})", texto_unificado)
    if not rad_match and filename_pdf:
        rad_match = re.search(r"(\d{2}\-\d{1,2}\-\d{4}\-\d{4,8})", filename_pdf)
    radicado = rad_match.group(1).strip() if rad_match else ""

    # 2. FECHA DE RADICACIÓN (Extraer la fecha pegada al sticker de radicación)
    fecha_rad_str = ""
    if radicado:
        # Busca fecha de tipo d/m/yyyy inmediatamente cerca del radicado
        sticker_match = re.search(r"(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})", texto_unificado)
        if sticker_match:
            d_rad = int(sticker_match.group(1))
            m_rad = int(sticker_match.group(2))
            a_rad = sticker_match.group(3)
            fecha_rad_str = f"{d_rad:02d} de {meses_dict.get(m_rad, 'julio')} de {a_rad}"

    if not fecha_rad_str:
        # Fallback a la fecha del encabezado (Ej: "Sogamoso, 28 Julio 2026")
        encabezado_match = re.search(r"(?:Sogamoso|Tunja|Duitama)?\s*,?\s*(\d{1,2})\s+(?:de\s+)?([a-zA-Z]+)\s+(?:de\s+)?(\d{4})", texto_unificado, re.IGNORECASE)
        if encabezado_match:
            d_enc = int(encabezado_match.group(1))
            m_enc = encabezado_match.group(2).lower()
            a_enc = encabezado_match.group(3)
            fecha_rad_str = f"{d_enc:02d} de {m_enc} de {a_enc}"

    # 3. CÉDULA DEL SOLICITANTE
    cedula_solicitante = ""
    ced_match = re.search(r"(?:Cedula|C\.C\.|Cédula)\s*(?:NO\.|No\.)?\s*([\d\.]+)", texto_unificado, re.IGNORECASE)
    if ced_match:
        cedula_solicitante = ced_match.group(1).replace(".", "").strip()

    # 4. PERIODOS CAUSADOS (Ej: vigencia 2024 al 2025 -> 16 de enero de 2024 al 15 de enero de 2025)
    p_inicio, p_fin = "", ""
    periodo_match = re.search(r"periodo\s+(?:comprendido\s+vigencia\s+)?(\d{4})\s*(?:a|al|\-)\s*(\d{4})", texto_unificado, re.IGNORECASE)
    if periodo_match:
        p_inicio = f"16 de enero de {periodo_match.group(1)}"
        p_fin = f"15 de enero de {periodo_match.group(2)}"

    # 5. FECHA DE INICIO DE DISFRUTE
    fecha_disfrute_obj = datetime.date(2026, 9, 7)
    disfrute_match = re.search(r"(?:entre|a partir del)\s+(\d{1,2})\s+DE\s+([a-zA-Z]+)(?:\s+Y\s+\d{1,2}\s+[a-zA-Z]+)?\s+(\d{4})", texto_unificado, re.IGNORECASE)
    if disfrute_match:
        dia = int(disfrute_match.group(1))
        mes_txt = disfrute_match.group(2).lower()
        anio = int(disfrute_match.group(3))
        fecha_disfrute_obj = datetime.date(anio, meses_nom.get(mes_txt, 9), dia)

    return {
        "radicado": radicado,
        "fecha_radicado": fecha_rad_str if fecha_rad_str else "29 de julio de 2026",
        "periodo_inicio": p_inicio if p_inicio else "16 de enero de 2024",
        "periodo_fin": p_fin if p_fin else "15 de enero de 2025",
        "fecha_inicio_obj": fecha_disfrute_obj,
        "cedula_solicitante": cedula_solicitante,
        "texto_completo": texto_unificado
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

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configuración y Bases de Datos")
    up_excel = st.file_uploader("Actualizar Excel Kactus", type=["xlsx", "xls"], key="up_excel")
    if up_excel:
        path_tmp_e = os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx")
        with open(path_tmp_e, "wb") as f:
            f.write(up_excel.getbuffer())
        file_excel_local = path_tmp_e
        st.success("✅ Excel KactuS cargado.")

    up_maestro = st.file_uploader("Actualizar Maestro por Dependencias (.pdf)", type=["pdf"], key="up_maestro")
    if up_maestro:
        path_tmp_m = os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf")
        with open(path_tmp_m, "wb") as f:
            f.write(up_maestro.getbuffer())
        file_maestro_local = path_tmp_m
        st.success("✅ Maestro de Cargos cargado.")

    up_word = st.file_uploader("Actualizar Plantilla Word", type=["docx"], key="up_word")
    if up_word:
        path_tmp_w = os.path.join(BASE_DIR, "Plantilla_Resolucion.docx")
        with open(path_tmp_w, "wb") as f:
            f.write(up_word.getbuffer())
        file_word_local = path_tmp_w
        st.success("✅ Plantilla Word cargada.")

# --- CONTENIDO PRINCIPAL ---
st.title("🏛️ Sistema Automático de Resoluciones de Vacaciones")

if not file_excel_local or not file_word_local:
    st.warning("⚠️ Debes subir el **Excel de KactuS** y la **Plantilla Word** en la barra lateral antes de continuar.")
else:
    archivo_pdf = st.file_uploader("Carga aquí la carta de solicitud (.pdf)", type=["pdf"], key="pdf_uploader")

    if archivo_pdf is not None:
        datos_carta = extraer_datos_pdf(archivo_pdf, archivo_pdf.name)
        
        xls = pd.ExcelFile(file_excel_local)
        nombre_hoja = 'KactuS - KNmVacac' if 'KactuS - KNmVacac' in xls.sheet_names else xls.sheet_names[0]
        df_kactus = pd.read_excel(file_excel_local, sheet_name=nombre_hoja)

        nombres_limpios = df_kactus['Nombre del Empleado'].fillna('').astype(str)
        apellidos_limpios = df_kactus['Apellidos Empleado'].fillna('').astype(str)
        df_kactus['Nombre_Completo'] = (nombres_limpios + " " + apellidos_limpios).str.strip().str.upper()
        
        lista_funcionarios = sorted([n for n in df_kactus['Nombre_Completo'].unique() if len(n) > 2])

        indice_sugerido = 0
        if datos_carta['cedula_solicitante']:
            filas_c = df_kactus[df_kactus['Identificación'].astype(str).str.contains(datos_carta['cedula_solicitante'])]
            if not filas_c.empty:
                nom_c = filas_c.iloc[0]['Nombre_Completo']
                if nom_c in lista_funcionarios:
                    indice_sugerido = lista_funcionarios.index(nom_c)

        st.success("📄 Carta analizada correctamente.")
        
        col_func, col_datos = st.columns([1, 1])
        with col_func:
            st.markdown("### 👤 Confirmación del Solicitante:")
            solicitante_elegido = st.selectbox(
                "Verifica el funcionario seleccionado:",
                options=lista_funcionarios,
                index=indice_sugerido
            )

        fila_encontrada = df_kactus[df_kactus['Nombre_Completo'] == solicitante_elegido].iloc[0]
        cedula_num = int(fila_encontrada['Identificación'])
        cedula_puntos = f"{cedula_num:,}".replace(",", ".")
        nombre_completo = solicitante_elegido

        st.markdown("---")
        st.markdown("### 📅 Ajusta o confirma la información extraída:")
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
            if 'F' in genero or nombre_completo.startswith(('CONSUELO', 'MARIA', 'BLANCA', 'ANGELA', 'NEILA', 'ENITH', 'YADIRA', 'MILENA')):
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

            doc = Document(file_word_local)
            
            reemplazos = {
                "[TITULO_DIRECTOR_COMPLETO]": "DIRECTORA (E) DE LA REGIONAL BOYACÁ DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\"",
                "[TEXTO_FUNCIONARIO]": texto_funcionario,
                "[TEXTO_FUNCIONARIO_A]": texto_funcionario_a,
                "[NOMBRE_EMPLEADO]": nombre_completo,
                "[CEDULA]": cedula_puntos,
                "[CARGO]": str(fila_encontrada.get('Cargo', 'Profesional Grado 2')),
                "[CENTRO_FORMACION]": "Centro Industrial de Mantenimiento y Manufactura de la regional Boyacá",
                "[RADICADO]": radicado_final,
                "[FECHA_RADICADO]": fecha_rad_final,
                "[FECHA_INICIO]": fecha_inicio_formateada,
                "[FECHA_FIN]": fecha_fin_str,
                "[PERIODO_INICIO]": p_ini_final,
                "[PERIODO_FIN]": p_fin_final,
                "[CIUDAD_CENTRO]": "Sogamoso",
                "[FECHA_HOY]": fecha_hoy_str,
                "[NOMBRE_JEFE_FIRMA]": "ENITH YADIRA RAMÍREZ CAMARGO",
                "[CARGO_JEFE_FIRMA]": "Directora (E) Regional Boyacá"
            }
            
            reemplazar_respetando_formato(doc, reemplazos)

            timestamp_unico = int(time.time())
            nombre_archivo_salida = f"Resolucion_{nombre_completo.replace(' ', '_')}_{timestamp_unico}.docx"
            salida_path = os.path.join(BASE_DIR, nombre_archivo_salida)
            doc.save(salida_path)

            st.balloons()
            st.success(f"✅ ¡Resolución generada para {nombre_completo}!")

            with open(salida_path, "rb") as file_docx:
                st.download_button(
                    label=f"📥 DESCARGAR RESOLUCIÓN DE {nombre_completo}",
                    data=file_docx,
                    file_name=f"Resolucion_{nombre_completo.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
