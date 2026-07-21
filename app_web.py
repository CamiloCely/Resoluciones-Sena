import os
import re
import datetime
import pandas as pd
from pypdf import PdfReader
from docx import Document
import streamlit as st

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Generador de Resoluciones SENA",
    page_icon="🏛️",
    layout="centered"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Carga de archivos base en el servidor
archivos_excel = [f for f in os.listdir(BASE_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
EXCEL_HISTORIAL = os.path.join(BASE_DIR, archivos_excel[0]) if archivos_excel else None

archivos_word = [f for f in os.listdir(BASE_DIR) if f.lower().endswith('.docx') and not f.startswith('~$')]
PLANTILLA_WORD = os.path.join(BASE_DIR, archivos_word[0]) if archivos_word else None

archivos_pdf_maestro = [f for f in os.listdir(BASE_DIR) if "MAESTRO" in f.upper() and f.lower().endswith('.pdf')]
MAESTRO_CARGOS = os.path.join(BASE_DIR, archivos_pdf_maestro[0]) if archivos_pdf_maestro else None

FESTIVOS_COLOMBIA = [
    datetime.date(2026, 1, 1),   datetime.date(2026, 1, 12),  datetime.date(2026, 3, 23),
    datetime.date(2026, 4, 2),   datetime.date(2026, 4, 3),   datetime.date(2026, 5, 1),
    datetime.date(2026, 5, 18),  datetime.date(2026, 6, 8),   datetime.date(2026, 6, 15),
    datetime.date(2026, 6, 29),  datetime.date(2026, 7, 20),  datetime.date(2026, 8, 7),
    datetime.date(2026, 10, 12), datetime.date(2026, 11, 2),  datetime.date(2026, 11, 16),
    datetime.date(2026, 12, 8),  datetime.date(2026, 12, 25),
]

def calcular_fecha_fin(fecha_inicio, dias_habiles=15):
    """Calcula 15 días hábiles omitiendo festivos y fines de semana."""
    fecha_actual = fecha_inicio
    dias_contados = 0
    while dias_contados < dias_habiles:
        if fecha_actual.weekday() < 5 and fecha_actual not in FESTIVOS_COLOMBIA:
            dias_contados += 1
        if dias_contados < dias_habiles:
            fecha_actual += datetime.timedelta(days=1)
    return fecha_actual

def extraer_datos_carta(file_bytes):
    lector = PdfReader(file_bytes)
    texto = ""
    for pag in lector.pages:
        txt = pag.extract_text()
        if txt:
            texto += txt + "\n"
        
    radicado = re.search(r"(?:No:|Radicado|No\.)\s*([\d\-]+)", texto)
    fecha_rad = re.search(r"(\d{1,2}\s+de\s+\w+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})", texto)
    periodo = re.search(r"(?:período|periodo)\s+(?:comprendido\s+)?entre\s+el\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+y\s+el\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    if not periodo:
        periodo = re.search(r"del\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+al\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
        
    fecha_disfrute = re.search(r"(?:partir\s+del|inicio\s+a\s+partir\s+del)\s+(?:día\s+)?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    nombre_firmante = re.search(r"Cordialmente,\s*\n+([\w\sÁÉÍÓÚáéíóúÑñ\.]+)\n", texto)
    cedula_match = re.search(r"(?:C\.C\.|cédula|cedula)\s*([\d\.]+)", texto, re.IGNORECASE)

    meses = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
    
    def parse_fecha(txt_fecha):
        partes = txt_fecha.lower().replace("de", "").split()
        if len(partes) >= 3 and partes[1].strip() in meses:
            return datetime.date(int(partes[2]), meses[partes[1].strip()], int(partes[0]))
        return datetime.date(2026, 6, 16)

    f_inicio_str = fecha_disfrute.group(1).strip() if fecha_disfrute else "16 de junio de 2026"
    solic_nombre = nombre_firmante.group(1).replace(".", "").strip() if nombre_firmante else "WILMAR AUGUSTO REINA ACERO"
    
    return {
        "radicado": radicado.group(1) if radicado else "15-1-2026-003231",
        "fecha_radicado": fecha_rad.group(1) if fecha_rad else "04 de mayo de 2026",
        "periodo_inicio": periodo.group(1) if periodo else "01 de marzo de 2023",
        "periodo_fin": periodo.group(2) if periodo else "28 de febrero de 2024",
        "fecha_inicio_texto": f_inicio_str,
        "fecha_inicio_obj": parse_fecha(f_inicio_str),
        "solicitante": solic_nombre,
        "cedula_extraida": cedula_match.group(1).replace(".", "").strip() if cedula_match else None
    }

def obtener_cargo_y_centro_oficial(nombre_empleado, cedula=None):
    centro_oficial = "Centro de Desarrollo Agropecuario y Agroindustrial de la regional Boyacá"
    cargo_oficial = "Profesional G04 (e)"

    if not MAESTRO_CARGOS or not os.path.exists(MAESTRO_CARGOS):
        return cargo_oficial, centro_oficial

    lector = PdfReader(MAESTRO_CARGOS)
    nombre_buscar = nombre_empleado.upper().strip()
    
    NOMBRES_CENTROS_LIMPIOS = {
        "9110": "Centro de Desarrollo Agropecuario y Agroindustrial de la regional Boyacá",
        "9111": "Centro Minero de la regional Boyacá",
        "9305": "Centro de Gestión Administrativa y Fortalecimiento Empresarial de la regional Boyacá",
        "9514": "Centro Industrial de Mantenimiento y Manufactura de la regional Boyacá",
        "1010": "Despacho Dirección Regional Boyacá"
    }

    codigo_dep_detectado = "9110"

    for pag in lector.pages:
        lineas = pag.extract_text().split("\n")
        for linea in lineas:
            if "DEPENDENCIA:" in linea:
                for cod in NOMBRES_CENTROS_LIMPIOS.keys():
                    if cod in linea:
                        codigo_dep_detectado = cod
            
            coincide_cedula = cedula and cedula in linea
            partes_nom = nombre_buscar.split()
            coincide_nombre = len(partes_nom) >= 2 and partes_nom[0] in linea.upper() and partes_nom[-1] in linea.upper()
            
            if coincide_cedula or coincide_nombre:
                centro_oficial = NOMBRES_CENTROS_LIMPIOS.get(codigo_dep_detectado, NOMBRES_CENTROS_LIMPIOS["9110"])
                match_cargo = re.search(r"(Instructor\s+G\d+|Profesional\s+G\d+(?:\s*\(e\))?|Tecnico\s+G\d+|Secretaria\s+G\d+|Auxiliar\s+G\d+|Subdirector\s+De\s+Centro|Oficial\s+Mantto[^\d]*G\d+)", linea, re.IGNORECASE)
                if match_cargo:
                    cargo_oficial = match_cargo.group(1).strip()
                return cargo_oficial, centro_oficial
                
    return cargo_oficial, centro_oficial

def reemplazar_en_parrafo_sin_perder_formato(parrafo, dic_reemplazos):
    """Reemplaza texto respetando estrictamente el formato original de la plantilla."""
    texto_parrafo = parrafo.text
    for buscar, reemplazar in dic_reemplazos.items():
        if buscar in texto_parrafo:
            # Buscar en qué run está presente el texto o reemplazar conservando el estilo del primer run
            for run in parrafo.runs:
                if buscar in run.text:
                    run.text = run.text.replace(buscar, str(reemplazar))
                    return
            
            # Si el texto estaba dividido entre varios runs, reemplazar preservando el formato del primer run
            if len(parrafo.runs) > 0:
                primer_run = parrafo.runs[0]
                nuevo_texto = texto_parrafo.replace(buscar, str(reemplazar))
                primer_run.text = nuevo_texto
                for run in parrafo.runs[1:]:
                    # No borrar si el run contiene imágenes
                    if 'graphic' not in run._element.xml:
                        run.text = ""

# --- INTERFAZ STREAMLIT ---
st.title("🏛️ Sistema Automático de Resoluciones de Vacaciones")
st.markdown("Carga la carta de solicitud enviada por el funcionario (PDF) para generar la resolución oficial en Word.")

if not EXCEL_HISTORIAL or not PLANTILLA_WORD:
    st.error("⚠️ Verifica que el archivo Excel de vacaciones y la plantilla .docx estén subidos al repositorio de GitHub.")
else:
    st.sidebar.header("📁 Archivos Base Activos")
    st.sidebar.success(f"Excel: {os.path.basename(EXCEL_HISTORIAL)}")
    st.sidebar.success(f"Plantilla: {os.path.basename(PLANTILLA_WORD)}")

    archivo_pdf = st.file_uploader("Arrastra aquí la carta de solicitud recibida (.pdf)", type=["pdf"])

    if archivo_pdf is not None:
        st.info("📄 Carta cargada con éxito. Haz clic abajo para procesar la resolución.")
        
        if st.button("⚡ Generar Resolución en Word"):
            with st.spinner("Generando resolución sobre tu plantilla original..."):
                datos_carta = extraer_datos_carta(archivo_pdf)
                
                xls = pd.ExcelFile(EXCEL_HISTORIAL)
                nombre_hoja = 'KactuS - KNmVacac' if 'KactuS - KNmVacac' in xls.sheet_names else xls.sheet_names[0]
                df_kactus = pd.read_excel(EXCEL_HISTORIAL, sheet_name=nombre_hoja)

                coincidencias = pd.DataFrame()
                if datos_carta['cedula_extraida']:
                    coincidencias = df_kactus[df_kactus['Identificación'].astype(str).str.contains(datos_carta['cedula_extraida'])]
                
                if coincidencias.empty:
                    pri_nom = datos_carta['solicitante'].upper().split()[0]
                    coincidencias = df_kactus[df_kactus['Nombre del Empleado'].str.upper().str.contains(pri_nom, na=False)]
                
                if coincidencias.empty:
                    st.error(f"❌ No se encontró a {datos_carta['solicitante']} en la base de datos de vacaciones.")
                else:
                    fila_emp = coincidencias.iloc[-1]
                    cedula_num = int(fila_emp['Identificación'])
                    cedula_puntos = f"{cedula_num:,}".replace(",", ".")
                    nombre_completo = f"{fila_emp['Nombre del Empleado']} {fila_emp['Apellidos Empleado']}".upper()
                    
                    cargo, centro = obtener_cargo_y_centro_oficial(nombre_completo, str(cedula_num))
                    
                    meses_esp = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                    
                    fecha_fin_obj = calcular_fecha_fin(datos_carta['fecha_inicio_obj'], 15)
                    
                    dia_fin_str = f"{fecha_fin_obj.day:02d}" if fecha_fin_obj.day < 10 else f"{fecha_fin_obj.day}"
                    fecha_fin_str = f"{dia_fin_str} de {meses_esp[fecha_fin_obj.month - 1]} de {fecha_fin_obj.year}"
                    
                    f_inicio_obj = datos_carta['fecha_inicio_obj']
                    dia_ini_str = f"{f_inicio_obj.day:02d}" if f_inicio_obj.day < 10 else f"{f_inicio_obj.day}"
                    fecha_inicio_formateada = f"{dia_ini_str} de {meses_esp[f_inicio_obj.month - 1]} de {f_inicio_obj.year}"

                    hoy = datetime.date.today()
                    fecha_hoy_str = f"{hoy.day:02d} de {meses_esp[hoy.month - 1]} de {hoy.year}"

                    doc = Document(PLANTILLA_WORD)
                    
                    # Diccionario con las etiquetas de plantilla o texto base
                    reemplazos = {
                        "[NOMBRE_EMPLEADO]": nombre_completo,
                        "[CEDULA]": cedula_puntos,
                        "[CARGO]": cargo,
                        "[CENTRO_FORMACION]": centro,
                        "[RADICADO]": datos_carta['radicado'],
                        "[FECHA_RADICADO]": datos_carta['fecha_radicado'],
                        "[FECHA_INICIO]": fecha_inicio_formateada,
                        "[FECHA_FIN]": fecha_fin_str,
                        "[PERIODO_INICIO]": datos_carta['periodo_inicio'],
                        "[PERIODO_FIN]": datos_carta['periodo_fin'],
                        "[FECHA_HOY]": fecha_hoy_str,
                        # Reemplazo sobre texto base si la plantilla tiene datos de ejemplo
                        "WILMAR AUGUSTO REINA ACERO": nombre_completo,
                        "7.176.576": cedula_puntos,
                        "15-1-2026-003231": datos_carta['radicado'],
                        "04 de mayo de 2026": datos_carta['fecha_radicado'],
                        "16 de junio de 2026": fecha_inicio_formateada,
                        "07 de julio de 2026": fecha_fin_str,
                        "01 de marzo de 2023": datos_carta['periodo_inicio'],
                        "28 de febrero de 2024": datos_carta['periodo_fin']
                    }
                    
                    # Reemplazar de forma limpia sin dañar firmas, tipos de letra ni márgenes
                    for parrafo in doc.paragraphs:
                        reemplazar_en_parrafo_sin_perder_formato(parrafo, reemplazos)
                                
                    for tabla in doc.tables:
                        for fila in tabla.rows:
                            for celda in fila.cells:
                                for parrafo in celda.paragraphs:
                                    reemplazar_en_parrafo_sin_perder_formato(parrafo, reemplazos)

                    salida_path = os.path.join(BASE_DIR, "Resolucion_Generada.docx")
                    doc.save(salida_path)

                    st.balloons()
                    st.success(f"✅ ¡Resolución generada exactamente sobre tu plantilla para {nombre_completo}!")
                    
                    with open(salida_path, "rb") as file_docx:
                        st.download_button(
                            label="📥 Descargar Resolución en Word (.docx)",
                            data=file_docx,
                            file_name=f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
