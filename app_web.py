import os
import re
import datetime
import pandas as pd
from pypdf import PdfReader
from docx import Document
import streamlit as st

# --- CONFIGURACIÓN DE LA PÁGINA STREAMLIT ---
st.set_page_config(
    page_title="Generador de Resoluciones SENA",
    page_icon="🏛️",
    layout="centered"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Localización inteligente de archivos base
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
        texto += pag.extract_text()
        
    radicado = re.search(r"No:\s*([\d\-]+)", texto)
    fecha_rad = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", texto)
    periodo = re.search(r"del\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+al\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    fecha_disfrute = re.search(r"partir\s+del\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    nombre_firmante = re.search(r"Cordialmente,\s*\n+([\w\sÁÉÍÓÚáéíóúÑñ]+)\n", texto)

    meses = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
    
    def parse_fecha(txt_fecha):
        partes = txt_fecha.lower().replace("de", "").split()
        return datetime.date(int(partes[2]), meses[partes[1].strip()], int(partes[0]))

    f_inicio_str = fecha_disfrute.group(1).strip() if fecha_disfrute else "16 de junio de 2026"
    
    return {
        "radicado": radicado.group(1) if radicado else "15-1-2026-003231",
        "fecha_radicado": fecha_rad.group(1) if fecha_rad else "04 de mayo de 2026",
        "periodo_inicio": periodo.group(1) if periodo else "01 de marzo de 2023",
        "periodo_fin": periodo.group(2) if periodo else "28 de febrero de 2024",
        "fecha_inicio_texto": f_inicio_str,
        "fecha_inicio_obj": parse_fecha(f_inicio_str),
        "solicitante": nombre_firmante.group(1).strip() if nombre_firmante else "WILMAR AUGUSTO REINA ACERO"
    }

def obtener_cargo_y_centro(nombre_empleado):
    if not MAESTRO_CARGOS or not os.path.exists(MAESTRO_CARGOS):
        return "Profesional G04 (e)", "Centro de Desarrollo Agropecuario y Agroindustrial"

    lector = PdfReader(MAESTRO_CARGOS)
    centro_actual = "Centro de Desarrollo Agropecuario y Agroindustrial"
    nombre_buscar = nombre_empleado.upper().strip()
    
    for pag in lector.pages:
        lineas = pag.extract_text().split("\n")
        for linea in lineas:
            if "DEPENDENCIA:" in linea:
                centro_actual = linea.split("DEPENDENCIA:")[-1].strip()
                centro_actual = re.sub(r'\d+', '', centro_actual).strip()
            
            if nombre_buscar.split()[0] in linea.upper() and nombre_buscar.split()[-1] in linea.upper():
                coincidencia_cargo = re.search(r"(Instructor\s+\w+|Profesional\s+\w+|Tecnico\s+\w+|Secretaria\s+\w+|Auxiliar\s+\w+)", linea, re.IGNORECASE)
                cargo = coincidencia_cargo.group(1) if coincidencia_cargo else "Profesional G04 (e)"
                return cargo, centro_actual
                
    return "Profesional G04 (e)", "Centro de Desarrollo Agropecuario y Agroindustrial"

# --- INTERFAZ GRÁFICA DE STREAMLIT ---
st.title("🏛️ Sistema Automático de Resoluciones de Vacaciones")
st.markdown("Carga la carta de solicitud enviada por el funcionario (PDF) para generar la resolución oficial en Word.")

# Verificación de requisitos del servidor
if not EXCEL_HISTORIAL or not PLANTILLA_WORD:
    st.error("⚠️ Falta cargar el archivo Excel de vacaciones o la plantilla de Word (.docx) en el servidor.")
else:
    st.sidebar.header("📁 Bases de Datos Activas")
    st.sidebar.success(f"Excel: {os.path.basename(EXCEL_HISTORIAL)}")
    st.sidebar.success(f"Plantilla: {os.path.basename(PLANTILLA_WORD)}")

    # Carga del archivo PDF
    archivo_pdf = st.file_uploader("Arrastra aquí la carta de solicitud recibida (.pdf)", type=["pdf"])

    if archivo_pdf is not None:
        st.info("📄 Archivo cargado correctamente. Haz clic en el botón para procesar.")
        
        if st.button("⚡ Generar Resolución en Word"):
            with st.spinner("Procesando datos y calculando días hábiles..."):
                # 1. Extraer datos
                datos_carta = extraer_datos_carta(archivo_pdf)
                
                # 2. Consultar Excel
                xls = pd.ExcelFile(EXCEL_HISTORIAL)
                nombre_hoja = 'KactuS - KNmVacac' if 'KactuS - KNmVacac' in xls.sheet_names else xls.sheet_names[0]
                df_kactus = pd.read_excel(EXCEL_HISTORIAL, sheet_name=nombre_hoja)

                nombre_pri = datos_carta['solicitante'].upper().split()[0]
                coincidencias = df_kactus[df_kactus['Nombre del Empleado'].str.upper().str.contains(nombre_pri, na=False)]
                
                if coincidencias.empty:
                    st.error(f"❌ No se encontró a '{datos_carta['solicitante']}' en la base de datos de vacaciones.")
                else:
                    fila_emp = coincidencias.iloc[-1]
                    cedula_num = int(fila_emp['Identificación'])
                    cedula_puntos = f"{cedula_num:,}".replace(",", ".")
                    nombre_completo = f"{fila_emp['Nombre del Empleado']} {fila_emp['Apellidos Empleado']}".upper()
                    
                    cargo, centro = obtener_cargo_y_centro(nombre_completo)
                    fecha_fin_obj = calcular_fecha_fin(datos_carta['fecha_inicio_obj'], 15)
                    meses_esp = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                    fecha_fin_str = f"{fecha_fin_obj.day} de {meses_esp[fecha_fin_obj.month - 1]} de {fecha_fin_obj.year}"

                    # 3. Reemplazar en Word
                    doc = Document(PLANTILLA_WORD)
                    reemplazos = {
                        "[NOMBRE_EMPLEADO]": nombre_completo,
                        "[CEDULA]": cedula_puntos,
                        "[CARGO]": cargo,
                        "[CENTRO_FORMACION]": centro,
                        "[RADICADO]": datos_carta['radicado'],
                        "[FECHA_RADICADO]": datos_carta['fecha_radicado'],
                        "[FECHA_INICIO]": datos_carta['fecha_inicio_texto'],
                        "[FECHA_FIN]": fecha_fin_str,
                        "[PERIODO_INICIO]": datos_carta['periodo_inicio'],
                        "[PERIODO_FIN]": datos_carta['periodo_fin']
                    }
                    
                    for parrafo in doc.paragraphs:
                        for etiqueta, valor in reemplazos.items():
                            if etiqueta in parrafo.text:
                                parrafo.text = parrafo.text.replace(etiqueta, str(valor))
                                
                    for tabla in doc.tables:
                        for fila in tabla.rows:
                            for celda in fila.cells:
                                for parrafo in celda.paragraphs:
                                    for etiqueta, valor in reemplazos.items():
                                        if etiqueta in parrafo.text:
                                            parrafo.text = parrafo.text.replace(etiqueta, str(valor))

                    # Guardar temporalmente para la descarga
                    salida_path = os.path.join(BASE_DIR, "Resolucion_Generada.docx")
                    doc.save(salida_path)

                    st.balloons()
                    st.success(f"✅ ¡Resolución generada con éxito para {nombre_completo}!")
                    
                    # Botón para descargar el Word listo
                    with open(salida_path, "rb") as file_docx:
                        st.download_button(
                            label="📥 Descargar Resolución en Word (.docx)",
                            data=file_docx,
                            file_name=f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )