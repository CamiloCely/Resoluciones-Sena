import os
import re
import time
import datetime
import pandas as pd
from pypdf import PdfReader
from docx import Document
from docx.shared import Cm, Pt
import streamlit as st

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="Generador de Resoluciones SENA",
    page_icon="🏛️",
    layout="centered"
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- SECCIÓN DE ADMINISTRACIÓN DE ARCHIVOS BASE ---
st.sidebar.title("⚙️ Administración de Bases de Datos")

nuevo_excel = st.sidebar.file_uploader("📊 Actualizar Excel Kactus / Vacaciones", type=["xlsx", "xls"])
if nuevo_excel is not None:
    path_nuevo_excel = os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx")
    with open(path_nuevo_excel, "wb") as f:
        f.write(nuevo_excel.getbuffer())
    st.sidebar.success("✅ Base de Vacaciones actualizada.")

nuevo_maestro = st.sidebar.file_uploader("📋 Actualizar Maestro por Dependencias (PDF)", type=["pdf"])
if nuevo_maestro is not None:
    path_nuevo_maestro = os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf")
    with open(path_nuevo_maestro, "wb") as f:
        f.write(nuevo_maestro.getbuffer())
    st.sidebar.success("✅ Maestro de Dependencias actualizado.")

st.sidebar.divider()

archivos_excel = [f for f in os.listdir(BASE_DIR) if f.lower().endswith(('.xlsx', '.xls'))]
EXCEL_HISTORIAL = os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx") if os.path.exists(os.path.join(BASE_DIR, "Kactus_Actualizado.xlsx")) else (os.path.join(BASE_DIR, archivos_excel[0]) if archivos_excel else None)

archivos_word = [f for f in os.listdir(BASE_DIR) if f.lower().endswith('.docx') and not f.startswith('~$')]
PLANTILLA_WORD = os.path.join(BASE_DIR, archivos_word[0]) if archivos_word else None

archivos_pdf_maestro = [f for f in os.listdir(BASE_DIR) if "MAESTRO" in f.upper() and f.lower().endswith('.pdf')]
MAESTRO_CARGOS = os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf") if os.path.exists(os.path.join(BASE_DIR, "MAESTRO_CARGOS_ACTUALIZADO.pdf")) else (os.path.join(BASE_DIR, archivos_pdf_maestro[0]) if archivos_pdf_maestro else None)

if st.sidebar.button("🔄 Reiniciar Memoria / Forzar Limpieza"):
    st.cache_data.clear()
    st.rerun()

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
        txt = pag.extract_text()
        if txt:
            texto += txt + "\n"
        
    radicado = re.search(r"(?:No:|Radicado|No\.)\s*([\d\-]+)", texto)
    fecha_rad = re.search(r"(\d{1,2}\s+de\s+\w+\s+de\s+\d{4}|\d{1,2}/\d{1,2}/\d{4})", texto)
    periodo = re.search(r"(?:período|periodo)\s+(?:comprendido\s+)?entre\s+el\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+y\s+el\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    if not periodo:
        periodo = re.search(r"del\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+al\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
        
    fecha_disfrute = re.search(r"(?:partir\s+del|inicio\s+a\s+partir\s+del)\s+(?:día\s+)?(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    
    todas_cedulas = re.findall(r"(?:C\.C\.|cédula|cedula|\bNo\.\b|\bcc\b)?\s*([\d\.]{7,12})", texto, re.IGNORECASE)
    cedula_limpia = None
    for c in todas_cedulas:
        num = c.replace(".", "").strip()
        if num.isdigit() and 7 <= len(num) <= 10:
            cedula_limpia = num
            break

    meses = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}
    
    def parse_fecha(txt_fecha):
        partes = txt_fecha.lower().replace("de", "").split()
        if len(partes) >= 3 and partes[1].strip() in meses:
            return datetime.date(int(partes[2]), meses[partes[1].strip()], int(partes[0]))
        return datetime.date.today()

    f_inicio_str = fecha_disfrute.group(1).strip() if fecha_disfrute else ""
    
    return {
        "radicado": radicado.group(1) if radicado else "15-1-2026-000000",
        "fecha_radicado": fecha_rad.group(1) if fecha_rad else "FECHA_PENDIENTE",
        "periodo_inicio": periodo.group(1) if periodo else "",
        "periodo_fin": periodo.group(2) if periodo else "",
        "fecha_inicio_texto": f_inicio_str,
        "fecha_inicio_obj": parse_fecha(f_inicio_str) if f_inicio_str else datetime.date.today(),
        "cedula_extraida": cedula_limpia,
        "texto_completo_pdf": texto.upper()
    }

def obtener_cargo_y_centro_oficial(nombre_empleado, cedula=None):
    centro_oficial = "Centro de Desarrollo Agropecuario y Agroindustrial de la regional Boyacá"
    cargo_oficial = "Profesional G06"
    cargo_director = "LA SUBDIRECTORA (E)"

    if not MAESTRO_CARGOS or not os.path.exists(MAESTRO_CARGOS):
        return cargo_oficial, centro_oficial, cargo_director

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
                
                if codigo_dep_detectado == "1010":
                    cargo_director = "EL DIRECTOR REGIONAL"
                else:
                    cargo_director = "LA SUBDIRECTORA (E)"

                match_cargo = re.search(r"(Instructor\s+G\d+|Profesional\s+G\d+(?:\s*\(e\))?|Tecnico\s+G\d+|Secretaria\s+G\d+|Auxiliar\s+G\d+|Subdirector\s+De\s+Centro|Oficial\s+Mantto[^\d]*G\d+)", linea, re.IGNORECASE)
                if match_cargo:
                    cargo_oficial = match_cargo.group(1).strip()
                return cargo_oficial, centro_oficial, cargo_director
                
    return cargo_oficial, centro_oficial, cargo_director

def forzar_formato_una_hoja(doc, dic_reemplazos):
    for section in doc.sections:
        section.top_margin = Cm(2.2)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.5)

    def procesar_parrafo(p):
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)

        for k, v in dic_reemplazos.items():
            if k in p.text:
                for r in p.runs:
                    if k in r.text:
                        r.text = r.text.replace(k, str(v))
                
                if k in p.text:
                    full_text = p.text.replace(k, str(v))
                    runs_no_imagen = [r for r in p.runs if not any(img in r._element.xml for img in ['w:drawing', 'w:pict', 'a:blip', 'v:shape'])]
                    if runs_no_imagen:
                        runs_no_imagen[0].text = full_text
                        for r in runs_no_imagen[1:]:
                            r.text = ""

    for p in doc.paragraphs:
        procesar_parrafo(p)

    for tabla in doc.tables:
        for fila in tabla.rows:
            for celda in fila.cells:
                for p in celda.paragraphs:
                    procesar_parrafo(p)

# --- INTERFAZ PRINCIPAL STREAMLIT ---
st.title("🏛️ Sistema Automático de Resoluciones de Vacaciones")
st.markdown("Carga la carta de solicitud enviada por el funcionario (PDF) para generar la resolución oficial en Word.")

if not EXCEL_HISTORIAL or not PLANTILLA_WORD:
    st.error("⚠️ Verifica que la plantilla .docx y la base Excel estén configuradas.")
else:
    archivo_pdf = st.file_uploader("Arrastra aquí la carta de solicitud recibida (.pdf)", type=["pdf"], key="pdf_uploader")

    if archivo_pdf is not None:
        st.info("📄 Carta cargada con éxito. Haz clic abajo para procesar la resolución.")
        
        if st.button("⚡ Generar Resolución en Word"):
            with st.spinner("Buscando coincidencia exacta en la base de datos Kactus..."):
                datos_carta = extraer_datos_carta(archivo_pdf)
                
                xls = pd.ExcelFile(EXCEL_HISTORIAL)
                nombre_hoja = 'KactuS - KNmVacac' if 'KactuS - KNmVacac' in xls.sheet_names else xls.sheet_names[0]
                df_kactus = pd.read_excel(EXCEL_HISTORIAL, sheet_name=nombre_hoja)

                fila_encontrada = None
                
                # 1. Buscar por Cédula extraída del PDF
                if datos_carta['cedula_extraida']:
                    filas = df_kactus[df_kactus['Identificación'].astype(str).str.contains(datos_carta['cedula_extraida'])]
                    if not filas.empty:
                        fila_encontrada = filas.iloc[0]
                
                # 2. Buscar por coincidencia de Nombres y Apellidos en el texto del PDF
                if fila_encontrada is None:
                    texto_pdf = datos_carta['texto_completo_pdf']
                    for idx, fila in df_kactus.iterrows():
                        nom = str(fila['Nombre del Empleado']).strip().upper()
                        ape = str(fila['Apellidos Empleado']).strip().upper()
                        p_nom = nom.split()[0] if nom else ""
                        p_ape = ape.split()[0] if ape else ""
                        
                        if len(p_nom) > 2 and len(p_ape) > 2:
                            if p_nom in texto_pdf and p_ape in texto_pdf:
                                fila_encontrada = fila
                                break

                if fila_encontrada is None:
                    st.error("❌ No se pudo identificar al funcionario del PDF en la base de datos Kactus. Verifica que el PDF contenga la cédula o nombres legibles.")
                else:
                    cedula_num = int(fila_encontrada['Identificación'])
                    cedula_puntos = f"{cedula_num:,}".replace(",", ".")
                    nombre_completo = f"{fila_encontrada['Nombre del Empleado']} {fila_encontrada['Apellidos Empleado']}".upper()
                    
                    cargo, centro, cargo_director = obtener_cargo_y_centro_oficial(nombre_completo, str(cedula_num))
                    
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
                    
                    reemplazos = {
                        "[CARGO_DIRECTOR]": cargo_director,
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
                        "[FECHA_HOY]": fecha_hoy_str
                    }
                    
                    forzar_formato_una_hoja(doc, reemplazos)

                    # NOMBRE DE ARCHIVO ÚNICO CON TIMESTAMP PARA EVITAR BUGS DE CACHÉ
                    timestamp_unico = int(time.time())
                    nombre_archivo_salida = f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}_{timestamp_unico}.docx"
                    salida_path = os.path.join(BASE_DIR, nombre_archivo_salida)
                    doc.save(salida_path)

                    st.balloons()
                    st.success(f"✅ ¡Resolución generada exitosamente!")
                    
                    # VISTA PREVIA DIRECTA
                    st.markdown("### 📋 Datos de la Resolución Generada:")
                    st.write(f"👤 **Funcionario:** {nombre_completo}")
                    st.write(f"🪪 **Cédula:** {cedula_puntos}")
                    st.write(f"💼 **Cargo:** {cargo}")
                    st.write(f"🏢 **Centro:** {centro}")
                    st.write(f"📅 **Período de Disfrute:** Del {fecha_inicio_formateada} al {fecha_fin_str}")

                    with open(salida_path, "rb") as file_docx:
                        st.download_button(
                            label=f"📥 DESCARGAR DOCUMENTO DE {nombre_completo}",
                            data=file_docx,
                            file_name=f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
