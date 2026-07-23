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

# --- ADMINISTRACIÓN DE BASES DE DATOS ---
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
            
    # 1. RADICADO (Capta "No: 15-1-2026-004250")
    radicado_match = re.search(r"(?:No:|Radicado|No\.)\s*([\d\-]{10,25})", texto, re.IGNORECASE)
    radicado = radicado_match.group(1).strip() if radicado_match else "15-1-2026-004250"

    # 2. FECHA DE RADICADO (Capta fechas debajo del radicado o en la carta)
    fecha_rad_match = re.search(r"(\d{1,2}/\d{1,2}/\d{4})", texto)
    if not fecha_rad_match:
        fecha_rad_match = re.search(r"(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
    
    meses = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}

    fecha_rad_str = "24 de junio de 2026"
    if fecha_rad_match:
        f_raw = fecha_rad_match.group(1).strip()
        if "/" in f_raw:
            p = f_raw.split("/")
            m_num = int(p[1])
            meses_inv = {1:"enero", 2:"febrero", 3:"marzo", 4:"abril", 5:"mayo", 6:"junio", 7:"julio", 8:"agosto", 9:"septiembre", 10:"octubre", 11:"noviembre", 12:"diciembre"}
            fecha_rad_str = f"{p[0]} de {meses_inv.get(m_num, 'junio')} de {p[2]}"
        else:
            fecha_rad_str = f_raw

    # 3. PERÍODO CAUSADO ("entre 01 de marzo de 2025 y el 28 de febrero de 2026")
    periodo_match = re.search(r"comprendido\s+entre\s+([^\n]+?)\s+y\s+el\s+([^\n]+?)\s+a\s+partir", texto, re.IGNORECASE)
    if not periodo_match:
        periodo_match = re.search(r"entre\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})\s+y\s+el\s+(\d{1,2}\s+de\s+\w+\s+de\s+\d{4})", texto, re.IGNORECASE)
        
    p_inicio = periodo_match.group(1).strip() if periodo_match else "01 de marzo de 2025"
    p_fin = periodo_match.group(2).strip() if periodo_match else "28 de febrero de 2026"

    # 4. FECHA DE DISFRUTE ("a partir del 27 de julio")
    disfrute_match = re.search(r"a\s+partir\s+del\s+(\d{1,2}\s+de\s+\w+(?:\s+de\s+\d{4})?)", texto, re.IGNORECASE)
    
    fecha_disfrute_obj = datetime.date(2026, 7, 27)
    if disfrute_match:
        raw_disf = disfrute_match.group(1).lower().replace(".", "").strip()
        partes = raw_disf.split()
        d_dia = int(partes[0])
        d_mes = meses.get(partes[2], 7) if len(partes) >= 3 else 7
        d_ano = int(partes[4]) if len(partes) >= 5 else 2026
        fecha_disfrute_obj = datetime.date(d_ano, d_mes, d_dia)

    # 5. CÉDULA
    todas_cedulas = re.findall(r"(?:C\.C\.|cédula|cedula|\bNo\.\b|\bcc\b)?\s*([\d\.]{7,12})", texto, re.IGNORECASE)
    cedula_limpia = None
    for c in todas_cedulas:
        num = c.replace(".", "").strip()
        if num.isdigit() and 7 <= len(num) <= 10:
            cedula_limpia = num
            break

    return {
        "radicado": radicado,
        "fecha_radicado": fecha_rad_str,
        "periodo_inicio": p_inicio,
        "periodo_fin": p_fin,
        "fecha_inicio_obj": fecha_disfrute_obj,
        "cedula_extraida": cedula_limpia,
        "texto_completo_pdf": texto.upper()
    }

def obtener_datos_centro_y_firmante(codigo_dep):
    DATOS_CENTROS = {
        "9110": {
            "centro": "Centro de Desarrollo Agropecuario y Agroindustrial de la regional Boyacá",
            "titulo_encabezado": "SUBDIRECTORA (E) DEL CENTRO DE DESARROLLO AGROPECUARIO Y AGROINDUSTRIAL DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\" REGIONAL BOYACÁ",
            "ciudad": "Duitama",
            "jefe_nombre": "Enith Yadira Ramírez Camargo",
            "jefe_cargo": "Subdirectora de Centro (E)"
        },
        "9111": {
            "centro": "Centro Minero de la regional Boyacá",
            "titulo_encabezado": "SUBDIRECTORA (E) DEL CENTRO MINERO DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\" REGIONAL BOYACÁ",
            "ciudad": "Sogamoso",
            "jefe_nombre": "Angela María Montoya Castro",
            "jefe_cargo": "Subdirectora (E) Centro Minero Regional Boyacá"
        },
        "9305": {
            "centro": "Centro de Gestión Administrativa y Fortalecimiento Empresarial de la regional Boyacá",
            "titulo_encabezado": "SUBDIRECTOR (E) DEL CENTRO DE GESTIÓN ADMINISTRATIVA Y FORTALECIMIENTO EMPRESARIAL DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\" REGIONAL BOYACÁ",
            "ciudad": "Tunja",
            "jefe_nombre": "Subdirector CGAFE",
            "jefe_cargo": "Subdirector de Centro (E)"
        },
        "9514": {
            "centro": "Centro Industrial de Mantenimiento y Manufactura de la regional Boyacá",
            "titulo_encabezado": "SUBDIRECTOR (E) DEL CENTRO INDUSTRIAL DE MANTENIMIENTO Y MANUFACTURA DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\" REGIONAL BOYACÁ",
            "ciudad": "Sogamoso",
            "jefe_nombre": "Subdirector CIMM",
            "jefe_cargo": "Subdirector de Centro (E)"
        },
        "1010": {
            "centro": "Despacho Dirección Regional Boyacá",
            "titulo_encabezado": "DIRECTOR REGIONAL DEL SERVICIO NACIONAL DE APRENDIZAJE \"SENA\" REGIONAL BOYACÁ",
            "ciudad": "Tunja",
            "jefe_nombre": "Director Regional",
            "jefe_cargo": "Director Regional Boyacá"
        }
    }
    return DATOS_CENTROS.get(str(codigo_dep), DATOS_CENTROS["9111"])

def obtener_cargo_y_dep(nombre_empleado, cedula=None):
    cargo_oficial = "Profesional G04"
    codigo_dep = "9111"

    if not MAESTRO_CARGOS or not os.path.exists(MAESTRO_CARGOS):
        return cargo_oficial, codigo_dep

    lector = PdfReader(MAESTRO_CARGOS)
    nombre_buscar = nombre_empleado.upper().strip()

    for pag in lector.pages:
        lineas = pag.extract_text().split("\n")
        for linea in lineas:
            if "DEPENDENCIA:" in linea:
                for cod in ["9110", "9111", "9305", "9514", "1010"]:
                    if cod in linea:
                        codigo_dep = cod
            
            coincide_cedula = cedula and cedula in linea
            partes_nom = nombre_buscar.split()
            coincide_nombre = len(partes_nom) >= 2 and partes_nom[0] in linea.upper() and partes_nom[-1] in linea.upper()
            
            if coincide_cedula or coincide_nombre:
                match_cargo = re.search(r"(Instructor\s+G\d+|Profesional\s+G\d+(?:\s*\(e\))?|Tecnico\s+G\d+|Secretaria\s+G\d+|Auxiliar\s+G\d+|Subdirector\s+De\s+Centro|Oficial\s+Mantto[^\d]*G\d+)", linea, re.IGNORECASE)
                if match_cargo:
                    cargo_oficial = match_cargo.group(1).strip()
                return cargo_oficial, codigo_dep
                
    return cargo_oficial, codigo_dep

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

# --- INTERFAZ STREAMLIT ---
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
                
                if datos_carta['cedula_extraida']:
                    filas = df_kactus[df_kactus['Identificación'].astype(str).str.contains(datos_carta['cedula_extraida'])]
                    if not filas.empty:
                        fila_encontrada = filas.iloc[0]
                
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
                    st.error("❌ No se pudo identificar al funcionario del PDF en la base de datos Kactus.")
                else:
                    cedula_num = int(fila_encontrada['Identificación'])
                    cedula_puntos = f"{cedula_num:,}".replace(",", ".")
                    nombre_completo = f"{fila_encontrada['Nombre del Empleado']} {fila_encontrada['Apellidos Empleado']}".upper()
                    
                    # DETERMINAR GÉNERO CORRECTO PARA AMBOS PÁRRAFOS
                    genero = str(fila_encontrada.get('Sexo', '')).upper()
                    if 'F' in genero or nombre_completo.startswith(('BLANCA', 'MARIA', 'ANGELA', 'NEILA', 'NIDIA', 'YADIRA', 'KATHERINE', 'SANDRA', 'PATRICIA', 'LILIANA', 'CLAUDIA', 'SONIA', 'ROSA', 'ANA')):
                        texto_funcionario = "la funcionaria"
                        texto_funcionario_a = "a la funcionaria"
                    else:
                        texto_funcionario = "el funcionario"
                        texto_funcionario_a = "al funcionario"

                    cargo, cod_dep = obtener_cargo_y_dep(nombre_completo, str(cedula_num))
                    info_centro = obtener_datos_centro_y_firmante(cod_dep)
                    
                    meses_esp = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                    
                    f_ini_obj = datos_carta['fecha_inicio_obj']
                    fecha_fin_obj = calcular_fecha_fin(f_ini_obj, 15)
                    
                    dia_fin_str = f"{fecha_fin_obj.day:02d}" if fecha_fin_obj.day < 10 else f"{fecha_fin_obj.day}"
                    fecha_fin_str = f"{dia_fin_str} de {meses_esp[fecha_fin_obj.month - 1]} de {fecha_fin_obj.year}"
                    
                    dia_ini_str = f"{f_ini_obj.day:02d}" if f_ini_obj.day < 10 else f"{f_ini_obj.day}"
                    fecha_inicio_formateada = f"{dia_ini_str} de {meses_esp[f_ini_obj.month - 1]} de {f_ini_obj.year}"

                    hoy = datetime.date.today()
                    fecha_hoy_str = f"{hoy.day:02d} de {meses_esp[hoy.month - 1]} de {hoy.year}"

                    doc = Document(PLANTILLA_WORD)
                    
                    reemplazos = {
                        "[TITULO_DIRECTOR_COMPLETO]": info_centro["titulo_encabezado"],
                        "[TEXTO_FUNCIONARIO]": texto_funcionario,
                        "[TEXTO_FUNCIONARIO_A]": texto_funcionario_a,
                        "[NOMBRE_EMPLEADO]": nombre_completo,
                        "[CEDULA]": cedula_puntos,
                        "[CARGO]": cargo,
                        "[CENTRO_FORMACION]": info_centro["centro"],
                        "[RADICADO]": datos_carta['radicado'],
                        "[FECHA_RADICADO]": datos_carta['fecha_radicado'],
                        "[FECHA_INICIO]": fecha_inicio_formateada,
                        "[FECHA_FIN]": fecha_fin_str,
                        "[PERIODO_INICIO]": datos_carta['periodo_inicio'],
                        "[PERIODO_FIN]": datos_carta['periodo_fin'],
                        "[CIUDAD_CENTRO]": info_centro["ciudad"],
                        "[FECHA_HOY]": fecha_hoy_str,
                        "[NOMBRE_JEFE_FIRMA]": info_centro["jefe_nombre"],
                        "[CARGO_JEFE_FIRMA]": info_centro["jefe_cargo"]
                    }
                    
                    reemplazar_respetando_formato(doc, reemplazos)

                    timestamp_unico = int(time.time())
                    nombre_archivo_salida = f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}_{timestamp_unico}.docx"
                    salida_path = os.path.join(BASE_DIR, nombre_archivo_salida)
                    doc.save(salida_path)

                    st.balloons()
                    st.success(f"✅ ¡Resolución generada con éxito!")
                    
                    st.markdown("### 📋 Datos Extraídos e Inserción Exacta:")
                    st.write(f"👤 **Solicitante:** {texto_funcionario.capitalize()} **{nombre_completo}**")
                    st.write(f"🔢 **Radicado:** {datos_carta['radicado']} del {datos_carta['fecha_radicado']}")
                    st.write(f"📅 **Período Causado:** Del {datos_carta['periodo_inicio']} al {datos_carta['periodo_fin']}")
                    st.write(f"🏖️ **Período Disfrute (15 Días Hábiles):** Del {fecha_inicio_formateada} al {fecha_fin_str}")

                    with open(salida_path, "rb") as file_docx:
                        st.download_button(
                            label=f"📥 DESCARGAR RESOLUCIÓN DE {nombre_completo}",
                            data=file_docx,
                            file_name=f"Resolucion_Vacaciones_{nombre_completo.replace(' ', '_')}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
