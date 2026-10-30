import shutil
import streamlit as st
import zipfile
import tempfile
import time
import os
from editar_sheets import *
from procesar_wsp import *
from generar_informe import generar_informe   # IMPORTANTE
import base64
import streamlit as st
from streamlit_extras.stylable_container import stylable_container
import pandas as pd

st.set_page_config(page_title="Sistema de Informes", layout="wide")

def modulo_datos_informe():

    with st.container(border=True):

        col_numero, col_texto = st.columns([0.7,2.8])

        with col_numero:
            st.markdown("""
            <div style="
                width:40px;
                height:40px;
                border-radius:40%;
                background:#2563EB;
                color:white;
                display:flex;
                align-items:center;
                justify-content:center;
                font-size:18px;
                font-weight:bold;
                margin-top:18px;
            ">
                1
            </div>
            """, unsafe_allow_html=True)

        with col_texto:
            st.markdown("""
            <h1 style="
                font-weight:700;
                font-size:20px;
            ">
                Datos del informe
            </h1>
                <p style="
                margin-top:-10px;
                font-weight:500;
                font-size:15px;
            ">
                Completa la infrmación general del informe
            """, unsafe_allow_html=True)
            
        sedes = leer_sedes()

        if not sedes:
            st.error("No hay sedes disponibles")
            return None

        opciones = [f"{s['EMPRESA']} - {s['SEDE']} - {s['CLIENTE']}" for s in sedes]
        seleccion = st.selectbox("Selecciona la sede", opciones)

        datos_sede = next(
                s for s in sedes
                if f"{s['EMPRESA']} - {s['SEDE']} - {s['CLIENTE']}" == seleccion
            )

        c1, c2 = st.columns(2)

        with c1:
                st.text_input("EMPRESA", datos_sede["EMPRESA"], disabled=True)
                st.text_input("SEDE", datos_sede["SEDE"], disabled=True)

        with c2:
                st.text_input("CLIENTE", datos_sede["CLIENTE"], disabled=True)
                st.text_input("DIRECCIÓN", datos_sede["DIRECCIÓN"], disabled=True)

        st.divider()

        st.markdown("""
        <style>

        /* EXPANDER AZUL */
        .st-key-expander_azul div[data-testid="stExpander"] details summary {
            background: #2563EB !important;
            color: white !important;
            border-radius: 10px;
            padding: 10px 18px;
            font-weight: 700;
            font-size: 16px;
        }

        /* HOVER */
        .st-key-expander_azul div[data-testid="stExpander"] details summary:hover {

            background: #1D4ED8 !important;
        }

        </style>
        """, unsafe_allow_html=True)


        with st.container(key="expander_azul"):

            with st.expander("Completar datos adicionales"):

                c3, c4 = st.columns(2)

                with c3:
                        inf = st.text_input("INF")
                        aten = st.text_input("ATENCIÓN")

                with c4:
                        aio = st.text_input("AIO")
                        cot = st.text_input("COTIZACIÓN")

                f1, f2 = st.columns(2)

                with f1:
                        fecha_inicio = st.date_input("FECHA INICIO")

                with f2:
                        fecha_fin = st.date_input("FECHA FIN")

                c5, c6 = st.columns(2)

                with c5:
                        elaborado = st.text_input("ELABORADO POR")

                with c6:
                        revisado = st.text_input("REVISADO POR")

    return {
        "sede": datos_sede,
        "inf": inf,
        "aten": aten,
        "aio": aio,
        "cot": cot,
        "fecha_inicio": fecha_inicio,
        "fecha_fin": fecha_fin,
        "elaborado": elaborado,
        "revisado": revisado,
    }

def modulo_zip():

    with st.container(border=True):
        col_numero, col_texto = st.columns([0.7,2.8])

        with col_numero:
            st.markdown("""
            <div style="
                width:40px;
                height:40px;
                border-radius:40%;
                background:#F59E0B;
                color:white;
                display:flex;
                align-items:center;
                justify-content:center;
                font-size:18px;
                font-weight:bold;
                margin-top:18px;
            ">
                2
            </div>
            """, unsafe_allow_html=True)

        with col_texto:
            st.markdown("""
            <h1 style="
                font-weight:700;
                font-size:20px;
            ">
                Carga de WhatsApp (ZIP)
            </h1>
                <p style="
                margin-top:-10px;
                font-weight:500;
                font-size:15px;
            ">
                Sube el archivo ZIP exportado de WhatsApp
            """, unsafe_allow_html=True)
            
        

        st.markdown("""
        <style>

        /* CONTENEDOR PRINCIPAL */
        [data-testid="stFileUploader"] {
            border: 2px dashed #FB923C;
            border-radius: 16px;
            padding: 30px 20px;
            background: rgba(251, 146, 60, 0.12);
        }

        /* HOVER */
        [data-testid="stFileUploader"]:hover {
            background: rgba(251, 146, 60, 0.28);
        }

        /* DROPZONE */
        [data-testid="stFileUploaderDropzone"] {
            background: transparent;
            border: none !important;
        }

        /* CONTENIDO INTERNO */
        [data-testid="stFileUploader"] section {
            padding: 0;

            display:flex;
            flex-direction:column;
            align-items:center;
            justify-content:center;
            
        }

        /* TEXTO */
        [data-testid="stFileUploaderDropzoneInstructions"] {
            width:100%;
        }

        [data-testid="stFileUploaderDropzoneInstructions"] div {
            width:100%;
            text-align:center !important;
        }

        /* BOTON */
        [data-testid="stBaseButton-secondary"] {
            border-radius:12px !important;
        }

        /* INSERTAR NUBE */
        [data-testid="stFileUploaderDropzone"]::before {

            content: "☁";

            font-size:100px;
            color:#F59E0B;

            display:block;
            text-align:center;

            margin-bottom:-70px;
                    
        }

        /* TEXTO */
        [data-testid="stFileUploaderDropzone"]::after {

            content: "Arrastra tu archivo ZIP aquí";

            display:block;

            text-align:center;

            font-size:14px;
            font-weight:700;

            margin-bottom:10px;
        }
                    

        </style>
        """, unsafe_allow_html=True)


        zip_file = st.file_uploader(" ", type=["zip"])

        if zip_file:
                with zipfile.ZipFile(zip_file, 'r') as z:
                    txt_file = next((f for f in z.namelist() if f.endswith(".txt")), None)
                    
                    if txt_file:
                        with z.open(txt_file) as f:
                            try:
                                contenido = f.read().decode("utf-8")
                            except:
                                contenido = f.read().decode("latin-1")
                                
                        st.markdown("""
                        <style>

                        .st-key-expander_anaranjado div[data-testid="stExpander"] details summary {

                            background: #F59E0B !important;

                            color: white !important;

                            padding: 10px 18px;

                            border-radius: 4px;

                            font-weight: 700;

                            font-size: 16px;
                        }

                        /* HOVER */
                        .st-key-expander_anaranjado div[data-testid="stExpander"] details summary:hover {

                            background: #D97706 !important;
                        }

                        </style>
                        """, unsafe_allow_html=True)


                        with st.container(key="expander_anaranjado"):
        
                            with st.expander("Vista del chat"):
                                
                                st.text_area("", contenido, height=250)
                    else:
                        st.error("No se encontró archivo .txt")

    return zip_file

def modulo_procesamiento(zip_file):

    with st.container(border=True):

        col_numero, col_texto = st.columns([0.7,2.8])

        with col_numero:
            st.markdown("""
            <div style="
                width:40px;
                height:40px;
                border-radius:40%;
                background:#22C55E;
                color:white;
                display:flex;
                align-items:center;
                justify-content:center;
                font-size:18px;
                font-weight:bold;
                margin-top:18px;
            ">
                3
            </div>
            """, unsafe_allow_html=True)

        with col_texto:
            st.markdown("""
            <h1 style="
                font-weight:700;
                font-size:20px;
            ">
                Procesamiento de datos
            </h1>
                <p style="
                margin-top:-10px;
                font-weight:500;
                font-size:15px;
            ">
                Extrae y procesa la información del chat 
            """, unsafe_allow_html=True)
            
        
        
        col1, col2, col3 = st.columns([0.01,4.5,0.2])
        with col2:
            if st.button("Procesar datos", use_container_width=True) and zip_file:

                # =========================
                # CARPETA FIJA
                # =========================
                carpeta_chat = "chat_temp"

                # crear carpeta si no existe
                os.makedirs(carpeta_chat, exist_ok=True)
                check_placeholder = st.empty()
                bar_placeholder = st.empty()

                log_placeholder = st.empty()
                st.markdown("""
                <style>

                /* CONTENEDOR DE LA BARRA */
                .stProgress > div > div > div {

                    height: 20px;
                    border-radius: 12px;
                }

                /* BARRA INTERNA */
                .stProgress > div > div > div > div {

                    background-color: #22C55E;

                    border-radius: 12px;
                }

                </style>
                """, unsafe_allow_html=True)              # BARRA
                progress_bar = st.progress(0, text="Iniciando procesamiento...")
                with bar_placeholder:
                    # =========================
                    # EXTRAER ZIP
                    # =========================
                    progress_bar.progress(10, text="10% - Extrayendo ZIP...")
                    extraer_zip_en_memoria(zip_file, carpeta_chat)
                    progress_bar.progress(30, text="30% - Buscando archivo TXT...")
                    txt_path = buscar_txt(carpeta_chat)

                    if txt_path:
                        progress_bar.progress(50,text="50% - Extrayendo tableros...")
                        tableros = extraer_tableros(txt_path)

                        progress_bar.progress(70,text="70% - Analizando imágenes con YOLO...")

                        tableros = clasificar_yolo(tableros,carpeta_chat,"best.pt",log_placeholder)

                        st.session_state["tableros"] = tableros
                        progress_bar.progress(95, text="95% - Guardando datos...")
                        progress_bar.progress(100, text="100% - Procesamiento completado")
                        # guardar JSON
                        guardar_json(tableros)
                        with check_placeholder:
                            st.markdown("""
                        <style>

                        .check-container {
                            display: flex;
                            justify-content: center;
                            align-items: center;
                            margin-top: 6px;
                            margin-bottom: 6px;
                        }

                        .check-circle {
                            width: 80px;
                            height: 80px;
                            border-radius: 50%;
                            background: #22C55E;
                            display: flex;
                            justify-content: center;
                            align-items: center;
                            animation: pop 0.4s ease;
                        }

                        .check-mark {
                            color: white;
                            font-size: 50px;
                            font-weight: bold;
                        }

                        @keyframes pop {
                            0% {
                                transform: scale(0);
                                opacity: 0;
                            }

                            80% {
                                transform: scale(1.15);
                                opacity: 1;
                            }

                            100% {
                                transform: scale(1);
                            }
                        }

                        </style>

                        <div class="check-container">
                            <div class="check-circle">
                                <div class="check-mark">✓</div>
                            </div>
                        </div>
                        """, unsafe_allow_html=True)
                        
                with st.container(border=True):
                    st.markdown(f"""
                    <h1 style="
                        margin-top:-20px;
                        font-weight:700;
                        font-size:15px;
                    ">
                        Archivo procesado:
                    </h1>

                    <p style="
                        margin-top:-10px;
                        font-weight:500;
                        font-size:12px;
                    ">
                        {zip_file.name}
                    </p>

                    <g style="
                        margin-top:-50px;
                        font-weight:700;
                        font-size:15px;
                    ">
                        Tablero detectados:
                    </g>

                    <h2 style="
                        margin-top:-30px;
                        font-weight:500;
                        font-size:12px;
                    ">
                        {len(tableros)}
                    
                    """, unsafe_allow_html=True)

def modulo_resultados():

    with st.container(border=True):
        col_numero, col_texto = st.columns([0.7,5])

        with col_numero:
            st.markdown("""
            <div style="
                width:40px;
                height:40px;
                border-radius:40%;
                background:#8B5CF6;
                color:white;
                display:flex;
                align-items:center;
                justify-content:center;
                font-size:18px;
                font-weight:bold;
                margin-top:18px;
            ">
                4
            </div>
            """, unsafe_allow_html=True)

        with col_texto:
            st.markdown("""
            <h1 style="
                font-weight:700;
                font-size:20px;
            ">
                Resultados
            </h1>
                <p style="
                margin-top:-10px;
                font-weight:500;
                font-size:15px;
            ">
                Revisa y gestiona los datos extraídos
            """, unsafe_allow_html=True)
            
        
        col1, col2 = st.columns(2)
       
        with col1:
            with st.container(border=True):
                col1_1, col1_2 = st.columns([0.3,1])
                with col1_1:
                    st.markdown("""
                    <div style="
                        width:70px;
                        height:70px;
                        border-radius:50%;
                        background:#F3F0FF;
                        color:#7C3AED;
                        display:flex;
                        align-items:center;
                        justify-content:center;
                        font-size:40px;
                        font-weight:bold;
                        margin-top:-5px;
                        margin-bottom:10px;
                    ">
                        🗔
                    </div>      
                    </div>
                    """, unsafe_allow_html=True)
                with col1_2:
                    st.markdown(f"""
                    <h1 style="
                        font-weight:700;
                        font-size:30px;
                        margin-top:-15px;
                        margin-bottom:10px;
                        margin-left:10px;
                    ">
                           {len(st.session_state.get("tableros", []))} 
                    </h1>
                        <p style="
                        margin-top:-30px;
                        font-weight:500;
                        font-size:15px;
                        margin-bottom:10px;
                        margin-left:10px;
                    ">
                        Tableros detectados
                    """, unsafe_allow_html=True)
                                 
        with col2:
            with st.container(border=True):
                col2_1, col2_2 = st.columns([0.3,1])
                with col2_1:
                        st.markdown("""
                        <div style="
                            width:70px;
                            height:70px;
                            border-radius:50%;
                            background:#DCFCE7;
                            color:#16A34A;
                            display:flex;
                            align-items:center;
                            justify-content:center;
                            font-size:50px;
                            font-weight:bold;
                            margin-top:-5px;
                            margin-bottom:10px;
                        ">
                            ▣
                        </div>
                        """, unsafe_allow_html=True)
                with col2_2:
                        st.markdown(f"""
                        <h1 style="
                            font-weight:700;
                            font-size:30px;
                            margin-top:-15px;
                            margin-bottom:10px;
                            margin-left:10px;
                        ">
                            {
                            sum(len(t.get("antes",[]))+len(t.get("durante",[]))+len(t.get("despues",[]))+len(t.get("descartadas",[]))for t in st.session_state.get("tableros",[]))
                            }
                        </h1>
                            <p style="
                            margin-top:-30px;
                            font-weight:500;
                            font-size:15px;
                            margin-bottom:10px;
                            margin-left:10px;
                        ">
                            Imágenes detectadas
                        """, unsafe_allow_html=True)
        if "tableros" not in st.session_state:
            st.info("Aún no hay datos procesados")
            return

        tableros = st.session_state["tableros"]

        st.success(f"Datos cargados")
        #lista = "\n".join([f"{i+1}. {t['tablero']}" for i, t in enumerate(tableros)])
        #st.text_area("Tableros", lista, height=200)

        # =========================
        # SUBIR A SHEETS
        # =========================
        if st.button("📗 Subir a Google Sheets"):

            subir_a_sheets(tableros, datos_informe)

            st.success("Datos subidos correctamente")


            st.session_state["esperando_confirmacion"] = True

        # =========================
        # CONFIRMACIÓN
        # =========================
        if st.session_state.get("esperando_confirmacion", False):

            opcion = st.radio(
                "¿Actualizaste los datos en el Sheet?",
                ["No", "Sí"]
            )

            st.markdown( "[Abrir Google Sheet](https://docs.google.com/spreadsheets/d/17J7v7-HWkjduQOH44RKujS-ZyFWQKes0zmWERuzp0sk/edit)" )

            if opcion == "Sí":

                #AQUÍ PASA LA MAGIA
                tableros = enriquecer_tableros_desde_sheets(tableros)

                #Guardar JSON final enriquecido
                guardar_json(tableros)

                st.session_state["tableros"] = tableros
                st.session_state["sheet_ok"] = True

                st.success("✔ Datos enriquecidos correctamente")
                #st.subheader("🧾 JSON generado")
                #st.json(st.session_state["tableros"])

            else:
                st.warning("⚠️ Actualiza primero el Sheet")
                st.session_state["sheet_ok"] = False

def modulo_generar_informe(datos, tableros):

    with st.container(border=True):
        col_numero, col_texto = st.columns([0.7,5])
        with col_numero:
            st.markdown("""
            <div style="
                width:40px;
                height:40px;
                border-radius:40%;
                background:#2563EB;
                color:white;
                display:flex;
                align-items:center;
                justify-content:center;
                font-size:18px;
                font-weight:bold;
                margin-top:18px;
            ">
                4
            </div>
            """, unsafe_allow_html=True)

        with col_texto:
            st.markdown("""
            <h1 style="
                font-weight:700;
                font-size:20px;
            ">
                Generación del informe
            </h1>
                <p style="
                margin-top:-10px;
                font-weight:500;
                font-size:15px;
            ">
                Genera y descarga el informe técnico en formato Word
            """, unsafe_allow_html=True)
      
        col1, col2 = st.columns(2)
        with col1:
            c1, c2, c3 = st.columns([1,2,1])

            with c2:

                st.markdown("""
                <div style="
                    margin-top:20px;
                    margin-bottom:10px;
                    text-align:center;
                ">
                """, unsafe_allow_html=True)

                st.image("imagen_1_carga.png")

                st.markdown("</div>", unsafe_allow_html=True)
        with col2:     
            if "tableros" not in st.session_state:
                st.warning("Primero debes procesar los datos")
                return

            if not datos:
                st.warning("Faltan datos del formulario")
                return

            # VALIDACIÓN BÁSICA
            if not datos["inf"] or not datos["cot"]:
                st.error("Completa INF y COTIZACIÓN")
                return

            # BOTÓN PRINCIPAL
            if st.button("🧾 Generar informe"):
                st.session_state["confirmar"] = True

            # CONFIRMACIÓN
            if st.session_state.get("confirmar", False):

                st.warning("⚠️ ¿Deseas generar el informe?")

                col1, col2 = st.columns(2)

                with col1:
                    if st.button("✅ Sí, generar"):

                        # =========================
                        # VISUALIZADOR
                        # =========================
                        progreso = st.progress(0)
                        estado = st.empty()

                        estado.info("⏳ Iniciando generación del informe...")

                        # Paso 1
                        progreso.progress(10)
                        time.sleep(0.3)

                        estado.info("📄 Construyendo estructura del documento...")
                        progreso.progress(25)
                        time.sleep(0.3)

                        # Paso 2
                        estado.info("🖼️ Insertando imágenes...")
                        progreso.progress(50)

                        # =========================
                        # GENERAR INFORME
                        # =========================
                        ruta = generar_informe(datos, tableros)

                        # =========================
                        # GENERAR EXCEL
                        # =========================

                        filas=[]

                        for i,t in enumerate(
                            tableros,
                            1
                        ):

                            filas.append({

                                "ITEM":i,

                                "DENOMINACIÓN":"Tablero Eléctrico",

                                "NOMBRE_TABLERO":t.get(
                                    "NOMBRE_TABLERO",
                                    t.get(
                                        "tablero",
                                        ""
                                    )
                                ),

                                "TIPO":t.get(
                                    "TIPO",
                                    ""
                                ),

                                "CAPACIDAD":t.get(
                                    "CAPACIDAD",
                                    ""
                                ),

                                "NIVEL":t.get(
                                    "NIVEL",
                                    ""
                                )
                            })

                        df=pd.DataFrame(
                            filas
                        )

                        excel_temp=tempfile.NamedTemporaryFile(
                            delete=False,
                            suffix=".xlsx"
                        )

                        with pd.ExcelWriter(
                            excel_temp.name,
                            engine="openpyxl"
                        ) as writer:

                            df.to_excel(

                                writer,

                                index=False,

                                sheet_name="Tableros"

                            )

                        ruta_excel=excel_temp.name

                        # Paso 3
                        estado.info("📊 Generando tablas y formato...")
                        progreso.progress(80)
                        time.sleep(0.3)

                        # Paso 4
                        estado.info("💾 Finalizando documento...")
                        progreso.progress(100)
                        time.sleep(0.5)

                        zip_temp = tempfile.NamedTemporaryFile(
                            delete=False,
                            suffix=".zip"
                        )

                        with zipfile.ZipFile(
                            zip_temp.name,
                            "w",
                            zipfile.ZIP_DEFLATED
                        ) as zipf:

                            zipf.write(
                                ruta,
                                arcname=os.path.basename(ruta)
                            )

                            zipf.write(
                                ruta_excel,
                                arcname=os.path.basename(ruta_excel)
                            )

                        # =========================
                        # BOTÓN ÚNICO
                        # =========================

                        with open(
                            zip_temp.name,
                            "rb"
                        ) as f_zip:

                            st.download_button(

                                label="📦 Descargar informe completo",

                                data=f_zip,

                                file_name=(
                                    f"INFORME_TECNICO_"
                                    f"{datos['inf']}_"
                                    f"{datos['sede']['CLIENTE']}.zip"
                                ),

                                mime="application/zip"

                            )
                        # =========================
                        # ELIMINAR TEMPORALES
                        # =========================
                        carpeta_chat = "chat_temp"

                        if os.path.exists(carpeta_chat):
                            shutil.rmtree(carpeta_chat)

                        st.success("✔ Proceso completado")

                        st.session_state["confirmar"] = False

                with col2:
                    if st.button("❌ Cancelar"):
                        st.info("Cancelado")
                        st.session_state["confirmar"] = False

if __name__ == "__main__":

    st.title("Sistema de automatización de informes técnicos")
    st.subheader("Genera informes técnicos de forma rápida y sencilla")
    # Crear dos columnas con la columna izquierda más ancha
    col1, col2, col3 = st.columns(3)     
    with col1:
        datos_informe = modulo_datos_informe()   
    with col2:
        zip_file = modulo_zip()
    with col3:
        modulo_procesamiento(zip_file)
    col4, col5 = st.columns(2)
    with col4:
        modulo_resultados()
    with col5:
        modulo_generar_informe(datos_informe, st.session_state.get("tableros", {}))

