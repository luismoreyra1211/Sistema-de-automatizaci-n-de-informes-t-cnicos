import zipfile
import os
import re
import json
from google.oauth2.service_account import Credentials
import gspread
from ultralytics import YOLO
from pathlib import Path

pat_tablero = re.compile(r"(sub\s*)?tablero", re.IGNORECASE)
pat_antes = re.compile(r"\b(antes|previo|pre)\b", re.IGNORECASE)
pat_durante = re.compile(r"\b(durante|proceso|intervencion)\b", re.IGNORECASE)
pat_despues = re.compile(r"\b(despues|después|post|final)\b", re.IGNORECASE)

PALABRAS_TABLERO = [

    "TABLERO",

    "SUBTABLERO",

    "SUB TABLERO",

    "TAB",

    "TBL",

    "TD",

    "TC",

    "T.D",

    "T.C",

    "TALERO"

]

def extraer_zip_en_memoria(zip_file, extract_path):
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(extract_path)

def buscar_txt(path):
    for root, _, files in os.walk(path):
        for file in files:
            if file.endswith(".txt"):
                return os.path.join(root, file)
    return None

def reconstruir_mensajes(txt_path):

    mensajes = []

    mensaje_actual = ""

    patron_inicio = re.compile(
        r"^\d{1,2}/\d{1,2}/\d{2,4},"
    )

    with open(
        txt_path,
        "r",
        encoding="utf-8"
    ) as f:

        for linea in f:

            linea = linea.strip()

            if linea == "":
                continue

            # NUEVO MENSAJE
            if patron_inicio.match(linea):

                if mensaje_actual:

                    mensajes.append(
                        mensaje_actual
                    )

                mensaje_actual = linea

            # CONTINUACIÓN DEL MENSAJE
            else:

                mensaje_actual += " " + linea

        if mensaje_actual:

            mensajes.append(
                mensaje_actual
            )

    return mensajes

def limpiar_nombre_tablero(nombre):
    # Quitar asteriscos y caracteres raros
    nombre = re.sub(r"[*_`~]", "", nombre)

    # Quitar emojis o caracteres no ASCII (opcional pero recomendado)
    nombre = re.sub(r"[^\w\s\-]", "", nombre)

    # Reemplazar múltiples espacios por uno solo
    nombre = re.sub(r"\s+", " ", nombre)

    # Quitar espacios al inicio/fin y pasar a mayúsculas
    nombre = nombre.strip().upper()

    return nombre

def parece_tablero(texto):

    texto = texto.upper().strip()

    # --------------------------
    # Ignorar imágenes
    # --------------------------

    if "IMG-" in texto:

        return False

    if ".JPG" in texto:

        return False

    if ".JPEG" in texto:

        return False

    if ".PNG" in texto:

        return False

    # --------------------------
    # Buscar palabras conocidas
    # --------------------------

    for palabra in PALABRAS_TABLERO:

        if palabra in texto:

            return True

    # --------------------------
    # Detectar:
    # IC-TE-02
    # AA-BB-01
    # QD-1-A
    # --------------------------

    if re.fullmatch(

        r'[A-Z]{1,5}(?:-[A-Z0-9]{1,5})+',

        texto

    ):

        return True

    # --------------------------
    # Detectar:
    # TD01
    # TD-01
    # TC05
    # QD3
    # --------------------------

    if re.fullmatch(

        r'[A-Z]{1,6}-?\d{1,4}',

        texto

    ):

        return True

    return False

def extraer_tableros(txt_path):

    patron_imagenes = "IMG-"

    tableros = []

    tablero_actual = None

    # =====================================
    # LEER TXT
    # =====================================

    with open(

        txt_path,

        "r",

        encoding="utf-8"

    ) as archivo:

        lineas_raw = archivo.readlines()

    # =====================================
    # RECONSTRUIR MENSAJES WHATSAPP
    # =====================================

    lineas = []

    mensaje_actual = ""

    patron_inicio = re.compile(

        r"^\d{1,2}/\d{1,2}/\d{2,4},"

    )

    for linea in lineas_raw:

        linea = linea.strip()

        if linea == "":

            continue

        # NUEVO MENSAJE

        if patron_inicio.match(

            linea

        ):

            if mensaje_actual:

                lineas.append(

                    mensaje_actual

                )

            mensaje_actual = linea

        # CONTINUACIÓN MENSAJE

        else:

            mensaje_actual += (

                " " + linea

            )

    if mensaje_actual:

        lineas.append(

            mensaje_actual

        )

    # =====================================
    # PROCESAMIENTO
    # =====================================

    for linea in lineas:

        linea_lower = linea.lower()

        # Ignorar basura

        if (

            "creó" in linea_lower

            or

            "realizados" in linea_lower

            or

            "realizado" in linea_lower

        ):

            continue

        # =====================================
        # EXTRAER MENSAJE
        # =====================================

        if ":" in linea:

            texto = linea.rsplit(
                ":",
                1
            )[1].strip()

        else:

            texto = linea.strip()

        texto = texto.replace(
            "\u200e",
            ""
        )

        texto = texto.strip("*")

        if not texto:

            continue

        # =====================================
        # IGNORAR SI NO ES TABLERO NI IMG
        # =====================================

        if not (

            parece_tablero(
                texto
            )

            or

            patron_imagenes in texto.upper()

        ):

            continue

        texto = texto.replace(

            "\u200e",

            ""

        )

        texto = texto.strip(

            "*"

        )

        if not texto:

            continue

        # =====================================
        # TABLERO
        # =====================================

        if parece_tablero(

            texto

        ):

            nombre = limpiar_nombre_tablero(

                texto

            )

            tablero_actual = {

                "tablero": nombre,

                "antes": [],

                "durante": [],

                "despues": [],

                "descartadas": [],

                "imagenes": []

            }

            tableros.append(

                tablero_actual

            )

            continue

        # =====================================
        # IMAGENES
        # =====================================

        if tablero_actual:

            match = re.search(

                r"(IMG[-_\dA-Za-z]+\.jpg)",

                texto,

                re.IGNORECASE

            )

            if match:

                tablero_actual[

                    "imagenes"

                ].append(

                    match.group(

                        1

                    ).upper()

                )

    return tableros

def mapear_imagenes(ruta):

    mapa={}

    for root,_,files in os.walk(ruta):

        for file in files:

            if file.upper().startswith(
                "IMG"
            ):

                mapa[
                    file.upper()
                ]=os.path.join(
                    root,
                    file
                )

    return mapa

def dividir_antes_despues(
    pendientes
):

    if len(
        pendientes
    )<=1:

        return pendientes,[]


    numeros=[]

    for ruta in pendientes:

        nombre=Path(
            ruta
        ).name

        match=re.search(
            r"WA(\d+)",
            nombre
        )

        if match:

            numeros.append(
                (
                    int(
                        match.group(1)
                    ),
                    ruta
                )
            )

    numeros.sort(
        key=lambda x:x[0]
    )

    mayor=0
    corte=0

    for i in range(
        len(numeros)-1
    ):

        salto=(
            numeros[i+1][0]
            -
            numeros[i][0]
        )

        if salto>mayor:

            mayor=salto

            corte=i+1

    if mayor<3:

        mitad=(
            len(numeros)+1
        )//2

        corte=mitad

    antes=[
        x[1]
        for x
        in numeros[:corte]
    ]

    despues=[
        x[1]
        for x
        in numeros[corte:]
    ]

    return antes,despues

def guardar_json(
    tableros,
    path="tableros.json"
):

    with open(
        path,
        "w",
        encoding="utf-8"
    ) as archivo:

        json.dump(

            tableros,

            archivo,

            indent=4,

            ensure_ascii=False

        )

    print(
        f"JSON guardado en: {path}"
    )

def clasificar_yolo(
    tableros,
    carpeta_imagenes,
    modelo="best.pt",
    log_placeholder=None
):

    import os

    model = YOLO(
        modelo
    )

    mapa = mapear_imagenes(
        carpeta_imagenes
    )

    logs = []

    for tablero in tableros:

        pendientes = []

        logs.append(
            f"\n===== {tablero['tablero']} ====="
        )

        if log_placeholder:

            log_placeholder.code(
                "\n".join(
                    logs[-20:]
                ),
                language=None
            )

        for img in tablero[
            "imagenes"
        ]:

            if img not in mapa:

                continue

            ruta = mapa[
                img
            ]

            results = model(

                ruta,

                conf=0.3,

                verbose=False

            )

            clases = set()

            for r in results:

                if r.boxes is None:

                    continue

                for box in r.boxes:

                    idx = int(
                        box.cls[0]
                    )

                    clases.add(
                        model.names[idx]
                    )

            # =====================
            # MOSTRAR DETECCIONES
            # =====================

            nombre_img = os.path.basename(
                ruta
            )

            texto = (

                f"{nombre_img}"

                f" -> "

                f"{list(clases)}"

            )

            logs.append(
                texto
            )

            if log_placeholder:

                log_placeholder.code(

                    "\n".join(
                        logs[-20:]
                    ),

                    language=None

                )

            # =====================
            # CLASIFICACION
            # =====================

            if (

                "ficha_tecnica"

                in clases

            ):

                tablero[
                    "descartadas"
                ].append(
                    ruta
                )

            elif (

                "mano"

                in clases

                or

                "herramienta"

                in clases

            ):

                tablero[
                    "durante"
                ].append(
                    ruta
                )

            else:

                pendientes.append(
                    ruta
                )

        # =====================
        # ANTES / DESPUES
        # =====================
        antes, despues = dividir_antes_despues(
            pendientes
        )

        # ---------------------
        # GARANTIZAR >=2
        # ---------------------

        if len(antes) < 2:

            faltan = 2 - len(antes)

            extras = [

                img

                for img in despues

                if img not in antes

            ][:faltan]

            antes.extend(
                extras
            )

        if len(despues) < 2:

            faltan = 2 - len(despues)

            extras = [

                img

                for img in antes

                if img not in despues

            ][:faltan]

            despues.extend(
                extras
            )

        tablero[
            "antes"
        ] = antes

        tablero[
            "despues"
        ] = despues

        tablero.pop(
            "imagenes",
            None
        )

    return tableros