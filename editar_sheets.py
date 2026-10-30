from google.oauth2.service_account import Credentials
import gspread

def conectar_sheets():
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        "credenciales.json",
        scopes=scope
    )

    client = gspread.authorize(creds)
    return client

def subir_a_sheets(
    tableros,
    datos
):

    client = conectar_sheets()

    sh = client.open_by_key(
        "17J7v7-HWkjduQOH44RKujS-ZyFWQKes0zmWERuzp0sk"
    )

    nombre_sede = str(

        datos["sede"]["SEDE"]

    ).strip()

    hojas = {

        ws.title.upper(): ws

        for ws in sh.worksheets()

    }

    encabezados = [

        "ITEM",

        "DENOMINACIÓN",

        "NOMBRE_TABLERO",

        "TIPO",

        "CAPACIDAD",

        "NIVEL",

        "L1-L2",

        "L2-L3",

        "L1-L3",

        "L1-N",

        "L2-N",

        "L3-N",

        "L1-T",

        "L2-T",

        "L3-T",

        "IG",

        "Is",

        "Ir",

        "It",

        "Observaciones",

        "Tcomp",

        "Tamb",

        "difT"

    ]

    # =======================
    # BUSCAR HOJA EXISTENTE
    # =======================

    if nombre_sede.upper() in hojas:

        ws = hojas[
            nombre_sede.upper()
        ]

    else:

        ws = sh.add_worksheet(

            title=nombre_sede[:100],

            rows=1000,

            cols=40

        )

        ws.update(

            "A1",

            [encabezados]

        )

    # =======================
    # CREAR DATOS A:C
    # =======================

    data=[]

    for i,t in enumerate(

        tableros,

        start=1

    ):

        data.append([

            i,

            "Tablero Electrico",

            t["tablero"]

        ])

    # =======================
    # SOLO BORRAR A:C
    # =======================

    ws.batch_clear([

        "A2:C1000"

    ])

    # =======================
    # ESCRIBIR A:C
    # =======================

    if data:

        ws.update(

            "A2",

            data

        )

    return True

def leer_sedes():

    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = Credentials.from_service_account_file(
        "credenciales.json",
        scopes=scope
    )

    client = gspread.authorize(creds)

    sh = client.open_by_key("17J7v7-HWkjduQOH44RKujS-ZyFWQKes0zmWERuzp0sk")

    ws = sh.worksheet("SEDES")

    data = ws.get_all_records(
    numericise_ignore=['all']
    )

    return data

def leer_tableros_sheets():

    client = conectar_sheets()

    sh = client.open_by_key(
        "17J7v7-HWkjduQOH44RKujS-ZyFWQKes0zmWERuzp0sk"
    )

    data = []

    for ws in sh.worksheets():

        if ws.title.upper() == "SEDES":
            continue

    registros = ws.get_all_records(
        numericise_ignore=['all']
    )

    data.extend(
        registros
    )

    return data

def enriquecer_tableros_desde_sheets(tableros):

    data_sheets = leer_tableros_sheets()

    # =========================
    # NOMBRES EXISTENTES EN SHEET
    # =========================

    nombres_sheet = set()

    for fila in data_sheets:

        nombre = str(

            fila.get(
                "NOMBRE_TABLERO",
                ""
            )

        ).strip().upper()

        nombres_sheet.add(
            nombre
        )

    # =========================
    # FILTRAR + ENRIQUECER
    # =========================

    tableros_finales = []

    for t in tableros:

        nombre = str(
            t["tablero"]
        ).strip().upper()

        # si no existe en sheet -> eliminar
        if nombre not in nombres_sheet:

            continue

        # fusionar datos
        for fila in data_sheets:

            nombre_sheet = str(

                fila.get(
                    "NOMBRE_TABLERO",
                    ""
                )

            ).strip().upper()

            if nombre == nombre_sheet:

                t.update(
                    fila
                )

                break

        tableros_finales.append(
            t
        )

    print(
        "Original:",
        len(tableros)
    )

    print(
        "Final:",
        len(tableros_finales)
    )

    return tableros_finales