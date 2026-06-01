import re
from pathlib import Path
from openpyxl import load_workbook
import pandas as pd
from logger import escribir_detalle



# =====================================================
# REGEX
# =====================================================

# ROMAN_RE = re.compile(r"^[IVXLCDM]+\.", re.I)
# NUM_RE = re.compile(r"^\d+\.")
# LETTER_RE = re.compile(r"^[a-z]\.", re.I)
PN_RE = re.compile(r"(PN[A-Z0-9]+)", re.I)
ROMAN_MAYOR_RE = re.compile(r'^[IVXLCDM]+\.',re.ASCII)
NUM_RE = re.compile(r'^\d+\.')
LETTER_RE = re.compile(r'^[a-z]\.')
ROMAN_MINOR_RE = re.compile(r'^(i|ii|iii|iv|v|vi|vii|viii|ix|x)\.',re.I)

# =====================================================
# UTILIDADES
# =====================================================

def limpiar(texto):
    if texto is None:
        return ""

    texto = str(texto).strip()
    texto = re.sub(r"\s+", " ", texto)

    return texto


def es_numero(valor):
    if valor is None:
        return False

    try:
        float(str(valor).replace(",", ""))
        return True
    except:
        return False


def obtener_pn(cell):
    try:
        if cell.hyperlink is None:
            return None

        target = cell.hyperlink.target
        if not target:
            return None

        m = PN_RE.search(target)
        if m:
            return m.group(1)
    except:
        pass

    return None


def tiene_hyperlink(cell):
    try:
        return cell.hyperlink is not None
    except:
        return False


def es_continuacion(ws, fila, col_desc):
    cell = ws.cell(fila, col_desc)
    texto = limpiar(cell.value)

    if not texto:
        return False

    # Si tiene PN, es descripción propia
    if obtener_pn(cell):
        return False

    # Si tiene hyperlink aunque no tenga PN
    if tiene_hyperlink(cell):
        return False

    # Si tiene datos numéricos asociados
    if tiene_datos(ws, fila, col_desc):
        return False

    # Si inicia una nueva jerarquía
    if obtener_nivel(texto) <= 4:
        return False

    return True


def obtener_descripcion_completa(ws, fila, col_desc, filas_consumidas):
    partes = []
    texto = limpiar(ws.cell(fila, col_desc).value)

    partes.append(texto)
    fila_actual = fila + 1

    while fila_actual <= ws.max_row:
        if not es_continuacion(ws, fila_actual, col_desc):
            break

        texto_extra = limpiar(ws.cell(fila_actual, col_desc).value)
        partes.append(texto_extra)
        filas_consumidas.add(fila_actual)
        fila_actual += 1

    return " ".join(partes)
# =====================================================
# JERARQUIA
# =====================================================

# def obtener_nivel(texto):
#     texto = texto.strip()

#     if ROMAN_RE.match(texto):
#         return 1

#     if NUM_RE.match(texto):
#         return 2

#     if LETTER_RE.match(texto):
#         return 3

#     return 4


def obtener_nivel(texto):
    texto = texto.strip()

    # I. II. III. IV.
    if ROMAN_MAYOR_RE.match(texto):
        if texto.startswith(("i.", "ii.", "iii.", "iv.", "v.")):
            pass
        else:
            return 1

    # 1. 2. 3.
    if NUM_RE.match(texto):
        return 2

    # a. b. c.
    if LETTER_RE.match(texto):
        if ROMAN_MINOR_RE.match(texto):
            pass
        else:
            return 3

    # i. ii. iii.
    if ROMAN_MINOR_RE.match(texto):
        return 4

    # textos auxiliares
    return 5

def construir_jerarquia(texto, stack):
    nivel = obtener_nivel(texto)

    while stack and stack[-1]["nivel"] >= nivel:
        stack.pop()

    stack.append({"nivel": nivel,"texto": texto})

    return " > ".join(item["texto"] for item in stack)


# =====================================================
# VALIDACION
# =====================================================

def tiene_datos(ws, fila, col_desc):
    encontrados = 0

    for col in range(col_desc + 1, ws.max_column + 1):
        valor = ws.cell(fila, col).value

        if es_numero(valor):
            encontrados += 1

        if encontrados >= 3:
            return True
        
    return False


# =====================================================
# PROCESAR CUADRO HORIZONTAL
# =====================================================

def procesar_horizontal(ruta_excel):

    wb = load_workbook(
        ruta_excel,
        data_only=False
    )

    ws = wb[wb.sheetnames[0]]

    numero_cuadro = re.search(
        r"(\d+)",
        Path(ruta_excel).stem
    ).group(1).zfill(3)

    resultados = []

    stack = []

    filas_consumidas = set()

    col_desc = 2  # Columna B

    print(f"\nProcesando cuadro {numero_cuadro}")

    for fila in range(1, ws.max_row + 1):

        # Saltar filas que ya fueron
        # utilizadas como continuación
        if fila in filas_consumidas:
            continue

        cell = ws.cell(fila, col_desc)

        texto = limpiar(cell.value)

        if not texto:
            continue

        tiene_serie = tiene_datos(
            ws,
            fila,
            col_desc
        )

        pn = obtener_pn(cell)

        # Si no tiene datos ni PN,
        # probablemente no sea descripción válida
        if not tiene_serie and not pn:
            continue

        # ==================================================
        # UNIR DESCRIPCIONES PARTIDAS
        # ==================================================

        texto_completo = obtener_descripcion_completa(
            ws,
            fila,
            col_desc,
            filas_consumidas
        )

        # ==================================================
        # CONSTRUIR JERARQUÍA
        # ==================================================

        descripcion = construir_jerarquia(
            texto_completo,
            stack
        )

        resultados.append({
            "N_CUADRO": f"cuadro-{numero_cuadro}",
            "CODIGO_SERIE": pn,
            "DESCRIPCION": descripcion
        })

        escribir_detalle(
            f"CUADRO={numero_cuadro} | "
            f"FILA={fila} | "
            f"PN={pn if pn else '-'} | "
            f"DESC={descripcion}"
        )

        print(
            f"Fila={fila:<4} "
            f"PN={pn if pn else '-'} "
            f"{descripcion}"
        )

    print(
        f"\nTotal encontrados: {len(resultados)}"
    )

    return resultados


# =====================================================
# PROCESAR TODOS LOS HORIZONTALES
# =====================================================

def procesar_carpeta():
    carpeta = Path("Nota_cuadros")
    resultados = []

    for archivo in carpeta.glob("*.xlsx"):
        print("\n" + "=" * 80)
        print(f"Procesando: {archivo.name}")
        print("=" * 80)

        try:
            datos = procesar_horizontal(archivo)
            resultados.extend(datos)
        except Exception as e:
            print(
                f"ERROR {archivo.name}: {e}"
            )

    df = pd.DataFrame(resultados)
    df.to_excel(
        "resultado_horizontal.xlsx",
        index=False
    )

    print(
        f"\nTotal registros: {len(df)}"
    )


# =====================================================
# MENU
# =====================================================

if __name__ == "__main__":
    print("\n1. Procesar cuadro-001")
    print("2. Procesar todos los horizontales")

    opcion = input("\nOpción: ").strip()
    if opcion == "1":
        datos = procesar_horizontal(
            "cuadro-001.xlsx"
        )

        df = pd.DataFrame(datos)

        df.to_excel(
            "resultado_prueba.xlsx",
            index=False
        )
    elif opcion == "2":
        procesar_carpeta()