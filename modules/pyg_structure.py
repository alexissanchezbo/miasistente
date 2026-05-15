"""
Define la jerarquía del PyG y las funciones para clasificar cuentas contables.
Basado en la estructura NIIF/NIC para empresa manufacturera.
"""

from collections import defaultdict


# ── Clasificación de cuentas ────────────────────────────────────────────────

def es_ingreso(cod):
    return str(cod).startswith("4")

def es_materiales(cod):
    return str(cod).startswith("5.1.1")

def es_mod(cod):
    return str(cod).startswith("5.1.2")

def es_moi(cod):
    return str(cod).startswith("5.1.3")

def es_cif(cod):
    return str(cod).startswith("5.1.4")

def es_costo_produccion(cod):
    return str(cod).startswith("5.1")

def es_gastos_ventas(cod):
    return str(cod).startswith("5.2.1.1")

def es_gastos_admin(cod):
    return str(cod).startswith("5.2.1.2")

def es_gastos_financieros(cod):
    return str(cod).startswith("5.2.1.3")

def es_gastos_no_op(cod):
    s = str(cod)
    return s.startswith("5.2.2") or s.startswith("5.2.3")

def es_impuestos(cod):
    return str(cod).startswith("5.2.4")


def clasificar(cod):
    """Retorna el bucket al que pertenece la cuenta."""
    if es_ingreso(cod):           return "ingreso"
    if es_materiales(cod):        return "materiales"
    if es_mod(cod):               return "mod"
    if es_moi(cod):               return "moi"
    if es_cif(cod):               return "cif"
    if es_gastos_ventas(cod):     return "gv"
    if es_gastos_admin(cod):      return "ga"
    if es_gastos_financieros(cod):return "gf"
    if es_gastos_no_op(cod):      return "gno"
    if es_impuestos(cod):         return "imp"
    return "otro"


# ── Cálculo de subtotales ────────────────────────────────────────────────────

def es_cuenta_hoja(cod, all_cods):
    """True si la cuenta no tiene hijos en el conjunto (es cuenta de transacción)."""
    prefix = str(cod) + "."
    return not any(c.startswith(prefix) for c in all_cods if c != str(cod))


def filtrar_ceros(filas, value_cols):
    """
    Elimina filas de tipo 'data' donde la suma de valores absolutos es cero.
    Las filas de subtotal nunca se eliminan.
    """
    result = []
    for fila in filas:
        if fila.get("tipo") == "subtotal":
            result.append(fila)
        else:
            total_abs = sum(abs(fila.get("values", {}).get(c, 0) or 0) for c in value_cols)
            if total_abs > 0.005:
                result.append(fila)
    return result


def calcular_subtotales(data_rows, value_cols):
    """
    Recibe lista de dicts con {'cod', 'concepto', 'values': {col: float}}
    Retorna dict de subtotales {nombre: {col: float}}.
    Solo acumula cuentas hoja (sin hijos) para evitar doble conteo.
    """
    all_cods = {str(row.get("cod", "")) for row in data_rows}
    sums = defaultdict(lambda: defaultdict(float))

    for row in data_rows:
        cod = str(row.get("cod", ""))
        if not es_cuenta_hoja(cod, all_cods):
            continue  # cuenta padre/grupo — omitir para no duplicar
        bucket = clasificar(cod)
        for col in value_cols:
            val = row.get("values", {}).get(col, 0.0) or 0.0
            sums[bucket][col] += val

    def s(bucket, col):
        return sums[bucket].get(col, 0.0)

    subtotales = {}
    for col in value_cols:
        ingreso   = s("ingreso", col)
        materiales = s("materiales", col)
        mod       = s("mod", col)
        moi       = s("moi", col)
        cif       = s("cif", col)
        gv        = s("gv", col)
        ga        = s("ga", col)
        gf        = s("gf", col)
        gno       = s("gno", col)
        imp       = s("imp", col)

        ub  = ingreso - materiales
        ubi = ub - mod - moi - cif
        uop = ubi - gv - ga
        uai = uop - gf
        un  = uai - gno - imp

        subtotales.setdefault("UTILIDAD BRUTA",              {})[col] = ub
        subtotales.setdefault("UTILIDAD BRUTA INDUSTRIAL",   {})[col] = ubi
        subtotales.setdefault("UTILIDAD OPERACIONAL",        {})[col] = uop
        subtotales.setdefault("EBITDA",                      {})[col] = uop  # sin D&A separada
        subtotales.setdefault("UTILIDAD ANTES DE IMPUESTOS", {})[col] = uai
        subtotales.setdefault("UTILIDAD NETA",               {})[col] = un

    return subtotales


# ── Inserción de subtotales en la lista de filas ─────────────────────────────

# Define qué sección "dispara" el subtotal cuando termina
TRIGGERS = [
    ("5.1.1", ["UTILIDAD BRUTA"]),
    ("5.1.4", ["UTILIDAD BRUTA INDUSTRIAL"]),
    ("5.2.1.2", ["UTILIDAD OPERACIONAL", "EBITDA"]),
    ("5.2.1.3", ["UTILIDAD ANTES DE IMPUESTOS"]),
]


def insertar_subtotales(data_rows, subtotales, value_cols):
    """
    Toma la lista de data_rows y retorna una nueva lista con las filas de
    subtotales insertadas en las posiciones correctas.
    Cada fila es un dict: {'tipo': 'data'|'subtotal', 'cod', 'concepto', 'values'}
    """
    result = []
    inserted = set()

    for i, row in enumerate(data_rows):
        cod = str(row.get("cod", ""))
        next_cod = str(data_rows[i + 1].get("cod", "")) if i + 1 < len(data_rows) else ""

        result.append({"tipo": "data", **row})

        for prefix, labels in TRIGGERS:
            if cod.startswith(prefix) and not next_cod.startswith(prefix):
                for label in labels:
                    if label not in inserted:
                        result.append({
                            "tipo": "subtotal",
                            "cod": "",
                            "concepto": label,
                            "values": subtotales.get(label, {c: 0.0 for c in value_cols}),
                        })
                        inserted.add(label)

    # UTILIDAD NETA siempre al final
    if "UTILIDAD NETA" not in inserted:
        result.append({
            "tipo": "subtotal",
            "cod": "",
            "concepto": "UTILIDAD NETA",
            "values": subtotales.get("UTILIDAD NETA", {c: 0.0 for c in value_cols}),
        })

    return result


def get_nivel(cod):
    """Retorna el nivel de indentación de una cuenta (0=raíz)."""
    return max(0, str(cod).count("."))
