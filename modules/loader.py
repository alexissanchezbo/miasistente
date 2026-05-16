import pandas as pd
import io


def _to_buffer(file_input):
    if hasattr(file_input, "read"):
        return io.BytesIO(file_input.read())
    return file_input


def load_pyg_mes(file_input):
    """
    Carga el archivo PyG por Mes.
    Retorna DataFrame con columnas: Cod, Concepto, + columnas de meses + Total
    """
    buf = _to_buffer(file_input)
    raw = pd.read_excel(buf, header=None, engine="xlrd")

    # Encuentra la fila de encabezados (contiene "enero" o el nombre del primer mes)
    header_row = None
    for i, row in raw.iterrows():
        vals = [str(v).lower() for v in row if pd.notna(v)]
        if any("enero" in v or "enero" in v for v in vals):
            header_row = i
            break

    if header_row is None:
        # Fallback: buscar la fila que tiene más de 4 valores no nulos
        for i, row in raw.iterrows():
            non_null = row.dropna()
            if len(non_null) >= 4 and i > 2:
                header_row = i
                break

    headers = raw.iloc[header_row].tolist()
    # Limpia encabezados vacíos
    clean_headers = []
    for h in headers:
        if pd.isna(h) or str(h).strip() == "":
            clean_headers.append(f"_col_{len(clean_headers)}")
        else:
            clean_headers.append(str(h).strip())

    data = raw.iloc[header_row + 1:].copy()
    data.columns = clean_headers
    data = data.reset_index(drop=True)

    # Renombra primeras dos columnas
    cols = list(data.columns)
    if len(cols) >= 2:
        data = data.rename(columns={cols[0]: "Cod", cols[1]: "Concepto"})

    # Filtra filas que tienen código de cuenta (empiezan con 4 o 5)
    def tiene_cuenta(v):
        s = str(v).strip()
        return s and s[0] in ("4", "5") and not s.startswith("nan")

    data = data[data["Cod"].apply(tiene_cuenta)].copy()

    # Convierte columnas numéricas
    for col in data.columns:
        if col not in ("Cod", "Concepto"):
            data[col] = pd.to_numeric(
                data[col].astype(str).str.replace(",", ".").str.replace(" ", ""),
                errors="coerce"
            ).fillna(0.0)

    data["Cod"] = data["Cod"].astype(str).str.strip()
    data["Concepto"] = data["Concepto"].astype(str).str.strip()

    return data


def load_pyg_cc(file_input):
    """
    Carga el archivo PyG por Centro de Costo.
    Retorna DataFrame con Cod, Concepto, + columnas de centros de costo
    """
    buf = _to_buffer(file_input)
    raw = pd.read_excel(buf, header=None, engine="openpyxl")

    # Encabezados en fila 4 (índice 3): nombres de centros de costo desde col 3
    header_row = 3
    headers = raw.iloc[header_row].tolist()

    clean_headers = []
    for h in headers:
        if pd.isna(h) or str(h).strip() in ("", "nan"):
            clean_headers.append(f"_col_{len(clean_headers)}")
        else:
            clean_headers.append(str(h).strip())

    data = raw.iloc[header_row + 1:].copy()
    data.columns = clean_headers
    data = data.reset_index(drop=True)

    cols = list(data.columns)
    data = data.rename(columns={cols[0]: "Cod", cols[1]: "Concepto"})

    def tiene_cuenta(v):
        s = str(v).strip()
        return s and s[0] in ("4", "5") and not s.startswith("nan")

    data = data[data["Cod"].apply(tiene_cuenta)].copy()

    for col in data.columns:
        if col not in ("Cod", "Concepto"):
            data[col] = pd.to_numeric(
                data[col].astype(str).str.replace(",", ".").str.replace(" ", ""),
                errors="coerce"
            ).fillna(0.0)

    data["Cod"] = data["Cod"].astype(str).str.strip()
    data["Concepto"] = data["Concepto"].astype(str).str.strip()

    # Elimina columnas _col_* vacías
    data = data.loc[:, ~data.columns.str.startswith("_col_")]

    return data


def load_mayor(file_input):
    """
    Carga un archivo de Mayor de Cuentas (ingresos o costos y gastos).
    Formato: filas de encabezado por cuenta (tiene Código, sin Fecha)
             seguidas de filas de detalle (tienen Fecha, Código vacío).

    Retorna DataFrame con columnas:
        Fecha, Codigo, Cuenta, CentroCosto, Proyecto, Debe, Haber, Monto
    donde Monto = Haber - Debe para cuentas 4.x (ingresos)
               = Debe - Haber para cuentas 5.x (costos/gastos)
    """
    engine = "openpyxl" if str(getattr(file_input, "name", file_input)).endswith("xlsx") else "xlrd"
    buf = _to_buffer(file_input)
    try:
        raw = pd.read_excel(buf, header=None, engine=engine)
    except Exception:
        buf.seek(0)
        raw = pd.read_excel(buf, header=None, engine="openpyxl")

    rows = []
    current_codigo = ""
    current_cuenta = ""

    for _, row in raw.iterrows():
        raw_fecha  = row.iloc[0]
        raw_codigo = row.iloc[1]
        raw_cuenta = row.iloc[2]
        raw_cc       = row.iloc[3]  if len(row) > 3  else None
        raw_proy     = row.iloc[4]  if len(row) > 4  else None
        raw_tipo_doc = row.iloc[5]  if len(row) > 5  else None
        raw_num_doc  = row.iloc[6]  if len(row) > 6  else None
        raw_tercero  = row.iloc[9]  if len(row) > 9  else None
        raw_debe     = row.iloc[11] if len(row) > 11 else None
        raw_haber    = row.iloc[12] if len(row) > 12 else None

        codigo = str(raw_codigo).strip() if pd.notna(raw_codigo) else ""
        cuenta = str(raw_cuenta).strip() if pd.notna(raw_cuenta) else ""
        fecha_es_vacio = pd.isna(raw_fecha) or str(raw_fecha).strip() in ("", "nan", "Fecha")

        # Fila de encabezado de cuenta: tiene código válido y sin fecha
        if codigo and codigo not in ("nan", "", "Código") and fecha_es_vacio:
            if codigo[0] in ("4", "5"):
                current_codigo = codigo
                current_cuenta = cuenta
            continue

        # Fila de detalle: tiene fecha
        if not fecha_es_vacio and current_codigo:
            try:
                fecha = pd.to_datetime(raw_fecha, dayfirst=True, errors="coerce")
            except Exception:
                fecha = pd.NaT

            _d = pd.to_numeric(raw_debe,  errors="coerce")
            _h = pd.to_numeric(raw_haber, errors="coerce")
            debe  = 0.0 if pd.isna(_d) else float(_d)
            haber = 0.0 if pd.isna(_h) else float(_h)

            cc       = str(raw_cc).strip()       if pd.notna(raw_cc)       else ""
            proy     = str(raw_proy).strip()     if pd.notna(raw_proy)     else ""
            tipo_doc = str(raw_tipo_doc).strip() if pd.notna(raw_tipo_doc) else ""
            num_doc  = str(raw_num_doc).strip()  if pd.notna(raw_num_doc)  else ""
            tercero  = str(raw_tercero).strip()  if pd.notna(raw_tercero)  else ""
            cc       = "" if cc       in ("nan", "Centro de Costo") else cc
            proy     = "" if proy     in ("nan", "Proyecto")        else proy
            tipo_doc = "" if tipo_doc in ("nan", "Tipo")            else tipo_doc
            num_doc  = "" if num_doc  in ("nan", "Número")          else num_doc
            tercero  = "" if tercero  in ("nan", "Tercero", "Nombre") else tercero

            # Monto según naturaleza de la cuenta
            if current_codigo.startswith("4"):
                monto = haber - debe   # ingreso neto
            else:
                monto = debe - haber   # costo/gasto neto

            rows.append({
                "Fecha":       fecha,
                "Codigo":      current_codigo,
                "Cuenta":      current_cuenta,
                "CentroCosto": cc,
                "Proyecto":    proy,
                "TipoDoc":     tipo_doc,
                "NumDoc":      num_doc,
                "Tercero":     tercero,
                "Debe":        debe,
                "Haber":       haber,
                "Monto":       monto,
            })

    df = pd.DataFrame(rows)
    if not df.empty:
        df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    return df


def load_balance(file_input):
    """
    Carga el Balance General (Estado de Situación Financiera).
    Retorna DataFrame con Cod, Concepto, Total
    """
    buf = _to_buffer(file_input)
    raw = pd.read_excel(buf, header=None, engine="xlrd")

    data_rows = []
    for _, row in raw.iterrows():
        cod = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        concepto = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ""
        valor = row.iloc[2] if len(row) > 2 else None

        if not cod or cod == "nan":
            continue
        # Filtra filas de encabezado
        if cod[0] in ("1", "2", "3") and concepto:
            v = pd.to_numeric(
                str(valor).replace(",", ".").replace(" ", "").replace("##########", "nan"),
                errors="coerce"
            ) if pd.notna(valor) else 0.0
            data_rows.append({"Cod": cod, "Concepto": concepto, "Total": v or 0.0})

    return pd.DataFrame(data_rows)
