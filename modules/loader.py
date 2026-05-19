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


def load_cartera(file_input):
    """
    Carga el archivo Cartera por Cobrar (Detallado).
    Encabezados en fila 4 (índice 4). Columnas:
        0=Cliente, 1=RazónSocial, 2=TipoDoc, 3=#Documento,
        4=F.Emisión, 5=F.Vencimiento, 6=Vendedor, 7=CentroCosto,
        9=PorVencer, 10=30d, 11=60d, 12=90d, 13=120d, 14=>120d,
        15=Total, 16=Descripción, 17=ValorDoc, 18=Retenciones, 19=Cobros

    Retorna DataFrame de filas de detalle (una por factura) con:
        Cliente, RazonSocial, TipoDoc, NumDoc, FechaEmision, FechaVencimiento,
        Vendedor, CentroCosto, PorVencer, D30, D60, D90, D120, D120p,
        Total, Descripcion, ValorDoc, Retenciones, Cobros, Obra,
        DiasVencido, FechaCorte
    """
    import re as _re

    buf = _to_buffer(file_input)
    try:
        raw = pd.read_excel(buf, header=None, engine="xlrd")
    except Exception:
        buf.seek(0)
        raw = pd.read_excel(buf, header=None, engine="openpyxl")

    # Fecha de corte desde la cabecera del archivo
    fecha_corte = pd.Timestamp.today().normalize()
    for i in range(min(6, len(raw))):
        for val in raw.iloc[i]:
            if pd.notna(val) and "Fecha de Corte" in str(val):
                m = _re.search(r'(\d{2}/\d{2}/\d{4})', str(val))
                if m:
                    fecha_corte = pd.to_datetime(m.group(1), dayfirst=True, errors="coerce")

    # Localizar fila de encabezados
    header_row = 4
    for i, row in raw.iterrows():
        vals = [str(v).strip() for v in row if pd.notna(v)]
        if "Cliente" in vals and "Total" in vals and "Cobros" in vals:
            header_row = i
            break

    rows = []
    for _, row in raw.iloc[header_row + 1:].iterrows():
        if len(row) < 20:
            continue
        # Omitir filas resumen de cliente (sin Razón Social ni # Documento)
        razon   = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
        num_doc = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        if not razon or razon in ("nan",) or not num_doc or num_doc in ("nan",):
            continue

        def _f(idx):
            return float(pd.to_numeric(row.iloc[idx], errors="coerce") or 0)

        cliente   = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        tipo_doc  = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
        f_emi     = pd.to_datetime(row.iloc[4], dayfirst=True, errors="coerce")
        f_venc    = pd.to_datetime(row.iloc[5], dayfirst=True, errors="coerce")
        vendedor  = str(row.iloc[6]).strip() if pd.notna(row.iloc[6]) else ""
        cc        = str(row.iloc[7]).strip() if pd.notna(row.iloc[7]) else ""
        desc      = str(row.iloc[16]).strip() if pd.notna(row.iloc[16]) else ""

        cc       = "" if cc       in ("nan", "None") else cc
        vendedor = "" if vendedor in ("nan", "None") else vendedor

        # Extraer OBRA de la Descripción
        obra = ""
        m = _re.search(r'OBRA:\s*(.+?)(?:\s+RQC\b|\s*$)', desc, _re.IGNORECASE)
        if m:
            obra = m.group(1).strip()

        # Días vencido (desde fecha vencimiento a fecha de corte)
        dias_vencido = int((fecha_corte - f_venc).days) if pd.notna(f_venc) else None

        rows.append({
            "Cliente":          cliente,
            "RazonSocial":      razon,
            "TipoDoc":          tipo_doc,
            "NumDoc":           num_doc,
            "FechaEmision":     f_emi,
            "FechaVencimiento": f_venc,
            "Vendedor":         vendedor,
            "CentroCosto":      cc,
            "PorVencer":        _f(9),
            "D30":              _f(10),
            "D60":              _f(11),
            "D90":              _f(12),
            "D120":             _f(13),
            "D120p":            _f(14),
            "Total":            _f(15),
            "Descripcion":      desc,
            "ValorDoc":         _f(17),
            "Retenciones":      _f(18),
            "Cobros":           _f(19),
            "Obra":             obra,
            "DiasVencido":      dias_vencido,
            "FechaCorte":       fecha_corte,
        })

    return pd.DataFrame(rows)


def load_transacciones(file_input):
    """
    Carga el archivo de Transacciones Detallado (cobros y pagos).
    Encabezados en fila 3 (índice 3). Columnas clave:
        0=Fecha, 2=Persona, 3=Tipo, 16=CentroCosto,
        17=CodComprobante, 18=FechaEmision, 20=Detalle, 21=Valor

    Retorna DataFrame con columnas:
        Fecha, Persona, Tipo, CodComprobante, FechaEmision,
        CentroCosto, Detalle, Valor, Obra, DiasCobranza
    """
    import re as _re

    buf = _to_buffer(file_input)
    try:
        raw = pd.read_excel(buf, header=None, engine="xlrd")
    except Exception:
        buf.seek(0)
        raw = pd.read_excel(buf, header=None, engine="openpyxl")

    # Localizar fila de encabezados (contiene "Fecha" y "Tipo")
    header_row = 3
    for i, row in raw.iterrows():
        vals = [str(v).strip() for v in row if pd.notna(v)]
        if "Fecha" in vals and "Tipo" in vals:
            header_row = i
            break

    rows = []
    for _, row in raw.iloc[header_row + 1:].iterrows():
        if len(row) < 22:
            continue
        tipo = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
        if tipo not in ("Cobro", "Pago"):
            continue

        fecha     = pd.to_datetime(row.iloc[0],  dayfirst=True, errors="coerce")
        persona   = str(row.iloc[2]).strip()  if pd.notna(row.iloc[2])  else ""
        cod_comp  = str(row.iloc[17]).strip() if pd.notna(row.iloc[17]) else ""
        fecha_emi = pd.to_datetime(row.iloc[18], errors="coerce")
        cc        = str(row.iloc[16]).strip() if pd.notna(row.iloc[16]) else ""
        detalle   = str(row.iloc[20]).strip() if pd.notna(row.iloc[20]) else ""
        valor     = float(pd.to_numeric(row.iloc[21], errors="coerce") or 0)

        cc = "" if cc in ("nan", "None", "NaN") else cc

        # Extraer nombre de OBRA del campo Detalle
        obra = ""
        m = _re.search(r'OBRA:\s*(.+?)(?:\s+RQC\b|\s*$)', detalle, _re.IGNORECASE)
        if m:
            obra = m.group(1).strip()

        # Días de cobranza (Cobros con ambas fechas válidas)
        dias = None
        if tipo == "Cobro" and pd.notna(fecha) and pd.notna(fecha_emi):
            dias = int((fecha - fecha_emi).days)

        rows.append({
            "Fecha":          fecha,
            "Persona":        persona,
            "Tipo":           tipo,
            "CodComprobante": cod_comp,
            "FechaEmision":   fecha_emi,
            "CentroCosto":    cc,
            "Detalle":        detalle,
            "Valor":          valor,
            "Obra":           obra,
            "DiasCobranza":   dias,
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


def load_activos(file_input):
    """
    Carga el registro de Activos Fijos exportado de Contifico.
    Encabezados en fila 2 (índice 1). Columnas clave:
        Fecha Compra, Codigo, Estado, Nombre, Categoria, Tipo,
        Ubicacion, Valor Inicial, Valor Depreciado, Valor Actual, Valor Venta

    Retorna DataFrame con esas columnas (solo activos Activos).
    """
    buf = _to_buffer(file_input)
    try:
        df = pd.read_excel(buf, header=1, engine="xlrd")
    except Exception:
        buf.seek(0)
        df = pd.read_excel(buf, header=1, engine="openpyxl")

    df.columns = [str(c).strip() for c in df.columns]

    for col in ["Valor Inicial", "Valor Depreciado", "Valor Actual"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return df
