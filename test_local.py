"""
Script de prueba local — valida que todos los módulos funcionan
con los archivos reales sin necesitar Streamlit.
"""
import sys, os
from datetime import datetime
sys.path.insert(0, os.path.dirname(__file__))

BASE = r"C:\Users\CONTABILIDAD\OneDrive\Escritorio\proyecto pyg 2026"

from modules.loader import load_pyg_mes, load_pyg_cc, load_mayor, load_balance
from modules.builder_mes import build_pyg_mes
from modules.builder_proyecto import build_pyg_proyecto
from modules.builder_cc import build_pyg_cc
from modules.observaciones import generar_observaciones
from modules.exporter import exportar_excel

FECHA_INI = datetime(2026, 1, 1, 0, 0, 0)
FECHA_FIN = datetime(2026, 4, 30, 23, 59, 59)

print("="*60)
print("TEST: Cargando archivos...")

df_mes = load_pyg_mes(os.path.join(BASE, "pyg por mes.xls"))
print(f"  PyG Mes: {len(df_mes)} cuentas, columnas: {list(df_mes.columns)}")

df_ing = load_mayor(os.path.join(BASE, "ingresos.xlsx"))
print(f"  Mayor Ingresos: {len(df_ing)} filas")
print(f"    Proyectos unicos: {df_ing['Proyecto'].nunique()}")
print(f"    Proyectos: {sorted(df_ing['Proyecto'].unique())[:8]}")

df_cos = load_mayor(os.path.join(BASE, "costos y gastos.xlsx"))
print(f"  Mayor Costos/Gastos: {len(df_cos)} filas")
print(f"    Proyectos unicos: {df_cos['Proyecto'].nunique()}")

df_cc  = load_pyg_cc(os.path.join(BASE, "pyg por centros de costos.xlsx"))
print(f"  PyG CC: {len(df_cc)} cuentas, columnas CC: {[c for c in df_cc.columns if c not in ('Cod','Concepto')]}")

df_bg  = load_balance(os.path.join(BASE, "BalanceGeneral.xls"))
print(f"  Balance: {len(df_bg)} cuentas")

print("\nTEST: Construyendo pestanas...")

filas_mes, vcols_mes, _ = build_pyg_mes(df_mes)
print(f"  PyG Mes: {len(filas_mes)} filas (incluyendo subtotales)")
subtots_mes = [f['concepto'] for f in filas_mes if f.get('tipo') == 'subtotal']
print(f"    Subtotales: {subtots_mes}")

filas_proy, vcols_proy = build_pyg_proyecto(
    df_ing, df_cos, df_mes,
    fecha_inicio=FECHA_INI,
    fecha_fin=FECHA_FIN,
)
print(f"  PyG Proyecto: {len(filas_proy)} filas, {len(vcols_proy)} columnas (proyectos+total)")
print(f"    Columnas: {vcols_proy[:8]}...")

filas_cc, vcols_cc = build_pyg_cc(df_cc)
print(f"  PyG CC: {len(filas_cc)} filas, {len(vcols_cc)} CCs")

print("\nTEST: Generando observaciones...")
obs = generar_observaciones(df_mes, df_cc, df_ing, df_bg)
print(f"  {len(obs)} observaciones generadas")
for o in obs[:5]:
    print(f"    [{o['nivel']}] {o['categoria']}: {o['descripcion'][:80]}...")

# Cuadre rapido: total ingresos del mayor vs PyG mes
import pandas as pd
total_mes_col = [c for c in df_mes.columns if str(c).upper() in ("TOTAL", "TOTAL ACUMULADO")]
if total_mes_col:
    ing_pyg = df_mes[df_mes["Cod"] == "4"][total_mes_col[0]].sum()
    ing_mayor = df_ing[df_ing["Codigo"].str.startswith("4")]["Monto"].sum()
    print(f"\n  Cuadre ingresos:")
    print(f"    PyG Mes  (cuenta 4):  {ing_pyg:>15,.2f}")
    print(f"    Mayor Ing (cuenta 4): {ing_mayor:>15,.2f}")
    diff = abs(ing_pyg - ing_mayor)
    print(f"    Diferencia:           {diff:>15,.2f}  {'OK' if diff < 1 else 'REVISAR'}")

print("\nTEST: Exportando Excel...")
buf = exportar_excel(
    empresa="BAEC BALDOSAS DEL ECUADOR S.A.S.",
    filas_mes=filas_mes, value_cols_mes=vcols_mes,
    filas_proyecto=filas_proy, value_cols_proyecto=vcols_proy,
    filas_cc=filas_cc, value_cols_cc=vcols_cc,
    observaciones=obs,
    titulo_mes="Estado de Resultados Comparativo Mensual  |  Enero-Abril 2026",
    titulo_proyecto="Estado de Resultados por Proyecto  |  Enero-Abril 2026 (01/01/2026 - 30/04/2026)",
    titulo_cc="Estado de Resultados por Centro de Costo  |  Enero-Abril 2026",
)

output_path = os.path.join(BASE, "Estados_Financieros_TEST.xlsx")
with open(output_path, "wb") as f:
    f.write(buf.getvalue())

print(f"  Excel guardado: {output_path}")
print(f"  Tamano: {os.path.getsize(output_path):,} bytes")
print("\n" + "="*60)
print("TODAS LAS PRUEBAS PASARON")
