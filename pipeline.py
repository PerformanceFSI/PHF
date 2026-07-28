import pandas as pd
import config as cfg
from jinja2 import Environment, FileSystemLoader
import base64
from pathlib import Path
import subprocess

subprocess.run(["git", "add", "."])
subprocess.run(["git", "commit", "-m", "Actualizacion diaria"])
subprocess.run(["git", "push", "origin", "main"])
# =========================
# 📂 OUTPUT
# =========================
cfg.OUTPUT_PATH.mkdir(parents=True, exist_ok=True)

# =========================
# 📥 LEER EXCEL
# =========================
print("📥 Leyendo Excel...")

df_plot = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_KPI)
df_pesos = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_PESOS)
df_chart = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_CHART)
df_avance = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_AVANCE)

df_familias = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_FAMILIAS)
df_familias_d1 = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_FAMILIASDIA)
df_dias = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_DIASTRABAJADOS)
df_ytd = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_YTD)
df_aperturado = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_APERTURADO)
df_fam_aperturado = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_FAM_APERTURADO)
df_ceco_depto = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_TGT_XCECO)
df_jnt = pd.read_excel(cfg.EXCEL_FILE, sheet_name=cfg.SHEET_JNT)

# =========================
# 🧹 LIMPIAR COLUMNAS
# =========================
df_plot.columns = df_plot.columns.str.strip()
df_pesos.columns = df_pesos.columns.str.strip()
df_chart.columns = df_chart.columns.str.strip()
df_avance.columns = df_avance.columns.str.strip()
df_familias.columns = df_familias.columns.str.strip()
df_familias_d1.columns = df_familias_d1.columns.str.strip()
df_dias.columns = df_dias.columns.str.strip()
df_ytd.columns = df_ytd.columns.str.strip()
df_aperturado.columns = df_aperturado.columns.str.strip()
df_fam_aperturado.columns = df_fam_aperturado.columns.str.strip()
df_ceco_depto.columns = df_ceco_depto.columns.str.strip()
df_jnt.columns = df_jnt.columns.str.strip()

# =========================
# 🔗 MAPEO CECO / DTO
# =========================

df_ceco_depto = df_ceco_depto.rename(columns={
    "Ceco": "CECO",
    "CECO": "CECO",
    "Dto": "DTO",
    "DTO": "DTO"
})

df_ceco_depto["CECO"] = df_ceco_depto["CECO"].astype(str).str.strip()
df_ceco_depto["DTO"] = df_ceco_depto["DTO"].astype(str).str.strip()
df_ceco_depto["UET"] = df_ceco_depto["UET"].astype(str).str.strip()
# =========================
# 📌 COLUMNAS DE DEPARTAMENTOS
# =========================
dept_cols = [
    cfg.COL_DIRECCION,
    cfg.COL_EMBUTICION,
    cfg.COL_SOLDADURA,
    cfg.COL_PINTURA,
    cfg.COL_PARAGOLPES,
    cfg.COL_MONTAJE,
    cfg.COL_DLI,
    cfg.COL_MANTENIMIENTO,
    cfg.COL_CALIDAD
]

# =========================
# 🔢 CONVERTIR NUMÉRICOS - KPI
# =========================
df_plot[cfg.COL_D1] = pd.to_numeric(df_plot[cfg.COL_D1], errors="coerce").fillna(0)
df_plot[cfg.COL_MTD] = pd.to_numeric(df_plot[cfg.COL_MTD], errors="coerce").fillna(0)
df_plot[cfg.COL_YTD] = pd.to_numeric(df_plot[cfg.COL_YTD], errors="coerce").fillna(0)
df_plot[cfg.COL_TGT] = pd.to_numeric(df_plot[cfg.COL_TGT], errors="coerce").fillna(0)
df_plot[cfg.COL_GAP] = pd.to_numeric(df_plot[cfg.COL_GAP], errors="coerce").fillna(0)

# =========================
# 🔢 CONVERTIR NUMÉRICOS - PESOS
# =========================

# Mantengo los importes numéricos, agregando BD K$
for col in [
    cfg.COL_REAL,
    cfg.COL_TARGET,
    cfg.COL_BD_MES
]:
    if col in df_pesos.columns:
        df_pesos[col] = pd.to_numeric(df_pesos[col], errors="coerce").fillna(0)

# IMPORTANTE:
# Dejo los porcentajes como tu pipeline original.
# Si Excel guarda 4,84% como 0.0484, esto lo convierte a 4.84.
df_pesos[cfg.COL_ESTADO_TGT] = pd.to_numeric(
    df_pesos[cfg.COL_ESTADO_TGT], errors="coerce"
).fillna(0) * 100

df_pesos[cfg.COL_ESTADO_BD] = pd.to_numeric(
    df_pesos[cfg.COL_ESTADO_BD], errors="coerce"
).fillna(0) * 100

# =========================
# 🔢 AVANCE
# =========================
avance_real = pd.to_numeric(
    df_avance[cfg.COL_AVANCE], errors="coerce"
).fillna(0).iloc[0]

# =========================
# 📊 KPI TOTAL DESDE PLOT EU_VH
# =========================
fila_total = df_plot[
    df_plot[cfg.COL_DEPTO].astype(str).str.strip() == "TOTAL VT"
]

if fila_total.empty:
    raise ValueError("No se encontró la fila 'TOTAL VT' en la hoja PLOT EU_VH")

row_total = fila_total.iloc[0]

kpi_total = {
    "D1": float(row_total[cfg.COL_D1]),
    "MTD": float(row_total[cfg.COL_MTD]),
    "YTD": float(row_total[cfg.COL_YTD]),
    "TGT": float(row_total[cfg.COL_TGT]),
}

# GAP = TARGET - REAL
gap_kpi = {
    "D1": float(row_total[cfg.COL_TGT]) - float(row_total[cfg.COL_D1]),
    "MTD": float(row_total[cfg.COL_TGT]) - float(row_total[cfg.COL_MTD]),
    "YTD": float(row_total[cfg.COL_TGT]) - float(row_total[cfg.COL_YTD]),
}

# =========================
# 📈 CHART MENSUAL - CIERRE MES
# =========================

# USINA / TOTAL
df_chart[cfg.COL_EUR_VH] = pd.to_numeric(
    df_chart[cfg.COL_EUR_VH], errors="coerce"
).fillna(0)

# Target
if cfg.COL_TGT in df_chart.columns:
    df_chart[cfg.COL_TGT] = pd.to_numeric(
        df_chart[cfg.COL_TGT], errors="coerce"
    ).fillna(cfg.TARGET_LINE)

# Departamentos
for col in dept_cols:
    if col in df_chart.columns:
        df_chart[col] = pd.to_numeric(df_chart[col], errors="coerce").fillna(0)

chart = {
    "mes": df_chart[cfg.COL_MES].astype(str).tolist(),
    "values": df_chart[cfg.COL_EUR_VH].astype(float).tolist(),
    "colors": [
        cfg.COLOR_BAD if v > cfg.TARGET_LINE else cfg.COLOR_OK
        for v in df_chart[cfg.COL_EUR_VH]
    ]
}

# Tabla completa para usar con filtro en JS
tabla_chart = df_chart.to_dict(orient="records")
# =========================
# 📅 ESTADO D-1 FABRICACION / JNT
# =========================
if cfg.COL_D1_FB_JNT in df_jnt.columns:
    d1_estado = str(df_jnt[cfg.COL_D1_FB_JNT].dropna().iloc[0]).strip().upper()
else:
    d1_estado = "FABRICACION"
# =========================
# 📊 FAMILIAS MES
# =========================
for col in dept_cols + ["TOTAL"]:
    if col in df_familias.columns:
        df_familias[col] = pd.to_numeric(
            df_familias[col], errors="coerce"
        ).fillna(0)

df_top_familias = df_familias.sort_values(
    by="TOTAL", ascending=False
).head(10)

top_familias = df_top_familias.to_dict(orient="records")

# =========================
# 📊 FAMILIAS D-1
# =========================
for col in dept_cols + ["TOTAL"]:
    if col in df_familias_d1.columns:
        df_familias_d1[col] = pd.to_numeric(
            df_familias_d1[col], errors="coerce"
        ).fillna(0)

df_top_familias_d1 = df_familias_d1.sort_values(
    by="TOTAL", ascending=False
).head(10)

top_familias_d1 = df_top_familias_d1.to_dict(orient="records")

# =========================
# 📊 DIAS TRABAJADOS
# =========================

# Convertir FECHA para que no rompa JSON
if "FECHA" in df_dias.columns:
    df_dias["FECHA"] = pd.to_datetime(
        df_dias["FECHA"], errors="coerce"
    ).dt.strftime("%Y-%m-%d")

# USINA
if cfg.COL_USINAXDIA in df_dias.columns:
    df_dias[cfg.COL_USINAXDIA] = pd.to_numeric(
        df_dias[cfg.COL_USINAXDIA], errors="coerce"
    ).fillna(0)

# Departamentos
for col in dept_cols:
    if col in df_dias.columns:
        df_dias[col] = pd.to_numeric(df_dias[col], errors="coerce").fillna(0)

dias_labels = pd.to_numeric(
    df_dias[cfg.COL_DIA], errors="coerce"
).fillna(0).astype(int).tolist()

valores_usina = df_dias[cfg.COL_USINAXDIA].tolist()

colores_dias = [
    "#1f7a3f" if str(v).lower() != "jnt" else "#3b82f6"
    for v in df_dias[cfg.COL_DIATRABAJADO]
]

chart_dias = {
    "dias": dias_labels,
    "values": valores_usina,
    "colors": colores_dias
}

# IMPORTANTE: no mandar FECHA al HTML
tabla_dias = df_dias.drop(
    columns=["FECHA"], errors="ignore"
).to_dict(orient="records")

# =========================
# 📈 YTD
# =========================
if cfg.COL_YTD in df_ytd.columns:
    df_ytd[cfg.COL_YTD] = pd.to_numeric(
        df_ytd[cfg.COL_YTD], errors="coerce"
    ).fillna(0)

for col in dept_cols:
    if col in df_ytd.columns:
        df_ytd[col] = pd.to_numeric(df_ytd[col], errors="coerce").fillna(0)

chart_ytd = {
    "mes": df_ytd[cfg.COL_MES].astype(str).tolist(),
    "values": df_ytd[cfg.COL_YTD].astype(float).tolist()
}

tabla_ytd = df_ytd.to_dict(orient="records")
# =========================
# 🖼️ LOGO EN BASE64
# =========================
logo_path = cfg.TEMPLATE_PATH / "logo.png"

if logo_path.exists():
    with open(logo_path, "rb") as img_file:
        logo_base64 = base64.b64encode(img_file.read()).decode("utf-8")

    logo_src = f"data:image/png;base64,{logo_base64}"
else:
    logo_src = ""
    print("⚠️ No se encontró logo.png en:", logo_path)
# =========================
# APERTURADO
# =========================

def limpiar_numero(valor):
    if pd.isna(valor):
        return 0

    if isinstance(valor, (int, float)):
        return float(valor)

    s = str(valor).strip()
    s = s.replace("€", "").replace("$", "").replace("%", "").strip()

    if s in ["", "-", "–"]:
        return 0

    if "," in s:
        s = s.replace(".", "").replace(",", ".")

    try:
        return float(s)
    except Exception:
        return 0


cols_aperturado_num = [
    cfg.COL_AP_D1,
    cfg.COL_AP_MTD,
    cfg.COL_AP_YTD,
    cfg.COL_AP_TGT_EUVH,

    cfg.COL_AP_REAL_MTD,
    cfg.COL_AP_REAL_K,
    cfg.COL_AP_TARGET_REAL,
    cfg.COL_AP_TARGET_K,
    cfg.COL_AP_BD_MES,

    cfg.COL_AP_VS_BD,
    cfg.COL_AP_VS_TGT,
    cfg.COL_AP_MTD_EU
]

for col in cols_aperturado_num:
    if col in df_aperturado.columns:
        df_aperturado[col] = df_aperturado[col].apply(limpiar_numero)

# =========================
# 📋 TABLAS PARA JINJA / JS
# =========================
tabla_kpi = df_plot.to_dict(orient="records")
tabla_plot = df_plot.to_dict(orient="records")
tabla_familias = df_familias.to_dict(orient="records")
tabla_familias_d1 = df_familias_d1.to_dict(orient="records")
tabla_aperturado = df_aperturado.to_dict(orient="records")
tabla_fam_aperturado = df_fam_aperturado.to_dict(orient="records")
tabla_ceco_depto = df_ceco_depto.to_dict(orient="records")
# =========================
# 🌐 JINJA
# =========================
env = Environment(loader=FileSystemLoader(cfg.TEMPLATE_PATH))

def eur(value):
    try:
        return f"€{float(value):,.2f}"
    except Exception:
        return "€0.00"

def ars(value):
    try:
        return f"${float(value):,.2f}"
    except Exception:
        return "$0.00"

env.filters["eur"] = eur
env.filters["ars"] = ars

template = env.get_template("PHF.html")
target_line = float(cfg.TARGET_LINE)

# =========================
# 🧩 RENDER
# =========================
html = template.render(
    tabla_kpi=tabla_kpi,
    tabla_pesos=df_pesos.to_dict(orient="records"),
    chart=chart,
    kpi_total=kpi_total,
    avance_real=avance_real,
    gap_kpi=gap_kpi,
    target_line=target_line,
    top_familias=top_familias,
    top_familias_d1=top_familias_d1,
    chart_dias=chart_dias,
    chart_ytd=chart_ytd,
    tabla_plot=tabla_plot,
    tabla_chart=tabla_chart,
    tabla_dias=tabla_dias,
    tabla_ytd=tabla_ytd,
    tabla_familias=tabla_familias,
    tabla_familias_d1=tabla_familias_d1,
    tabla_aperturado=tabla_aperturado,
tabla_fam_aperturado=tabla_fam_aperturado,
d1_estado=d1_estado,
logo_src=logo_src,
tabla_ceco_depto=tabla_ceco_depto

)

# =========================
# 💾 OUTPUT
# =========================
output_file = cfg.OUTPUT_PATH / "dashboardPHF.html"

with open(output_file, "w", encoding="utf-8") as f:
    f.write(html)

print("✅ Dashboard generado:", output_file)
