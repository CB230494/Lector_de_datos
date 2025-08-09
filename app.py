import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import math
from datetime import time
import matplotlib.pyplot as plt

st.set_page_config(page_title="Encuesta – Unificación y Dashboard", layout="wide")

st.title("🧮 Encuesta: Unificación de respuestas y Dashboard")
st.write("Sube 1 o varios archivos con una hoja **'Hoja 1'** con columnas como: "
         "`Timestamp, Seguridad, Preocupacion, Descripcion_Delito, Lugares_Evitados, Peticion, Fuerza_Publica`.")

# ------------------ utilidades ------------------
def _normalize_text(s: str) -> str:
    if pd.isna(s): return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def _tokenize(text: str):
    return re.findall(r"[a-z0-9áéíóúñ]+", text)

def _score_match(text: str, keywords):
    """puntaje por # de keywords que aparecen; acepta subfrases."""
    if not text: return 0
    score = 0
    for kw in keywords:
        if re.search(rf"\b{re.escape(kw)}", text):
            score += 1
    return score

def _classify_by_scores(text, rules, threshold=1):
    best_label, best_score = "otro", 0
    for label, kws in rules:
        s = _score_match(text, kws)
        if s > best_score:
            best_label, best_score = label, s
    return best_label if best_score >= threshold else "otro"

def _save_fig(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=160)
    buf.seek(0)
    return buf

def _pct_series(counts):
    total = counts.sum()
    if total == 0:
        return (counts * 0), total
    return (counts / total * 100).round(1), total

# ------------------ taxonomías (ampliadas) ------------------
TAX_PREOCUPACION = [
    ("robos/asaltos", ["robo", "asalt", "hurto", "aran", "sustra", "ladron"]),
    ("homicidios/violencia", ["homicid", "asesin", "herida", "agresion", "violencia", "pelea"]),
    ("drogas", ["droga", "narco", "venta de droga", "microtr", "consumo"]),
    ("ruidos/convivencia", ["ruido", "escandalo", "molestia", "convivencia", "bulla", "fiesta"]),
    ("iluminacion", ["luz", "ilumin", "alumbrado", "farol"]),
    ("tránsito", ["transit", "trafico", "velocidad", "exceso", "accidente", "mezcalcoh", "ebrio"]),
    ("pandillas", ["pandilla", "marero"]),
    ("espacios/lotebaldio", ["lote", "baldio", "parque", "zonaverde", "camino solo", "solar"]),
]

TAX_DELITO = [
    ("robo/asalto", ["robo", "asalt", "arrebato", "hurto"]),
    ("drogas", ["droga", "narco", "microtr", "consumo"]),
    ("homicidio", ["homicid", "asesin"]),
    ("vandalismo", ["vandal", "graffi", "daño", "romp"]),
    ("violencia intrafamiliar", ["intrafamiliar", "domestic", "pareja"]),
    ("estafas/amenazas", ["estafa", "amenaza", "acoso", "hostig"]),
]

TAX_LUGAR = [
    ("paradas/buses", ["parada", "bus", "terminal"]),
    ("parques", ["parque", "plaza"]),
    ("lotes/zonas baldías", ["lote", "baldio", "solar"]),
    ("barrios/residenciales", ["resid", "barrio", "urbaniz", "vecindario"]),
    ("colegios/escuelas", ["colegio", "escuela", "liceo"]),
    ("comercios", ["comerc", "tienda", "super", "pulperia", "centro comercial"]),
    ("calles/avenidas", ["calle", "avenida", "ruta", "carretera", "tramo"]),
    ("nocturno/madrugada", ["noche", "madrug", "oscuro", "tarde noche", "despues de", "pm"]),
]

TAX_PETICION = [
    ("mas patrullaje", ["recorrido", "patrull", "presencia", "mas policia", "presencia policial", "rondas"]),
    ("camaras/tecnologia", ["camara", "cctv", "drone", "dron", "tecnolog", "monitoreo"]),
    ("iluminacion", ["luz", "ilumin", "alumbrado", "farol"]),
    ("organizacion/comunidad", ["comunal", "comunidad", "comite", "red", "vecinal", "organizacion"]),
    ("intervencion social", ["social", "prevencion", "programa", "convivencia", "juvenil"]),
    ("otros", ["limpieza", "basura", "semaforo", "parqueo", "señal"]),
]

# ------------------ carga y limpieza ------------------
@st.cache_data
def cargar_y_unificar(files, score_threshold=1):
    frames = []
    for f in files:
        try:
            xls = pd.ExcelFile(f)
            sheet = "Hoja 1" if "Hoja 1" in xls.sheet_names else xls.sheet_names[0]
            df = pd.read_excel(xls, sheet_name=sheet)
            frames.append(df)
        except Exception as e:
            st.error(f"❌ Error con '{f.name}': {e}")
    if not frames:
        return pd.DataFrame()

    raw = pd.concat(frames, ignore_index=True).drop_duplicates()

    # normalizar texto en campos abiertos
    for col in ["Seguridad","Preocupacion","Descripcion_Delito","Lugares_Evitados","Peticion","Fuerza_Publica"]:
        if col in raw.columns:
            raw[col] = raw[col].astype(str).fillna("").map(_normalize_text)

    # parse de tiempo
    if "Timestamp" in raw.columns:
        try:
            raw["Timestamp"] = pd.to_datetime(raw["Timestamp"], errors="coerce")
        except:
            pass

    out = raw.copy()

    # seguridad → categorías homogéneas
    rep_seg = {
        "muy seguro":"muy seguro","seguro":"seguro",
        "ni seguro ni inseguro":"ni seguro ni inseguro",
        "inseguro":"inseguro","muy inseguro":"muy inseguro",
    }
    out["Seguridad_Descriptor"] = out.get("Seguridad", "").map(lambda x: rep_seg.get(x, x or "sin dato"))

    # clasificadores por puntaje
    out["Preocupacion_Descriptor"] = out.get("Preocupacion", "").map(lambda t: _classify_by_scores(t, TAX_PREOCUPACION, score_threshold) if t else "sin dato")
    out["Delito_Descriptor"]       = out.get("Descripcion_Delito", "").map(lambda t: _classify_by_scores(t, TAX_DELITO, score_threshold) if t else "sin dato")
    out["Lugar_Descriptor"]        = out.get("Lugares_Evitados", "").map(lambda t: _classify_by_scores(t, TAX_LUGAR, score_threshold) if t else "sin dato")
    out["Peticion_Descriptor"]     = out.get("Peticion", "").map(lambda t: _classify_by_scores(t, TAX_PETICION, score_threshold) if t else "sin dato")

    # post-reglas para bajar "otro"
    def _post_rules(row):
        # Si lugar “otro” pero texto tiene pistas
        if row.get("Lugar_Descriptor") == "otro":
            t = row.get("Lugares_Evitados","")
            if "barrio" in t or "resid" in t or "vecind" in t:
                row["Lugar_Descriptor"] = "barrios/residenciales"
            elif "noche" in t or "madrug" in t or "pm" in t or "oscuro" in t:
                row["Lugar_Descriptor"] = "nocturno/madrugada"
        # Si preocupación “otro” pero menciona luz/ruidos, etc.
        if row.get("Preocupacion_Descriptor") == "otro":
            t = row.get("Preocupacion","")
            if "luz" in t or "ilumin" in t or "alumbrado" in t:
                row["Preocupacion_Descriptor"] = "iluminacion"
            elif "ruido" in t or "bulla" in t or "fiesta" in t:
                row["Preocupacion_Descriptor"] = "ruidos/convivencia"
        return row

    out = out.apply(_post_rules, axis=1)

    # hora y franja
    if "Timestamp" in out.columns:
        out["Fecha"] = out["Timestamp"].dt.date
        out["Hora"]  = out["Timestamp"].dt.hour
        out["Franja"] = out["Hora"].map(lambda h: "nocturno" if (h>=18 or h<6) else "diurno")

    return out

# ------------------ UI carga ------------------
files = st.file_uploader("📁 Sube archivos .xlsx / .xlsm", type=["xlsx","xlsm"], accept_multiple_files=True)
score_threshold = st.sidebar.slider("Umbral de coincidencias por descriptor", 1, 3, 1, help="A mayor umbral, más estricto; bajarlo reduce 'otro'.")
chart_type = st.sidebar.selectbox("Tipo de gráfico principal", ["Barras","Pastel"])
order_top_k = st.sidebar.slider("Top categorías a mostrar", 3, 12, 8)

if not files:
    st.info("Sube tus archivos para comenzar.")
    st.stop()

df = cargar_y_unificar(files, score_threshold)
if df.empty:
    st.warning("No se encontraron datos válidos.")
    st.stop()

st.success(f"✅ {len(df)} respuestas tras deduplicar y unificar.")

with st.expander("🔎 Ver tabla unificada"):
    st.dataframe(df, use_container_width=True, height=420)

# exportar limpio
out_buf = io.BytesIO()
with pd.ExcelWriter(out_buf, engine="openpyxl") as w:
    df.to_excel(w, index=False, sheet_name="Unificado")
st.download_button("📥 Descargar datos unificados (Excel)", data=out_buf.getvalue(),
                   file_name="encuesta_unificada.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ------------------ helpers de gráficos ------------------
def grafico_barras(counts, titulo, key):
    counts = counts.sort_values(ascending=False).head(order_top_k)
    pct, total = _pct_series(counts)
    fig, ax = plt.subplots(figsize=(8, 5), constrained_layout=True)
    pct.plot(kind="bar", ax=ax)
    ax.set_ylabel("%")
    ax.set_title(f"{titulo} (n={total})")
    # margen alto para etiquetas
    ax.margins(y=0.15)
    # etiquetas encima
    for i, v in enumerate(pct.values):
        ax.text(i, v + max(2, pct.max()*0.02), f"{v:.1f}%", ha="center", va="bottom", fontsize=10, rotation=0)
    st.pyplot(fig)
    st.download_button(f"🖼️ Descargar '{titulo}' (PNG)", data=_save_fig(fig),
                       file_name=f"{key}.png", mime="image/png")

def grafico_pastel(counts, titulo, key):
    counts = counts.sort_values(ascending=False).head(order_top_k)
    pct, total = _pct_series(counts)
    fig, ax = plt.subplots(figsize=(6, 6), constrained_layout=True)
    ax.pie(pct.values, labels=pct.index, autopct="%1.1f%%", startangle=90)
    ax.set_title(f"{titulo} (n={total})")
    ax.axis("equal")
    st.pyplot(fig)
    st.download_button(f"🖼️ Descargar '{titulo}' (PNG)", data=_save_fig(fig),
                       file_name=f"{key}.png", mime="image/png")

def grafico_general(counts, titulo, key):
    if chart_type == "Barras":
        grafico_barras(counts, titulo, key)
    else:
        grafico_pastel(counts, titulo, key)

# ------------------ dashboard ------------------
col1, col2 = st.columns(2)
with col1:
    if "Preocupacion_Descriptor" in df.columns:
        grafico_general(df["Preocupacion_Descriptor"].value_counts(), "Principales preocupaciones", "preocupaciones")
with col2:
    if "Delito_Descriptor" in df.columns:
        grafico_general(df["Delito_Descriptor"].value_counts(), "Delitos percibidos", "delitos")

st.markdown("---")
col3, col4 = st.columns(2)
with col3:
    if "Lugar_Descriptor" in df.columns:
        grafico_general(df["Lugar_Descriptor"].value_counts(), "Lugares/condiciones evitadas", "lugares")
with col4:
    if "Peticion_Descriptor" in df.columns:
        grafico_general(df["Peticion_Descriptor"].value_counts(), "Peticiones a la autoridad", "peticiones")

st.markdown("---")
# Percepción de seguridad y Fuerza Pública
col5, col6 = st.columns(2)
with col5:
    if "Seguridad_Descriptor" in df.columns:
        grafico_general(df["Seguridad_Descriptor"].value_counts(), "Percepción de seguridad", "seguridad")
with col6:
    if "Fuerza_Publica" in df.columns:
        grafico_general(df["Fuerza_Publica"].replace({"si":"sí"}).value_counts(), "Participación de Fuerza Pública", "fuerza_publica")

st.markdown("---")
# Temporal (línea de tiempo) y franjas/horas
st.header("🕒 Análisis temporal")
if "Timestamp" in df.columns and df["Timestamp"].notna().any():
    # línea por fecha
    serie_fecha = df.groupby("Fecha").size().sort_index()
    fig, ax = plt.subplots(figsize=(9, 4), constrained_layout=True)
    ax.plot(serie_fecha.index, serie_fecha.values, marker="o")
    ax.set_title(f"Evolución de respuestas por fecha (n={serie_fecha.sum()})")
    ax.set_xlabel("Fecha")
    ax.set_ylabel("Respuestas")
    ax.grid(True, alpha=0.3)
    st.pyplot(fig)
    st.download_button("🖼️ Descargar línea temporal (PNG)", data=_save_fig(fig),
                       file_name="linea_temporal.png", mime="image/png")

    # franja horaria
    if "Franja" in df.columns:
        grafico_general(df["Franja"].value_counts(), "Franja horaria (diurno vs nocturno)", "franja_horaria")

    # por hora
    if "Hora" in df.columns:
        counts_hora = df["Hora"].value_counts().sort_index()
        fig, ax = plt.subplots(figsize=(9, 4), constrained_layout=True)
        ax.plot(counts_hora.index, counts_hora.values, marker="o")
        ax.set_xticks(range(0,24,2))
        ax.set_xlabel("Hora del día")
        ax.set_ylabel("Respuestas")
        ax.set_title("Distribución por hora")
        ax.grid(True, alpha=0.3)
        st.pyplot(fig)
        st.download_button("🖼️ Descargar curva por hora (PNG)", data=_save_fig(fig),
                           file_name="por_hora.png", mime="image/png")
else:
    st.info("No hay columna Timestamp válida para el análisis temporal.")

st.caption("💡 Si aún ves mucho 'otro', baja el *Umbral de coincidencias* en la barra lateral o dime frases reales para añadir reglas.")



