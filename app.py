# app.py
import streamlit as st
import pandas as pd
import io
import re
import unicodedata
import matplotlib.pyplot as plt

st.set_page_config(page_title="Encuesta - Consolidador y Dashboard", layout="wide")

st.title("🧮 Encuesta: Unificación de respuestas y Dashboard")
st.write("Sube 1 o varios archivos con una hoja llamada **'Hoja 1'** con las columnas: "
         "`Timestamp, Seguridad, Preocupacion, Descripcion_Delito, Lugares_Evitados, Peticion, Fuerza_Publica`.")

# ============== Utilidades ==============
def _normalize_text(s: str) -> str:
    """minúsculas, sin acentos, espacios colapsados."""
    if pd.isna(s):
        return ""
    s = str(s).strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def _any(text, patterns):
    """¿Alguna keyword/patrón aparece en el texto? (buscar por palabra)."""
    for p in patterns:
        # palabra completa o subfrase relevante
        if re.search(rf"\b{re.escape(p)}\b", text):
            return True
    return False

def _first_match(text, rules):
    """
    Devuelve el primer descriptor cuya lista de keywords haga match.
    'rules' es una lista de tuplas (descriptor, [keywords...]) evaluadas en orden.
    """
    for label, keywords in rules:
        if _any(text, keywords):
            return label
    return "otro"

def _multi_match(text, rules):
    """Versión multi-etiqueta (si quieres asignar más de un descriptor). No se usa por defecto."""
    labels = [label for label, keywords in rules if _any(text, keywords)]
    return labels if labels else ["otro"]

def _save_fig_to_bytes(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=160)
    buf.seek(0)
    return buf

def _pct_series(counts):
    total = counts.sum()
    if total == 0:
        return (counts * 0), total
    return (counts / total * 100).round(1), total

# ============== TAXONOMÍA (EDITABLE) ==============
# Ajusta/expande las palabras clave según tu realidad local.
TAX_PREOCUPACION = [
    ("robos/asaltos", ["robo", "asalt", "hurto"]),
    ("homicidios/violencia", ["homicid", "asesin", "agresion", "violencia"]),
    ("drogas", ["droga", "narco", "venta de droga"]),
    ("ruidos/convivencia", ["ruido", "escandalo", "molestia"]),
    ("iluminacion", ["luz", "ilumin", "alumbrado"]),
    ("tránsito", ["transit", "trafico", "velocidad", "accidente"]),
    ("pandillas", ["pandilla"]),
    ("espacios/lotebaldio", ["lote", "baldio", "parque", "zonaverde"]),
]

TAX_DELITO = [
    ("robo/asalto", ["robo", "asalt"]),
    ("drogas", ["droga", "narco"]),
    ("homicidio", ["homicid", "asesin"]),
    ("vandalismo", ["vandal", "daño"]),
    ("violencia intrafamiliar", ["intrafamiliar", "domestic"]),
    ("otros", ["estafa", "acoso", "amenaza", "hostig"]),
]

TAX_LUGAR = [
    ("paradas/buses", ["parada", "bus"]),
    ("parques", ["parque"]),
    ("lotes/zonas baldías", ["lote", "baldio"]),
    ("barrios/residenciales", ["resid", "barrio", "urbaniz"]),
    ("colegios/escuelas", ["colegio", "escuela"]),
    ("comercios", ["comerc", "tienda", "super"]),
    ("calles/avenidas", ["calle", "avenida", "ruta"]),
    ("nocturno/madrugada", ["noche", "madrug", "6pm", "7pm", "8pm", "9pm", "10pm", "11pm", "12am"]),
]

TAX_PETICION = [
    ("mas patrullaje", ["recorrido", "patrull", "presencia", "mas policia", "presencia policial"]),
    ("camaras/tecnologia", ["camara", "cctv", "drones", "tecnolog"]),
    ("iluminacion", ["luz", "ilumin", "alumbrado"]),
    ("organizacion/comunidad", ["comunal", "comunidad", "comite", "red", "vecinal"]),
    ("intervencion social", ["social", "prevencion", "programa", "convivencia"]),
    ("otros", ["limpieza", "basura", "semaforo", "parqueo"]),
]

# ============== Parser principal ==============
@st.cache_data
def cargar_y_procesar(files):
    frames = []
    for f in files:
        try:
            xls = pd.ExcelFile(f)
            if "Hoja 1" not in xls.sheet_names:
                st.warning(f"⚠️ '{f.name}' no tiene hoja 'Hoja 1'. Se omite.")
                continue
            df = pd.read_excel(xls, sheet_name="Hoja 1")
            frames.append(df)
        except Exception as e:
            st.error(f"❌ Error con '{f.name}': {e}")

    if not frames:
        return pd.DataFrame()

    raw = pd.concat(frames, ignore_index=True)

    # Deduplicación exacta: todas las columnas iguales
    raw = raw.drop_duplicates()

    # Normalización columnas texto
    for col in ["Seguridad","Preocupacion","Descripcion_Delito","Lugares_Evitados","Peticion","Fuerza_Publica"]:
        if col in raw.columns:
            raw[col] = raw[col].astype(str).fillna("").map(_normalize_text)

    # Clasificación a descriptores
    out = raw.copy()

    # Seguridad: ya viene categórica, solo homogeneizamos algunos valores
    rep_seg = {
        "muy seguro": "muy seguro",
        "seguro": "seguro",
        "ni seguro ni inseguro": "ni seguro ni inseguro",
        "inseguro": "inseguro",
        "muy inseguro": "muy inseguro",
    }
    out["Seguridad_Descriptor"] = out["Seguridad"].map(lambda x: rep_seg.get(x, x or "sin dato"))

    out["Preocupacion_Descriptor"] = out["Preocupacion"].map(lambda t: _first_match(t, TAX_PREOCUPACION) if t else "sin dato")
    out["Delito_Descriptor"] = out["Descripcion_Delito"].map(lambda t: _first_match(t, TAX_DELITO) if t else "sin dato")
    out["Lugar_Descriptor"] = out["Lugares_Evitados"].map(lambda t: _first_match(t, TAX_LUGAR) if t else "sin dato")
    out["Peticion_Descriptor"] = out["Peticion"].map(lambda t: _first_match(t, TAX_PETICION) if t else "sin dato")

    return out

# ============== UI: Carga ==============
files = st.file_uploader("📁 Sube archivos .xlsx / .xlsm (mismo formato de columnas)", type=["xlsx", "xlsm"], accept_multiple_files=True)

if not files:
    st.info("Sube tus archivos para comenzar.")
    st.stop()

df = cargar_y_procesar(files)

if df.empty:
    st.warning("No se encontraron datos válidos.")
    st.stop()

st.success(f"✅ {len(df)} respuestas (tras deduplicar).")

# ============== Vista de datos limpios ==============
with st.expander("🔎 Ver tabla limpia (con descriptores)"):
    st.dataframe(df, use_container_width=True, height=400)

# Descargar Excel limpio
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Limpio")
st.download_button("📥 Descargar datos limpios (Excel)", data=output.getvalue(),
                   file_name="encuesta_limpia.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ============== DASHBOARD ==============
st.header("📊 Dashboard")

def grafico_barras_porcentaje(serie_counts, titulo):
    pct, total = _pct_series(serie_counts)
    fig, ax = plt.subplots(figsize=(7, 4))
    pct.sort_values(ascending=False).plot(kind="bar", ax=ax)
    ax.set_ylabel("%")
    ax.set_title(f"{titulo} (n={total})")
    for i, v in enumerate(pct.sort_values(ascending=False).values):
        ax.text(i, v + 1, f"{v:.1f}%", ha="center", va="bottom", fontsize=9)
    fig.tight_layout()
    return fig

def bloque_grafico_y_descarga(counts, titulo, key):
    fig = grafico_barras_porcentaje(counts, titulo)
    st.pyplot(fig)
    png = _save_fig_to_bytes(fig)
    st.download_button(f"🖼️ Descargar '{titulo}' (PNG)", data=png, file_name=f"{key}.png", mime="image/png")

col1, col2 = st.columns(2)

with col1:
    # Seguridad
    if "Seguridad_Descriptor" in df.columns:
        counts = df["Seguridad_Descriptor"].value_counts()
        bloque_grafico_y_descarga(counts, "Percepción de seguridad", "seguridad")

with col2:
    # Fuerza Pública
    if "Fuerza_Publica" in df.columns:
        counts = df["Fuerza_Publica"].replace({
            "si": "sí", "no estoy seguro/a":"no estoy seguro/a"
        }).value_counts()
        bloque_grafico_y_descarga(counts, "Participación de Fuerza Pública", "fuerza_publica")

st.markdown("---")

col3, col4 = st.columns(2)
with col3:
    if "Preocupacion_Descriptor" in df.columns:
        counts = df["Preocupacion_Descriptor"].value_counts()
        bloque_grafico_y_descarga(counts, "Principales preocupaciones", "preocupaciones")

with col4:
    if "Delito_Descriptor" in df.columns:
        counts = df["Delito_Descriptor"].value_counts()
        bloque_grafico_y_descarga(counts, "Delitos percibidos", "delitos")

st.markdown("---")

col5, col6 = st.columns(2)
with col5:
    if "Lugar_Descriptor" in df.columns:
        counts = df["Lugar_Descriptor"].value_counts()
        bloque_grafico_y_descarga(counts, "Lugares/condiciones evitadas", "lugares")

with col6:
    if "Peticion_Descriptor" in df.columns:
        counts = df["Peticion_Descriptor"].value_counts()
        bloque_grafico_y_descarga(counts, "Peticiones a la autoridad", "peticiones")

st.caption("Tip: ajusta los **TAX_*** al inicio para mejorar la unificación de textos.")




