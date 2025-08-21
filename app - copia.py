# =========================
# 📋 Asistencia – Público (Supabase Storage CSV)
# =========================
import streamlit as st
import pandas as pd
from io import BytesIO
from supabase import create_client
import sys

st.set_page_config(page_title="Asistencia - Público", layout="wide")
st.markdown("## 📋 Lista de asistencia – Seguimiento de líneas de acción (Público)")

# ---------- CONFIG STORAGE (usa tu proyecto) ----------
SUPABASE_URL = "https://fuqenmijstetuwhdulax.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImZ1cWVubWlqc3RldHV3aGR1bGF4Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3NTMyODM4MzksImV4cCI6MjA2ODg1OTgzOX0.9JdF70hcLCVCa0-lCd7yoSFKtO72niZbahM-u2ycAVg"
BUCKET = "asistencia"
OBJECT = "asistencia.csv"

HEADER = ["Nombre","Cédula de Identidad","Institución","Cargo","Teléfono","Género","Sexo","Rango de Edad"]
supabase = create_client(SUPABASE_URL, SUPABASE_KEY)

def _ensure_remote_csv():
    """Si no existe el archivo, lo crea vacío con encabezados."""
    try:
        supabase.storage.from_(BUCKET).download(OBJECT)
    except Exception as e:
        # Si el objeto no existe, lo creamos; si el bucket no existe, mostramos error claro.
        msg = str(e).lower()
        if "not found" in msg or "object not found" in msg:
            empty = pd.DataFrame(columns=HEADER).to_csv(index=False).encode("utf-8")
            try:
                supabase.storage.from_(BUCKET).upload(OBJECT, empty, {"content-type": "text/csv"}, upsert=True)
            except Exception as e2:
                st.error("No se pudo crear el archivo remoto. Verifica que exista el bucket "
                         f"**{BUCKET}** en Supabase Storage.")
                st.stop()
        elif "bucket" in msg:
            st.error(f"No existe el bucket **{BUCKET}** en Supabase Storage. Créalo y vuelve a intentar.")
            st.stop()

def fetch_all_df():
    _ensure_remote_csv()
    data = supabase.storage.from_(BUCKET).download(OBJECT)
    df = pd.read_csv(BytesIO(data), dtype=str).fillna("")
    if df.empty:
        return pd.DataFrame(columns=["Nº"] + HEADER)
    df.insert(0, "Nº", range(1, len(df)+1))
    return df

def append_row(d):
    _ensure_remote_csv()
    data = supabase.storage.from_(BUCKET).download(OBJECT)
    df = pd.read_csv(BytesIO(data), dtype=str) if data else pd.DataFrame(columns=HEADER)
    df = pd.concat([df, pd.DataFrame([d], columns=HEADER)], ignore_index=True)
    supabase.storage.from_(BUCKET).upload(
        OBJECT, df.to_csv(index=False).encode("utf-8"),
        {"content-type": "text/csv"}, upsert=True
    )

# ---------- Formulario ----------
with st.form("form_asistencia_publico", clear_on_submit=True):
    c1, c2, c3 = st.columns([1.2, 1, 1])
    nombre      = c1.text_input("Nombre")
    cedula      = c2.text_input("Cédula de Identidad")
    institucion = c3.text_input("Institución")

    c4, c5 = st.columns([1, 1])
    cargo    = c4.text_input("Cargo")
    telefono = c5.text_input("Teléfono")

    st.markdown("#### ")
    gcol, scol, ecol = st.columns([1.1, 1.5, 1.5])
    genero = gcol.radio("Género", ["F", "M", "LGBTIQ+"], horizontal=True)
    sexo   = scol.radio("Sexo (Hombre, Mujer o Intersex)", ["H", "M", "I"], horizontal=True)
    edad   = ecol.radio("Rango de Edad", ["18 a 35 años", "36 a 64 años", "65 años o más"], horizontal=True)

    submitted = st.form_submit_button("➕ Agregar", use_container_width=True)
    if submitted:
        if not nombre.strip():
            st.warning("Ingresa al menos el nombre.")
        else:
            fila = {
                "Nombre": nombre.strip(),
                "Cédula de Identidad": cedula.strip(),
                "Institución": institucion.strip(),
                "Cargo": cargo.strip(),
                "Teléfono": telefono.strip(),
                "Género": genero,
                "Sexo": sexo,
                "Rango de Edad": edad
            }
            append_row(fila)
            st.success("Registro guardado.")

# ---------- Vista (solo lectura) ----------
st.markdown("### 📥 Registros recibidos")
df_all = fetch_all_df()
if not df_all.empty:
    st.dataframe(df_all[["Nº"] + HEADER], use_container_width=True, hide_index=True)
else:
    st.info("Aún no hay registros guardados.")







