import streamlit as st
from datetime import datetime, timedelta
from database import init_db, registrar_cita, obtener_citas

init_db()

st.title("📅 Agenda de Citas - Barbería Carlos")
st.markdown("Horario: **Lunes a Sábado, 8:00am - 7:00pm**")

dias_semana = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado']

# Calendario de selección
fecha = st.date_input("Selecciona una fecha:", min_value=datetime.today())

# Mostrar espacios disponibles
hora_inicio = datetime.strptime("08:00", "%H:%M")
hora_fin = datetime.strptime("19:00", "%H:%M")
intervalo = timedelta(minutes=30)

horas_disponibles = []
actual = hora_inicio
while actual < hora_fin:
    horas_disponibles.append(actual.strftime("%H:%M"))
    actual += intervalo

# Mostrar horario
hora = st.selectbox("Selecciona una hora:", horas_disponibles)
nombre = st.text_input("¿Cuál es tu nombre completo?")

if st.button("Agendar cita"):
    registrar_cita(nombre, fecha.strftime("%Y-%m-%d"), hora)
    st.success(f"Cita agendada para {nombre} el {fecha.strftime('%d-%m-%Y')} a las {hora}")

st.subheader("👀 Citas Agendadas")
citas = obtener_citas()
for c in citas:
    st.write(f"📌 {c[0]} - {c[1]} - {c[2]}")
