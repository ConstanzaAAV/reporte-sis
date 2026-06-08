# -*- coding: utf-8 -*-
"""Página web (Streamlit): subir PDF del mes -> descargar Excel con datos y gráficos."""
import streamlit as st
from procesar import procesar_pdf

st.set_page_config(page_title="Reporte SIS → Excel", page_icon="📊", layout="centered")


# ---------------- Contraseña ----------------
def check_password():
    pw_correcta = st.secrets.get("password", None)
    if pw_correcta is None:
        st.error("⚠️ Falta configurar la contraseña (Secrets). Avisa al administrador.")
        return False
    if st.session_state.get("acceso_ok"):
        return True

    def _verificar():
        if st.session_state.get("pw_input", "") == pw_correcta:
            st.session_state["acceso_ok"] = True
            st.session_state["pw_input"] = ""
        else:
            st.session_state["acceso_ok"] = False

    st.title("📊 Reporte SIS → Excel")
    st.text_input("Contraseña", type="password", key="pw_input", on_change=_verificar)
    if st.session_state.get("acceso_ok") is False:
        st.error("Contraseña incorrecta.")
    st.caption("Herramienta interna. Solicita la contraseña a quien la administra.")
    return False


if not check_password():
    st.stop()

# ---------------- App ----------------
st.title("📊 Reporte SIS → Excel")
st.write(
    "Sube el **PDF del mes** (Plan de Intervención · Mejor Niñez) y descarga el "
    "**Excel** con los datos ordenados y los gráficos de análisis "
    "(visitas reales, tipos de intervención y evento, panel por niño)."
)

archivo = st.file_uploader("Arrastra aquí el PDF del mes (o haz clic para buscarlo)", type=["pdf"])

if archivo is not None:
    try:
        with st.spinner("Procesando… puede tardar 1–2 minutos en reportes grandes. No cierres la página."):
            xlsx, stats = procesar_pdf(archivo.getvalue())
    except Exception as e:
        st.error("No se pudo procesar el PDF. ¿Es el reporte correcto (Plan de Intervención)?")
        st.exception(e)
        st.stop()

    st.success("¡Listo! Excel generado%s." % ((" para " + stats["mes"]) if stats.get("mes") else ""))
    c1, c2, c3 = st.columns(3)
    c1.metric("Niños / jóvenes", stats["ninos"])
    c2.metric("Intervenciones", stats["filas"])
    c3.metric("Visitas reales", stats["visitas"])

    nombre = archivo.name.rsplit(".", 1)[0] + " - EXTRAIDO.xlsx"
    st.download_button(
        "⬇️  Descargar Excel",
        data=xlsx,
        file_name=nombre,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
    st.caption(
        "🔒 Este archivo contiene datos sensibles. Guárdalo en un lugar seguro y no lo "
        "compartas por canales abiertos. Ábrelo en Microsoft Excel para ver bien los gráficos."
    )
