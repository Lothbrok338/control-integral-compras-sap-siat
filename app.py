import os
import tempfile
from pathlib import Path

import streamlit as st

from motor_auditoria import run_audit, __project__, __version__

st.set_page_config(
    page_title="Control Integral de Compras SAP–SIAT",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
.block-container {max-width: 980px; padding-top: 2rem; padding-bottom: 3rem;}
.main-title {font-size: 2rem; font-weight: 750; margin-bottom: .15rem;}
.sub-title {color: #667085; margin-bottom: 1.4rem;}
[data-testid="stFileUploader"] {border: 1px solid #d0d5dd; border-radius: 12px; padding: .35rem .55rem;}
.stButton > button {width: 100%; min-height: 3rem; font-weight: 700; border-radius: 10px;}
[data-testid="stDownloadButton"] > button {width: 100%; min-height: 3rem; font-weight: 700; border-radius: 10px;}
.small-note {color:#667085; font-size:.86rem;}
.footer {color:#98a2b3; font-size:.76rem; text-align:center; margin-top:2.2rem;}
</style>
""", unsafe_allow_html=True)

# Acceso opcional: si defines AUDITOR_PASSWORD en variables de entorno o Secrets,
# la app pedirá contraseña. Si no se define, entra directamente.
def _configured_password():
    pwd = os.getenv("AUDITOR_PASSWORD", "")
    if pwd:
        return pwd
    try:
        return str(st.secrets.get("AUDITOR_PASSWORD", ""))
    except Exception:
        return ""

password = _configured_password()
if password:
    st.markdown('<div class="main-title">Control Integral de Compras SAP–SIAT</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Acceso restringido</div>', unsafe_allow_html=True)
    entered = st.text_input("Contraseña", type="password")
    if entered != password:
        if entered:
            st.error("Contraseña incorrecta.")
        st.stop()

st.markdown('<div class="main-title">Control Integral de Compras SAP–SIAT</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Auditoría automática de facturas y diferencias campo por campo</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    sap_file = st.file_uploader("1. Archivo SAP", type=["xlsx", "xls"], key="sap")
    if sap_file:
        st.success(f"SAP cargado: {sap_file.name}")
with col2:
    siat_file = st.file_uploader("2. Archivo SIAT", type=["csv"], key="siat")
    if siat_file:
        st.success(f"SIAT cargado: {siat_file.name}")

with st.expander("Configuración avanzada"):
    round_tol = st.number_input(
        "Tolerancia para revisar redondeo (Bs)",
        min_value=0.00,
        max_value=10.00,
        value=0.05,
        step=0.01,
        format="%.2f",
        help="Las diferencias monetarias hasta este valor se clasifican como REVISAR REDONDEO y no como error.",
    )

ready = sap_file is not None and siat_file is not None
if ready:
    st.info("Ambos archivos están listos. Puedes iniciar la auditoría.")
else:
    st.caption("Carga ambos archivos para habilitar la auditoría.")

run = st.button("AUDITAR Y EXPORTAR", type="primary", disabled=not ready)

if run:
    st.session_state.pop("result_bytes", None)
    st.session_state.pop("metrics", None)
    progress = st.progress(0, text="Preparando archivos...")
    try:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            sap_path = td / (sap_file.name or "SAP.xlsx")
            siat_path = td / (siat_file.name or "SIAT.csv")
            out_path = td / "CONTROL_INTEGRAL_COMPRAS_SAP_SIAT.xlsx"

            sap_path.write_bytes(sap_file.getvalue())
            siat_path.write_bytes(siat_file.getvalue())
            progress.progress(15, text="Validando y normalizando archivos...")

            # El motor realiza internamente reparación SIAT, emparejamiento,
            # comparación campo por campo y creación del Excel final.
            progress.progress(35, text="Reparando estructura SIAT y emparejando facturas...")
            _, metrics = run_audit(
                sap_path,
                siat_path,
                output_path=out_path,
                round_tol=float(round_tol),
            )
            progress.progress(90, text="Generando reporte de auditoría...")
            result_bytes = out_path.read_bytes()

        st.session_state["result_bytes"] = result_bytes
        st.session_state["metrics"] = metrics
        progress.progress(100, text="Auditoría terminada")
        st.success("Auditoría completada correctamente.")
    except Exception as exc:
        progress.empty()
        st.error(f"No se pudo completar la auditoría: {exc}")

if st.session_state.get("metrics"):
    m = st.session_state["metrics"]
    st.markdown("### Resultado")
    a, b, c, d = st.columns(4)
    a.metric("SAP", f"{m['sap']:,}".replace(',', '.'))
    b.metric("SIAT", f"{m['siat']:,}".replace(',', '.'))
    c.metric("Emparejadas", f"{m['matched']:,}".replace(',', '.'))
    d.metric("OK", f"{m['ok']:,}".replace(',', '.'))

    a, b, c, d = st.columns(4)
    a.metric("Con diferencias", f"{m['diff']:,}".replace(',', '.'))
    b.metric("Redondeos", f"{m['rounding']:,}".replace(',', '.'))
    c.metric("Solo SAP", f"{m['solo_sap']:,}".replace(',', '.'))
    d.metric("Solo SIAT", f"{m['solo_siat']:,}".replace(',', '.'))

    if m.get("repairs", 0):
        st.warning(f"El SIAT contenía {m['repairs']} fila(s) reparada(s) automáticamente. El detalle queda registrado en el Excel.")

if st.session_state.get("result_bytes"):
    st.download_button(
        "DESCARGAR REPORTE EXCEL",
        data=st.session_state["result_bytes"],
        file_name="CONTROL_INTEGRAL_COMPRAS_SAP_SIAT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

st.markdown(
    f'<div class="footer">{__project__} · v{__version__}</div>',
    unsafe_allow_html=True,
)
