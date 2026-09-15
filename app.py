import hashlib
import os
import tempfile
from pathlib import Path

import streamlit as st

from motor_auditoria import run_audit, __project__, __version__

# =========================================================
# METADATOS DEL PROYECTO
# =========================================================
__author__ = "Gabriel Torrico Torrejon"
__app_name__ = "Control Integral de Compras SAP–SIAT"

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Control Integral de Compras SAP–SIAT | Univalle",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# =========================================================
# IDENTIDAD VISUAL UNIVALLE
# Paleta institucional inspirada en la identidad visual actual:
# magenta + grafito + fondos claros.
# =========================================================
st.markdown(
    """
<style>
    :root {
        --uv-magenta: #C5004E;
        --uv-magenta-dark: #98003C;
        --uv-magenta-soft: #FBEAF1;
        --uv-graphite: #31343A;
        --uv-gray: #6F737B;
        --uv-gray-soft: #EEF0F3;
        --uv-border: #E1E4E8;
        --uv-bg: #F7F8FA;
        --uv-white: #FFFFFF;
        --uv-success: #18794E;
        --uv-success-bg: #EAF7F0;
        --uv-warning: #A56300;
        --uv-warning-bg: #FFF6DF;
        --uv-danger: #B42318;
        --uv-danger-bg: #FDECEC;
    }

    /* Página */
    .stApp {
        background:
            radial-gradient(circle at 90% 0%, rgba(197, 0, 78, 0.06), transparent 24rem),
            var(--uv-bg);
        color: var(--uv-graphite);
    }

    .block-container {
        max-width: 1180px;
        padding-top: 1.35rem;
        padding-bottom: 3rem;
    }

    /* Limpieza visual de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header[data-testid="stHeader"] {
        background: transparent;
    }

    /* Encabezado institucional */
    .uv-hero {
        position: relative;
        overflow: hidden;
        background: linear-gradient(115deg, #2F3136 0%, #3A3338 45%, #6A1235 100%);
        border-radius: 22px;
        padding: 1.55rem 1.75rem 1.45rem 1.75rem;
        margin-bottom: 1.15rem;
        box-shadow: 0 14px 34px rgba(49, 52, 58, 0.12);
        border: 1px solid rgba(255,255,255,0.08);
    }

    .uv-hero::after {
        content: "";
        position: absolute;
        width: 240px;
        height: 240px;
        border-radius: 50%;
        background: rgba(197, 0, 78, 0.28);
        right: -80px;
        top: -95px;
    }

    .uv-brand-row {
        display: flex;
        align-items: center;
        gap: .7rem;
        margin-bottom: .7rem;
    }

    .uv-brand-dot {
        width: 12px;
        height: 12px;
        border-radius: 50%;
        background: #E30059;
        box-shadow: 0 0 0 6px rgba(227, 0, 89, 0.13);
    }

    .uv-brand {
        color: #FFFFFF;
        font-size: .82rem;
        font-weight: 800;
        letter-spacing: .15em;
        text-transform: uppercase;
    }

    .uv-title {
        color: #FFFFFF;
        font-size: clamp(1.75rem, 3vw, 2.55rem);
        font-weight: 800;
        line-height: 1.08;
        margin: 0;
        max-width: 850px;
    }

    .uv-subtitle {
        color: rgba(255,255,255,.82);
        font-size: 1rem;
        margin-top: .6rem;
        margin-bottom: 0;
        max-width: 760px;
    }

    .uv-accent-line {
        width: 78px;
        height: 4px;
        background: #E30059;
        border-radius: 999px;
        margin-top: 1rem;
    }

    /* Texto introductorio */
    .uv-intro {
        background: var(--uv-white);
        border: 1px solid var(--uv-border);
        border-radius: 14px;
        padding: .95rem 1.05rem;
        margin: .25rem 0 1rem 0;
        color: var(--uv-gray);
        font-size: .94rem;
        box-shadow: 0 4px 14px rgba(31, 35, 40, 0.035);
    }

    .uv-intro strong {
        color: var(--uv-graphite);
    }

    /* File uploaders */
    [data-testid="stFileUploader"] {
        background: var(--uv-white);
        border: 1px solid var(--uv-border);
        border-radius: 16px;
        padding: .55rem .75rem .35rem .75rem;
        box-shadow: 0 5px 16px rgba(31, 35, 40, 0.045);
    }

    [data-testid="stFileUploader"] label p {
        color: var(--uv-graphite) !important;
        font-weight: 750 !important;
        font-size: .97rem !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        background: #FBFBFC;
        border: 1.5px dashed #C9CDD3;
        border-radius: 12px;
    }

    [data-testid="stFileUploaderDropzone"]:hover {
        border-color: var(--uv-magenta);
        background: #FFF9FB;
    }

    [data-testid="stFileUploaderDropzone"] button {
        background: var(--uv-white) !important;
        color: var(--uv-magenta) !important;
        border: 1px solid #E6B7C9 !important;
        border-radius: 9px !important;
        font-weight: 750 !important;
    }

    /* Botones */
    .stButton > button {
        width: 100%;
        min-height: 3.15rem;
        border-radius: 11px !important;
        font-weight: 800 !important;
        letter-spacing: .015em;
        transition: all .18s ease;
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(90deg, var(--uv-magenta), #D80859) !important;
        color: white !important;
        border: 1px solid var(--uv-magenta) !important;
        box-shadow: 0 8px 18px rgba(197, 0, 78, .20);
    }

    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(90deg, var(--uv-magenta-dark), var(--uv-magenta)) !important;
        transform: translateY(-1px);
        box-shadow: 0 10px 22px rgba(197, 0, 78, .26);
    }

    .stButton > button:disabled {
        background: #E6E8EB !important;
        color: #979CA3 !important;
        border-color: #E1E4E8 !important;
        box-shadow: none !important;
    }

    [data-testid="stDownloadButton"] > button {
        width: 100%;
        min-height: 3.15rem;
        border-radius: 11px !important;
        font-weight: 800 !important;
        background: var(--uv-graphite) !important;
        color: white !important;
        border: 1px solid var(--uv-graphite) !important;
    }

    [data-testid="stDownloadButton"] > button:hover {
        background: #1F2125 !important;
        border-color: #1F2125 !important;
    }

    /* Expander / configuración */
    [data-testid="stExpander"] {
        background: var(--uv-white);
        border: 1px solid var(--uv-border);
        border-radius: 13px;
        box-shadow: 0 4px 14px rgba(31, 35, 40, 0.035);
    }

    [data-testid="stExpander"] summary p {
        font-weight: 700 !important;
        color: var(--uv-graphite) !important;
    }

    /* Inputs */
    [data-testid="stNumberInput"] input,
    [data-testid="stTextInput"] input {
        border-radius: 10px !important;
    }

    /* Alertas */
    [data-testid="stAlert"] {
        border-radius: 12px;
        border-width: 1px;
    }

    /* Métricas */
    [data-testid="stMetric"] {
        background: var(--uv-white);
        border: 1px solid var(--uv-border);
        border-radius: 14px;
        padding: .85rem .95rem;
        min-height: 104px;
        box-shadow: 0 4px 14px rgba(31, 35, 40, .04);
    }

    [data-testid="stMetricLabel"] {
        color: var(--uv-gray);
        font-weight: 700;
    }

    [data-testid="stMetricValue"] {
        color: var(--uv-graphite);
        font-weight: 800;
    }

    /* Progreso */
    [data-testid="stProgress"] > div > div > div > div {
        background-color: var(--uv-magenta) !important;
    }

    /* Títulos de sección */
    .uv-section-title {
        display: flex;
        align-items: center;
        gap: .55rem;
        color: var(--uv-graphite);
        font-size: 1.12rem;
        font-weight: 800;
        margin-top: .8rem;
        margin-bottom: .45rem;
    }

    .uv-section-title::before {
        content: "";
        width: 5px;
        height: 22px;
        border-radius: 4px;
        background: var(--uv-magenta);
    }

    /* Estados */
    .uv-status-row {
        display: flex;
        flex-wrap: wrap;
        gap: .5rem;
        margin: .25rem 0 .9rem 0;
    }

    .uv-chip {
        display: inline-flex;
        align-items: center;
        gap: .35rem;
        padding: .36rem .66rem;
        border-radius: 999px;
        font-size: .78rem;
        font-weight: 750;
        border: 1px solid transparent;
    }

    .uv-chip-ok {
        color: var(--uv-success);
        background: var(--uv-success-bg);
        border-color: #B9E3CD;
    }

    .uv-chip-warn {
        color: var(--uv-warning);
        background: var(--uv-warning-bg);
        border-color: #F2D58A;
    }

    .uv-chip-diff {
        color: var(--uv-magenta-dark);
        background: var(--uv-magenta-soft);
        border-color: #E9B7CA;
    }

    /* Pie */
    .uv-footer {
        margin-top: 2.4rem;
        padding-top: 1.2rem;
        border-top: 1px solid var(--uv-border);
        color: #969BA3;
        text-align: center;
        font-size: .76rem;
        line-height: 1.6;
    }

    /* Responsive */
    @media (max-width: 780px) {
        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
            padding-top: .85rem;
        }
        .uv-hero {
            padding: 1.25rem 1.1rem;
            border-radius: 17px;
        }
        .uv-title {
            font-size: 1.65rem;
        }
    }
</style>
""",
    unsafe_allow_html=True,
)


# =========================================================
# FUNCIONES AUXILIARES DE INTERFAZ
# =========================================================
def _configured_password() -> str:
    """Lee contraseña opcional desde variable de entorno o Secrets."""
    pwd = os.getenv("AUDITOR_PASSWORD", "")
    if pwd:
        return pwd
    try:
        return str(st.secrets.get("AUDITOR_PASSWORD", ""))
    except Exception:
        return ""


def _format_count(value) -> str:
    try:
        return f"{int(value):,}".replace(",", ".")
    except Exception:
        return str(value)


def _format_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 ** 2:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 ** 2):.2f} MB"


def _fingerprint(sap_file, siat_file, tolerance: float) -> str:
    """Evita mostrar resultados viejos si el usuario cambia los archivos."""
    h = hashlib.sha256()
    h.update(sap_file.getvalue())
    h.update(b"||SAP-SIAT||")
    h.update(siat_file.getvalue())
    h.update(f"|{tolerance:.4f}".encode("utf-8"))
    return h.hexdigest()


# =========================================================
# ACCESO OPCIONAL
# =========================================================
password = _configured_password()
if password:
    st.markdown(
        """
        <div class="uv-hero">
            <div class="uv-brand-row">
                <span class="uv-brand-dot"></span>
                <span class="uv-brand">UNIVALLE · CONTROL INTERNO</span>
            </div>
            <h1 class="uv-title">Control Integral de Compras SAP–SIAT</h1>
            <p class="uv-subtitle">Acceso restringido a la herramienta de auditoría.</p>
            <div class="uv-accent-line"></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    entered = st.text_input("Contraseña de acceso", type="password", placeholder="Ingresa la contraseña")
    if entered != password:
        if entered:
            st.error("Contraseña incorrecta.")
        st.stop()


# =========================================================
# ENCABEZADO
# =========================================================
st.markdown(
    """
    <div class="uv-hero">
        <div class="uv-brand-row">
            <span class="uv-brand-dot"></span>
            <span class="uv-brand">UNIVALLE · CONTROL INTERNO</span>
        </div>
        <h1 class="uv-title">Control Integral de Compras SAP–SIAT</h1>
        <p class="uv-subtitle">
            Auditoría automática de facturas, validación campo por campo y control previo de observaciones.
        </p>
        <div class="uv-accent-line"></div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="uv-intro">
        <strong>Flujo de trabajo:</strong> carga el archivo de compras SAP y el CSV del SIAT.
        El sistema normaliza los datos, repara incidencias estructurales detectables del SIAT,
        empareja las facturas y genera un reporte Excel con diferencias, redondeos y registros sin correspondencia.
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================================================
# CARGA DE ARCHIVOS
# =========================================================
st.markdown('<div class="uv-section-title">Archivos de entrada</div>', unsafe_allow_html=True)

col1, col2 = st.columns(2, gap="large")

with col1:
    sap_file = st.file_uploader(
        "1. Archivo de compras SAP",
        type=["xlsx", "xls"],
        key="sap",
        help="Selecciona el archivo Excel exportado desde SAP.",
    )
    if sap_file:
        st.success(f"✓ SAP cargado · {sap_file.name} · {_format_size(sap_file.size)}")

with col2:
    siat_file = st.file_uploader(
        "2. Archivo de compras SIAT",
        type=["csv"],
        key="siat",
        help="Selecciona el archivo CSV exportado desde SIAT.",
    )
    if siat_file:
        st.success(f"✓ SIAT cargado · {siat_file.name} · {_format_size(siat_file.size)}")

# =========================================================
# CONFIGURACIÓN
# =========================================================
with st.expander("⚙️ Configuración avanzada", expanded=False):
    round_tol = st.number_input(
        "Tolerancia para revisar redondeo (Bs)",
        min_value=0.00,
        max_value=10.00,
        value=0.05,
        step=0.01,
        format="%.2f",
        help=(
            "Las diferencias monetarias iguales o menores a este valor se clasifican "
            "como REVISAR REDONDEO y no como error real."
        ),
    )
    st.caption(
        "Valor recomendado: Bs 0,05. El cambio de esta tolerancia no modifica los importes originales; "
        "solo cambia la clasificación de la observación."
    )

ready = sap_file is not None and siat_file is not None

if ready:
    current_fingerprint = _fingerprint(sap_file, siat_file, float(round_tol))
    previous_fingerprint = st.session_state.get("input_fingerprint")

    if previous_fingerprint and previous_fingerprint != current_fingerprint:
        st.session_state.pop("result_bytes", None)
        st.session_state.pop("metrics", None)
        st.session_state.pop("input_fingerprint", None)

    st.markdown(
        """
        <div class="uv-status-row">
            <span class="uv-chip uv-chip-ok">● Archivos listos</span>
            <span class="uv-chip uv-chip-warn">● Centavos: revisar redondeo</span>
            <span class="uv-chip uv-chip-diff">● Diferencias reales: observar</span>
        </div>
        """,
        unsafe_allow_html=True,
    )
else:
    st.info("Carga ambos archivos para habilitar la auditoría.")

# =========================================================
# EJECUCIÓN
# =========================================================
run = st.button(
    "AUDITAR Y GENERAR REPORTE",
    type="primary",
    use_container_width=True,
    disabled=not ready,
)

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

            progress.progress(15, text="Validando estructura y normalizando datos...")
            progress.progress(32, text="Revisando y reparando incidencias del CSV SIAT...")
            progress.progress(48, text="Identificando y emparejando facturas...")

            _, metrics = run_audit(
                sap_path,
                siat_path,
                output_path=out_path,
                round_tol=float(round_tol),
            )

            progress.progress(88, text="Construyendo el reporte Excel de auditoría...")
            result_bytes = out_path.read_bytes()

        st.session_state["result_bytes"] = result_bytes
        st.session_state["metrics"] = metrics
        st.session_state["input_fingerprint"] = current_fingerprint

        progress.progress(100, text="Auditoría finalizada correctamente")
        st.success("✓ Auditoría completada. Revisa el resumen y descarga el reporte Excel.")

    except Exception as exc:
        progress.empty()
        st.error("No se pudo completar la auditoría. Revisa los archivos de entrada y vuelve a intentarlo.")
        with st.expander("Detalle técnico del error"):
            st.code(f"{type(exc).__name__}: {exc}")

# =========================================================
# RESULTADOS
# =========================================================
if st.session_state.get("metrics"):
    m = st.session_state["metrics"]

    st.markdown('<div class="uv-section-title">Resumen de auditoría</div>', unsafe_allow_html=True)

    a, b, c, d = st.columns(4, gap="medium")
    a.metric("Registros SAP", _format_count(m["sap"]))
    b.metric("Registros SIAT", _format_count(m["siat"]))
    c.metric("Emparejadas", _format_count(m["matched"]))
    d.metric("Sin diferencias", _format_count(m["ok"]))

    a, b, c, d = st.columns(4, gap="medium")
    a.metric("Con diferencias", _format_count(m["diff"]))
    b.metric("Revisar redondeo", _format_count(m["rounding"]))
    c.metric("Solo SAP", _format_count(m["solo_sap"]))
    d.metric("Solo SIAT", _format_count(m["solo_siat"]))

    if m.get("repairs", 0):
        st.warning(
            f"El CSV del SIAT contenía {m['repairs']} incidencia(s) estructural(es) reparada(s) "
            "automáticamente. La trazabilidad de la reparación está incluida en el Excel."
        )

# =========================================================
# DESCARGA
# =========================================================
if st.session_state.get("result_bytes"):
    st.markdown('<div class="uv-section-title">Reporte final</div>', unsafe_allow_html=True)

    st.download_button(
        "DESCARGAR REPORTE EXCEL",
        data=st.session_state["result_bytes"],
        file_name="CONTROL_INTEGRAL_COMPRAS_SAP_SIAT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.caption(
        "El archivo incluye resumen, control general, detalle de diferencias, redondeos, "
        "registros exclusivos de SAP/SIAT, reparaciones estructurales y bases normalizadas."
    )

# =========================================================
# PIE
# =========================================================
st.markdown(
    f"""
    <div class="uv-footer">
        <strong>UNIVALLE</strong> · {__project__} · v{__version__}<br>
        Herramienta interna de control y apoyo a la revisión de compras.
    </div>
    """,
    unsafe_allow_html=True,
)
