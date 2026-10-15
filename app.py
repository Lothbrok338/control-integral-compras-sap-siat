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
    st.session_state.pop("metadata_diagnostico", None)

    progress = st.progress(0, text="Preparando archivos...")

    try:
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            sap_path = td / Path(sap_file.name or "SAP.xlsx").name
            siat_path = td / Path(siat_file.name or "SIAT.csv").name
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
            # Postproceso: sólo 3 hojas; usuario y asiento SAP se buscan en el Excel fuente.
            operativo_path = td / "REPORTE_OPERATIVO_COMPRAS_SAP_SIAT.xlsx"
            diagnostico = generar_reporte_operativo(
                reporte_original=out_path,
                sap_original=sap_path,
                reporte_salida=operativo_path,
            )
            result_bytes = operativo_path.read_bytes()

        st.session_state["result_bytes"] = result_bytes
        st.session_state["metrics"] = metrics
        st.session_state["metadata_diagnostico"] = diagnostico
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

    diag = st.session_state.get("metadata_diagnostico") or {}
    if diag:
        st.caption(
            f"Trazabilidad de SAP: {diag.get('filas_con_usuario', 0)} con usuario · "
            f"{diag.get('filas_con_asiento', 0)} con asiento. "
            "Los valores sin identificación segura quedan vacíos."
        )
        for advertencia in diag.get("advertencias", []):
            st.warning(advertencia)

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
        file_name="REPORTE_OPERATIVO_COMPRAS_SAP_SIAT.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.caption(
        "Excel operativo con solo tres pestañas: CONTROL GENERAL, POR CONCILIAR y RESUMEN. "
        "Incluye usuario SAP y asiento cuando se verifican contra el archivo original."
    )

# =========================================================
# PIE
# =========================================================
st.markdown(
    f"""
    <div class="uv-footer">
