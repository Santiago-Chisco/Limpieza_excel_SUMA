"""Limpiador de Informe General -> MATRIZ (Streamlit).

Ejecutar:  streamlit run app.py
"""
import hmac
import os
from pathlib import Path

import pandas as pd
import streamlit as st

import cleaner
import informe

st.set_page_config(page_title="Informe → MATRIZ", page_icon="🧹", layout="wide")


def require_login() -> None:
    """Pide usuario y contraseña si APP_PASSWORD está definida (uso en la nube)."""
    password = os.environ.get("APP_PASSWORD")
    if not password or st.session_state.get("autenticado"):
        return
    user = os.environ.get("APP_USER", "admin")
    st.title("🔒 Informe → MATRIZ")
    with st.form("login"):
        u = st.text_input("Usuario")
        p = st.text_input("Contraseña", type="password")
        ok = st.form_submit_button("Entrar")
    if ok:
        if hmac.compare_digest(u, user) and hmac.compare_digest(p, password):
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Usuario o contraseña incorrectos.")
    st.stop()


require_login()
st.title("🧹 Informe General → MATRIZ")
st.caption(
    "Sube el Excel con filas en rojo. La aplicación elimina las rojas, conserva las demás "
    "y descarga la MATRIZ con el formato verde."
)

tab_main, tab_damaged = st.tabs(["📂 Procesar informe", "⚠️ Excel dañados"])

with tab_main:
    uploaded = st.file_uploader("Sube el Informe General (.xlsx)", type=["xlsx", "xlsm"])
    incluir_elim = st.checkbox("Incluir hoja con las filas eliminadas (para revisar)", value=False)

    if uploaded is not None:
        raw = uploaded.getvalue()
        try:
            data = informe.read_informe(raw)
        except cleaner.DamagedFileError as exc:
            saved = cleaner.save_damaged(raw, uploaded.name, str(exc))
            st.error(f"❌ No se pudo procesar **{uploaded.name}**.")
            st.caption(f"Motivo: {exc}")
            st.warning(f"Se guardó automáticamente en `{saved.parent}/` ({saved.name}).")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Registros en el informe", data["total"])
            c2.metric("Filas rojas eliminadas", len(data["deleted"]))
            c3.metric("Filas conservadas", len(data["keep"]))

            out = informe.build_matriz(data["keep"], data["deleted"] if incluir_elim else None)

            st.subheader("Vista previa (lo que queda)")
            st.dataframe(
                pd.DataFrame(
                    [
                        {
                            "FECHA SERVICIO": r["fecha"].strftime("%Y/%m/%d"),
                            "MANIFIESTO": r["manifiesto"],
                            "ACTA": r["acta"],
                            "CLIENTE": r["cliente"],
                            "RESIDUO": r["residuo"],
                            "TRATAMIENTO": r["tratamiento"],
                            "PESO": r["peso"],
                            "UNIDAD": r["unidad"],
                            "PROVEEDOR": r["proveedor"],
                        }
                        for r in data["keep"]
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
            st.download_button(
                "⬇️ Descargar MATRIZ",
                data=out,
                file_name=f"MATRIZ_{Path(uploaded.name).stem}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )

with tab_damaged:
    st.subheader("Archivos dañados guardados")
    st.caption(f"Se guardan en `{cleaner.DAMAGED_DIR}/` con el motivo.")
    items = cleaner.list_damaged()
    if not items:
        st.info("Aún no hay archivos dañados guardados.")
    for i, item in enumerate(items):
        a, b, c = st.columns([6, 2, 1])
        a.markdown(
            f"**{item['nombre']}**  \n{item['fecha']} · {item['tamano_kb']} KB"
            + (f"  \n_Motivo: {item['motivo']}_" if item["motivo"] else "")
        )
        b.download_button("⬇️ Descargar", data=item["ruta"].read_bytes(), file_name=item["nombre"], key=f"dl_{i}")
        if c.button("🗑️", key=f"del_{i}", help="Eliminar"):
            cleaner.delete_damaged(item["ruta"])
            st.rerun()
        st.divider()
