"""
app.py — Matriz de Seguridad Grow
Interfaz Streamlit en dos etapas: generar equivalencias → calcular roles.
"""

import streamlit as st
import pandas as pd
import io
from motor import (
    load_agr_users, load_agr_tcodes, load_iam_apps, load_st03n,
    generate_equivalencias_excel, load_equivalencias_completadas,
    run_motor, export_excel,
)

st.set_page_config(page_title="Matriz de Seguridad Grow", page_icon="🔐",
                   layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #F5F7FA; }
[data-testid="stHeader"] { background: transparent; }
#MainMenu, footer, header { visibility: hidden; }

.app-header {
    background: linear-gradient(135deg, #1F3864 0%, #2E75B6 100%);
    border-radius: 12px; padding: 24px 36px; margin-bottom: 20px;
    display: flex; align-items: center; gap: 18px;
}
.app-header-icon { font-size: 36px; }
.app-header-title { color: white; font-size: 22px; font-weight: 700; margin: 0; }
.app-header-sub   { color: rgba(255,255,255,0.75); font-size: 12px; margin: 3px 0 0; }

/* Etapa buttons */
.etapa-btn {
    display: inline-block; padding: 10px 20px; border-radius: 8px;
    font-size: 13px; font-weight: 600; cursor: pointer; text-align: center;
    border: 2px solid #2E75B6; transition: all 0.15s;
}
.etapa-btn-active {
    background: #1F3864; color: white !important; border-color: #1F3864;
}
.etapa-btn-inactive {
    background: white; color: #1F3864 !important; border-color: #2E75B6;
}

.step-card {
    background: white; border-radius: 10px; border: 1px solid #E0E7EF;
    padding: 18px 22px; margin-bottom: 14px;
    box-shadow: 0 1px 4px rgba(31,56,100,0.06);
}
.step-header { display: flex; align-items: center; gap: 12px; margin-bottom: 12px; }
.step-num {
    background: #CBD5E0; color: white; width: 28px; height: 28px;
    border-radius: 50%; display: flex; align-items: center; justify-content: center;
    font-size: 12px; font-weight: 700; flex-shrink: 0;
}
.step-num.done   { background: #00B050; }
.step-num.active { background: #2E75B6; }
.step-title { font-size: 14px; font-weight: 600; color: #1F3864; }
.step-sub   { font-size: 12px; color: #666; margin-top: 2px; }

.upload-label {
    font-size: 11px; font-weight: 600; color: #444;
    margin-bottom: 4px; text-transform: uppercase; letter-spacing: 0.3px;
}
.metric-row { display: flex; gap: 10px; margin: 12px 0; flex-wrap: wrap; }
.metric-card {
    flex: 1; min-width: 100px; background: white; border-radius: 8px;
    border: 1px solid #E0E7EF; padding: 12px 14px; text-align: center;
}
.metric-val { font-size: 28px; font-weight: 700; color: #1F3864; }
.metric-val.green  { color: #00B050; }
.metric-val.orange { color: #ED7D31; }
.metric-val.red    { color: #C00000; }
.metric-lbl {
    font-size: 10px; color: #888; margin-top: 2px;
    text-transform: uppercase; letter-spacing: 0.3px;
}
.info-box {
    background: #EEF3FA; border-left: 4px solid #2E75B6;
    padding: 10px 16px; border-radius: 0 8px 8px 0;
    font-size: 13px; color: #333; margin: 10px 0;
}
.warn-box {
    background: #FFF9E6; border-left: 4px solid #ED7D31;
    padding: 10px 16px; border-radius: 0 8px 8px 0;
    font-size: 13px; color: #7E3000; margin: 10px 0;
}
/* Fix Streamlit's file uploader label color */
[data-testid="stFileUploader"] label { color: #333 !important; }
/* Fix st.info background */
[data-testid="stAlert"] { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── HEADER ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="app-header">
    <div class="app-header-icon">🔐</div>
    <div>
        <div class="app-header-title">Matriz de Seguridad Grow</div>
        <div class="app-header-sub">
            Migración de autorizaciones: Private Cloud / On-Premise
            → SAP S/4HANA Cloud Public Edition
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── SESSION STATE ─────────────────────────────────────────────────────────────
for k, v in {
    "cliente": "", "proyecto": "", "etapa": 1,
    "df_users": None, "df_tcodes": None, "df_iam": None, "df_st03n": None,
    "df_equiv": None, "result": None, "excel_bytes": None,
    "equiv_bytes_etapa1": None, "files_were_ready": False,
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


def snum(n, state="pending"):
    # state: "done" | "active" | "pending"
    cls = state
    lbl = "✓" if state == "done" else str(n)
    return f'<div class="step-num {cls}">{lbl}</div>'


# ── BLOQUE 0: CLIENTE / PROYECTO / ETAPA ─────────────────────────────────────
with st.container():
    st.markdown('<div class="step-card" style="padding:14px 22px;">', unsafe_allow_html=True)

    col_meta1, col_meta2, col_meta3 = st.columns([2, 2, 3])

    with col_meta1:
        st.markdown('<div class="upload-label">Cliente</div>', unsafe_allow_html=True)
        cliente = st.text_input("cli", value=st.session_state.cliente,
                                placeholder="Ej: Triunfo",
                                label_visibility="collapsed", key="inp_cli")
        st.session_state.cliente = cliente

    with col_meta2:
        st.markdown('<div class="upload-label">Proyecto</div>', unsafe_allow_html=True)
        proyecto = st.text_input("pro", value=st.session_state.proyecto,
                                 placeholder="Ej: SAP ERP Cloud",
                                 label_visibility="collapsed", key="inp_pro")
        st.session_state.proyecto = proyecto

    with col_meta3:
        st.markdown('<div class="upload-label">Etapa de trabajo</div>', unsafe_allow_html=True)
        col_e1, col_e2 = st.columns(2)
        with col_e1:
            if st.button("📋  Etapa 1\nGenerar equivalencias",
                         use_container_width=True,
                         type="primary" if st.session_state.etapa == 1 else "secondary"):
                st.session_state.etapa = 1
                st.rerun()
        with col_e2:
            if st.button("▶  Etapa 2\nCalcular roles",
                         use_container_width=True,
                         type="primary" if st.session_state.etapa == 2 else "secondary"):
                st.session_state.etapa = 2
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

etapa_num = st.session_state.etapa
suffix = " ".join(filter(None, [proyecto, cliente])).strip()

# ── BLOQUE 1: ARCHIVOS BASE ───────────────────────────────────────────────────
files_ready_prev = all(st.session_state[k] is not None
                       for k in ("df_users", "df_tcodes", "df_iam", "df_st03n"))

step1_state = "done" if files_ready_prev else "active"

st.markdown(f"""
<div class="step-card">
<div class="step-header">
    {snum(1, step1_state)}
    <div>
        <div class="step-title">Cargar archivos del sistema origen y Grow</div>
        <div class="step-sub">
            AGR_USERS · AGR_TCODES de SE16N (.xlsx o .txt) ·
            IAM Apps del Starter · ST03N 3 meses
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.markdown('<div class="upload-label">AGR_USERS *</div>', unsafe_allow_html=True)
    f = st.file_uploader("u", type=["xlsx","xls","txt"], key="up_users",
                          label_visibility="collapsed")
    if f:
        try:
            st.session_state.df_users = load_agr_users(f)
            st.success(f"✓ {st.session_state.df_users['User'].nunique()} usuarios")
        except Exception as e:
            st.error(str(e)); st.session_state.df_users = None

with col2:
    st.markdown('<div class="upload-label">AGR_TCODES *</div>', unsafe_allow_html=True)
    f = st.file_uploader("t", type=["xlsx","xls","txt"], key="up_tcodes",
                          label_visibility="collapsed")
    if f:
        try:
            st.session_state.df_tcodes = load_agr_tcodes(f)
            st.success(f"✓ {len(st.session_state.df_tcodes):,} pares Rol→TCode")
        except Exception as e:
            st.error(str(e)); st.session_state.df_tcodes = None

with col3:
    st.markdown('<div class="upload-label">IAM Apps — Starter Grow *</div>',
                unsafe_allow_html=True)
    f = st.file_uploader("i", type=["xlsx","xls"], key="up_iam",
                          label_visibility="collapsed")
    if f:
        try:
            st.session_state.df_iam = load_iam_apps(f)
            st.success(f"✓ {st.session_state.df_iam['BRT_ID'].nunique()} Role Templates")
        except Exception as e:
            st.error(str(e)); st.session_state.df_iam = None

with col4:
    st.markdown('<div class="upload-label">ST03N — últimos 3 meses *</div>',
                unsafe_allow_html=True)
    fs = st.file_uploader("s", type=["xlsx","xls"], key="up_st03n",
                           label_visibility="collapsed", accept_multiple_files=True)
    if fs:
        try:
            st.session_state.df_st03n = load_st03n(fs)
            st.success(f"✓ {len(st.session_state.df_st03n):,} TCodes con uso")
        except Exception as e:
            st.error(str(e)); st.session_state.df_st03n = None

# Recompute after uploaders
files_ready = all(st.session_state[k] is not None
                  for k in ("df_users", "df_tcodes", "df_iam", "df_st03n"))

if not files_ready:
    missing = [n for n, k in [("AGR_USERS","df_users"),("AGR_TCODES","df_tcodes"),
                                ("IAM Apps","df_iam"),("ST03N","df_st03n")]
               if st.session_state[k] is None]
    st.caption(f"* Obligatorios — faltan: {', '.join(missing)}")

st.markdown('</div>', unsafe_allow_html=True)

# Trigger rerun once when files first become ready
if files_ready and not st.session_state.files_were_ready:
    st.session_state.files_were_ready = True
    st.rerun()


# ═══════════════════════════════════════════════════════════════════════════════
# ETAPA 1: GENERAR EQUIVALENCIAS
# ═══════════════════════════════════════════════════════════════════════════════
if etapa_num == 1:

    done1 = st.session_state.equiv_bytes_etapa1 is not None
    step2_state = "done" if done1 else ("active" if files_ready else "pending")

    st.markdown(f"""
    <div class="step-card">
    <div class="step-header">
        {snum(2, step2_state)}
        <div>
            <div class="step-title">Generar archivo de equivalencias</div>
            <div class="step-sub">
                Excel para que los consultores completen el mapeo TCode → App Grow.
                Incluye desplegable con todas las apps válidas del tenant.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not files_ready:
        st.info("⬆️ Cargá los 4 archivos del Paso 1 para continuar.")
    else:
        n_active = st.session_state.df_st03n["TCode"].nunique()
        n_apps   = st.session_state.df_iam["Transaction_Code"].nunique()
        st.markdown(f"""<div class="info-box">
            <strong>El archivo incluirá:</strong>
            {n_active:,} TCodes con uso real (ST03N) ·
            Desplegable con {n_apps:,} apps del Starter ·
            Los TCodes con resolución automática ya vendrán propuestos en verde
        </div>""", unsafe_allow_html=True)

        if st.button("📋 Generar archivo de equivalencias", type="primary",
                     use_container_width=False):
            with st.spinner("Generando..."):
                try:
                    st.session_state.equiv_bytes_etapa1 = generate_equivalencias_excel(
                        df_users=st.session_state.df_users,
                        df_tcodes=st.session_state.df_tcodes,
                        df_iam=st.session_state.df_iam,
                        df_st03n=st.session_state.df_st03n,
                        cliente=st.session_state.cliente,
                        proyecto=st.session_state.proyecto,
                    )
                    st.rerun()
                except Exception as e:
                    st.error(f"Error generando archivo: {e}")
                    import traceback; st.code(traceback.format_exc())

        if done1:
            fname = f"Equivalencias a Completar{' - ' + suffix if suffix else ''}.xlsx"
            st.success(f"✅ Listo: **{fname}**")
            st.download_button(
                label=f"⬇️ Descargar: {fname}",
                data=st.session_state.equiv_bytes_etapa1,
                file_name=fname,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True,
            )
            st.markdown(f"""<div class="info-box">
                <strong>Próximos pasos:</strong><br>
                1. Compartir <strong>{fname}</strong> con los consultores funcionales<br>
                2. Completar la columna <strong>App_Grow_ID</strong> usando el desplegable
                   (formato: <em>F1443A — Manage Cost Centers</em>)<br>
                3. <span style="background:#FFFACD;padding:1px 6px;border-radius:3px;">
                   Amarillo</span> = pendiente &nbsp;|&nbsp;
                   <span style="background:#E2EFDA;padding:1px 6px;border-radius:3px;">
                   Verde</span> = ya propuesto automáticamente<br>
                4. Cuando esté completo: hacer clic en
                   <strong>▶ Etapa 2 — Calcular roles</strong> (arriba a la derecha)
                   y subir el archivo
            </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# ETAPA 2: CALCULAR ROLES
# ═══════════════════════════════════════════════════════════════════════════════
if etapa_num == 2:

    # ── Paso 2: cargar equivalencias ─────────────────────────────────────────
    equiv_loaded_prev = st.session_state.df_equiv is not None
    step2e_state = "done" if equiv_loaded_prev else ("active" if files_ready else "pending")

    st.markdown(f"""
    <div class="step-card">
    <div class="step-header">
        {snum(2, step2e_state)}
        <div>
            <div class="step-title">Cargar equivalencias completadas por los consultores</div>
            <div class="step-sub">
                El archivo "Equivalencias a Completar" generado en la Etapa 1,
                con la columna App_Grow_ID completada
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    f_eq = st.file_uploader(
        "Equivalencias a Completar (.xlsx)",
        type=["xlsx"], key="up_equiv",
        help="Archivo con la columna App_Grow_ID completada por los consultores")

    if f_eq:
        try:
            st.session_state.df_equiv = load_equivalencias_completadas(f_eq)
            df_e = st.session_state.df_equiv
            n_ok   = (df_e["App_Grow_ID"] != "").sum()
            n_pend = len(df_e) - n_ok
            st.success(f"✓ {len(df_e)} TCodes cargados — "
                       f"{n_ok} con equivalencia asignada, {n_pend} sin completar")
            if n_pend > 0:
                st.markdown(
                    f'<div class="warn-box">⚠️ <strong>{n_pend} TCodes</strong> no tienen '
                    f'equivalencia — se excluirán del cálculo de roles.</div>',
                    unsafe_allow_html=True)
        except Exception as e:
            st.error(f"Error al cargar equivalencias: {e}")
            st.session_state.df_equiv = None

    # Recompute after uploader
    equiv_loaded = st.session_state.df_equiv is not None

    # Trigger rerun if equiv just became loaded
    if equiv_loaded and not equiv_loaded_prev:
        st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # ── Paso 3: calcular ─────────────────────────────────────────────────────
    result_ready = st.session_state.result is not None
    step3_state = ("done" if result_ready
                   else ("active" if (files_ready and equiv_loaded) else "pending"))

    st.markdown(f"""
    <div class="step-card">
    <div class="step-header">
        {snum(3, step3_state)}
        <div>
            <div class="step-title">Calcular asignaciones de roles</div>
            <div class="step-sub">
                Cruza usuarios, roles, TCodes, equivalencias y mapa IAM →
                produce el OUTPUT por usuario
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if not files_ready:
        st.info("⬆️ Cargá los archivos del Paso 1 primero.")
    elif not equiv_loaded:
        st.info("⬆️ Cargá el archivo de equivalencias en el Paso 2.")
    else:
        if st.button("▶  CALCULAR ROLES", type="primary",
                     use_container_width=False):
            pb = st.progress(0)
            st_txt = st.empty()

            def cb(msg, pct):
                pb.progress(pct / 100)
                st_txt.markdown(
                    f'<div style="font-size:13px;color:#555;">{msg}</div>',
                    unsafe_allow_html=True)

            try:
                st.session_state.result = run_motor(
                    st.session_state.df_users,
                    st.session_state.df_tcodes,
                    st.session_state.df_iam,
                    st.session_state.df_equiv,
                    st.session_state.df_st03n,
                    cb,
                )
                st.session_state.excel_bytes = None
                pb.empty(); st_txt.empty()
                st.rerun()
            except Exception as e:
                pb.empty(); st_txt.empty()
                st.error(f"Error: {e}")
                import traceback; st.code(traceback.format_exc())

        if result_ready:
            s = st.session_state.result["stats"]
            pct = int(s["complete_users"] / s["total_users"] * 100) if s["total_users"] else 0

            st.markdown(f"""
            <div class="metric-row">
                <div class="metric-card">
                    <div class="metric-val">{s["total_users"]}</div>
                    <div class="metric-lbl">Usuarios</div>
                </div>
                <div class="metric-card">
                    <div class="metric-val green">{s["complete_users"]}</div>
                    <div class="metric-lbl">Completos</div>
                </div>
                <div class="metric-card">
                    <div class="metric-val orange">{s["total_users"]-s["complete_users"]}</div>
                    <div class="metric-lbl">Con gaps</div>
                </div>
                <div class="metric-card">
                    <div class="metric-val">{s["total_roles_assigned"]}</div>
                    <div class="metric-lbl">Roles únicos</div>
                </div>
                <div class="metric-card">
                    <div class="metric-val red">{s["gap_tcodes"]}</div>
                    <div class="metric-lbl">Sin equiv.</div>
                </div>
            </div>
            <div style="background:#E2EFDA;border-radius:8px;padding:10px 16px;margin:8px 0;">
                <strong style="color:#375623;">{pct}% de usuarios completamente mapeados</strong>
                <div style="background:#C6EFCE;height:6px;border-radius:3px;margin-top:6px;">
                    <div style="background:#00B050;height:6px;border-radius:3px;
                                width:{pct}%;"></div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander("👁 Vista previa del OUTPUT", expanded=True):
                st.dataframe(
                    st.session_state.result["df_output"].head(20),
                    use_container_width=True, hide_index=True, height=260)
                n_total = len(st.session_state.result["df_output"])
                if n_total > 20:
                    st.caption(f"Mostrando 20 de {n_total} usuarios.")

            st.markdown(
                "<hr style='border:none;border-top:1px solid #E0E7EF;margin:14px 0;'>",
                unsafe_allow_html=True)
            st.markdown(
                "<div style='font-size:13px;font-weight:600;color:#333;"
                "margin:14px 0 6px;'>⬇️ Descargas</div>",
                unsafe_allow_html=True)

            # Generate Excel once, cache in session state
            if st.session_state.excel_bytes is None:
                try:
                    with st.spinner("Generando Excel..."):
                        st.session_state.excel_bytes = export_excel(
                            st.session_state.result,
                            st.session_state.df_equiv,
                            cliente=st.session_state.cliente,
                            proyecto=st.session_state.proyecto,
                        )
                except Exception as e:
                    st.error(f"Error generando Excel: {e}")

            cd1, cd2, cd3 = st.columns([2, 2, 3])

            with cd1:
                if st.session_state.excel_bytes:
                    fname = f"Matriz Seguridad Grow{' - ' + suffix if suffix else ''}.xlsx"
                    st.download_button(
                        "⬇️ Excel completo (.xlsx)",
                        data=st.session_state.excel_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", use_container_width=True, key="dl_excel",
                    )

            with cd2:
                try:
                    buf = io.BytesIO()
                    st.session_state.df_equiv.to_excel(
                        buf, index=False, sheet_name="Equivalencias")
                    buf.seek(0)
                    fname_eq = f"Equivalencias Grow{' - ' + suffix if suffix else ''}.xlsx"
                    st.download_button(
                        "💾 Guardar equivalencias",
                        data=buf.getvalue(),
                        file_name=fname_eq,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary", use_container_width=True, key="dl_eq",
                        help="Reutilizá estas equivalencias en el próximo cliente",
                    )
                except Exception as e:
                    st.error(f"Error: {e}")

            with cd3:
                st.markdown("""
                <div style="font-size:12px;color:#555;padding:8px 0;">
                    <strong>El Excel contiene:</strong>
                    OUTPUT · Equivalencias · GAP_Activo · Mapeo_Detalle · MOTOR<br>
                    <span style="color:#888;">
                    Próximo cliente: guardá las equivalencias y subilas en la Etapa 1.
                    </span>
                </div>""", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


# ── FOOTER ────────────────────────────────────────────────────────────────────
st.markdown("<br>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center;font-size:11px;color:#AAA;padding:8px 0;">
    Matriz de Seguridad Grow — Process Technologies &nbsp;|&nbsp;
    Datos procesados localmente, nada sale de tu máquina
</div>
""", unsafe_allow_html=True)
