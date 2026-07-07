# -*- coding: utf-8 -*-
"""
app.py — Visita a Clientes 
"""
from datetime import datetime

import pandas as pd
import streamlit as st
import io

from utils.helpers import (
    load_css, safe_str, safe_float, fmt_money, slug,
    cargar_excel, CRITERIOS_DEF, CLIENTE_VISITADO_OPCIONES,
    hay_borrador, guardar_borrador, cargar_borrador, borrar_borrador,
    registrar_historial, leer_historial, ahora_peru,
    calcular_resultado, criterios_seleccionados_lista,
    generar_word, generar_pdf, guardar_reporte_en_carpeta,
    sincronizar_historial_onedrive,
    reporte_consolidado_por_agencia, reporte_consolidado_por_cliente,
    # NUEVO — para la pantalla de Búsqueda igualando el mockup:
    iniciales, clase_calificacion, clientes_similares, solo_digitos,
    to_excel, reporte_anexo07, to_excel_anexo07,
)

st.set_page_config(
    page_title="Visita a Clientes",
    page_icon="🏦",
    layout="centered",
    initial_sidebar_state="collapsed",
)
load_css("assets/style.css")

# ACCESO — solo cuentas de Microsoft 
TENANT_ID_PERMITIDO = "f3831aea-ec1b-461b-b42f-ca26f9f78551"
DOMINIO_PERMITIDO = "@cajaarequipa.pe"


def _login_configurado() -> bool:
    try:
        return "auth" in st.secrets
    except Exception:
        return False


def pantalla_login():
    st.markdown(
        """<div class="login-wrap">
                <div class="login-icon">🏦</div>
                <h1>Visita a Clientes</h1>
                <p>Auditoría Interna · Caja Arequipa. Inicia sesión con tu
                cuenta corporativa de Microsoft para continuar.</p>
            </div>""",
        unsafe_allow_html=True,
    )
    st.markdown('<div class="btn-microsoft">', unsafe_allow_html=True)
    st.button("🔑  Iniciar sesión con Microsoft", use_container_width=True,
               on_click=st.login, args=["microsoft"])
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="login-badge">🔒 Acceso restringido a personal de Caja Arequipa</div>'
        '<div class="login-footer">Gerencia de Auditoría Interna</div>',
        unsafe_allow_html=True,
    )


if not _login_configurado():
    st.warning("⚠️")
elif not st.user.is_logged_in:
    pantalla_login()
    st.stop()
else:
    _correo = (st.user.get("email") or "").lower()
    _tid = st.user.get("tid") or ""
    if not _correo.endswith(DOMINIO_PERMITIDO) and _tid != TENANT_ID_PERMITIDO:
        st.error(
            "🚫 Acceso restringido a personal de Caja Arequipa. "
            "La cuenta con la que iniciaste sesión no pertenece a la organización."
        )
        st.button("Cerrar sesión", on_click=st.logout)
        st.stop()

# inicio
DEFAULTS = {
    "usuario": "",
    "view": "busqueda",
    "df": None,
    "hoja_usada": "",
    "cliente_actual": None,
    "visitas": {},
    "garantias": [],
    "rcc": [],
    "borrador_prompt": False,
    "ultimo_archivo": None,
    "cliente_visitado": "",

   # estado vista carga
    "archivo_meta": {},          # nombre, tamaño y fecha del Excel cargado
    "mostrar_preview": False,    # toggle de "Vista previa de datos"
    "uploader_key_version": 0,    # ver comentario en "btn_limpiar_base"
}
for k, v in DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v


# componentes
def header(icono, titulo, subtitulo="", icono_derecha=None):
    #Encabezado tipo app 
    extra = f'<div class="icon-box-derecha">{icono_derecha}</div>' if icono_derecha else ""
    st.markdown(
        f"""<div class="app-header">
                <div class="icon-box">{icono}</div>
                <div class="titles">
                    <h1>{titulo}</h1>
                    <p>{subtitulo}</p>
                </div>
                {extra}
            </div>""",
        unsafe_allow_html=True,
    )


def badge(texto, clase):
    st.markdown(f'<span class="badge {clase}">{texto}</span>', unsafe_allow_html=True)


PASOS = ["busqueda", "evaluacion", "ficha", "ubicacion", "reporte"]
PASOS_LABEL = {
    "busqueda": ("🔍", "Buscar"),
    "evaluacion": ("⚠️", "Criterio"),
    "ficha": ("👤", "Cliente"),
    "ubicacion": ("📍", "Visita"),
    "reporte": ("📄", "Reporte"),
}


def barra_usuario():
    """Chip con el nombre del usuario autenticado y botón de cerrar sesión."""
    if not _login_configurado() or not getattr(st.user, "is_logged_in", False):
        return
    nombre = st.user.get("name") or st.user.get("email") or "Usuario"
    correo = st.user.get("email") or ""
    c1, c2 = st.columns([3.2, 1])
    with c1:
        st.markdown(
            f"""<div class="user-chip">
                    <div class="avatar-circle">{iniciales(nombre)}</div>
                    <div>
                        <div class="u-name">{nombre}</div>
                        <div class="u-mail">{correo}</div>
                    </div>
                </div>""",
            unsafe_allow_html=True,
        )
    with c2:
        st.button("Salir", key="btn_logout_topbar", use_container_width=True,
                   on_click=st.logout)


def top_menu():
    """Barra de navegación.  "Resumen / Clientes / Visitas / Más" del mockup."""
    barra_usuario()
    st.markdown('<div class="top-menu-spacer"></div>', unsafe_allow_html=True)
    pasos_visibles = list(PASOS) if st.session_state.cliente_actual is not None else ["busqueda"]
    mostrar_consolidado = st.session_state.df is not None
    n = len(pasos_visibles) + (1 if mostrar_consolidado else 0)
    st.markdown('<div class="marker-bottomnav"></div>', unsafe_allow_html=True)
    cols = st.columns(n)
    for i, paso in enumerate(pasos_visibles):
        icono, label = PASOS_LABEL[paso]
        activo = " ✅" if st.session_state.view == paso else ""
        if cols[i].button(f"{icono} {label}", key=f"nav_{paso}", use_container_width=True,
                           help=f"Ir a {label}", type="primary" if st.session_state.view == paso else "secondary"):
            st.session_state.view = paso
            st.rerun()
    if mostrar_consolidado:
        if cols[-1].button("📊 Resumen", key="nav_consolidado", use_container_width=True,
                            help="Reporte consolidado por agencia y cliente",
                            type="primary" if st.session_state.view == "consolidado" else "secondary"):
            st.session_state.view = "consolidado"
            st.rerun()


def ir_a(paso):
    st.session_state.view = paso
    st.rerun()


def cliente():
    return st.session_state.cliente_actual or {}


def guardar_avance():
    c = cliente()
    if c:
        guardar_borrador(st.session_state.usuario, safe_str(c.get("DOCPEN")), c)


# PANTALLA 1 — BÚSQUEDA Y CARGA
# "Evaluación de Crédito" / Carga de Base de Datos, y "Buscar Cliente" / Búsqueda Inteligente).

def pantalla_busqueda():
    header("👤", "Ingrese su usuario", "Inserte su usuario para registrar la visita", icono_derecha="❓")

    #  Usuario avance/historial
    cu1, cu2 = st.columns([1, 2.2])
    with cu1:
        st.markdown('<div class="usuario-bar-label">👤 Auditor</div>', unsafe_allow_html=True)
    with cu2:
        usuario = st.text_input(
            "Tu nombre / usuario", value=st.session_state.usuario, key="input_usuario",
            placeholder="Ej: ASDL", label_visibility="collapsed",
        )
        st.session_state.usuario = usuario.strip()
    if not st.session_state.usuario:
        st.info("Escribe tu nombre de usuario para continuar.")
        return

    # carga bd
    with st.container(border=True):
        st.markdown("**Carga de Base de Datos**")
        st.caption("Carga tu archivo Excel con la cartera de clientes")

        if st.session_state.df is None:

            from utils.onedrive import (
                credenciales_configuradas as _od_ok,
                descargar_excel_base, ONEDRIVE_BASE_SHARE_URL,
            )
            if _od_ok():
                st.markdown("**☁️ Recomendado en celular**")
                if st.button("📥 Cargar base desde OneDrive", use_container_width=True,
                             type="primary", key="btn_cargar_onedrive"):
                    with st.spinner("Descargando desde OneDrive..."):
                        contenido, nombre_o_error, fecha = descargar_excel_base(ONEDRIVE_BASE_SHARE_URL)
                    if contenido is None:
                        st.error(f"No se pudo cargar desde OneDrive: {nombre_o_error}")
                    else:
                        df, hoja_usada, faltantes = cargar_excel(contenido)
                        st.session_state.df = df
                        st.session_state.hoja_usada = hoja_usada
                        st.session_state.archivo_meta = {
                            "nombre": nombre_o_error,
                            "tamano_mb": len(contenido) / (1024 * 1024),
                            "fecha": fecha or ahora_peru().strftime("%d/%m/%Y %I:%M %p"),
                        }
                        st.rerun()
                st.markdown(
                    '<div style="text-align:center;color:#8A97AC;font-size:0.78rem;'
                    'margin:0.5rem 0;">— o sube el archivo manualmente —</div>',
                    unsafe_allow_html=True,
                )

            archivo = st.file_uploader(
                "Seleccionar archivo Excel — Formatos permitidos: .xlsx, .xls",
                type=["xlsx", "xls"],
                key=f"file_uploader_main_{st.session_state.uploader_key_version}",
            )
            if archivo is not None:
                df, hoja_usada, faltantes = cargar_excel(archivo.getvalue())
                st.session_state.df = df
                st.session_state.hoja_usada = hoja_usada
                st.session_state.archivo_meta = {
                    "nombre": archivo.name,
                    "tamano_mb": len(archivo.getvalue()) / (1024 * 1024),
                    "fecha": ahora_peru().strftime("%d/%m/%Y %I:%M %p"),
                }
                if hoja_usada != "MUESTRA_FINAL":
                    st.warning("No se encontró la hoja 'MUESTRA_FINAL'; se usó la primera hoja del archivo.")
                if faltantes:
                    st.caption(
                        "Columnas no encontradas (quedarán vacías): "
                        + ", ".join(faltantes[:8]) + ("..." if len(faltantes) > 8 else "")
                    )
                st.rerun()
        else:
            df_actual = st.session_state.df
            meta = st.session_state.get("archivo_meta", {}) or {}
            st.markdown(
                f"""<div class="processed-panel">
                        <div class="processed-head"> Archivo procesado correctamente</div>
                        <div class="processed-file">
                            <span class="file-icon"></span>
                            <div>
                                <div class="file-name">{meta.get('nombre', 'archivo.xlsx')}</div>
                                <div class="file-meta">{meta.get('tamano_mb', 0):.1f} MB · {meta.get('fecha', '')}</div>
                            </div>
                        </div>
                        <div class="processed-count">{len(df_actual):,}</div>
                        <div class="processed-count-label">Registros cargados</div>
                    </div>""",
                unsafe_allow_html=True,
            )

            st.markdown('<div class="marker-vistaprevia"></div>', unsafe_allow_html=True)
            if st.button(" Vista previa de datos", use_container_width=True, key="btn_vista_previa"):
                st.session_state.mostrar_preview = not st.session_state.mostrar_preview
            if st.session_state.mostrar_preview:
                st.dataframe(df_actual.head(20), use_container_width=True, hide_index=True)
                st.caption(f"Mostrando 20 de {len(df_actual):,} registros.")

            st.markdown('<div class="marker-limpiar"></div>', unsafe_allow_html=True)
            if st.button("🗑️ Limpiar base — esto eliminará la base actual cargada",
                         use_container_width=True, key="btn_limpiar_base"):
                st.session_state.df = None
                st.session_state.hoja_usada = ""
                st.session_state.archivo_meta = {}
                st.session_state.mostrar_preview = False
                st.session_state.uploader_key_version += 1
                st.rerun()

    # estado del sistema
    with st.container(border=True):
        st.markdown("**Estado del Sistema**")
        conectado = st.session_state.df is not None
        st.markdown(
            f'<span class="estado-sistema-dot">{"🟢" if conectado else "⚪"}</span> '
            f'**{"Conectado" if conectado else "Sin datos cargados"}**',
            unsafe_allow_html=True,
        )
        meta = st.session_state.get("archivo_meta", {}) or {}
        if meta:
            st.caption(f"Última actualización: {meta.get('fecha', '')}")

    df = st.session_state.df
    if df is None:
        st.info("Sube el archivo Excel para poder buscar clientes.")
        return

    # busqueda
    with st.container(border=True):
        st.markdown("**🔎 Búsqueda de Cliente**")
        st.caption("Encuentra a tu cliente por nombre, DNI o código")

        c_busq, c_scan = st.columns([2.6, 1])
        with c_busq:
            busqueda = st.text_input(
                "Buscar", placeholder="Nombre, DNI o código de cliente",
                label_visibility="collapsed", key="input_busqueda",
            )
        with c_scan:
            st.markdown('<div class="marker-buscar"></div>', unsafe_allow_html=True)
            st.button("Buscar", use_container_width=True, key="btn_buscar")

    if not busqueda:
        return

    st.markdown(
        '<div class="live-indicator"><span class="live-dot"></span>'
        'Buscando coincidencias en tiempo real...</div>',
        unsafe_allow_html=True,
    )

    b = busqueda.strip().lower()
    b_digitos = solo_digitos(busqueda)
    mask = (
        df.get("DOCPEN", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        | df.get("BCCTA", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        | df.get("CODCLI", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
        | df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
    )
    resultados = df[mask].head(10)

    # (DNI completo o nombre completo iguales) 
    exacto = None
    for _, row in resultados.iterrows():
        dni_row = solo_digitos(row.get("DOCPEN"))
        nombre_row = safe_str(row.get("CLIENTE")).strip().lower()
        if (b_digitos and dni_row == b_digitos) or (b and nombre_row == b):
            exacto = row
            break
    if exacto is None and len(resultados) == 1:
        exacto = resultados.iloc[0]

    if exacto is not None:
        render_cliente_encontrado(exacto, df)
    elif len(resultados):
        st.markdown("**No encontramos coincidencias exactas**")
        st.caption("Te mostramos clientes similares de tu agencia")
        render_lista_similares(resultados.head(5))
    else:
        st.warning("No se encontraron coincidencias.")

    st.markdown(
        """<div class="tip-box">
                💡 <b>Consejo</b><br>
                Puedes buscar por nombre completo, parcial, DNI o código de
                cliente.
            </div>""",
        unsafe_allow_html=True,
    )


def render_cliente_encontrado(row, df):
    """Tarjeta compacta: nombre, DNI, dirección y agencia,
     botón Confirmar este cliente."""
    nombre   = safe_str(row.get("CLIENTE"), "Sin nombre")
    dni      = safe_str(row.get("DOCPEN"), "-")
    agencia  = safe_str(row.get("AGENCIA"), "-")
    direccion = safe_str(row.get("DIRECCION_DOM")) or safe_str(row.get("DIRECCION_NEG"), "No registrada")

    st.markdown(
        f"""<div class="cliente-card">
                <div class="cliente-card-top">
                    <div class="avatar-circle">{iniciales(nombre)}</div>
                    <div class="cliente-nombre-wrap">
                        <div class="cliente-nombre">{nombre}</div>
                    </div>
                    <span class="badge-encontrado">Cliente encontrado</span>
                </div>
                <div style="padding:0.5rem 0.2rem;font-size:0.85rem;color:#475569;">
                    <div class="dato-fila"> <b>DNI:</b> {dni}</div>
                    <div class="dato-fila"> <b>Dirección:</b> {direccion}</div>
                    <div class="dato-fila"> <b>Agencia:</b> {agencia}</div>
                </div>
            </div>""",
        unsafe_allow_html=True,
    )

    st.markdown('<div class="marker-confirmar"></div>', unsafe_allow_html=True)
    if st.button("Confirmar este cliente", use_container_width=True,
                 key="btn_confirmar_cliente", type="primary"):
        seleccionar_cliente(row.to_dict())
    st.caption(f"Continuar con la evaluación de {nombre}")

    similares = clientes_similares(df, row)
    if len(similares):
        st.write("")
        st.markdown("**¿No es este cliente?**")
        render_lista_similares(similares)


def render_lista_similares(resultados):
    """lista de clientes parecidos por busqueda"""
    for idx, row in resultados.iterrows():
        nombre = safe_str(row.get("CLIENTE"), "Sin nombre")
        st.markdown(
            f"""<div class="similar-item">
                    <div class="avatar-circle avatar-chica">{iniciales(nombre)}</div>
                    <div class="similar-info">
                        <div class="similar-nombre">{nombre}</div>
                        <div class="similar-dni">DNI: {safe_str(row.get('DOCPEN'), '-')}</div>
                    </div>
                    <div class="similar-saldo">{fmt_money(row.get('SALDO_MN'))}</div>
                </div>""",
            unsafe_allow_html=True,
        )
        st.markdown('<div class="marker-abrir-similar"></div>', unsafe_allow_html=True)
        if st.button("Abrir →", key=f"abrir_similar_{idx}", use_container_width=True):
            seleccionar_cliente(row.to_dict())


def seleccionar_cliente(fila):
    st.session_state.cliente_actual = fila
    dni = safe_str(fila.get("DOCPEN"))
    if hay_borrador(st.session_state.usuario, dni):
        st.session_state.borrador_prompt = True
    else:
        st.session_state.visitas = {}
        st.session_state.garantias = []
        st.session_state.rcc = []
        st.session_state.cliente_visitado = ""

        ir_a("evaluacion")
    st.rerun()


def prompt_borrador():
    c = cliente()
    with st.container(border=True):
        st.warning(f"Encontramos un avance guardado para **{safe_str(c.get('CLIENTE'))}** con tu usuario.")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("🔄 Continuar avance", use_container_width=True):
                cargar_borrador(st.session_state.usuario, safe_str(c.get("DOCPEN")))
                st.session_state.borrador_prompt = False
                ir_a("evaluacion")
        with c2:
            if st.button("🆕 Iniciar nuevo", use_container_width=True):
                borrar_borrador(st.session_state.usuario, safe_str(c.get("DOCPEN")))
                st.session_state.visitas = {}
                st.session_state.garantias = []
                st.session_state.rcc = []
                st.session_state.cliente_visitado = ""
        
                st.session_state.borrador_prompt = False
                ir_a("evaluacion")


# visita criterios vista2
def pantalla_evaluacion():
    c = cliente()
    header("⚠️", "Evaluación de Crédito", f"Cliente: {safe_str(c.get('CLIENTE'))}")
    st.caption("Marca los criterios identificados para esta visita.")

    for categoria, items in CRITERIOS_DEF.items():
        keys = [f"chk_{slug(categoria)}_{slug(item)}" for item in items]
        activo = any(st.session_state.get(k, False) for k in keys)
        icono = "🔴" if activo else "⚪"
        with st.container(border=True):
            with st.expander(f"{icono} {categoria}", expanded=activo):
                for item, key in zip(items, keys):
                    st.checkbox(item, key=key)
                    if item == "Calificación diferente a normal" and st.session_state.get(key):
                        st.text_input("Indicar la calificación a la fecha de revisión", key="calif_revision")

    n_marcados = sum(
        1 for cat, items in CRITERIOS_DEF.items() for item in items
        if st.session_state.get(f"chk_{slug(cat)}_{slug(item)}", False)
    )
    if n_marcados:
        badge(f"⚠️ {n_marcados} criterio(s) marcado(s)", "badge-pend")
    else:
        badge("Sin criterios de riesgo marcados", "badge-ok")

    st.write("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Volver a buscar", use_container_width=True):
            ir_a("busqueda")
    with c2:
        if st.button("Guardar y continuar ➡️", use_container_width=True, type="primary"):
            # Guardar snapshot explícito de los criterios marcados para que
            # la vista Reporte los pueda leer 
            st.session_state["_criterios_snapshot"] = {
                k: v for k, v in st.session_state.items() if k.startswith("chk_")
            }
            guardar_avance()
            ir_a("ficha")


# ficha cliente
def pantalla_ficha():
    c = cliente()
    header("👤", "Cliente y Crédito", "Ficha de identidad (solo lectura)")

    st.markdown(
        f"""<div class="banner-cliente">
                <div class="nombre">{safe_str(c.get('CLIENTE'))}</div>
                <div class="dni">DNI: {safe_str(c.get('DOCPEN'))} · Cuenta: {safe_str(c.get('BCCTA'))}</div>
            </div>""",
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown("**Información del crédito**")
        chips = [
            ("N° de cuenta", safe_str(c.get("BCCTA"), "-")),
            ("Tipo de crédito", safe_str(c.get("PRODUCTO_CAJA"), "-")),
            ("Calificación", safe_str(c.get("CATEG_RESULTANTE"), "-")),
            ("Importe desembolsado", fmt_money(c.get("IMPDESEMB_MN"))),
            ("Saldo actual", fmt_money(c.get("SALDO_MN"))),
            ("Fecha de desembolso", safe_str(c.get("FECDES"), "-")),
            ("Último pago", safe_str(c.get("FECHA_UTLPAGO"), "-")),
        ]
        chips_html = "".join(
            f'<div class="chip"><div class="lbl">{lbl}</div><div class="val">{val}</div></div>'
            for lbl, val in chips
        )
        st.markdown(f'<div class="info-credito">{chips_html}</div>', unsafe_allow_html=True)

        imp = safe_float(c.get("IMPDESEMB_MN"))
        saldo = safe_float(c.get("SALDO_MN"))
        if imp > 0:
            usado_pct = max(0.0, min(1.0, 1 - (saldo / imp)))
            st.progress(usado_pct, text=f"{usado_pct*100:.0f}% pagado del importe original")

    with st.expander("ℹ️ Información adicional"):
        info = [
            ("Agencia", c.get("AGENCIA")), ("Analista vigente", c.get("ANALISTA")),
            ("Analista evaluador", c.get("ANALISTA_EVAL")), ("Tipo SBS", c.get("TIPO_SBS")),
            ("Rubro / Actividad", c.get("ACTIVIDAD_ECON")), ("Segmentación MYPE", c.get("SEGMENTACION_MYPE")),
            ("Cuenta aval", c.get("CUENTA_AVAL")), ("Estado del crédito", c.get("ESTADO_CREDITO")),
        ]
        for label, val in info:
            st.write(f"**{label}:** {safe_str(val, '-')}")

    st.write("")
    st.markdown('<div class="nav-pie">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Criterio", use_container_width=True):
            ir_a("evaluacion")
    with c2:
        if st.button("Ir a la visita ➡️", use_container_width=True, type="primary"):
            guardar_avance()
            ir_a("ubicacion")
    st.markdown('</div>', unsafe_allow_html=True)


# vista ubicacion
TIPOS_VISITA = {
    "negocio": ("💼", "Negocio", "DIRECCION_NEG", "DISTRITO_NEG", "PROVINCIA_NEG", "DEPARTAMENTO_NEG", True),
    "laboral": ("🏢", "Centro laboral", None, None, None, None, False),
    "aval": ("🧾", "Aval", None, None, None, None, False),
    "domicilio": ("🏠", "Domicilio", "DIRECCION_DOM", "DISTRITO_DOM", "PROVINCIA_DOM", "DEPARTAMENTO_DOM", False),
}


def pantalla_ubicacion():
    c = cliente()
    header("📍", "Nueva Visita", "Verificación: negocio (obligatorio), laboral, aval y domicilio (opcionales)")

    tabs = st.tabs([f"{TIPOS_VISITA[t][0]} {TIPOS_VISITA[t][1]}" for t in TIPOS_VISITA])
    for tab, clave in zip(tabs, TIPOS_VISITA):
        with tab:
            render_visita(clave, c)

    st.write("")
    st.markdown('<div class="nav-pie">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Cliente", use_container_width=True, key="back_ubic"):
            ir_a("ficha")
    with c2:
        if st.button("Ir al reporte ➡️", use_container_width=True, type="primary", key="next_ubic"):
            guardar_avance()
            ir_a("reporte")
    st.markdown('</div>', unsafe_allow_html=True)


def render_visita(clave, c):
    icono, etiqueta, k_dir, k_dist, k_prov, k_depto, obligatoria = TIPOS_VISITA[clave]
    visitas = st.session_state.visitas
    data = visitas.get(clave, {})
    
    # Solicitar GPS automáticamente al entrar
    if f"gps_iniciado_{clave}" not in st.session_state:
         st.session_state[f"gps_iniciado_{clave}"] = True
         st.session_state[f"solicitar_gps_{clave}"] = True

    with st.container(border=True):
        # PASO 3 — Datos del lugar (primero)
        st.markdown("**Paso 3 · Datos del lugar**")
        valor_dir = data.get("direccion") or (safe_str(c.get(k_dir)) if k_dir else "")
        direccion = st.text_input("Dirección", value=valor_dir, key=f"dir_{clave}")
        cc1, cc2 = st.columns(2)
        with cc1:
            valor_dist = data.get("distrito") or (safe_str(c.get(k_dist)) if k_dist else "")
            distrito = st.text_input("Distrito", value=valor_dist, key=f"dist_{clave}")
            valor_prov = data.get("provincia") or (safe_str(c.get(k_prov)) if k_prov else "")
            provincia = st.text_input("Provincia", value=valor_prov, key=f"prov_{clave}")
        with cc2:
            valor_depto = data.get("departamento") or (safe_str(c.get(k_depto)) if k_depto else "")
            departamento = st.text_input("Departamento", value=valor_depto, key=f"depto_{clave}")
            referencia = st.text_input("Referencia", value=data.get("referencia", ""), key=f"ref_{clave}")

        # PASO 4 — Entrevista, comentarios y cliente visitado
        st.markdown("**Paso 4 · Observaciones**")
        ahora_v = ahora_peru()
        entrevista_con = st.text_input("Entrevista con", value=data.get("entrevista_con", ""), key=f"entrevista_{clave}")
        comentarios = st.text_area("Comentarios", value=data.get("comentarios", ""), key=f"comentarios_{clave}")

        # cliente visitado 
        if clave == "negocio":
            st.markdown("**Cliente visitado**")
            opciones_cv = ["— Selecciona —"] + CLIENTE_VISITADO_OPCIONES
            idx_cv = 0
            if st.session_state.get("cliente_visitado") in CLIENTE_VISITADO_OPCIONES:
                idx_cv = CLIENTE_VISITADO_OPCIONES.index(st.session_state["cliente_visitado"]) + 1
            opcion_cv = st.selectbox(
                "Resultado de la visita", opciones_cv, index=idx_cv,
                key=f"cv_{clave}", label_visibility="collapsed",
            )
            if opcion_cv in CLIENTE_VISITADO_OPCIONES:
                st.session_state["cliente_visitado"] = opcion_cv

        # FOTO + GPS AUTOMÁTICO
        etiqueta_foto = "Foto de verificación (obligatoria)" if obligatoria else "Foto de verificación (opcional)"
        st.markdown(f"**📷 {etiqueta_foto}**")
        st.caption("La ubicación GPS se solicita automáticamente al entrar. Tome también la fotografía de verificación")

        foto_final = st.camera_input(
        "📷 Tomar foto de verificación",
         key=f"camara_{clave}"
         )
        if foto_final is None and data.get("foto_bytes"):
            st.image(data["foto_bytes"], caption="Foto guardada previamente", width=200)
        if obligatoria and foto_final is None and not data.get("foto_bytes"):
            st.warning("⚠ Esta sección requiere foto de verificación antes de guardar.")

        # GPS — se activa automáticamente en cuanto hay foto
        lat = st.session_state.get(
            f"lat_{clave}",
            data.get("lat")
        )

        lon = st.session_state.get(
            f"lon_{clave}",
            data.get("lon")
        )

        precision = st.session_state.get(
            f"precision_{clave}",
            data.get("precision")
        )
        # La ubicación se solicita automáticamente al entrar a la vista,
        # igual que la cámara se activa sola al mostrarse.
        if st.session_state.get(f"solicitar_gps_{clave}") and not (lat and lon):
            try:
                from streamlit_js_eval import get_geolocation
                # OJO: el parámetro correcto es "component_key", NO "key".
                loc = get_geolocation(component_key=f"geo_{clave}")

                if loc and "coords" in loc:
                    # Éxito: el usuario aceptó el permiso y llegaron las coordenadas
                    lat = loc["coords"]["latitude"]
                    lon = loc["coords"]["longitude"]
                    precision = loc["coords"].get("accuracy")

                    st.session_state[f"lat_{clave}"] = lat
                    st.session_state[f"lon_{clave}"] = lon
                    st.session_state[f"precision_{clave}"] = precision
                    st.session_state[f"solicitar_gps_{clave}"] = False

                elif loc and "error" in loc:
                    # El navegador respondió, pero con un error (permiso denegado,
                    # posición no disponible, timeout, o geolocalización no soportada)
                    codigo = loc["error"].get("code")
                    mensaje = loc["error"].get("message", "Error desconocido")
                    explicacion = {
                        1: "Permiso de ubicación denegado. Debes permitirlo en la "
                           "configuración del sitio en tu navegador (icono de "
                           "candado/ubicación junto a la barra de direcciones).",
                        2: "La posición no está disponible. Verifica que el GPS del "
                           "dispositivo esté activado y que tengas señal.",
                        3: "Se agotó el tiempo de espera al intentar obtener la "
                           "ubicación. Intenta de nuevo en un lugar con mejor señal.",
                        0: "Este navegador no soporta geolocalización.",
                    }.get(codigo, mensaje)
                    st.warning(f"⚠️ No se pudo obtener la ubicación: {explicacion}")
                    st.session_state[f"solicitar_gps_{clave}"] = False

                else:
                    # loc es None: el componente todavía no ha respondido (primera
                    # carga). Streamlit volverá a ejecutar el script automáticamente
                    # en cuanto el navegador resuelva la promesa de JS.
                    st.info("📡 Obteniendo ubicación del dispositivo... acepta el permiso si el navegador lo solicita.")

            except Exception as e:
                st.warning(
                    f"⚠️ Error inesperado al solicitar la ubicación ({e}). "
                    "Presiona 'Reintentar ubicación' para intentar de nuevo."
                )
                st.session_state[f"solicitar_gps_{clave}"] = False

        st.markdown("**Ubicación:**")
        if lat and lon:
            ubicacion_txt = f"{lat:.6f}, {lon:.6f}"
            if precision:
                ubicacion_txt += f" (±{precision:.0f} m)"
            st.success(ubicacion_txt)
            st.map(
                pd.DataFrame({"lat": [lat], "lon": [lon]}),
                zoom=15,
                height=160
            )
        elif st.session_state.get(f"solicitar_gps_{clave}"):
             st.info("📡 Solicitando ubicación del dispositivo...")
        else:
           st.warning(
               "⚠️ No se obtuvo la ubicación GPS. Asegúrate de haber aceptado el permiso "
               "de ubicación del navegador y que el GPS esté activado, luego presiona "
               "'Reintentar ubicación'."
           )
           if st.button("📍 Reintentar ubicación", key=f"retry_gps_{clave}", use_container_width=True):
               st.session_state[f"solicitar_gps_{clave}"] = True
               st.rerun()

        hay_foto = (
            foto_final is not None
            or bool(data.get("foto_bytes"))
        )
        
        puede_guardar = (
            hay_foto
            and lat is not None
            and lon is not None
        
        )
        if st.button(f"💾 Guardar visita de {etiqueta}", key=f"guardar_{clave}",
                     use_container_width=True, type="primary", disabled=not puede_guardar):
            st.session_state.visitas[clave] = {
                "direccion": direccion, 
                "distrito": distrito,        
                "provincia": provincia,      
                "departamento": departamento,
                "referencia": referencia,    
                "fecha": ahora_v.strftime("%d/%m/%Y"), 
                "hora": ahora_v.strftime("%H:%M:%S"),
                "entrevista_con": entrevista_con, 
                "comentarios": comentarios,
                "lat": lat, "lon": lon, "precision": precision,
                "foto_bytes": foto_final.getvalue() if foto_final is not None else data.get("foto_bytes"),
            }
            guardar_avance()
            st.success(f"Visita de {etiqueta} guardada correctamente.")
        
        if not puede_guardar:
            st.caption("Se requiere foto y ubicación GPS para guardar la visita")

        if clave in visitas:
            badge("Registrada", "badge-ok")
        else:
            badge("Pendiente", "badge-pend")


# generar reporte
def pantalla_reporte():
    c = cliente()
    header("📄", "Generación de Reporte", "Revisión final y descarga del documento")

    visitas = st.session_state.visitas
    secciones = [("negocio", "Negocio", True), ("laboral", "Laboral", False),
                 ("aval", "Aval", False), ("domicilio", "Domicilio", False)]
    completas = sum(1 for k, _, _ in secciones if k in visitas)

    with st.container(border=True):
        st.markdown(f"**Resumen de calidad** — {completas} de {len(secciones)} visitas registradas")
        cols = st.columns(4)
        for col, (clave, etiqueta, obligatoria) in zip(cols, secciones):
            ok = clave in visitas
            badge_clase = "badge-ok" if ok else ("badge-pend" if obligatoria else "badge-warn")
            texto = "Foto capturada" if ok else ("Falta (obligatoria)" if obligatoria else "Opcional")
            col.markdown(
                f"""<div style="text-align:center;padding:0.5rem 0.2rem;border-radius:10px;background:{'#F0FDF4' if ok else '#FEF2F2'};">
                        <div style="font-size:1.3rem;">{'✅' if ok else ('⚠️' if obligatoria else '➖')}</div>
                        <div style="font-weight:700;font-size:0.8rem;">{etiqueta}</div>
                        <span class="badge {badge_clase}" style="font-size:0.65rem;">{texto}</span>
                    </div>""",
                unsafe_allow_html=True,
            )

    if "negocio" not in visitas:
        st.warning("Acción requerida — falta la visita obligatoria al **Negocio**. Puedes generar el reporte igual; quedará indicado como pendiente.")

    # leer criterios desde el snapshot guardado al avanzar de pantalla
    # session_state normal como respaldo
    criterios_dict = st.session_state.get("_criterios_snapshot") or \
                     {k: v for k, v in st.session_state.items() if k.startswith("chk_")}
    criterios_txt = criterios_seleccionados_lista(criterios_dict, st.session_state.get("calif_revision", ""))
    observacion_criterio = ""  # campo eliminado
    ing = {k: st.session_state.get(k, 0.0) for k in [
        "ingreso_principal", "otros_ingresos", "op_alquiler", "op_servicios", "op_transporte",
        "op_mercaderia", "op_publicidad", "op_otros", "fam_alimentacion", "fam_vivienda",
        "fam_servicios", "fam_educacion", "fam_salud", "fam_otros",
    ]}
    calc = calcular_resultado(ing)
    cliente_visitado = st.session_state.get("cliente_visitado", "")

    with st.container(border=True):
        st.markdown("**Resumen de la evaluación**")
        st.write(f"**Cuenta cliente:** {safe_str(c.get('BCCTA'))}")
        st.write(f"**N° de operación:** {safe_str(c.get('BCOPER'))}")
        st.write(f"**Nombre del cliente:** {safe_str(c.get('CLIENTE'))}")
        st.write(f"**Módulo:** {safe_str(c.get('MODULO'))}")
        st.write(f"**Analista vigente:** {safe_str(c.get('ANALISTA'))}")
        st.write(f"**Analista evaluador:** {safe_str(c.get('ANALISTA_EVAL'))}")
        st.write(f"**Auditor:** {st.session_state.usuario}")
        st.write(f"**Fecha de visita:** {ahora_peru().strftime('%d/%m/%Y %H:%M')} (hora Perú)")
        st.write(f"**Cliente visitado:** {cliente_visitado or '—'}")
        st.write(f"**Utilidad neta:** {fmt_money(calc['utilidad_neta'])}")

        st.markdown("---")
        st.markdown("**Criterio seleccionado**:")
        if criterios_txt:
            for ct in criterios_txt:
                st.markdown(
                    f'<div style="background:#FEF2F2;border-left:4px solid #C8102E;'
                    f'border-radius:6px;padding:0.3rem 0.6rem;margin:0.25rem 0;font-size:0.85rem;">'
                    f'✔️ {ct}</div>',
                    unsafe_allow_html=True,
                )
        else:
            st.caption("Ninguno marcado en la vista Criterio.")

    with st.container(border=True):
        st.markdown("**Generar y descargar reporte**")
        st.caption("Disponible en Word (.docx) y PDF. Se guarda automáticamente en la carpeta de reportes configurada.")

        base_nombre = f"Visita_{slug(c.get('CLIENTE'))}_{ahora_peru().strftime('%Y%m%d_%H%M')}"

        # pre-generar los bytes para Word y PDF 
        with st.spinner("Preparando archivos..."):
            buf_word = generar_word(c, criterios_txt, calc, ing, visitas,
                                    st.session_state.garantias, st.session_state.rcc,
                                    st.session_state.usuario, cliente_visitado, observacion_criterio)
            buf_pdf  = generar_pdf(c, criterios_txt, calc, ing, visitas,
                                   st.session_state.garantias, st.session_state.rcc,
                                   st.session_state.usuario, cliente_visitado, observacion_criterio)

        nombre_word = base_nombre + ".docx"
        nombre_pdf  = base_nombre + ".pdf"
        mime_word   = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        mime_pdf    = "application/pdf"

        c1, c2 = st.columns(2)
        with c1:
            descargado_word = st.download_button(
                "📝 Generar y descargar Word",
                data=buf_word.getvalue(), file_name=nombre_word, mime=mime_word,
                use_container_width=True, type="primary", key="dl_word",
            )
        with c2:
            descargado_pdf = st.download_button(
                "📕 Generar y descargar PDF",
                data=buf_pdf.getvalue(), file_name=nombre_pdf, mime=mime_pdf,
                use_container_width=True, type="primary", key="dl_pdf",
            )

        # registro historial
        if descargado_word:
            guardado = guardar_reporte_en_carpeta(nombre_word, buf_word.getvalue())
            sincronizar_historial_onedrive()
            n_ag, n_gen = registrar_historial(st.session_state.usuario, c, "Word", nombre_word,
                                               "; ".join(criterios_txt), cliente_visitado,
                                               guardado.get("online") or guardado.get("local"))
            st.session_state["ultimo_conteo"] = (n_ag, n_gen, guardado)

        if descargado_pdf:
            guardado = guardar_reporte_en_carpeta(nombre_pdf, buf_pdf.getvalue())
            sincronizar_historial_onedrive()
            n_ag, n_gen = registrar_historial(st.session_state.usuario, c, "PDF", nombre_pdf,
                                               "; ".join(criterios_txt), cliente_visitado,
                                               guardado.get("online") or guardado.get("local"))
            st.session_state["ultimo_conteo"] = (n_ag, n_gen, guardado)

        if st.session_state.get("ultimo_conteo"):
            n_ag, n_gen, guardado = st.session_state["ultimo_conteo"]
            agencia_txt = safe_str(c.get("AGENCIA"), "-")
            st.success(f"✅ Descargado. Visita N° {n_ag} en agencia **{agencia_txt}** (N° {n_gen} general).")
            if isinstance(guardado, dict):
                if guardado.get("online"):
                    st.markdown(f"☁️ **Subido a OneDrive:** [Abrir archivo]({guardado['online']})")
                if guardado.get("error"):
                    st.caption(f"⚠ {guardado['error']}")


    with st.expander("🗂️ Ver historial de reportes generados"):
        hist = leer_historial()
        if len(hist):
            st.dataframe(hist.tail(20), use_container_width=True, hide_index=True)
        else:
            st.caption("Aún no se ha generado ningún reporte.")

    st.write("")
    st.markdown('<div class="nav-pie">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("⬅️ Visita", use_container_width=True):
            ir_a("ubicacion")
    with c2:
        if st.button("🏁 Terminar y volver a buscar", use_container_width=True):
            c_dni = safe_str(c.get("DOCPEN"))
            borrar_borrador(st.session_state.usuario, c_dni)
            st.session_state.cliente_actual = None
            st.session_state.visitas = {}
            st.session_state.garantias = []
            st.session_state.rcc = []
            st.session_state.ultimo_archivo = None
            st.session_state.cliente_visitado = ""
    
            ir_a("busqueda")
    st.markdown('</div>', unsafe_allow_html=True)


# reporte consolidado agencia y cliente
def pantalla_consolidado():
    header("📊", "Reporte Consolidado", "Visitas realizadas por agencia y por cliente")
    st.caption(
        ""
    )

    resumen_agencia = reporte_consolidado_por_agencia()
    with st.container(border=True):
        st.markdown("**Por agencia** — clientes visitados y reportes generados")
        if len(resumen_agencia):
            st.dataframe(resumen_agencia, use_container_width=True, hide_index=True)
        else:
            st.caption("Aún no hay reportes generados para consolidar.")

    agencias_disponibles = ["(Todas)"] + (
        sorted(resumen_agencia["Agencia"].dropna().astype(str).unique().tolist())
        if len(resumen_agencia) else []
    )
    
    with st.container(border=True):
        st.markdown("**Detalle por cliente**")
        agencia_sel = st.selectbox("Filtrar por agencia", agencias_disponibles, key="sel_agencia_consolidado")
        filtro = None if agencia_sel == "(Todas)" else agencia_sel
        detalle = reporte_consolidado_por_cliente(filtro)
        
        if len(detalle):
            st.dataframe(detalle, use_container_width=True, hide_index=True)

            # descarga 
            if filtro:
                # Se filtró por una agencia 
                detalle_anexo07 = reporte_anexo07(filtro)
                excel_data = to_excel_anexo07(detalle_anexo07, filtro)
                nombre_archivo = f"ResultadoVisitas_{slug(agencia_sel)}_{datetime.now().strftime('%Y%m%d')}.xlsx"
            else:
                excel_data = to_excel(detalle)
                nombre_archivo = f"Reporte_Consolidado_Todas_{datetime.now().strftime('%Y%m%d')}.xlsx"

            st.download_button(
                label="📥 Descargar Detalle en Excel",
                data=excel_data,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.caption("Sin datos para este filtro todavía.")

    st.write("")
    if st.button("⬅️ Volver a buscar", use_container_width=True):
        ir_a("busqueda")


# router
top_menu()

if st.session_state.borrador_prompt:
    prompt_borrador()
elif st.session_state.view == "consolidado" and st.session_state.df is not None:
    pantalla_consolidado()
elif st.session_state.view == "busqueda" or st.session_state.cliente_actual is None:
    pantalla_busqueda()
elif st.session_state.view == "evaluacion":
    pantalla_evaluacion()
elif st.session_state.view == "ficha":
    pantalla_ficha()
elif st.session_state.view == "ubicacion":
    pantalla_ubicacion()
elif st.session_state.view == "reporte":
    pantalla_reporte()
else:
    pantalla_busqueda()
