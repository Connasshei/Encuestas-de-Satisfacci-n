import streamlit as st
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re

# ══════════════════════════════════════════════════════════
# CONFIGURACIÓN
# ══════════════════════════════════════════════════════════
st.set_page_config(
    page_title="Evaluación a Docentes",
    layout="wide",
)

# Colores para escala de 3 niveles
COLOR_MAP = {
    "Sí": "#1a9641",
    "A veces": "#f39c12",
    "No": "#d7191c",
}
SCALE = ["Sí", "A veces", "No"]

QUESTION_SHORT = {
    1: "Claridad programa y objetivos",
    2: "Organización de contenidos",
    3: "Bibliografía útil",
    4: "Puntualidad del docente",
    5: "Dominio de contenidos",
    6: "Claridad de explicaciones",
    7: "Estrategias de participación",
    8: "Participación estudiantil",
    9: "Relación con ejemplos prof.",
    10: "Atención de dudas",
    11: "Clima de respeto y confianza",
    12: "Orientaciones de contenido",
    13: "Coherencia de evaluaciones",
    14: "Explicación de criterios",
    15: "Retroalimentación",
    16: "Aplicación práctica",
    17: "Acceso y navegación AV",
    18: "Plataforma adecuada",
}

DIMENSIONS = [
    "Planificacion e inicio de la sesion",
    "Desarrollo didactico de los contenidos",
    "Interaccion y participacion estudiantil",
    "Uso pedagogico de recursos tecnologicos",
    "Evaluacion formativa y retroalimentacion",
    "Gestion del tiempo y clima de aprendizaje",
]

DIMENSION_COLORS = {
    DIMENSIONS[0]: "#4e79a7",
    DIMENSIONS[1]: "#f28e2b",
    DIMENSIONS[2]: "#e15759",
    DIMENSIONS[3]: "#76b7b2",
    DIMENSIONS[4]: "#59a14f",
    DIMENSIONS[5]: "#edc948",
}

DIMENSION_MAP = {
    1: [DIMENSIONS[0]],
    2: [DIMENSIONS[0]],
    3: [DIMENSIONS[0]],
    4: [DIMENSIONS[5]],
    5: [DIMENSIONS[1]],
    6: [DIMENSIONS[1]],
    7: [DIMENSIONS[2]],
    8: [DIMENSIONS[2]],
    9: [DIMENSIONS[5]],
    10: [DIMENSIONS[2]],
    11: [DIMENSIONS[5]],
    12: [DIMENSIONS[1]],
    13: [DIMENSIONS[4]],
    14: [DIMENSIONS[4]],
    15: [DIMENSIONS[4]],
    16: [DIMENSIONS[1]],
    17: [DIMENSIONS[3]],
    18: [DIMENSIONS[3]],
}

# ══════════════════════════════════════════════════════════
# PARSING DEL EXCEL
# ══════════════════════════════════════════════════════════


def extract_meta_value(cell_val: str, prefix: str) -> str:
    """Extrae el valor después del prefijo en celdas de metadatos."""
    if cell_val and cell_val.startswith(prefix):
        return cell_val[len(prefix) :].strip()
    return ""


def parse_excel(file_bytes: bytes, filename: str) -> dict:
    """Parsea un Excel de feedback Moodle y devuelve un dict estructurado."""
    wb = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    ws = wb.active

    rows = list(ws.iter_rows(values_only=True))

    meta = {
        "carrera": "",
        "sede": "",
        "modulo": "",
        "grupo": "",
        "docente": "",
        "fecha_inicio": "",
        "fecha_fin": "",
        "respuestas": 0,
        "archivo": filename,
    }
    questions = []
    comments = []

    in_questions = False
    in_comments = False

    for row in rows:
        first = str(row[0] or "").strip()

        # ── Metadatos ──────────────────────────
        if first.startswith("CARRERA:"):
            meta["carrera"] = extract_meta_value(first, "CARRERA:")
        elif first.startswith("SEDE:"):
            meta["sede"] = extract_meta_value(first, "SEDE:")
        elif first.startswith("MODULO:"):
            meta["modulo"] = extract_meta_value(first, "MODULO:")
        elif first.startswith("GRUPO:"):
            meta["grupo"] = extract_meta_value(first, "GRUPO:")
        elif first.startswith("DOCENTE:"):
            meta["docente"] = extract_meta_value(first, "DOCENTE:")
        elif first.startswith("FECHA INICIO:"):
            meta["fecha_inicio"] = extract_meta_value(first, "FECHA INICIO:")
        elif first.startswith("FECHA FIN:"):
            meta["fecha_fin"] = extract_meta_value(first, "FECHA FIN:")
        elif first.startswith("RESPUESTAS ENVIADAS:"):
            try:
                meta["respuestas"] = int(
                    extract_meta_value(first, "RESPUESTAS ENVIADAS:")
                )
            except:
                pass

        # ── Encabezado de tabla ─────────────────
        elif first == "Etiqueta":
            in_questions = True
            continue

        # ── Filas de preguntas ──────────────────
        elif in_questions and not in_comments:
            pregunta = str(row[1] or "").strip()
            if "observaciones" in pregunta.lower():
                in_questions = False
                in_comments = True
                continue

            if pregunta and len(pregunta) > 5:
                num_match = re.match(r"^(\d+)\.", pregunta)
                num = int(num_match.group(1)) if num_match else len(questions) + 1

                def safe(val):
                    try:
                        return float(val or 0)
                    except:
                        return 0.0

                questions.append(
                    {
                        "num": num,
                        "pregunta": pregunta,
                        "short": QUESTION_SHORT.get(num, f"P{num}"),
                        "dimensiones": DIMENSION_MAP.get(num, []),
                        "Sí": safe(row[2]),
                        "pct_Sí": round(safe(row[3]) * 100, 1),
                        "A veces": safe(row[4]),
                        "pct_A veces": round(safe(row[5]) * 100, 1),
                        "No": safe(row[6]),
                        "pct_No": round(safe(row[7]) * 100, 1),
                    }
                )

        # ── Comentarios ──────
        elif in_comments:
            for cell in row[2:]:
                val = str(cell or "").strip()
                if val and val.lower() not in ("ninguno", "ninguna", "n/a", ""):
                    comments.append(val)

    # Preservar el valor original del Excel (puede estar vacío)
    meta["docente_excel"] = meta["docente"]

    # Fallback: usar nombre de archivo como título si no hay docente
    if not meta["docente"]:
        name_part = filename.replace("feedback_ENCUESTA_ROL_FACILITADOR_--_", "")
        name_part = name_part.replace(".xlsx", "").replace("_", " ")
        meta["docente"] = name_part.strip()

    # Puntaje ponderado (Sí=3, A veces=2, No=1)
    for q in questions:
        total = q["Sí"] + q["A veces"] + q["No"]
        if total > 0:
            score = (q["Sí"] * 3 + q["A veces"] * 2 + q["No"] * 1) / total
            q["score"] = round(score, 2)
        else:
            q["score"] = 0.0

    return {"meta": meta, "questions": questions, "comments": comments}


def label_for(entry: dict) -> str:
    """Etiqueta corta para un archivo cargado."""
    m = entry["meta"]
    base = m["modulo"] if m["modulo"] else m["archivo"]
    return base[:60]


# ══════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════
if "data" not in st.session_state:
    st.session_state.data = {}  # filename -> parsed dict

# ══════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## Evaluación a Docentes")
    st.markdown("---")
    page = st.radio(
        "Navegación",
        ["Cargar Datos", "Ver Resultados"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    n_arch = len(st.session_state.data)
    n_resp = sum(d["meta"]["respuestas"] for d in st.session_state.data.values())
    st.metric("Archivos cargados", n_arch)
    st.metric("Total encuestados", n_resp)

# ══════════════════════════════════════════════════════════
# PÁGINA 1 — CARGA
# ══════════════════════════════════════════════════════════
if page == "Cargar Datos":
    st.title("Carga de Archivos de Encuesta")
    st.markdown(
        "Sube uno o varios archivos **Excel (.xlsx)** exportados desde Moodle. "
        "Cada archivo corresponde a una asignatura/módulo."
    )

    uploaded = st.file_uploader(
        "Selecciona los archivos Excel:",
        type=["xlsx"],
        accept_multiple_files=True,
        help="Ctrl+clic o Cmd+clic para seleccionar varios a la vez.",
    )

    if uploaded:
        nuevos, errores = 0, []
        for f in uploaded:
            if f.name not in st.session_state.data:
                try:
                    parsed = parse_excel(f.read(), f.name)
                    st.session_state.data[f.name] = parsed
                    nuevos += 1
                except Exception as e:
                    errores.append((f.name, str(e)))

        if nuevos:
            st.success(f"{nuevos} archivo(s) cargado(s) correctamente.")
        for fname, err in errores:
            st.error(f"Error en **{fname}**: {err}")

    # Lista de archivos cargados
    if st.session_state.data:
        st.markdown("### Archivos en sesión")

        for fname, entry in list(st.session_state.data.items()):
            m = entry["meta"]
            with st.container(border=True):
                col1, col2 = st.columns([8, 1])
                with col1:
                    st.markdown(f"**{m['modulo'] or fname}**")
                    cols_meta = st.columns(4)
                    cols_meta[0].caption(m["docente"] or "\u2014")
                    cols_meta[1].caption(m["carrera"] or "\u2014")
                    cols_meta[2].caption(m["sede"] or "\u2014")
                    cols_meta[3].caption(f"{m['respuestas']} respuestas")
                if col2.button("Borrar", key=f"del_{fname}"):
                    del st.session_state.data[fname]
                    st.rerun()

        st.markdown("---")
        colA, colB = st.columns([5, 2])
        colA.info(
            f"**{n_resp} encuestados totales** en **{n_arch} modulo(s)**. Vaya a **Ver Resultados**."
        )
        if colB.button("Limpiar todo", type="secondary"):
            st.session_state.data = {}
            st.rerun()
    else:
        st.info("Aun no hay datos. Suba al menos un archivo Excel para comenzar.")

# ══════════════════════════════════════════════════════════
# PÁGINA 2 — RESULTADOS
# ══════════════════════════════════════════════════════════
elif page == "Ver Resultados":
    st.title("Resultados de Evaluacion del Docente")

    if not st.session_state.data:
        st.warning("No hay datos cargados. Vaya primero a **Cargar Datos**.")
        st.stop()

    all_entries = list(st.session_state.data.values())

    # ── Selector de módulo ────────────────────
    opciones = {"Comparar todos los módulos": None}
    for fname, entry in st.session_state.data.items():
        opciones[label_for(entry)] = fname

    sel_label = st.selectbox("Modulo a visualizar:", list(opciones.keys()))
    sel_fname = opciones[sel_label]

    if sel_fname is not None:
        # ─── VISTA INDIVIDUAL ──────────────────────
        entry = st.session_state.data[sel_fname]
        meta = entry["meta"]
        qs = entry["questions"]
        comments = entry["comments"]

        # Ficha del módulo
        with st.container(border=True):
            guion = "\u2014"
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.markdown("**Modulo**")
                st.markdown(
                    f"<div style='word-break:break-word;white-space:normal;line-height:1.4'>{meta['modulo'] or guion}</div>",
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown("**Titulo archivo**")
                st.markdown(
                    f"<div style='word-break:break-word;white-space:normal;line-height:1.4'>{meta['archivo']}</div>",
                    unsafe_allow_html=True,
                )
            with c3:
                st.markdown("**Carrera**")
                st.markdown(
                    f"<div style='word-break:break-word;white-space:normal;line-height:1.4'>{meta['carrera'] or guion}</div>",
                    unsafe_allow_html=True,
                )
            with c4:
                st.metric("Respuestas enviadas", meta["respuestas"])
            c5, c6, c7 = st.columns(3)
            c5.caption(f"Inicio: {meta['fecha_inicio']}")
            c6.caption(f"Fin: {meta['fecha_fin']}")
            if meta.get("docente_excel"):
                c7.markdown(f"**Docente:** {meta['docente_excel']}")

        st.divider()

        # ── KPIs generales ────────────────────
        st.subheader("Indicadores Generales")
        total_resp = meta["respuestas"] or 1

        total_si = sum(q["Sí"] for q in qs)
        total_aveces = sum(q["A veces"] for q in qs)
        total_no = sum(q["No"] for q in qs)
        grand_total = total_si + total_aveces + total_no or 1

        avg_score = sum(q["score"] for q in qs) / len(qs)
        pct_si = total_si / grand_total * 100
        pct_aveces = total_aveces / grand_total * 100

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Puntaje promedio", f"{avg_score:.2f} / 3.00")
        k2.metric("Sí", f"{pct_si:.1f}%")
        k3.metric("A veces", f"{pct_aveces:.1f}%")
        k4.metric("No", f"{total_no / grand_total * 100:.1f}%")

        st.divider()

        # ── Resumen por Dimensión ──────────────
        st.subheader("Resumen por Dimension")
        dim_rows = []
        for dim in DIMENSIONS:
            dim_qs = [q for q in qs if dim in q.get("dimensiones", [])]
            if dim_qs:
                avg_dim = sum(q["score"] for q in dim_qs) / len(dim_qs)
                dim_rows.append(
                    {"Dimension": dim, "Puntaje promedio": round(avg_dim, 2)}
                )

        if dim_rows:
            df_dim = pd.DataFrame(dim_rows)
            cols_dim = st.columns(len(DIMENSIONS))
            for i, row in enumerate(dim_rows):
                cols_dim[i].metric(
                    row["Dimension"][:25],
                    row["Puntaje promedio"],
                )

            fig_dim = px.bar(
                df_dim,
                x="Dimension",
                y="Puntaje promedio",
                color="Dimension",
                color_discrete_map=DIMENSION_COLORS,
                text="Puntaje promedio",
                height=380,
            )
            fig_dim.update_traces(texttemplate="%{text:.2f}", textposition="outside")
            fig_dim.update_layout(
                yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
                showlegend=False,
                xaxis_tickangle=-20,
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_dim, use_container_width=True)

        st.divider()

        # ── Tabla resumen ─────────────────────
        st.subheader("Resumen por Pregunta")
        rows_sum = []
        for q in qs:
            dim_label = q.get("dimensiones", ["—"])[0] if q.get("dimensiones") else "—"
            rows_sum.append(
                {
                    "Dimension": dim_label,
                    "Pregunta": q["short"],
                    "Sí": f"{q['pct_Sí']:.0f}%",
                    "A veces": f"{q['pct_A veces']:.0f}%",
                    "No": f"{q['pct_No']:.0f}%",
                    "Puntaje (/3)": q["score"],
                }
            )
        df_sum = pd.DataFrame(rows_sum)
        st.dataframe(
            df_sum.style.background_gradient(
                subset=["Puntaje (/3)"], cmap="RdYlGn", vmin=1, vmax=3
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.divider()

        # ── Barras 100% apiladas ──────────────
        st.subheader("Distribucion de Respuestas por Pregunta")

        fig_stack = go.Figure()
        for cat in SCALE:
            pct_col = f"pct_{cat}"
            fig_stack.add_trace(
                go.Bar(
                    name=cat,
                    y=[q["short"] for q in qs],
                    x=[q[pct_col] for q in qs],
                    orientation="h",
                    marker_color=COLOR_MAP[cat],
                    text=[f"{q[pct_col]:.0f}%" if q[pct_col] > 3 else "" for q in qs],
                    textposition="inside",
                    insidetextanchor="middle",
                )
            )
        fig_stack.update_layout(
            barmode="stack",
            xaxis=dict(title="% de respuestas", ticksuffix="%", range=[0, 100]),
            height=400,
            legend=dict(
                orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1
            ),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=180, r=20, t=60, b=40),
        )
        st.plotly_chart(fig_stack, use_container_width=True)

        # ── Barras de puntaje promedio ─────────
        st.subheader("Puntaje Ponderado por Pregunta (escala 1-3)")
        df_score = pd.DataFrame(
            {
                "Pregunta": [q["short"] for q in qs],
                "Puntaje": [q["score"] for q in qs],
                "Dimension": [
                    q.get("dimensiones", ["—"])[0] if q.get("dimensiones") else "—"
                    for q in qs
                ],
            }
        )
        fig_score = px.bar(
            df_score,
            x="Puntaje",
            y="Pregunta",
            orientation="h",
            color="Dimension",
            color_discrete_map=DIMENSION_COLORS,
            text="Puntaje",
            height=380,
        )
        fig_score.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig_score.update_layout(
            xaxis=dict(range=[0, 3.5], title="Puntaje promedio"),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
            margin=dict(l=180, r=80, t=20, b=40),
        )
        st.plotly_chart(fig_score, use_container_width=True)

        st.divider()

        # ── Gráficos de torta ─────────────────
        st.subheader("Detalle por Pregunta")
        for i in range(0, len(qs), 2):
            cols = st.columns(2)
            for j, q in enumerate(qs[i : i + 2]):
                labels = [s for s in SCALE if q[s] > 0]
                values = [q[s] for s in SCALE if q[s] > 0]
                fig_pie = px.pie(
                    names=labels,
                    values=values,
                    title=q["short"],
                    color=labels,
                    color_discrete_map=COLOR_MAP,
                    hole=0.42,
                    height=290,
                )
                fig_pie.update_traces(textinfo="percent+label")
                fig_pie.update_layout(
                    showlegend=False,
                    margin=dict(t=50, b=10, l=10, r=10),
                    paper_bgcolor="rgba(0,0,0,0)",
                )
                dim_label = (
                    q.get("dimensiones", ["—"])[0] if q.get("dimensiones") else ""
                )
                cols[j].plotly_chart(fig_pie, use_container_width=True)
                cols[j].caption(f"{q['pregunta']}  |  _{dim_label}_")

        st.divider()

        # ── Radar ─────────────────────────────
        st.subheader("Perfil del Docente")
        labels_r = [q["short"] for q in qs]
        scores_r = [q["score"] for q in qs]
        fig_rad = go.Figure(
            go.Scatterpolar(
                r=scores_r + [scores_r[0]],
                theta=labels_r + [labels_r[0]],
                fill="toself",
                fillcolor="rgba(26, 150, 65, 0.2)",
                line=dict(color="#1a9641", width=2),
                name="Puntaje",
            )
        )
        fig_rad.update_layout(
            polar=dict(
                radialaxis=dict(visible=True, range=[0, 3], tickfont=dict(size=9))
            ),
            height=480,
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_rad, use_container_width=True)

        st.divider()

        # ── Comentarios ───────────────────────
        st.subheader("Comentarios y Observaciones")
        if comments:
            for i, c in enumerate(comments, 1):
                st.markdown(f"**{i}.** {c}")
        else:
            st.info("No hay comentarios escritos para este modulo.")

    else:
        # ─── VISTA COMPARATIVA ────────────────────────────────────────
        st.subheader("Comparacion entre Modulos")

        if len(all_entries) < 2:
            st.info("Cargue al menos 2 archivos para ver la comparacion.")
            st.stop()

        labels_all = [label_for(e)[:40] for e in all_entries]

        # Tabla resumen comparativa
        st.markdown("#### Puntajes promedio por modulo y pregunta")
        comp_rows = []
        for e, lbl in zip(all_entries, labels_all):
            row = {"Modulo": lbl, "Respuestas": e["meta"]["respuestas"]}
            for q in e["questions"]:
                row[q["short"]] = q["score"]
            row["Promedio global"] = round(
                sum(q["score"] for q in e["questions"]) / len(e["questions"]), 2
            )
            comp_rows.append(row)

        df_comp = pd.DataFrame(comp_rows)
        score_cols = [c for c in df_comp.columns if c not in ("Modulo", "Respuestas")]
        st.dataframe(
            df_comp.style.background_gradient(
                subset=score_cols, cmap="RdYlGn", vmin=1, vmax=3
            ),
            use_container_width=True,
            hide_index=True,
        )

        st.divider()

        # ── Comparación por Dimensión ──────────
        st.markdown("#### Comparacion por Dimension")
        dim_comp_rows = []
        for e, lbl in zip(all_entries, labels_all):
            for dim in DIMENSIONS:
                dim_qs = [q for q in e["questions"] if dim in q.get("dimensiones", [])]
                if dim_qs:
                    avg_dim = sum(q["score"] for q in dim_qs) / len(dim_qs)
                    dim_comp_rows.append(
                        {"Modulo": lbl, "Dimension": dim, "Puntaje": round(avg_dim, 2)}
                    )

        if dim_comp_rows:
            df_dim_comp = pd.DataFrame(dim_comp_rows)

            fig_dim_comp = px.bar(
                df_dim_comp,
                x="Dimension",
                y="Puntaje",
                color="Modulo",
                barmode="group",
                text="Puntaje",
                height=400,
                color_discrete_sequence=px.colors.qualitative.Set2,
            )
            fig_dim_comp.update_traces(
                texttemplate="%{text:.2f}", textposition="outside"
            )
            fig_dim_comp.update_layout(
                yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
                xaxis_tickangle=-20,
                plot_bgcolor="rgba(0,0,0,0)",
                paper_bgcolor="rgba(0,0,0,0)",
            )
            st.plotly_chart(fig_dim_comp, use_container_width=True)

            fig_dim_heat = px.imshow(
                df_dim_comp.pivot(
                    index="Modulo", columns="Dimension", values="Puntaje"
                ),
                color_continuous_scale=["#d7191c", "#f39c12", "#1a9641"],
                zmin=1,
                zmax=3,
                text_auto=".2f",
                aspect="auto",
                height=max(250, len(all_entries) * 60 + 120),
            )
            fig_dim_heat.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", coloraxis_colorbar_title="Puntaje"
            )
            st.plotly_chart(fig_dim_heat, use_container_width=True)

            st.markdown("##### Tabla por dimension")
            dim_pivot = df_dim_comp.pivot(
                index="Modulo", columns="Dimension", values="Puntaje"
            )
            dim_pivot["Promedio"] = dim_pivot.mean(axis=1).round(2)
            st.dataframe(
                dim_pivot.style.background_gradient(cmap="RdYlGn", vmin=1, vmax=3),
                use_container_width=True,
            )

        st.divider()

        # Barras agrupadas — puntaje por módulo
        st.markdown("#### Puntaje por pregunta y modulo")
        long_rows = []
        for e, lbl in zip(all_entries, labels_all):
            for q in e["questions"]:
                dim_label = (
                    q.get("dimensiones", ["—"])[0] if q.get("dimensiones") else ""
                )
                long_rows.append(
                    {
                        "Modulo": lbl,
                        "Pregunta": q["short"],
                        "Dimension": dim_label,
                        "Puntaje": q["score"],
                    }
                )
        df_long = pd.DataFrame(long_rows)

        fig_comp = px.bar(
            df_long,
            x="Pregunta",
            y="Puntaje",
            color="Modulo",
            barmode="group",
            text="Puntaje",
            height=480,
        )
        fig_comp.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig_comp.update_layout(
            xaxis_tickangle=-25,
            yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_comp, use_container_width=True)

        st.divider()

        # Heatmap
        st.markdown("#### Mapa de calor — Puntaje por modulo y pregunta")
        piv_cols = [q["short"] for q in all_entries[0]["questions"]]
        pivot = df_long.pivot(index="Modulo", columns="Pregunta", values="Puntaje")[
            piv_cols
        ]
        fig_heat = px.imshow(
            pivot,
            color_continuous_scale=[
                "#d7191c",
                "#f39c12",
                "#1a9641",
            ],
            zmin=1,
            zmax=3,
            text_auto=".2f",
            aspect="auto",
            height=max(300, len(all_entries) * 70 + 150),
        )
        fig_heat.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", coloraxis_colorbar_title="Puntaje"
        )
        st.plotly_chart(fig_heat, use_container_width=True)

        st.divider()

        # Radar comparativo
        st.markdown("#### Radar comparativo")
        fig_multi_rad = go.Figure()
        palette = px.colors.qualitative.Set2
        for idx, (e, lbl) in enumerate(zip(all_entries, labels_all)):
            scores_r = [q["score"] for q in e["questions"]]
            labels_r = [q["short"] for q in e["questions"]]
            color = palette[idx % len(palette)]
            fig_multi_rad.add_trace(
                go.Scatterpolar(
                    r=scores_r + [scores_r[0]],
                    theta=labels_r + [labels_r[0]],
                    fill="toself",
                    name=lbl,
                    line=dict(color=color, width=2),
                )
            )
        fig_multi_rad.update_layout(
            polar=dict(radialaxis=dict(range=[0, 3])),
            height=520,
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_multi_rad, use_container_width=True)

        st.divider()

        # Promedio global ranking
        st.markdown("#### Ranking de modulos por puntaje global")
        rank_rows = []
        for e, lbl in zip(all_entries, labels_all):
            avg = sum(q["score"] for q in e["questions"]) / len(e["questions"])
            rank_rows.append(
                {
                    "Modulo": lbl,
                    "Puntaje global": round(avg, 2),
                    "Respuestas": e["meta"]["respuestas"],
                }
            )
        df_rank = (
            pd.DataFrame(rank_rows)
            .sort_values("Puntaje global", ascending=False)
            .reset_index(drop=True)
        )
        df_rank.index += 1

        fig_rank = px.bar(
            df_rank,
            x="Modulo",
            y="Puntaje global",
            color="Puntaje global",
            color_continuous_scale=["#d7191c", "#fdae61", "#1a9641"],
            range_color=[1, 3],
            text="Puntaje global",
            height=380,
        )
        fig_rank.update_traces(texttemplate="%{text:.2f}", textposition="outside")
        fig_rank.update_layout(
            yaxis=dict(range=[0, 3.5]),
            coloraxis_showscale=False,
            xaxis_tickangle=-20,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_rank, use_container_width=True)
        st.dataframe(df_rank, use_container_width=True)

    # ── Exportar (siempre visible) ────────────
    st.divider()
    st.subheader("Exportar")

    if sel_fname is not None:
        entry = st.session_state.data[sel_fname]
        rows_exp = []
        for q in entry["questions"]:
            rows_exp.append(
                {
                    "Pregunta": q["pregunta"],
                    "Sí": q["Sí"],
                    "%Sí": q["pct_Sí"],
                    "A veces": q["A veces"],
                    "%A veces": q["pct_A veces"],
                    "No": q["No"],
                    "%No": q["pct_No"],
                    "Puntaje(/3)": q["score"],
                }
            )
        csv = pd.DataFrame(rows_exp).to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar resumen del modulo (CSV)",
            csv,
            "resumen_modulo.csv",
            "text/csv",
        )
    else:
        all_rows = []
        for e, lbl in zip(all_entries, labels_all):
            for q in e["questions"]:
                all_rows.append(
                    {
                        "Modulo": lbl,
                        "Pregunta": q["pregunta"],
                        "Sí": q["Sí"],
                        "%Sí": q["pct_Sí"],
                        "A veces": q["A veces"],
                        "%A veces": q["pct_A veces"],
                        "No": q["No"],
                        "%No": q["pct_No"],
                        "Puntaje(/3)": q["score"],
                    }
                )
        csv = pd.DataFrame(all_rows).to_csv(index=False).encode("utf-8")
        st.download_button(
            "Descargar comparativa completa (CSV)",
            csv,
            "comparativa_modulos.csv",
            "text/csv",
        )
