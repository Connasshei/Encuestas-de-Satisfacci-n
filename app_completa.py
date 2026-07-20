import streamlit as st
import openpyxl
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
import unicodedata
from collections import defaultdict

# ══════════════════════════════════════════════════════════
# CONFIGURACIÓN GLOBAL
# ══════════════════════════════════════════════════════════
st.set_page_config(page_title="Evaluaciones - Encuestas de Satisfacción", layout="wide")

COLOR_MAP = {"Sí": "#1a9641", "A veces": "#f39c12", "No": "#d7191c"}
SCALE = ["Sí", "A veces", "No"]

# ── Dimensiones de Docente ──────────────────────────────
DOCENTE_DIMS = [
    "Planificacion e inicio de la sesion",
    "Desarrollo didactico de los contenidos",
    "Interaccion y participacion estudiantil",
    "Uso pedagogico de recursos tecnologicos",
    "Evaluacion formativa y retroalimentacion",
    "Gestion del tiempo y clima de aprendizaje",
]
DOCENTE_DIM_COLORS = {
    DOCENTE_DIMS[0]: "#4e79a7",
    DOCENTE_DIMS[1]: "#f28e2b",
    DOCENTE_DIMS[2]: "#e15759",
    DOCENTE_DIMS[3]: "#76b7b2",
    DOCENTE_DIMS[4]: "#59a14f",
    DOCENTE_DIMS[5]: "#edc948",
}

# ── Dimensiones de Facilitador ──────────────────────────
FACILITADOR_DIMS = [
    "Apoyo y coordinacion",
    "Seguimiento academico",
    "Evaluacion y retroalimentacion",
    "Comunicacion efectiva",
    "Actitud profesional y gestion",
]
FACILITADOR_DIM_COLORS = {
    FACILITADOR_DIMS[0]: "#4e79a7",
    FACILITADOR_DIMS[1]: "#f28e2b",
    FACILITADOR_DIMS[2]: "#e15759",
    FACILITADOR_DIMS[3]: "#76b7b2",
    FACILITADOR_DIMS[4]: "#59a14f",
}

# ── Dimensiones de Supervisor ───────────────────────────
SUPERVISOR_DIMS = [
    "Coordinacion y planificacion",
    "Seguimiento y supervision",
    "Comunicacion",
    "Gestion de recursos y evaluacion",
    "Actitud profesional",
]
SUPERVISOR_DIM_COLORS = {
    SUPERVISOR_DIMS[0]: "#4e79a7",
    SUPERVISOR_DIMS[1]: "#f28e2b",
    SUPERVISOR_DIMS[2]: "#e15759",
    SUPERVISOR_DIMS[3]: "#76b7b2",
    SUPERVISOR_DIMS[4]: "#59a14f",
}

# ── Dimensiones de Asignatura ───────────────────────────
ASIGNATURA_DIMS = [
    "Induccion y planificacion",
    "Comunicacion y apoyo",
    "Desarrollo curricular",
    "Recursos y plataforma",
    "Logros y satisfaccion",
]
ASIGNATURA_DIM_COLORS = {
    ASIGNATURA_DIMS[0]: "#4e79a7",
    ASIGNATURA_DIMS[1]: "#f28e2b",
    ASIGNATURA_DIMS[2]: "#e15759",
    ASIGNATURA_DIMS[3]: "#76b7b2",
    ASIGNATURA_DIMS[4]: "#59a14f",
}

# ══════════════════════════════════════════════════════════
# CONFIGURACIÓN DE LOS 4 TIPOS DE EVALUADOS
# ══════════════════════════════════════════════════════════

EVALUATED_TYPES = {
    "docente": {
        "label": "Docente",
        "color": "#4e79a7",
        "evaluators": {
            "estudiante": {
                "label": "Estudiante",
                "icon": "🎓",
                "question_short": {
                    1: "Claridad programa y objetivos",
                    2: "Organizacion de contenidos",
                    3: "Bibliografia util",
                    4: "Puntualidad del docente",
                    5: "Dominio de contenidos",
                    6: "Claridad de explicaciones",
                    7: "Estrategias de participacion",
                    8: "Relacion con ejemplos prof.",
                    9: "Atencion de dudas",
                    10: "Clima de respeto y confianza",
                    11: "Orientaciones de contenido",
                    12: "Coherencia de evaluaciones",
                    13: "Explicacion de criterios",
                    14: "Retroalimentacion",
                    15: "Aplicacion practica",
                    16: "Acceso y navegacion AV",
                    17: "Plataforma adecuada",
                },
                "dimensions": DOCENTE_DIMS,
                "dimension_colors": DOCENTE_DIM_COLORS,
                "dimension_map": {
                    1: [DOCENTE_DIMS[0]],
                    2: [DOCENTE_DIMS[0]],
                    3: [DOCENTE_DIMS[0]],
                    4: [DOCENTE_DIMS[5]],
                    5: [DOCENTE_DIMS[1]],
                    6: [DOCENTE_DIMS[1]],
                    7: [DOCENTE_DIMS[2]],
                    8: [DOCENTE_DIMS[1]],
                    9: [DOCENTE_DIMS[2]],
                    10: [DOCENTE_DIMS[5]],
                    11: [DOCENTE_DIMS[1]],
                    12: [DOCENTE_DIMS[4]],
                    13: [DOCENTE_DIMS[4]],
                    14: [DOCENTE_DIMS[4]],
                    15: [DOCENTE_DIMS[1]],
                    16: [DOCENTE_DIMS[3]],
                    17: [DOCENTE_DIMS[3]],
                },
                "dedup": True,
            },
            "carrera": {
                "label": "Equipo de la Carrera",
                "icon": "📋",
                "question_short": {
                    1: "Comunicacion de objetivos",
                    2: "Contenidos planificados",
                    3: "Orden logico de contenidos",
                    4: "Atencion de inquietudes",
                    5: "Consignas claras",
                    6: "Recursos digitales",
                    7: "Ambiente participativo",
                    8: "Cumplimiento horario",
                    9: "Retroalimentacion",
                    10: "Coordinacion facilitadores",
                    11: "Planificacion universitaria",
                    12: "Contenido en plataforma",
                    13: "Seguimiento plataforma",
                },
                "dimensions": DOCENTE_DIMS,
                "dimension_colors": DOCENTE_DIM_COLORS,
                "dimension_map": {
                    1: [DOCENTE_DIMS[0]],
                    2: [DOCENTE_DIMS[1]],
                    3: [DOCENTE_DIMS[1]],
                    4: [DOCENTE_DIMS[2]],
                    5: [DOCENTE_DIMS[0]],
                    6: [DOCENTE_DIMS[3]],
                    7: [DOCENTE_DIMS[2]],
                    8: [DOCENTE_DIMS[5]],
                    9: [DOCENTE_DIMS[4]],
                    10: [DOCENTE_DIMS[5]],
                    11: [DOCENTE_DIMS[0]],
                    12: [DOCENTE_DIMS[3]],
                    13: [DOCENTE_DIMS[4]],
                },
                "dedup": False,
            },
        },
    },
    "facilitador": {
        "label": "Facilitador",
        "color": "#f28e2b",
        "evaluators": {
            "supervisor": {
                "label": "Supervisor",
                "icon": "👁",
                "question_short": {
                    1: "Apoyo oportuno",
                    2: "Coordinacion con docente",
                    3: "Coordinacion oportuna",
                    4: "Seguimiento participacion",
                    5: "Identifica estudiantes riesgo",
                    6: "Plazos calificacion",
                    7: "Registro academico actualizado",
                    8: "Responsabilidad y compromiso",
                    9: "Participa reunion coordinacion",
                    10: "Fomenta reflexion",
                    11: "Comunicacion efectiva",
                },
                "dimensions": FACILITADOR_DIMS,
                "dimension_colors": FACILITADOR_DIM_COLORS,
                "dimension_map": {
                    1: [FACILITADOR_DIMS[0]],
                    2: [FACILITADOR_DIMS[0]],
                    3: [FACILITADOR_DIMS[0]],
                    4: [FACILITADOR_DIMS[1]],
                    5: [FACILITADOR_DIMS[1]],
                    6: [FACILITADOR_DIMS[2]],
                    7: [FACILITADOR_DIMS[2]],
                    8: [FACILITADOR_DIMS[4]],
                    9: [FACILITADOR_DIMS[4]],
                    10: [FACILITADOR_DIMS[4]],
                    11: [FACILITADOR_DIMS[3]],
                },
                "dedup": False,
            },
            "estudiantes": {
                "label": "Estudiantes",
                "icon": "🎓",
                "question_short": {
                    1: "Apoyo ante dificultades",
                    2: "Habilita actividades",
                    3: "Seguimiento participacion",
                    4: "Claridad y prontitud",
                    5: "Revision tareas",
                    6: "Retroalimentacion clara",
                    7: "Publica calificaciones",
                    8: "Comunicacion respetuosa",
                },
                "dimensions": FACILITADOR_DIMS,
                "dimension_colors": FACILITADOR_DIM_COLORS,
                "dimension_map": {
                    1: [FACILITADOR_DIMS[0]],
                    2: [FACILITADOR_DIMS[0]],
                    3: [FACILITADOR_DIMS[1]],
                    5: [FACILITADOR_DIMS[1]],
                    6: [FACILITADOR_DIMS[2]],
                    7: [FACILITADOR_DIMS[2]],
                    4: [FACILITADOR_DIMS[3]],
                    8: [FACILITADOR_DIMS[3]],
                },
                "dedup": False,
            },
        },
    },
    "supervisor": {
        "label": "Supervisor",
        "color": "#e15759",
        "evaluators": {
            "equipo_pedagogico": {
                "label": "Equipo Pedagogico",
                "icon": "📚",
                "question_short": {
                    1: "Coordina procesos academicos",
                    2: "Seguimiento academico",
                    3: "Coordina revision programas",
                    4: "Supervisa implementacion",
                    5: "Comunicacion equipo pedagogico",
                    6: "Disposicion a mejoras",
                    7: "Seguimiento cronogramas",
                    8: "Promueve evaluacion",
                    9: "Actitud profesional",
                    10: "Gestion clase espejo",
                    11: "Comunicacion respetuosa",
                    12: "Coordina entrega guion",
                    13: "Articula equipo",
                    14: "Coordina revision materiales",
                },
                "dimensions": SUPERVISOR_DIMS,
                "dimension_colors": SUPERVISOR_DIM_COLORS,
                "dimension_map": {
                    1: [SUPERVISOR_DIMS[0]],
                    3: [SUPERVISOR_DIMS[0]],
                    12: [SUPERVISOR_DIMS[0]],
                    14: [SUPERVISOR_DIMS[0]],
                    2: [SUPERVISOR_DIMS[1]],
                    4: [SUPERVISOR_DIMS[1]],
                    7: [SUPERVISOR_DIMS[1]],
                    5: [SUPERVISOR_DIMS[2]],
                    11: [SUPERVISOR_DIMS[2]],
                    13: [SUPERVISOR_DIMS[2]],
                    6: [SUPERVISOR_DIMS[3]],
                    8: [SUPERVISOR_DIMS[3]],
                    10: [SUPERVISOR_DIMS[3]],
                    9: [SUPERVISOR_DIMS[4]],
                },
                "dedup": False,
            },
            "facilitador": {
                "label": "Facilitador",
                "icon": "👥",
                "question_short": {
                    1: "Coordina con facilitadores",
                    2: "Induccion plataforma",
                    3: "Verifica recursos",
                    4: "Comunicacion clara",
                    5: "Seguimiento sincronicas",
                    6: "Coordina progreso estudiantes",
                    7: "Seguimiento foros",
                    8: "Evalua pertinencia",
                    9: "Coordina encuestas",
                    10: "Fomenta reflexion",
                    11: "Reuniones de cierre",
                },
                "dimensions": SUPERVISOR_DIMS,
                "dimension_colors": SUPERVISOR_DIM_COLORS,
                "dimension_map": {
                    1: [SUPERVISOR_DIMS[0]],
                    2: [SUPERVISOR_DIMS[0]],
                    3: [SUPERVISOR_DIMS[0]],
                    5: [SUPERVISOR_DIMS[1]],
                    6: [SUPERVISOR_DIMS[1]],
                    7: [SUPERVISOR_DIMS[1]],
                    4: [SUPERVISOR_DIMS[2]],
                    8: [SUPERVISOR_DIMS[3]],
                    9: [SUPERVISOR_DIMS[3]],
                    10: [SUPERVISOR_DIMS[3]],
                    11: [SUPERVISOR_DIMS[3]],
                },
                "dedup": False,
            },
        },
    },
    "asignatura": {
        "label": "Asignatura",
        "color": "#76b7b2",
        "evaluators": {
            "docente": {
                "label": "Docente",
                "icon": "👨‍🏫",
                "question_short": {
                    1: "Participacion induccion",
                    2: "Informacion induccion",
                    3: "Uso programa asignatura",
                    4: "Planificacion facilito trabajo",
                    5: "Acompanamiento supervisor",
                    6: "Comunicacion supervisora",
                    7: "Comunicacion facilitadora",
                    8: "Apoyo facilitadora",
                    9: "Desarrollo planificado",
                    10: "Cumplio requisitos",
                    11: "Plataforma facilito",
                    12: "Soporte tecnico",
                    13: "Logros estudiantes",
                    14: "Experiencia satisfactoria",
                },
                "dimensions": ASIGNATURA_DIMS,
                "dimension_colors": ASIGNATURA_DIM_COLORS,
                "dimension_map": {
                    1: [ASIGNATURA_DIMS[0]],
                    2: [ASIGNATURA_DIMS[0]],
                    3: [ASIGNATURA_DIMS[0]],
                    4: [ASIGNATURA_DIMS[0]],
                    5: [ASIGNATURA_DIMS[1]],
                    6: [ASIGNATURA_DIMS[1]],
                    7: [ASIGNATURA_DIMS[1]],
                    8: [ASIGNATURA_DIMS[1]],
                    9: [ASIGNATURA_DIMS[2]],
                    10: [ASIGNATURA_DIMS[2]],
                    11: [ASIGNATURA_DIMS[3]],
                    12: [ASIGNATURA_DIMS[3]],
                    13: [ASIGNATURA_DIMS[4]],
                    14: [ASIGNATURA_DIMS[4]],
                },
                "dedup": False,
            },
        },
    },
}


def get_evaluator_pairs(ev_type_key):
    """Retorna lista de (ev_key, ev_config) para un tipo evaluado."""
    return list(EVALUATED_TYPES[ev_type_key]["evaluators"].items())


# ══════════════════════════════════════════════════════════
# PARSING
# ══════════════════════════════════════════════════════════


def extract_meta_value(cell_val, prefix):
    if cell_val and str(cell_val).startswith(prefix):
        return str(cell_val)[len(prefix) :].strip()
    return ""


def safe(val):
    try:
        return float(val or 0)
    except:
        return 0.0


def parse_evaluacion(file_bytes, filename, ev_type_key, ev_key, config):
    """Parsea un Excel de evaluacion usando el config del evaluador."""
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
        "tipo_evaluado": ev_type_key,
        "tipo_evaluador": ev_key,
    }
    questions = []
    comments = []
    in_questions = False
    in_comments = False

    dedup = config.get("dedup", False)
    seen_texts = {}

    for row in rows:
        first = str(row[0] or "").strip()

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
        elif first == "Etiqueta":
            in_questions = True
            continue
        elif in_questions and not in_comments:
            pregunta = str(row[1] or "").strip()
            if "observaciones" in pregunta.lower() or "comentarios" in pregunta.lower():
                in_questions = False
                in_comments = True
                continue

            if pregunta and len(pregunta) > 5:
                num_match = re.match(r"^(\d+)[\.\-]", pregunta)
                clean_text = re.sub(r"^\d+[\.\-]\s*", "", pregunta).strip()

                si = safe(row[2])
                aveces = safe(row[4])
                no = safe(row[6])

                if dedup and clean_text in seen_texts:
                    idx = seen_texts[clean_text]
                    prev = st.session_state._parsing_buffer[idx]
                    prev["Sí"] = (prev["Sí"] + si) / 2
                    prev["A veces"] = (prev["A veces"] + aveces) / 2
                    prev["No"] = (prev["No"] + no) / 2
                    continue

                num = (
                    int(num_match.group(1))
                    if num_match
                    else (len(seen_texts) + 1 if dedup else len(questions) + 1)
                )

                max_q = len(config["question_short"])
                if num > max_q:
                    continue

                q_entry = {
                    "num": num,
                    "pregunta": pregunta,
                    "short": config["question_short"].get(num, f"P{num}"),
                    "dimensiones": config["dimension_map"].get(num, []),
                    "Sí": si,
                    "A veces": aveces,
                    "No": no,
                }

                if dedup:
                    seen_texts[clean_text] = num
                    if "_parsing_buffer" not in st.session_state:
                        st.session_state._parsing_buffer = {}
                    st.session_state._parsing_buffer[num] = q_entry
                else:
                    questions.append(q_entry)

        elif in_comments:
            for cell in row[2:]:
                val = str(cell or "").strip()
                if val and val.lower() not in (
                    "ninguno",
                    "ninguna",
                    "n/a",
                    "",
                    ".",
                    "-",
                    "..",
                    "ninguna observacion",
                    "ninguna observación",
                    "sin observaciones",
                    "sin comentarios",
                    "sin observacion",
                    "sin comentario",
                    "ninguna observacion ",
                    "ninguna observación ",
                    "sin observaciones.",
                    "nada",
                    "0",
                    "ningún",
                    "ningun",
                ):
                    comments.append(val)

    if not meta["docente"]:
        name_part = filename.replace(".xlsx", "").replace("_", " ")
        meta["docente"] = name_part.strip()

    if dedup:
        buf = getattr(st.session_state, "_parsing_buffer", {})
        for num in sorted(buf.keys()):
            q = buf[num]
            total = q["Sí"] + q["A veces"] + q["No"]
            q["score"] = (
                min(
                    3.0,
                    round((q["Sí"] * 3 + q["A veces"] * 2 + q["No"] * 1) / total, 2),
                )
                if total > 0
                else 0.0
            )
            questions.append(q)
        if "_parsing_buffer" in st.session_state:
            del st.session_state._parsing_buffer

    for q in questions:
        total = q["Sí"] + q["A veces"] + q["No"]
        q["score"] = (
            min(3.0, round((q["Sí"] * 3 + q["A veces"] * 2 + q["No"] * 1) / total, 2))
            if total > 0
            else 0.0
        )

    return {"meta": meta, "questions": questions, "comments": comments}


# ══════════════════════════════════════════════════════════
# FUNCIONES DE VISUALIZACIÓN
# ══════════════════════════════════════════════════════════


def get_dimension_info(entry):
    """Obtiene dimensions y dimension_colors desde el entry."""
    ev_type = entry["meta"]["tipo_evaluado"]
    ev_key = entry["meta"]["tipo_evaluador"]
    cfg = EVALUATED_TYPES[ev_type]["evaluators"][ev_key]
    return cfg["dimensions"], cfg["dimension_colors"]


def render_individual(entry):
    """Muestra la vista de resultados para una entrada."""
    meta = entry["meta"]
    qs = entry["questions"]
    comments = entry["comments"]

    ev_type_key = meta["tipo_evaluado"]
    ev_key = meta["tipo_evaluador"]
    ev_type_label = EVALUATED_TYPES[ev_type_key]["label"]
    ev_label = EVALUATED_TYPES[ev_type_key]["evaluators"][ev_key]["label"]

    dimensions, dim_colors = get_dimension_info(entry)

    with st.container(border=True):
        guion = "\u2014"
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown("**Módulo**")
            st.markdown(
                f"<div style='word-break:break-word;white-space:normal;line-height:1.4'>{meta['modulo'] or guion}</div>",
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown("**Evaluado**")
            st.markdown(
                f"<div style='word-break:break-word;white-space:normal;line-height:1.4'>{ev_type_label}</div>",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown("**Evaluador**")
            st.markdown(
                f"<div style='word-break:break-word;white-space:normal;line-height:1.4'>{ev_label}</div>",
                unsafe_allow_html=True,
            )
        with c4:
            st.metric("Respuestas", meta["respuestas"])
        c5, c6, c7 = st.columns(3)
        c5.caption(f"Inicio: {meta['fecha_inicio']}")
        c6.caption(f"Fin: {meta['fecha_fin']}")
        if meta.get("docente"):
            c7.markdown(f"**Docente:** {meta['docente']}")

    st.divider()

    # ── KPIs ──
    st.subheader("Indicadores Generales")
    total_si = sum(q["Sí"] for q in qs)
    total_aveces = sum(q["A veces"] for q in qs)
    total_no = sum(q["No"] for q in qs)
    grand_total = total_si + total_aveces + total_no or 1
    avg_score = sum(q["score"] for q in qs) / len(qs) if qs else 0
    pct_si = total_si / grand_total * 100
    pct_aveces = total_aveces / grand_total * 100
    pct_no = total_no / grand_total * 100

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Puntaje promedio", f"{avg_score:.2f} / 3.00")
    k2.metric("Sí", f"{pct_si:.1f}%")
    k3.metric("A veces", f"{pct_aveces:.1f}%")
    k4.metric("No", f"{pct_no:.1f}%")

    st.divider()

    # ── Resumen por Dimensión ──
    st.subheader("Resumen por Dimensión")
    dim_rows = []
    for dim in dimensions:
        dim_qs = [q for q in qs if dim in q.get("dimensiones", [])]
        if dim_qs:
            avg_dim = sum(q["score"] for q in dim_qs) / len(dim_qs)
            dim_rows.append({"Dimensión": dim, "Puntaje promedio": round(avg_dim, 2)})

    if dim_rows:
        df_dim = pd.DataFrame(dim_rows)
        cols_dim = st.columns(len(dimensions))
        for i, row in enumerate(dim_rows):
            cols_dim[i].metric(row["Dimensión"][:25], row["Puntaje promedio"])

        fig_dim = px.bar(
            df_dim,
            x="Dimensión",
            y="Puntaje promedio",
            color="Dimensión",
            color_discrete_map=dim_colors,
            text="Puntaje promedio",
            height=380,
        )
        fig_dim.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside",
            hovertemplate="%{x}: %{y:.2f}",
        )
        fig_dim.update_layout(
            yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
            showlegend=False,
            xaxis_tickangle=-20,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_dim, use_container_width=True)

    st.divider()

    # ── Tabla resumen ──
    st.subheader("Resumen por Pregunta")
    rows_sum = []
    for q in qs:
        dim_label = (
            q.get("dimensiones", ["\u2014"])[0] if q.get("dimensiones") else "\u2014"
        )
        total_q = q["Sí"] + q["A veces"] + q["No"] or 1
        rows_sum.append(
            {
                "Dimensión": dim_label,
                "Pregunta": q["short"],
                "Sí": f"{q['Sí'] / total_q * 100:.0f}%",
                "A veces": f"{q['A veces'] / total_q * 100:.0f}%",
                "No": f"{q['No'] / total_q * 100:.0f}%",
                "Puntaje (/3)": q["score"],
            }
        )
    df_sum = pd.DataFrame(rows_sum)
    styler = df_sum.style.background_gradient(
        subset=["Puntaje (/3)"], cmap="RdYlGn", vmin=1, vmax=3
    ).format({"Puntaje (/3)": "{:.2f}"})
    st.dataframe(styler, use_container_width=True, hide_index=True)

    st.divider()

    # ── Barras 100% apiladas ──
    st.subheader("Distribución de Respuestas por Pregunta")
    fig_stack = go.Figure()
    for cat in SCALE:
        pct_vals = [
            round(q[cat] / (q["Sí"] + q["A veces"] + q["No"] or 1) * 100, 2) for q in qs
        ]
        fig_stack.add_trace(
            go.Bar(
                name=cat,
                y=[q["short"] for q in qs],
                x=pct_vals,
                orientation="h",
                marker_color=COLOR_MAP[cat],
                text=[f"{pct:.0f}%" if pct > 3 else "" for pct in pct_vals],
                textposition="inside",
                insidetextanchor="middle",
                hovertemplate=f"%{{y}}: %{{x:.1f}}%<extra>{cat}</extra>",
            )
        )
    fig_stack.update_layout(
        barmode="stack",
        xaxis=dict(title="% de respuestas", ticksuffix="%", range=[0, 100]),
        height=400,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=180, r=20, t=60, b=40),
    )
    st.plotly_chart(fig_stack, use_container_width=True)

    # ── Barras de puntaje ──
    st.subheader("Puntaje Ponderado por Pregunta (escala 1-3)")
    df_score = pd.DataFrame(
        {
            "Pregunta": [q["short"] for q in qs],
            "Puntaje": [q["score"] for q in qs],
            "Dimensión": [
                q.get("dimensiones", ["\u2014"])[0]
                if q.get("dimensiones")
                else "\u2014"
                for q in qs
            ],
        }
    )
    fig_score = px.bar(
        df_score,
        x="Puntaje",
        y="Pregunta",
        orientation="h",
        color="Dimensión",
        color_discrete_map=dim_colors,
        text="Puntaje",
        height=380,
    )
    fig_score.update_traces(
        texttemplate="%{text:.2f}",
        textposition="outside",
        hovertemplate="%{y}: %{x:.2f}",
    )
    fig_score.update_layout(
        xaxis=dict(range=[0, 3.5], title="Puntaje promedio"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=180, r=80, t=20, b=40),
    )
    st.plotly_chart(fig_score, use_container_width=True)

    st.divider()

    # ── Gráficos de torta ──
    st.subheader("Detalle por Pregunta")
    if len(qs) > 0:
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
                    q.get("dimensiones", ["\u2014"])[0] if q.get("dimensiones") else ""
                )
                cols[j].plotly_chart(fig_pie, use_container_width=True)
                cols[j].caption(f"{q['pregunta']}  |  _{dim_label}_")

    st.divider()

    # ── Radar ──
    st.subheader(f"Perfil del {ev_type_label}")
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
            hovertemplate="%{theta}: %{r:.2f}",
        )
    )
    fig_rad.update_layout(
        polar=dict(radialaxis=dict(visible=True, range=[0, 3], tickfont=dict(size=9))),
        height=480,
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_rad, use_container_width=True)

    st.divider()

    # ── Comentarios ──
    st.subheader("Comentarios y Observaciones")
    if comments:
        for i, c in enumerate(comments, 1):
            st.markdown(f"**{i}.** {c}")
    else:
        st.info("No hay comentarios escritos para esta evaluación.")


def normalize_mod(name):
    if not name:
        return ""
    n = re.sub(r"\s+", " ", name).strip().upper()
    n = "".join(
        c for c in unicodedata.normalize("NFKD", n) if not unicodedata.combining(c)
    )
    return n


def render_comparacion():
    """Agrupa entradas por (tipo_evaluado, módulo), promedia evaluadores y muestra gráficas del promedio combinado."""

    groups = defaultdict(list)
    for ev_type_key in EVALUATED_TYPES:
        for ev_key in EVALUATED_TYPES[ev_type_key]["evaluators"]:
            store = st.session_state.data.get(ev_type_key, {}).get(ev_key, {})
            for entry in store.values():
                mod_norm = normalize_mod(entry["meta"]["modulo"])
                if mod_norm:
                    groups[(ev_type_key, mod_norm)].append(entry)

    if not groups:
        st.info("No hay datos cargados.")
        return

    combined = []
    for (ev_type_key, _mod_norm), entries in groups.items():
        entry_avgs = []
        for e in entries:
            qs = e["questions"]
            entry_avgs.append(sum(q["score"] for q in qs) / len(qs) if qs else 0)
        combined_avg = round(sum(entry_avgs) / len(entry_avgs), 2)
        first = entries[0]

        ev_breakdown = {}
        for e in entries:
            ek = e["meta"]["tipo_evaluador"]
            el = EVALUATED_TYPES[ev_type_key]["evaluators"][ek]["label"]
            qs = e["questions"]
            ev_breakdown[el] = round(
                sum(q["score"] for q in qs) / len(qs) if qs else 0, 2
            )

        combined.append(
            {
                "ev_type_key": ev_type_key,
                "modulo": first["meta"]["modulo"],
                "docente": first["meta"]["docente"],
                "combined_avg": combined_avg,
                "entries": entries,
                "ev_breakdown": ev_breakdown,
                "respuestas": sum(e["meta"]["respuestas"] for e in entries),
            }
        )

    combined.sort(key=lambda x: x["combined_avg"], reverse=True)

    guion = "\u2014"
    type_labels = {k: v["label"] for k, v in EVALUATED_TYPES.items()}
    type_colors = {v["label"]: v["color"] for v in EVALUATED_TYPES.values()}

    # ───────────────────────────────
    # Fila 1 — Métricas combinadas
    # ───────────────────────────────
    st.markdown("#### Métricas Combinadas")
    overall_avg = round(sum(d["combined_avg"] for d in combined) / len(combined), 2)
    best_mod = combined[0]
    worst_mod = combined[-1]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Promedio Combinado", f"{overall_avg} / 3")
    col2.metric("Total Módulos Únicos", len(combined))
    col3.metric(
        "Mejor Módulo",
        (best_mod["modulo"] or best_mod["docente"])[:25],
        f"{best_mod['combined_avg']:.2f}",
    )
    col4.metric(
        "Menor Módulo",
        (worst_mod["modulo"] or worst_mod["docente"])[:25],
        f"{worst_mod['combined_avg']:.2f}",
    )

    st.divider()

    # ───────────────────────────────
    # Fila 2 — Ranking Combinado
    # ───────────────────────────────
    st.markdown("#### Ranking Combinado por Módulo")
    rank_rows = []
    for d in combined:
        rank_rows.append(
            {
                "Módulo": d["modulo"] or d["docente"] or guion,
                "Tipo": type_labels[d["ev_type_key"]],
                "Promedio Combinado": d["combined_avg"],
                **{k: v for k, v in d["ev_breakdown"].items()},
                "Respuestas": d["respuestas"],
            }
        )
    df_rank = pd.DataFrame(rank_rows).reset_index(drop=True)
    df_rank.index += 1

    fig_rank = px.bar(
        df_rank,
        x="Módulo",
        y="Promedio Combinado",
        color="Tipo",
        color_discrete_map=type_colors,
        text="Promedio Combinado",
        height=380,
    )
    fig_rank.update_traces(
        texttemplate="%{text:.2f}",
        textposition="outside",
        hovertemplate="%{x}: %{y:.2f}",
    )
    fig_rank.update_layout(
        yaxis=dict(range=[0, 3.5]),
        xaxis_tickangle=-20,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
    )
    st.plotly_chart(fig_rank, use_container_width=True)
    float_cols = df_rank.select_dtypes(include="float").columns.tolist()
    st.dataframe(
        df_rank.style.format({c: "{:.2f}" for c in float_cols}),
        use_container_width=True,
    )

    st.divider()

    # ───────────────────────────────
    # Fila 3 — Promedio por Tipo
    # ───────────────────────────────
    st.markdown("#### Promedio Combinado por Tipo Evaluado")
    tipo_rows = []
    for ev_type_key, cfg in EVALUATED_TYPES.items():
        scores = [
            d["combined_avg"] for d in combined if d["ev_type_key"] == ev_type_key
        ]
        if scores:
            tipo_rows.append(
                {
                    "Tipo Evaluado": cfg["label"],
                    "Promedio Combinado": round(sum(scores) / len(scores), 2),
                    "Cantidad Módulos": len(scores),
                }
            )
    if tipo_rows:
        df_tipo = pd.DataFrame(tipo_rows)
        fig_tipo = px.bar(
            df_tipo,
            x="Tipo Evaluado",
            y="Promedio Combinado",
            color="Tipo Evaluado",
            color_discrete_map=type_colors,
            text="Promedio Combinado",
            height=350,
        )
        fig_tipo.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside",
            hovertemplate="%{x}: %{y:.2f}",
        )
        fig_tipo.update_layout(
            yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
            showlegend=False,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_tipo, use_container_width=True)

    st.divider()

    # ───────────────────────────────
    # Fila 4 — Dimensión combinada
    # ───────────────────────────────
    st.markdown("#### Promedio por Dimensión (Combinado)")
    dim_data = []
    for d in combined:
        ev_type_key = d["ev_type_key"]
        for e in d["entries"]:
            ek = e["meta"]["tipo_evaluador"]
            cfg = EVALUATED_TYPES[ev_type_key]["evaluators"][ek]
            for dim in cfg["dimensions"]:
                dim_qs = [q for q in e["questions"] if dim in q.get("dimensiones", [])]
                if dim_qs:
                    dim_data.append(
                        {
                            "Dimensión": dim,
                            "Tipo": type_labels[ev_type_key],
                            "Promedio": round(
                                sum(q["score"] for q in dim_qs) / len(dim_qs), 2
                            ),
                        }
                    )
    if dim_data:
        df_dim = pd.DataFrame(dim_data)
        df_dim_grouped = (
            df_dim.groupby(["Dimensión", "Tipo"], as_index=False)["Promedio"]
            .mean()
            .round(2)
        )
        fig_dim = px.bar(
            df_dim_grouped,
            x="Dimensión",
            y="Promedio",
            color="Tipo",
            barmode="group",
            color_discrete_map=type_colors,
            text="Promedio",
            height=400,
        )
        fig_dim.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside",
            hovertemplate="%{x}: %{y:.2f}",
        )
        fig_dim.update_layout(
            yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
            xaxis_tickangle=-20,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_dim, use_container_width=True)
    else:
        st.info("No hay datos de dimensiones.")

    st.divider()

    # ───────────────────────────────
    # Fila 5 — Desglose por módulo
    # ───────────────────────────────
    st.markdown("#### Desglose por Módulo")
    detail_rows = []
    for d in combined:
        row = {
            "Tipo": type_labels[d["ev_type_key"]],
            "Módulo": d["modulo"] or d["docente"] or guion,
            "Prom. Combinado": d["combined_avg"],
        }
        for el, sc in d["ev_breakdown"].items():
            row[el] = sc
        row["Respuestas"] = d["respuestas"]
        detail_rows.append(row)
    df_detail = pd.DataFrame(detail_rows).sort_values(
        ["Tipo", "Prom. Combinado"], ascending=[True, False]
    )
    float_cols = df_detail.select_dtypes(include="float").columns.tolist()
    fmt_dict = {c: "{:.2f}" for c in float_cols}
    styler = df_detail.style.background_gradient(
        subset=[c for c in df_detail.columns if c in ("Prom. Combinado",)],
        cmap="RdYlGn",
        vmin=1,
        vmax=3,
    ).format(fmt_dict)
    st.dataframe(styler, use_container_width=True, hide_index=True)


def render_diagnostico():
    """Dashboard consolidado con todos los tipos cargados."""
    st.subheader("Diagnóstico General")

    all_entries = []
    for ev_type_key in EVALUATED_TYPES:
        for ev_key in EVALUATED_TYPES[ev_type_key]["evaluators"]:
            store = st.session_state.data.get(ev_type_key, {}).get(ev_key, {})
            all_entries.extend(store.values())

    if not all_entries:
        st.info("No hay datos cargados.")
        return

    # ── Métricas globales ──
    cols_metrics = st.columns(len(EVALUATED_TYPES))
    for i, (ev_type_key, cfg) in enumerate(EVALUATED_TYPES.items()):
        total = 0
        resp = 0
        for ev_key in cfg["evaluators"]:
            store = st.session_state.data.get(ev_type_key, {}).get(ev_key, {})
            total += len(store)
            resp += sum(e["meta"]["respuestas"] for e in store.values())
        cols_metrics[i].metric(f"{cfg['label']}", f"{total} archivos", f"{resp} resp.")

    st.divider()

    # ── Filtro por tipo evaluado ──
    tipo_opts = {"Todos": None}
    for ev_type_key, cfg in EVALUATED_TYPES.items():
        tipo_opts[cfg["label"]] = ev_type_key
    sel_tipo_label = st.selectbox("Filtrar por tipo evaluado:", list(tipo_opts.keys()))
    sel_tipo_key = tipo_opts[sel_tipo_label]

    filtered = (
        all_entries
        if sel_tipo_key is None
        else [e for e in all_entries if e["meta"]["tipo_evaluado"] == sel_tipo_key]
    )

    if not filtered:
        st.info("No hay datos para el filtro seleccionado.")
        return

    # ── Ranking ──
    st.markdown("#### Ranking de Evaluaciones")
    rank_rows = []
    for entry in filtered:
        qs = entry["questions"]
        avg = sum(q["score"] for q in qs) / len(qs) if qs else 0
        ev_type_label = EVALUATED_TYPES[entry["meta"]["tipo_evaluado"]]["label"]
        ev_label = EVALUATED_TYPES[entry["meta"]["tipo_evaluado"]]["evaluators"][
            entry["meta"]["tipo_evaluador"]
        ]["label"]
        rank_rows.append(
            {
                "Módulo": entry["meta"]["modulo"] or entry["meta"]["archivo"],
                "Evaluado": ev_type_label,
                "Evaluador": ev_label,
                "Docente": entry["meta"]["docente"] or "\u2014",
                "Puntaje global": min(3.0, round(avg, 2)),
                "Respuestas": entry["meta"]["respuestas"],
            }
        )

    if rank_rows:
        df_rank = pd.DataFrame(rank_rows)
        df_rank_grouped = (
            df_rank.groupby(
                ["Módulo", "Evaluado", "Evaluador", "Docente"], as_index=False
            )
            .agg({"Puntaje global": "mean", "Respuestas": "sum"})
            .round(2)
            .sort_values("Puntaje global", ascending=False)
            .reset_index(drop=True)
        )
        df_rank_grouped["Puntaje global"] = df_rank_grouped["Puntaje global"].clip(
            upper=3.0
        )
        df_rank_grouped.index += 1
        type_colors = {
            cfg["label"]: cfg["color"] for ev_type_key, cfg in EVALUATED_TYPES.items()
        }
        fig_rank = px.bar(
            df_rank_grouped,
            x="Módulo",
            y="Puntaje global",
            color="Evaluado",
            color_discrete_map=type_colors,
            text="Puntaje global",
            height=380,
            barmode="group",
        )
        fig_rank.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside",
            hovertemplate="%{x}: %{y:.2f}",
        )
        fig_rank.update_layout(
            yaxis=dict(range=[0, 3.5]),
            xaxis_tickangle=-20,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_rank, use_container_width=True)
        float_cols = df_rank_grouped.select_dtypes(include="float").columns.tolist()
        st.dataframe(
            df_rank_grouped.style.format({c: "{:.2f}" for c in float_cols}),
            use_container_width=True,
        )

    st.divider()

    # ── Promedios por Dimensión y Tipo ──
    st.markdown("#### Promedio por Dimensión y Tipo")
    dim_data = []
    for entry in filtered:
        ev_type_key = entry["meta"]["tipo_evaluado"]
        ev_key = entry["meta"]["tipo_evaluador"]
        cfg = EVALUATED_TYPES[ev_type_key]["evaluators"][ev_key]
        ev_label = cfg["label"]
        for dim in cfg["dimensions"]:
            dim_qs = [q for q in entry["questions"] if dim in q.get("dimensiones", [])]
            if dim_qs:
                dim_data.append(
                    {
                        "Dimensión": dim,
                        "Tipo": ev_label,
                        "Promedio": round(
                            sum(q["score"] for q in dim_qs) / len(dim_qs), 2
                        ),
                    }
                )
    if dim_data:
        df_dim = pd.DataFrame(dim_data)
        df_dim_grouped = (
            df_dim.groupby(["Dimensión", "Tipo"], as_index=False)["Promedio"]
            .mean()
            .round(2)
        )
        fig_dim = px.bar(
            df_dim_grouped,
            x="Dimensión",
            y="Promedio",
            color="Tipo",
            barmode="group",
            text="Promedio",
            height=400,
        )
        fig_dim.update_traces(
            texttemplate="%{text:.2f}",
            textposition="outside",
            hovertemplate="%{x}: %{y:.2f}",
        )
        fig_dim.update_layout(
            yaxis=dict(range=[0, 3.5], title="Puntaje /3"),
            xaxis_tickangle=-20,
            plot_bgcolor="rgba(0,0,0,0)",
            paper_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(fig_dim, use_container_width=True)

    st.divider()

    # ── Fortalezas y Áreas de Mejora ──
    st.markdown("#### Fortalezas y Áreas de Mejora")
    fortalezas = []
    debilidades = []
    for entry in filtered:
        ev_type_label = EVALUATED_TYPES[entry["meta"]["tipo_evaluado"]]["label"]
        for q in entry["questions"]:
            item = {
                "Módulo": entry["meta"]["modulo"] or entry["meta"]["archivo"],
                "Evaluado": ev_type_label,
                "Pregunta": q["short"],
                "Puntaje": q["score"],
            }
            if q["score"] >= 2.8:
                fortalezas.append(item)
            elif q["score"] <= 2.0:
                debilidades.append(item)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Fortalezas** (puntaje ≥ 2.8)")
        if fortalezas:
            df_fort = pd.DataFrame(fortalezas).sort_values("Puntaje", ascending=False)
            st.dataframe(
                df_fort.style.format({"Puntaje": "{:.2f}"}),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Sin fortalezas destacadas.")
    with col_b:
        st.markdown("**Áreas de mejora** (puntaje ≤ 2.0)")
        if debilidades:
            df_deb = pd.DataFrame(debilidades).sort_values("Puntaje")
            st.dataframe(
                df_deb.style.format({"Puntaje": "{:.2f}"}),
                use_container_width=True,
                hide_index=True,
            )
        else:
            st.info("Sin áreas de mejora detectadas.")

    st.divider()

    # ── Estadísticas Descriptivas ──
    st.markdown("#### Estadísticas Descriptivas")
    stats_rows = []
    for entry in filtered:
        scores = [q["score"] for q in entry["questions"]]
        if scores:
            ev_type_label = EVALUATED_TYPES[entry["meta"]["tipo_evaluado"]]["label"]
            stats_rows.append(
                {
                    "Módulo": entry["meta"]["modulo"] or entry["meta"]["archivo"],
                    "Evaluado": ev_type_label,
                    "Media": round(sum(scores) / len(scores), 2),
                    "Mediana": round(pd.Series(scores).median(), 2),
                    "Mín": round(min(scores), 2),
                    "Máx": round(max(scores), 2),
                    "Std": round(pd.Series(scores).std(), 2) if len(scores) > 1 else 0,
                }
            )
    if stats_rows:
        df_stats = pd.DataFrame(stats_rows)
        float_cols = df_stats.select_dtypes(include="float").columns.tolist()
        st.dataframe(
            df_stats.style.format({c: "{:.2f}" for c in float_cols}),
            use_container_width=True,
            hide_index=True,
        )


def label_for(entry):
    m = entry["meta"]
    base = m["modulo"] if m["modulo"] else m["archivo"]
    return base[:60]


# ══════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════
if "data" not in st.session_state:
    st.session_state.data = {}
    for ev_type_key in EVALUATED_TYPES:
        st.session_state.data[ev_type_key] = {}
        for ev_key in EVALUATED_TYPES[ev_type_key]["evaluators"]:
            st.session_state.data[ev_type_key][ev_key] = {}

# ══════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("## Evaluaciones")
    st.markdown("Encuestas de Satisfacción")
    st.markdown("---")
    page = st.radio(
        "Navegación",
        ["Cargar Datos", "Resultados", "Comparación", "Diagnóstico General"],
        label_visibility="collapsed",
    )
    st.markdown("---")

    for ev_type_key, cfg in EVALUATED_TYPES.items():
        total_archivos = 0
        total_resp = 0
        for ev_key in cfg["evaluators"]:
            store = st.session_state.data.get(ev_type_key, {}).get(ev_key, {})
            total_archivos += len(store)
            total_resp += sum(d["meta"]["respuestas"] for d in store.values())
        st.metric(
            f"{cfg['label']}", f"{total_archivos} archivos", f"{total_resp} respuestas"
        )

# ══════════════════════════════════════════════════════════
# PÁGINA 1 — CARGAR DATOS
# ══════════════════════════════════════════════════════════
if page == "Cargar Datos":
    st.title("Carga de Archivos de Encuesta")
    st.markdown(
        "Selecciona la **persona evaluada** y sube los archivos Excel (.xlsx) exportados desde Moodle."
    )

    # Selector de tipo evaluado
    ev_type_opts = {cfg["label"]: key for key, cfg in EVALUATED_TYPES.items()}
    sel_label = st.segmented_control(
        "Persona evaluada",
        list(ev_type_opts.keys()),
        default=list(ev_type_opts.keys())[0],
        key="upload_type_selector",
    )
    sel_type_key = ev_type_opts[sel_label]
    sel_config = EVALUATED_TYPES[sel_type_key]

    st.markdown(f"**{sel_config['label']}** evaluado por:")

    evaluator_list = get_evaluator_pairs(sel_type_key)
    tab_labels = [f"{cfg['icon']} {cfg['label']}" for _, cfg in evaluator_list]
    tabs = st.tabs(tab_labels)

    for tab_idx, (tab, (ev_key, ev_cfg)) in enumerate(zip(tabs, evaluator_list)):
        with tab:
            st.markdown(
                f"**Evaluación del {ev_cfg['label']} al {sel_config['label']}**"
            )
            st.caption(
                "Escala: Sí -> 3, A veces -> 2, No -> 1"
            )

            store_key = (sel_type_key, ev_key)
            uploaded = st.file_uploader(
                f"Seleccionar archivo(s) de {ev_cfg['label']}:",
                type=["xlsx"],
                accept_multiple_files=True,
                key=f"upload_{sel_type_key}_{ev_key}",
            )

            if uploaded:
                nuevos = 0
                for f in uploaded:
                    if f.name not in st.session_state.data[sel_type_key][ev_key]:
                        try:
                            parsed = parse_evaluacion(
                                f.read(), f.name, sel_type_key, ev_key, ev_cfg
                            )
                            st.session_state.data[sel_type_key][ev_key][f.name] = parsed
                            nuevos += 1
                        except Exception as e:
                            st.error(f"Error en **{f.name}**: {e}")
                if nuevos:
                    st.success(f"{nuevos} archivo(s) de {ev_cfg['label']} cargado(s).")

            store = st.session_state.data[sel_type_key][ev_key]
            if store:
                st.markdown(f"### Archivos de {ev_cfg['label']} cargados")
                for fname, entry in list(store.items()):
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
                        if col2.button(
                            "Borrar", key=f"del_{sel_type_key}_{ev_key}_{fname}"
                        ):
                            del st.session_state.data[sel_type_key][ev_key][fname]
                            st.rerun()

    # Acciones globales por tipo
    st.markdown("---")
    total_files_type = sum(
        len(st.session_state.data[sel_type_key][ev_key])
        for ev_key in sel_config["evaluators"]
    )
    total_resp_type = sum(
        sum(
            d["meta"]["respuestas"]
            for d in st.session_state.data[sel_type_key][ev_key].values()
        )
        for ev_key in sel_config["evaluators"]
    )
    colA, colB = st.columns([3, 2])
    if total_files_type > 0:
        colA.info(
            f"**{total_resp_type}** encuestados en **{total_files_type}** archivos de **{sel_config['label']}**."
        )
    if colB.button(f"Limpiar todo {sel_config['label']}", type="secondary"):
        for ev_key in sel_config["evaluators"]:
            st.session_state.data[sel_type_key][ev_key] = {}
        st.rerun()

# ══════════════════════════════════════════════════════════
# PÁGINA 2 — RESULTADOS INDIVIDUALES
# ══════════════════════════════════════════════════════════
elif page == "Resultados":
    st.title("Resultados de Evaluación")

    # Seleccionar tipo evaluado
    ev_type_opts = {cfg["label"]: key for key, cfg in EVALUATED_TYPES.items()}
    sel_type_label = st.selectbox(
        "Tipo de evaluado:", list(ev_type_opts.keys()), key="res_type"
    )
    sel_type_key = ev_type_opts[sel_type_label]

    # Seleccionar evaluador
    ev_cfgs = EVALUATED_TYPES[sel_type_key]["evaluators"]
    ev_opts = {cfg["label"]: key for key, cfg in ev_cfgs.items()}
    sel_ev_label = st.selectbox("Evaluador:", list(ev_opts.keys()), key="res_eval")
    sel_ev_key = ev_opts[sel_ev_label]

    store = st.session_state.data.get(sel_type_key, {}).get(sel_ev_key, {})
    if not store:
        st.warning(
            f"No hay datos de {sel_type_label} evaluado por {sel_ev_label}. Vaya a **Cargar Datos**."
        )
        st.stop()

    opciones = {label_for(v): k for k, v in store.items()}
    sel_label = st.selectbox(f"Módulo a visualizar:", list(opciones.keys()))
    entry = store[opciones[sel_label]]

    render_individual(entry)

    # Exportar
    st.divider()
    st.subheader("Exportar")
    rows_exp = []
    for q in entry["questions"]:
        total_q = q["Sí"] + q["A veces"] + q["No"] or 1
        rows_exp.append(
            {
                "Pregunta": q["pregunta"],
                "Sí": q["Sí"],
                "%Sí": round(q["Sí"] / total_q * 100, 1),
                "A veces": q["A veces"],
                "%A veces": round(q["A veces"] / total_q * 100, 1),
                "No": q["No"],
                "%No": round(q["No"] / total_q * 100, 1),
                "Puntaje(/3)": q["score"],
            }
        )
    csv = pd.DataFrame(rows_exp).to_csv(index=False).encode("utf-8")
    st.download_button(
        "Descargar resumen (CSV)",
        csv,
        f"resumen_{sel_type_key}_{sel_ev_key}.csv",
        "text/csv",
    )

# ══════════════════════════════════════════════════════════
# PÁGINA 3 — COMPARACIÓN
# ══════════════════════════════════════════════════════════
elif page == "Comparación":
    st.title("Comparación")
    st.markdown(
        "Agrupa evaluadores por módulo, promedia sus puntajes y muestra el desempeño combinado."
    )

    render_comparacion()

# ══════════════════════════════════════════════════════════
# PÁGINA 4 — DIAGNÓSTICO GENERAL
# ══════════════════════════════════════════════════════════
elif page == "Diagnóstico General":
    st.title("Diagnóstico General")
    st.markdown("Dashboard consolidado de todas las evaluaciones cargadas.")

    render_diagnostico()
