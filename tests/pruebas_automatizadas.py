# -*- coding: utf-8 -*-
"""Pruebas automatizadas de app_completa.py con datos sintéticos.

Ejecuta:  venv\\Scripts\\python tests\\pruebas_automatizadas.py [--solo S1,S2]
"""

from __future__ import annotations

import json
import os
import sys

from streamlit.testing.v1 import AppTest

from generar_datos import (
    generar_excel,
    _preguntas_asignatura_docente,
    _preguntas_carrera_docente,
    _preguntas_docente_estudiante,
    _preguntas_facilitador_supervisor,
)

HERE = os.path.dirname(os.path.abspath(__file__))
APP = os.path.abspath(os.path.join(HERE, "..", "app_completa.py"))
DATOS = os.path.join(HERE, "datos")
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

HALLAZGOS = []
_actual = {"escenario": "", "descripcion": ""}


def escenario(nombre, descripcion):
    _actual["escenario"] = nombre
    _actual["descripcion"] = descripcion


def hallazgo(veredicto, detalle):
    HALLAZGOS.append(
        {
            "escenario": _actual["escenario"],
            "descripcion": _actual["descripcion"],
            "veredicto": veredicto,
            "detalle": detalle,
        }
    )


class AppDriver:
    def __init__(self):
        self.at = AppTest.from_file(APP, default_timeout=120)
        self.at.run()
        self._excepciones = []

    def _run(self, expect_ok=True):
        self.at.run()
        exc = [str(e.value) for e in self.at.exception]
        if exc:
            self._excepciones.extend(exc)

    def ir_a(self, pagina):
        self.at.radio[0].set_value(pagina)
        self._run()

    def set_tipo(self, tipo_label):
        self.at.button_group[0].set_value(tipo_label)
        self._run()

    def add_persona(self, tipo_key, ev_key, nombre, archivos):
        self.at.button(key=f"add_card_{tipo_key}_{ev_key}").click()
        self._run()
        prefix = f"persona_{tipo_key}_{ev_key}_card_"
        cid = max(
            int(sb.key.rsplit("_", 1)[1])
            for sb in self.at.selectbox
            if sb.key.startswith(prefix)
        )
        self.at.selectbox(key=f"{prefix}{cid}").set_value("[Nueva persona]")
        self._run()
        self.at.text_input(
            key=f"nuevo_nombre_{tipo_key}_{ev_key}_card_{cid}"
        ).set_value(nombre)
        self._run()
        payload = [(os.path.basename(f), open(f, "rb").read(), MIME) for f in archivos]
        self.at.file_uploader(key=f"upload_{tipo_key}_{ev_key}_card_{cid}").set_value(
            payload
        )
        self._run()

    @property
    def data(self):
        return self.at.session_state["data"]

    @property
    def excepciones(self):
        return list(self._excepciones)

    def plotly(self):
        figs = []
        for ch in self.at.get("plotly_chart"):
            try:
                figs.append(json.loads(ch.proto.spec))
            except Exception:
                pass
        return figs

    def dataframes(self):
        out = []
        for df in self.at.dataframe:
            v = df.value
            if hasattr(v, "data"):
                v = v.data
            out.append(v)
        return out

    def mensajes(self, tipo):
        return [m.value for m in self.at.get(tipo)]

    def markdowns(self):
        return [m.value for m in self.at.markdown]

    def metricas(self):
        return {m.label: m.value for m in self.at.metric}

    def opciones_selectbox(self, key=None, label=None):
        for sb in self.at.selectbox:
            if key is not None and sb.key != key:
                continue
            if label is not None and getattr(sb, "label", None) != label:
                continue
            return list(sb.options)
        return []


def fig_ranking(figs):
    for f in figs:
        if any(
            t.get("type") == "scatter"
            and t.get("marker", {}).get("symbol") == "diamond"
            for t in f["data"]
        ):
            return f
    return None


def barras_nombres(fig):
    if not fig:
        return []
    return [t.get("name") for t in fig["data"] if t.get("type") == "bar"]


def dataframe_con_columnas(dfs, col):
    for d in dfs:
        if hasattr(d, "columns") and col in d.columns:
            return d
    return None


def total_respuestas(d):
    if d is not None and "Respuestas" in d.columns:
        return int(d["Respuestas"].sum())
    return None


def metric_int(metricas, label):
    try:
        return int(metricas.get(label))
    except (TypeError, ValueError):
        return None


def contar_filas(d):
    return len(d) if d is not None else None


# ─────────────────────────────────────────────────────────────
# ESCENARIOS
# ─────────────────────────────────────────────────────────────


def s1_varios_docentes_una_asignatura():
    """S1: 3 docentes evaluados por estudiantes, misma asignatura."""
    escenario(
        "S1",
        "3 docentes, 1 asignatura (mismo módulo), evaluados por estudiantes",
    )
    nombres = ["Ana Perez", "Luis Gomez", "Maria Diaz"]
    archivos = []
    for i, nombre in enumerate(nombres):
        fp = os.path.join(DATOS, "s1", f"docente_{i + 1}.xlsx")
        generar_excel(
            fp,
            titulo="EVALUACIÓN DE LOS CURSANTES A MÓDULOS DE PROGRAMAS",
            carrera="Primer Año",
            sede="Virtual",
            modulo="FILOSOFÍA (FIL101) - A",
            respuestas=120 + i * 10,
            preguntas=_preguntas_docente_estudiante(nivel=-0.4 + i * 0.35),
            comentarios=["Muy buena clase"],
        )
        archivos.append(fp)

    drv = AppDriver()
    for nombre, fp in zip(nombres, archivos):
        drv.add_persona("docente", "estudiante", nombre, [fp])

    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entradas = drv.data["docente"]["estudiante"]
    if len(entradas) != 3:
        hallazgo("FALLA", f"Se cargaron {len(entradas)} archivos, se esperaban 3")
    for fname, e in entradas.items():
        if len(e["questions"]) != 17:
            hallazgo("FALLA", f"{fname}: {len(e['questions'])} preguntas (esperaba 17)")

    drv.ir_a("Comparación")
    figs = drv.plotly()
    rank = fig_ranking(figs)
    barras = barras_nombres(rank)
    if len(barras) != 3:
        hallazgo("FALLA", f"Ranking con {len(barras)} barras (esperaba 3): {barras}")
    faltan = [n for n in nombres if not any(b.startswith(n) for b in barras)]
    if faltan:
        hallazgo("FALLA", f"Docentes sin barra: {faltan}")
    metricas = drv.metricas()
    n_mod = metric_int(metricas, "Total Módulos Únicos")
    if n_mod != 1:
        hallazgo(
            "FALLA",
            f"Total Módulos Únicos = {n_mod} (esperaba 1)",
        )
    resumen = [m for m in drv.markdowns() if "obtuvo un puntaje promedio" in m]
    if len(resumen) != 3:
        hallazgo(
            "FALLA", f"Resumen de Resultados con {len(resumen)} párrafos (esperaba 3)"
        )

    drv.ir_a("Diagnóstico General")
    dfs = drv.dataframes()
    d_rank = dataframe_con_columnas(dfs, "Respuestas")
    if d_rank is not None:
        n = contar_filas(d_rank)
        if n != 1:
            hallazgo(
                "FALLA",
                f"Diagnóstico muestra {n} filas (esperaba 1: el mismo módulo con 3 "
                "docentes se agrupa en un promedio)",
            )
        total = total_respuestas(d_rank)
        if total != 390:
            hallazgo("FALLA", f"Diagnóstico suma {total} respuestas (esperaba 390)")

    if not any(h["escenario"] == "S1" for h in HALLAZGOS):
        hallazgo("OK", "3 docentes en 1 asignatura: todo correcto.")
    return drv


def s2_docentes_con_carrera():
    """S2: los 3 docentes también evaluados por Equipo de la Carrera."""
    escenario(
        "S2",
        "3 docentes evaluados por estudiantes Y por Equipo de la Carrera (mismo módulo)",
    )
    nombres = ["Ana Perez", "Luis Gomez", "Maria Diaz"]
    est = []
    car = []
    for i, nombre in enumerate(nombres):
        f1 = os.path.join(DATOS, "s2", f"est_{i + 1}.xlsx")
        f2 = os.path.join(DATOS, "s2", f"car_{i + 1}.xlsx")
        generar_excel(
            f1,
            titulo="X",
            carrera="Primer Año",
            sede="Virtual",
            modulo="MATEMÁTICA (MAT101) - A",
            respuestas=100 + i * 20,
            preguntas=_preguntas_docente_estudiante(nivel=-0.4 + i * 0.35),
        )
        generar_excel(
            f2,
            titulo="X",
            carrera="Primer Año",
            sede="Virtual",
            modulo="MATEMÁTICA (MAT101) - A",
            respuestas=5 + i,
            preguntas=_preguntas_carrera_docente(nivel=-0.3 + i * 0.3),
        )
        est.append(f1)
        car.append(f2)

    drv = AppDriver()
    for nombre, fp in zip(nombres, est):
        drv.add_persona("docente", "estudiante", nombre, [fp])
    for nombre, fp in zip(nombres, car):
        drv.add_persona("docente", "carrera", nombre, [fp])

    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    drv.ir_a("Comparación")
    rank = fig_ranking(drv.plotly())
    barras = barras_nombres(rank)
    if len(barras) != 6:
        hallazgo("FALLA", f"Ranking con {len(barras)} barras (esperaba 6): {barras}")
    if (
        "Ana Perez (Estudiante)" not in barras
        or "Ana Perez (Equipo de la Carrera)" not in barras
    ):
        hallazgo("FALLA", f"No aparecen ambos evaluadores por persona: {barras}")

    drv.ir_a("Diagnóstico General")
    d_rank = dataframe_con_columnas(drv.dataframes(), "Respuestas")
    if d_rank is not None and contar_filas(d_rank) != 2:
        hallazgo(
            "FALLA",
            f"Diagnóstico muestra {contar_filas(d_rank)} filas (esperaba 2: un "
            "promedio por módulo y evaluador)",
        )

    if not any(h["escenario"] == "S2" for h in HALLAZGOS):
        hallazgo("OK", "Docente con doble evaluador por módulo: 6 barras correctas.")
    return drv


def s3_varios_facilitadores_una_asignatura():
    """S3: 3 facilitadores evaluados por supervisor, misma asignatura."""
    escenario(
        "S3",
        "3 facilitadores, 1 asignatura (mismo módulo), evaluados por supervisor",
    )
    nombres = ["Fac 1 Carlos", "Fac 2 Rosa", "Fac 3 Juan"]
    archivos = []
    for i, nombre in enumerate(nombres):
        fp = os.path.join(DATOS, "s3", f"fac_{i + 1}.xlsx")
        generar_excel(
            fp,
            titulo="EVALUACIÓN DEL SUPERVISOR AL FACILITADOR",
            carrera="Segundo Año",
            sede="Virtual",
            modulo="MICROECONOMÍA (GSC212) - A",
            respuestas=8 + i,
            preguntas=_preguntas_facilitador_supervisor(nivel=-0.3 + i * 0.3),
        )
        archivos.append(fp)

    drv = AppDriver()
    drv.set_tipo("Facilitador")
    for nombre, fp in zip(nombres, archivos):
        drv.add_persona("facilitador", "supervisor", nombre, [fp])

    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    drv.ir_a("Comparación")
    barras = barras_nombres(fig_ranking(drv.plotly()))
    if len(barras) != 3:
        hallazgo("FALLA", f"Ranking con {len(barras)} barras (esperaba 3): {barras}")
    faltan = [n for n in nombres if not any(b.startswith(n) for b in barras)]
    if faltan:
        hallazgo("FALLA", f"Facilitadores sin barra: {faltan}")

    if not any(h["escenario"] == "S3" for h in HALLAZGOS):
        hallazgo("OK", "3 facilitadores en 1 asignatura: todo correcto.")
    return drv


def s4_asignatura_evaluada_por_3_docentes():
    """S4: 1 asignatura evaluada por 3 docentes."""
    escenario(
        "S4",
        "1 asignatura (tipo Asignatura) evaluada por 3 docentes",
    )
    nombres = ["Doc Ana", "Doc Luis", "Doc Maria"]
    archivos = []
    for i, nombre in enumerate(nombres):
        fp = os.path.join(DATOS, "s4", f"asig_{i + 1}.xlsx")
        generar_excel(
            fp,
            titulo="CUIESTIONARIO DE EVALUACIÓN DEL DOCENTE A LA ASIGNATURA",
            carrera="Segundo Año",
            sede="Virtual",
            modulo="MACROECONOMÍA (GSC211) - A",
            respuestas=1,
            preguntas=_preguntas_asignatura_docente(nivel=-0.3 + i * 0.3),
            fila_comentarios="¿Qué aspectos considera que deben mejorarse?.",
        )
        archivos.append(fp)

    drv = AppDriver()
    drv.set_tipo("Asignatura")
    for nombre, fp in zip(nombres, archivos):
        drv.add_persona("asignatura", "docente", nombre, [fp])

    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entradas = drv.data["asignatura"]["docente"]
    if len(entradas) != 3:
        hallazgo("FALLA", f"Se cargaron {len(entradas)} archivos (esperaba 3)")
    for fname, e in entradas.items():
        if len(e["questions"]) != 14:
            hallazgo(
                "FALLA",
                f"{fname}: {len(e['questions'])} preguntas (esperaba 14)",
            )

    drv.ir_a("Comparación")
    barras = barras_nombres(fig_ranking(drv.plotly()))
    if len(barras) != 3:
        hallazgo("FALLA", f"Ranking con {len(barras)} barras (esperaba 3): {barras}")

    if not any(h["escenario"] == "S4" for h in HALLAZGOS):
        hallazgo("OK", "Asignatura evaluada por 3 docentes: 3 barras correctas.")
    return drv


def s5_archivo_duplicado_mismo_contenido():
    """S5: mismo archivo/contenido subido con 2 nombres de archivo."""
    escenario(
        "S5",
        "Mismo facilitador con 2 archivos de contenido idéntico (mismo módulo y respuestas)",
    )
    preguntas = _preguntas_facilitador_supervisor(nivel=0.1)
    f1 = os.path.join(DATOS, "s5", "jhomara_sociologia_a.xlsx")
    f2 = os.path.join(DATOS, "s5", "jhomara_sociologia_b.xlsx")
    for fp in (f1, f2):
        generar_excel(
            fp,
            titulo="X",
            carrera="Primer Año",
            sede="Virtual",
            modulo="SOCIOLOGÍA (GBJ108) - A",
            respuestas=82,
            preguntas=preguntas,
        )

    drv = AppDriver()
    drv.set_tipo("Facilitador")
    drv.add_persona("facilitador", "supervisor", "Jhomara Rojas", [f1, f2])

    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entradas = drv.data["facilitador"]["supervisor"]
    if len(entradas) != 1:
        hallazgo(
            "FALLA",
            f"Se cargaron {len(entradas)} archivos (esperaba 1; el duplicado debe "
            "omitirse para no inflar los resultados)",
        )
    avisos = drv.mensajes("warning") + drv.mensajes("info")
    if len(entradas) == 1 and not any("duplicado" in a.lower() for a in avisos):
        hallazgo("FALLA", "No se mostró aviso de archivo duplicado.")

    drv.ir_a("Comparación")
    barras = barras_nombres(fig_ranking(drv.plotly()))
    dupes = [b for b in set(barras) if barras.count(b) > 1]
    if dupes:
        hallazgo(
            "REPETIDO",
            f"Barras duplicadas en ranking para la misma persona/módulo: {dupes}",
        )

    drv.ir_a("Diagnóstico General")
    d_rank = dataframe_con_columnas(drv.dataframes(), "Respuestas")
    total = total_respuestas(d_rank)
    if total != 82:
        hallazgo(
            "FALLA",
            f"Diagnóstico suma {total} respuestas (esperaba 82, sin inflar por "
            "archivo duplicado)",
        )

    if not any(h["escenario"] == "S5" for h in HALLAZGOS):
        hallazgo("OK", "Archivo duplicado detectado y omitido con aviso.")
    return drv


def s6_mismo_nombre_archivo_2_personas():
    """S6: mismo nombre de archivo subido para 2 personas distintas."""
    escenario(
        "S6",
        "Mismo nombre de archivo (conflicto.xlsx) subido a 2 personas distintas",
    )
    f1 = os.path.join(DATOS, "s6a", "conflicto.xlsx")
    f2 = os.path.join(DATOS, "s6b", "conflicto.xlsx")
    generar_excel(
        f1,
        titulo="X",
        carrera="Primer Año",
        sede="Virtual",
        modulo="MÓDULO A",
        respuestas=50,
        preguntas=_preguntas_docente_estudiante(nivel=0.2),
    )
    generar_excel(
        f2,
        titulo="X",
        carrera="Segundo Año",
        sede="Virtual",
        modulo="MÓDULO B",
        respuestas=60,
        preguntas=_preguntas_docente_estudiante(nivel=-0.2),
    )

    drv = AppDriver()
    drv.add_persona("docente", "estudiante", "Ana Perez", [f1])
    drv.add_persona("docente", "estudiante", "Luis Gomez", [f2])

    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entradas = drv.data["docente"]["estudiante"]
    nombres_cargados = [e["meta"]["persona_nombre"] for e in entradas.values()]
    if len(entradas) != 2:
        hallazgo(
            "FALLA",
            f"Se esperaban 2 archivos almacenados (uno por persona); hay "
            f"{len(entradas)}. Claves: {list(entradas)}. "
            f"Personas esperadas: Ana Perez y Luis Gomez; cargadas: {nombres_cargados}",
        )
    else:
        if "Ana Perez" not in nombres_cargados or "Luis Gomez" not in nombres_cargados:
            hallazgo("FALLA", f"Personas faltantes: {nombres_cargados}")
    advertencias = drv.mensajes("warning") + drv.mensajes("info")
    if len(entradas) == 2 and not any("conflicto" in a.lower() for a in advertencias):
        hallazgo(
            "FALLA",
            "No se mostró ninguna advertencia sobre el nombre de archivo repetido "
            "entre dos personas distintas.",
        )

    if not any(h["escenario"] == "S6" for h in HALLAZGOS):
        hallazgo(
            "OK", "Mismo nombre de archivo en 2 personas: ambos almacenados con aviso."
        )
    return drv


def s7_cuestionario_sin_numeros():
    """S7: cuestionario sin numeración (estilo Filosofía MÓDULO)."""
    escenario(
        "S7",
        "Cuestionario sin números de pregunta y con orden/texto distinto al config",
    )
    textos = [
        "¿El/la docente demostró dominio de los contenidos?",
        "¿La explicación del docente fue clara?",
        "¿El/la docente promovió la participación de los estudiantes?",
        "¿El/la docente utilizó ejemplos de la realidad profesional?",
        "¿El/la docente respondió con claridad a las dudas?",
        "¿El docente proporciona bibliografía actualizada?",
        "¿El docente comparte presentaciones, infografías y recursos?",
        "¿Las orientaciones y consignas para las actividades fueron claras?",
        "¿Considera que los objetivos del curso fueron alcanzados?",
        "¿Los contenidos del curso fueron interesantes y pertinentes?",
        "¿La metodología y formas de trabajo empleadas fueron adecuadas?",
        "¿El docente fomenta la participación activa en foros?",
        "¿El docente atiende de manera oportuna las consultas?",
        "¿El docente brinda retroalimentación constructiva?",
        "¿El docente evalúa de manera justa y transparente?",
        "¿El docente entrega calificaciones en los plazos establecidos?",
        "¿El curso contó con recursos digitales de calidad?",
    ]
    preguntas = [(t, 60 + i, 30 - i // 2, 10) for i, t in enumerate(textos)]
    fp = os.path.join(DATOS, "s7", "sin_numeros.xlsx")
    generar_excel(
        fp,
        titulo="EVALUACIÓN DE LOS CURSANTES A MÓDULOS",
        carrera="Primer Año",
        sede="Virtual",
        modulo="FILOSOFÍA (FIL101) - A",
        respuestas=100,
        preguntas=preguntas,
        numeradas=False,
    )

    drv = AppDriver()
    drv.add_persona("docente", "estudiante", "Ana Perez", [fp])
    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entry = drv.data["docente"]["estudiante"]["sin_numeros.xlsx"]
    advertencias = entry["meta"].get("advertencias", [])
    if len(entry["questions"]) != 17:
        hallazgo(
            "FALLA", f"Se parsearon {len(entry['questions'])} preguntas (esperaba 17)"
        )
    if not any("numeración" in a or "no coincide" in a for a in advertencias):
        hallazgo(
            "FALLA",
            "No se generó ninguna advertencia para un cuestionario sin numeración "
            "cuyo texto no coincide con la plantilla esperada.",
        )
    else:
        hallazgo(
            "OK",
            "Cuestionario sin numeración/no coincidente: se advierte al usuario sobre "
            "el posible desalineamiento de etiquetas y dimensiones.",
        )
    return drv


def s8_mas_preguntas_que_config():
    """S8: cuestionario con 18 preguntas únicas (config soporta 17)."""
    escenario(
        "S8",
        "Cuestionario con 18 preguntas únicas para docente/estudiante (config=17)",
    )
    textos = _preguntas_docente_estudiante(nivel=0.0)
    textos = [t[0] for t in textos] + ["Pregunta extra número 18 de la encuesta."]
    preguntas = [(t, 80 - i, 20 + i // 2, 5) for i, t in enumerate(textos)]
    fp = os.path.join(DATOS, "s8", "18_preguntas.xlsx")
    generar_excel(
        fp,
        titulo="X",
        carrera="Primer Año",
        sede="Virtual",
        modulo="MÓDULO X",
        respuestas=105,
        preguntas=preguntas,
    )

    drv = AppDriver()
    drv.add_persona("docente", "estudiante", "Ana Perez", [fp])
    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entry = drv.data["docente"]["estudiante"]["18_preguntas.xlsx"]
    advertencias = entry["meta"].get("advertencias", [])
    nums = [q["num"] for q in entry["questions"]]
    if len(entry["questions"]) != 17:
        hallazgo(
            "FALLA",
            f"Se guardaron {len(entry['questions'])} preguntas (esperaba 17 por config)",
        )
    if "Pregunta extra número 18" not in [q["pregunta"] for q in entry["questions"]]:
        if not any("descartaron" in a for a in advertencias):
            hallazgo(
                "FALLA",
                "La pregunta 18 se descartó sin ninguna advertencia al usuario.",
            )
        else:
            hallazgo(
                "OK",
                "Pregunta extra descartada (config=17) pero con advertencia al usuario.",
            )

    if not any(h["escenario"] == "S8" for h in HALLAZGOS):
        hallazgo("OK", "Número de preguntas dentro de lo esperado.")
    return drv


def s9_comentarios_basura():
    """S9: comentarios irrelevantes que deberían filtrarse."""
    escenario(
        "S9",
        "Comentarios abiertos con texto irrelevante (Ningno, Bien, Conforme, nada, 0)",
    )
    preguntas = _preguntas_facilitador_supervisor(nivel=0.0)[:8]
    fp = os.path.join(DATOS, "s9", "comentarios.xlsx")
    generar_excel(
        fp,
        titulo="X",
        carrera="Primer Año",
        sede="Virtual",
        modulo="SOCIOLOGÍA (GBJ108) - A",
        respuestas=30,
        preguntas=preguntas,
        fila_comentarios="9. Otras observaciones y comentarios:",
        comentarios=[
            "Ningno",
            "Bien",
            "Conforme",
            "nada",
            "0",
            "Ninguna observación",
            "Muy buena facilitadora",
        ],
    )

    drv = AppDriver()
    drv.set_tipo("Facilitador")
    drv.add_persona("facilitador", "supervisor", "Carlos Ruiz", [fp])
    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    entry = drv.data["facilitador"]["supervisor"]["comentarios.xlsx"]
    comentarios = entry["comments"]
    ruido = [
        c
        for c in comentarios
        if c.lower() in ("ningno", "bien", "conforme", "ninguna observacion")
        or len(c) < 4
    ]
    if ruido:
        hallazgo(
            "INNECESARIO",
            f"Aparecen como comentarios textos irrelevantes: {ruido}. "
            "El filtro no cubre variantes como 'Ningno', 'Bien', 'Conforme'.",
        )

    if not any(h["escenario"] == "S9" for h in HALLAZGOS):
        hallazgo("OK", "Filtro de comentarios sin ruido.")
    return drv


def s10_modulos_casi_identicos():
    """S10: módulos casi idénticos con distinto texto."""
    escenario(
        "S10",
        "Módulos casi idénticos: 'FILOSOFÍA 2026', 'Filosofia 2026', 'FILOSOFÍA 2026 (tarde)'",
    )
    combos = [
        ("Ana Perez", "FILOSOFÍA 2026"),
        ("Luis Gomez", "Filosofia 2026"),
        ("Maria Diaz", "FILOSOFÍA 2026 (tarde)"),
    ]
    archivos = []
    for i, (nombre, mod) in enumerate(combos):
        fp = os.path.join(DATOS, "s10", f"m{i + 1}.xlsx")
        generar_excel(
            fp,
            titulo="X",
            carrera="Primer Año",
            sede="Virtual",
            modulo=mod,
            respuestas=100,
            preguntas=_preguntas_docente_estudiante(nivel=0.1),
        )
        archivos.append(fp)

    drv = AppDriver()
    for (nombre, _mod), fp in zip(combos, archivos):
        drv.add_persona("docente", "estudiante", nombre, [fp])
    if drv.excepciones:
        hallazgo("FALLA", f"Excepción al cargar: {drv.excepciones[0][:200]}")
        return drv

    drv.ir_a("Comparación")
    metricas = drv.metricas()
    n_mod = metric_int(metricas, "Total Módulos Únicos")
    if n_mod != 2:
        hallazgo("FALLA", f"Total Módulos Únicos = {n_mod} (esperaba 2)")
    else:
        hallazgo(
            "RIESGO",
            "'FILOSOFÍA 2026' y 'Filosofia 2026' se agrupan (bien), pero "
            "'FILOSOFÍA 2026 (tarde)' se separa. Cualquier diferencia de texto en "
            "MODULO (sufijos, paréntesis, grupos) fragmenta el mismo módulo en "
            "grupos separados de la Comparación.",
        )
    return drv


# ─────────────────────────────────────────────────────────────
# EJECUCIÓN Y REPORTE
# ─────────────────────────────────────────────────────────────

ESCENARIOS = [
    s1_varios_docentes_una_asignatura,
    s2_docentes_con_carrera,
    s3_varios_facilitadores_una_asignatura,
    s4_asignatura_evaluada_por_3_docentes,
    s5_archivo_duplicado_mismo_contenido,
    s6_mismo_nombre_archivo_2_personas,
    s7_cuestionario_sin_numeros,
    s8_mas_preguntas_que_config,
    s9_comentarios_basura,
    s10_modulos_casi_identicos,
]


def main():
    os.makedirs(DATOS, exist_ok=True)
    solo = None
    if "--solo" in sys.argv:
        idx = sys.argv.index("--solo")
        solo = {x.strip().upper() for x in sys.argv[idx + 1].split(",")}

    for fn in ESCENARIOS:
        nombre = fn.__name__.split("_")[0].upper()
        if solo and nombre not in solo:
            continue
        print(f"▶ Ejecutando {nombre} ...", flush=True)
        try:
            fn()
        except Exception as e:  # noqa: BLE001
            hallazgo("FALLA", f"El escenario {nombre} abortó con excepción: {e!r}")

    print("\n" + "=" * 90)
    print("REPORTE DE PRUEBAS")
    print("=" * 90)
    for h in HALLAZGOS:
        print(f"[{h['veredicto']}] {h['escenario']} — {h['descripcion']}")
        print(f"      {h['detalle']}")
    print("=" * 90)

    resumen = {}
    for h in HALLAZGOS:
        resumen[h["veredicto"]] = resumen.get(h["veredicto"], 0) + 1
    print("Resumen:", ", ".join(f"{k}: {v}" for k, v in resumen.items()))

    with open(os.path.join(HERE, "reporte_pruebas.json"), "w", encoding="utf-8") as f:
        json.dump(
            {"resumen": resumen, "hallazgos": HALLAZGOS},
            f,
            ensure_ascii=False,
            indent=2,
        )
    print("Reporte guardado en tests/reporte_pruebas.json")


if __name__ == "__main__":
    main()
