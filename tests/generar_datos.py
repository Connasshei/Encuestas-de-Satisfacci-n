# -*- coding: utf-8 -*-
"""Generador de archivos Excel sintéticos con formato Moodle (export feedback).

Replica la estructura que parsea app_completa.py:
  fila 3..11  -> CARRERA / SEDE / MODULO / GRUPO / DOCENTE / FECHAS / RESPUESTAS
  fila 14     -> cabecera "Etiqueta" con columnas Excelentes/% Muy Buenos/% ...
  fila 15+    -> preguntas  (si, a veces, no) en columnas 2,4,6
                 (regular/insuficiente en 8,10 opcional)
  fila final  -> fila de comentarios abiertos (col 1 texto, col 2+ comentarios)
"""

from __future__ import annotations

import os

import openpyxl

HEADER_STRUCT = [
    "Etiqueta",
    "Pregunta",
    "Excelentes",
    "%",
    "Muy Buenos",
    "%",
    "Bueno",
    "%",
    "Regular",
    "%",
    "Insuficiente",
    "%",
]


def _set_row(ws, idx, values):
    for j, v in enumerate(values):
        if v is not None:
            ws.cell(row=idx, column=j + 1, value=v)


def generar_excel(
    path: str,
    *,
    titulo: str,
    carrera: str,
    sede: str,
    modulo: str,
    grupo: str = "",
    docente: str = "",
    fecha_inicio: str = "Monday, 30 de March de 2026, 00:00",
    fecha_fin: str = "Sunday, 7 de June de 2026, 23:55",
    respuestas: int,
    preguntas: list[tuple],
    numeradas: bool = True,
    cabecera_preguntas: str = "Etiqueta",
    fila_comentarios: str | None = "Otras observaciones y comentarios:",
    comentarios: list[str] | None = None,
) -> str:
    """Genera un .xlsx.

    preguntas: lista de tuplas (texto, si, aveces, no[, regular=0[, insuficiente=0]])
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    _set_row(ws, 1, [titulo])
    _set_row(ws, 3, [f"CARRERA: {carrera}"])
    _set_row(ws, 4, [f"SEDE: {sede}"])
    _set_row(ws, 5, [f"MODULO: {modulo}"])
    _set_row(ws, 6, [f"GRUPO: {grupo}"])
    _set_row(ws, 7, [f"DOCENTE: {docente}"])
    _set_row(ws, 8, [f"FECHA INICIO: {fecha_inicio}"])
    _set_row(ws, 9, [f"FECHA FIN: {fecha_fin}"])
    _set_row(ws, 10, [f"RESPUESTAS ENVIADAS: {respuestas}"])
    _set_row(ws, 11, ["Wednesday, 3 de June de 2026, 15:26"])
    _set_row(ws, 12, [f"Preguntas: {len(preguntas)}"])
    _set_row(ws, 13, [""])
    _set_row(ws, 14, HEADER_STRUCT)

    r = 15
    for i, q in enumerate(preguntas):
        texto, si, av, no = q[0], q[1], q[2], q[3]
        regular = q[4] if len(q) > 4 else None
        insuf = q[5] if len(q) > 5 else None
        prefijo = f"{i + 1}." if numeradas else ""
        _set_row(
            ws,
            r,
            [
                None,
                f"{prefijo} {texto}".strip(),
                si or None,
                None,
                av or None,
                None,
                no or None,
                None,
                regular or None,
                None,
                insuf or None,
                None,
            ],
        )
        r += 1

    if fila_comentarios is not None:
        _set_row(ws, r, [None, fila_comentarios, None])
        r += 1
        for c in comentarios or []:
            _set_row(ws, r, [None, None, c])
            r += 1

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    return path


def _nivel_pct(nivel: float, base: float) -> float:
    """Desplaza el % de 'Sí' según el nivel de desempeño (-0.5..0.5)."""
    pct = base + nivel * 0.08
    return max(0.30, min(0.95, pct))


def _preguntas_docente_estudiante(nivel: float = 0.0) -> list[tuple]:
    """17 preguntas estándar docente/estudiante con respuestas variables."""
    textos = [
        "¿El programa, los objetivos y el cronograma se presentaron con claridad?",
        "¿Los contenidos, actividades y recursos de aprendizaje se presentaron organizados?",
        "¿La bibliografía proporcionada fue útil para su aprendizaje?",
        "¿El docente inicia puntual las clases?",
        "¿El/la docente demostró dominio de contenidos durante la asignatura?",
        "¿Las explicaciones del/de la docente fueron claras, didácticas y coherentes?",
        "¿El/la docente utilizó estrategias que promovieron la participación?",
        "¿El/la docente relacionó los contenidos con ejemplos o situaciones de la profesión?",
        "¿El/la docente atendió dudas o dificultades académicas durante la clase?",
        "¿El/la docente promovió un clima de respeto, motivación y confianza?",
        "¿El/la docente brindó orientaciones sobre el contenido durante el proceso?",
        "¿Las evaluaciones fueron coherentes con los contenidos desarrollados?",
        "¿El/la docente explicó las consignas, instrucciones y criterios de evaluación?",
        "¿El/la docente brindó retroalimentación durante el proceso de aprendizaje?",
        "¿Las evaluaciones incluyeron actividades de aplicación práctica de los contenidos?",
        "¿El acceso y navegación en el aula virtual fueron adecuados?",
        "¿La plataforma utilizada permitió el desarrollo de la asignatura?",
    ]
    out = []
    for i, t in enumerate(textos):
        pct = _nivel_pct(nivel, (0.70 + 0.015 * i) % 0.95)
        si = int(120 * pct)
        av = int(120 * (0.85 - pct))
        no = max(0, 120 - si - av)
        out.append((t, si, av, no))
    return out


def _preguntas_carrera_docente(nivel: float = 0.0) -> list[tuple]:
    textos = [
        "¿El docente comunicó los objetivos de la asignatura?",
        "¿Los contenidos se desarrollaron conforme lo planificado?",
        "¿El orden lógico de los contenidos fue el adecuado?",
        "¿El docente atendió inquietudes del equipo de carrera?",
        "¿Las consignas entregadas fueron claras?",
        "¿El docente usó los recursos digitales institucionales?",
        "¿Promovió un ambiente participativo?",
        "¿Cumplió con el horario establecido?",
        "¿Brindó retroalimentación oportuna?",
        "¿Coordinó con los facilitadores de la asignatura?",
        "¿Consideró la planificación universitaria?",
        "¿El contenido estuvo disponible en la plataforma?",
        "¿Realizó seguimiento del uso de la plataforma?",
    ]
    out = []
    for i, t in enumerate(textos):
        base = 0.80 - 0.02 * i
        si = int(40 * _nivel_pct(nivel, base))
        av = max(0, min(40, int(40 * (0.90 - _nivel_pct(nivel, base)))))
        no = max(0, 40 - si - av)
        out.append((t, si, av, no))
    return out


def _preguntas_facilitador_supervisor(nivel: float = 0.0) -> list[tuple]:
    textos = [
        "¿El facilitador brindó apoyo oportuno al docente?",
        "¿El facilitador mantuvo coordinación con el docente?",
        "¿El facilitador coordinó oportunamente las actividades?",
        "¿El facilitador realizó seguimiento de la participación?",
        "¿El facilitador identificó estudiantes en riesgo?",
        "¿El facilitador cumplió con los plazos de calificación?",
        "¿El facilitador mantuvo actualizado el registro académico?",
        "¿El facilitador mostró responsabilidad y compromiso?",
        "¿El facilitador participó activamente en reuniones de coordinación?",
        "¿El facilitador fomentó la reflexión entre los docentes?",
        "¿El facilitador promovió una comunicación efectiva?",
    ]
    out = []
    for i, t in enumerate(textos):
        base = 0.80 - 0.015 * i
        si = int(8 * _nivel_pct(nivel, base))
        av = max(0, min(8, int(8 * (0.95 - _nivel_pct(nivel, base)))))
        no = max(0, 8 - si - av)
        out.append((t, si, av, no))
    return out


def _preguntas_asignatura_docente(nivel: float = 0.0) -> list[tuple]:
    textos = [
        "¿Usted como docente participó en la inducción de la asignatura?",
        "¿En la reunión de inducción se le brindó la información necesaria?",
        "¿Utilizó la propuesta de programa de asignatura?",
        "¿La planificación y lineamientos facilitaron su trabajo?",
        "¿El acompañamiento del supervisor fue oportuno?",
        "¿La comunicación con el o la supervisora fue clara?",
        "¿La comunicación con el o la facilitadora fue efectiva?",
        "¿El apoyo brindado por el o la facilitadora fue suficiente?",
        "¿Logró desarrollar la asignatura conforme lo planificado?",
        "¿Usted cumplió con todos los requisitos académicos?",
        "¿La plataforma y herramientas digitales facilitaron su labor?",
        "¿El soporte técnico e institucional fue oportuno?",
        "¿Considera que los estudiantes lograron los resultados esperados?",
        "En general, ¿su experiencia docente en la asignatura fue satisfactoria?",
    ]
    out = []
    for i, t in enumerate(textos):
        pct = 0.70 + 0.02 * i
        si = 1 if _nivel_pct(nivel, pct) > 0.75 else 0
        av = 0 if si else 1
        no = 0
        out.append((t, si, av, no))
    return out
