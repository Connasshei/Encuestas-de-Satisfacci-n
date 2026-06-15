# 🎓 Evaluación a Docentes — App Streamlit

Herramienta para analizar encuestas de evaluación del rol docente exportadas desde Moodle (formato Excel .xlsx).

---

## ✅ Requisitos previos

- **Python 3.10 o superior** instalado.
  Verifica: `python --version`
  Descarga: https://www.python.org/downloads/

---

## 🚀 Primera vez

```bash
# 1. Entra a la carpeta
cd encuestas_app

# 2. Crea entorno virtual
python -m venv venv

# 3. Actívalo
venv\Scripts\activate          # Windows
source venv/bin/activate       # Mac / Linux

# 4. Instala dependencias (solo una vez)
pip install -r requirements.txt

# 5. Ejecuta
streamlit run app.py
```

Se abre en el navegador en http://localhost:8501

---

## 🔁 Ejecuciones siguientes

```bash
cd encuestas_app
venv\Scripts\activate          # Windows
streamlit run app.py
```

---

## 📁 Flujo de uso

1. **Cargar Datos** → Sube uno o varios `.xlsx` de Moodle.
   La app extrae automáticamente: Módulo, Docente, Carrera, Sede, Fechas, Nº de respuestas.

2. **Ver Resultados** → Selecciona un módulo o "Comparar todos":

   ### Vista por módulo:
   - Ficha con metadatos del módulo
   - KPIs: puntaje promedio, % Excelente+Muy Bueno, etc.
   - Tabla resumen con degradado de color
   - Barras 100% apiladas (Excelente / Muy Bueno / Bueno / Regular / Insuficiente)
   - Barras de puntaje ponderado (escala 1–5)
   - 8 gráficos de torta (uno por pregunta)
   - Radar del perfil del docente
   - Listado de comentarios abiertos
   - Exportar CSV

   ### Vista comparativa (2+ módulos):
   - Tabla cruzada de puntajes
   - Barras agrupadas por pregunta
   - Mapa de calor
   - Radar superpuesto de todos los módulos
   - Ranking global con exportación CSV

---

## 📊 Escala de evaluación

| Nivel        | Puntaje |
|--------------|---------|
| Excelente    | 5       |
| Muy Bueno    | 4       |
| Bueno        | 3       |
| Regular      | 2       |
| Insuficiente | 1       |
