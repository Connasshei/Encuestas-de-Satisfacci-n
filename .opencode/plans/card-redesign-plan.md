# Plan: Rediseño de Tarjetas Autocontenidas

## Estado: APROBADO - Listo para implementar

## Cambios a realizar

### 1. Agregar `person_cards` al session_state
- Ubicación: después de `person_registry` init
- Estructura: `{ (ev_type_key, ev_key): ["card_0", "card_1", ...] }`
- Cada tipo+evaluador empieza con una tarjeta "card_0"

### 2. Reescribir Página 1 "Cargar Datos"
- Eliminar el flujo actual de tabs con selector arriba
- Nuevo diseño:
  - Tipo evaluado (segmented control) - se mantiene
  - Tabs de evaluador - se mantienen
  - Dentro de cada tab: tarjetas autocontenidas
  - Cada tarjeta: nombre + file uploader + archivos + borrar
  - Botón "+ Agregar otra persona" al final

### 3. Funciones helper necesarias
- `add_person_card(type_key, ev_key)` - agrega tarjeta
- `remove_person_card(type_key, ev_key, card_id)` - borra tarjeta + archivos

### 4. No cambia
- Página 2 (Resultados)
- Página 3 (Comparación)
- Página 4 (Diagnóstico)
- parse_evaluacion
- person_registry
