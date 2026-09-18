# ES2027 MVP — Generador de Informes

MVP monolítico para:

1. importar el cronograma Excel,
2. detectar jornadas/recorridos,
3. completar calles y observaciones,
4. registrar si hubo artículos promocionales,
5. generar un informe Word.

## Stack

- FastAPI
- Jinja2
- SQLite
- OpenPyXL
- python-docx
- HTML/CSS/JS
- Un solo repositorio y un solo deploy

## Ejecutar en GitHub Codespaces

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

Abre el puerto 8000 desde la pestaña **Ports** de Codespaces.

## Flujo

- `/` → importar cronograma.
- `/cronograma/{id}` → revisar jornadas detectadas.
- `/informe/{id}` → completar recorridos.
- `/informe/{id}/word` → generar `.docx`.

## Cronograma compatible

El importador está preparado para el archivo de agosto ES2027:

- fila 11: encabezados,
- datos desde fila 13,
- columnas:
  - A funcionario,
  - B día,
  - C fecha,
  - D hora,
  - E provincia,
  - F cantón,
  - G parroquia,
  - H sector.

También soporta filas donde fecha/provincia/cantón vienen vacías y deben heredarse de la fila anterior.

## Importante

La generación Word incluida es funcional para el MVP, pero aún no replica píxel por píxel
la plantilla institucional. El siguiente paso será mapear el `MODELO INFORME ES2027.docx`
para conservar exactamente encabezado, pie, tablas y estilos oficiales.
