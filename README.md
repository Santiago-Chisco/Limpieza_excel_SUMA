# 🧹 Informe General → MATRIZ

1. Instalar: `pip install -r requirements.txt`
2. Ejecutar: `streamlit run app.py` (abre http://localhost:8501)
3. Subir el Excel con filas en rojo/rosado.
4. La app elimina las rojas, conserva las demás y descarga `MATRIZ_<nombre>.xlsx`
   con las columnas de la MATRIZ y el formato verde (título, "Fecha documento", encabezado verde).

Archivos: `app.py` (interfaz), `informe.py` (lógica rojas -> MATRIZ), `cleaner.py` (archivos dañados),
`render.yaml` (despliegue en Render).
