# Visor PEO – Huachipa 8

App de Streamlit para visualizar y actualizar el plan PEO desde un Excel alojado en OneDrive/SharePoint **sincronizado** en la PC.

## Requisitos
- Python 3.10+
- Acceso local a la biblioteca de OneDrive/SharePoint (sincronizada)

## Instalación
```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
# source .venv/bin/activate

pip install -r requirements.txt
```

## Configuración
1. Copia `.env.example` a `.env` y ajusta:
   - `EXCEL_PATH` → ruta local al Excel (no es un link https).
   - `EXCEL_SHEET` → 0 ó nombre de hoja.
   - `UPLOAD_DIR` → carpeta existente para evidencias.
2. Asegúrate de que el Excel esté **cerrado** cuando presiones “Guardar cambios” en la app.

## Ejecución
```bash
streamlit run app.py
```

## Notas
- La app sobrescribe el Excel y pinta en **morado** las celdas cambiadas.
- Si falla por permisos, cierra el Excel abierto y vuelve a intentar.
- No subas `.env` al repositorio, usa `.env.example`.