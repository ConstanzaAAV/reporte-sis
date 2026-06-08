# Reporte SIS → Excel

Herramienta web interna para procesar el **Reporte SIS** (Plan de Intervención · Mejor Niñez).
Subes el PDF del mes y descargas un Excel con los datos ordenados y gráficos de análisis:
visitas reales, tipos de intervención y de evento, y un panel interactivo por niño.

## ⚠️ Datos sensibles
El PDF contiene datos de niños, niñas y adolescentes. **Nunca** se suben PDF ni Excel a este
repositorio (ver `.gitignore`). El acceso a la app está protegido con contraseña.

## ¿Qué hace?
1. Lee cada página del PDF y extrae: Nombre, Tipo Intervención, Fecha, Tipo Evento, Descripción, Técnico.
2. Une descripciones que continúan entre páginas y elimina reimpresiones duplicadas.
3. Calcula las **visitas reales** (niño + fecha + técnico) y arma las hojas:
   `Eventos`, `Resumen por niño`, `Panel` (interactivo) y `General` (gráficos).

## Usar localmente
```bash
pip install -r requirements.txt
# crear .streamlit/secrets.toml con:  password = "TU_CLAVE"
streamlit run app.py
```

## Publicar gratis (Streamlit Community Cloud)
1. Sube este repo a GitHub (sin datos).
2. Entra a https://share.streamlit.io → New app → elige el repo y `app.py`.
3. En **Advanced settings → Secrets** agrega:
   ```toml
   password = "TU_CLAVE"
   ```
4. Deploy. Comparte el link y la contraseña solo con quien corresponda.

## Archivos
- `app.py` — página web (Streamlit).
- `procesar.py` — motor: PDF → Excel (toda la lógica).
- `requirements.txt` — dependencias.
