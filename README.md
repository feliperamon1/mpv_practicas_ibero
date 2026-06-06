# Dashboard SG-SST IMA Company SAS - Streamlit

Aplicación Streamlit para visualizar los instrumentos digitales del SG-SST construidos en Excel/Google Sheets: presupuesto, formación, autodiagnóstico Res. 0312/2019, plan anual PHVA, matriz legal, reportes laborales, proveedores, EMO, accidentalidad, enfermedad laboral, indicadores, matriz de peligros GTC-45, mantenimiento preventivo y APCM.

## Estructura recomendada del repositorio

```text
streamlit_sgsst_dashboard/
├── app.py
├── requirements.txt
├── README.md
├── .streamlit/
│   └── config.toml
└── data/
    ├── Autodiagnostico_Res_0312_2019_2023_2025.xlsx
    ├── Base Reporte Laboral Autogestión - IMA Company SAS (1).xlsx
    ├── Formato_Asignacion_Recursos_Presupuesto_SST_2023_2026.xlsx
    ├── Gestion_Proveedores_Contratistas_Vigilancia_2023_2026.xlsx
    ├── Matriz_APCM_Vigilancia_2023_2026.xlsx
    ├── Matriz_Accidentalidad_2023_2026.xlsx
    ├── Matriz_Enfermedad_Laboral_2023_2026.xlsx
    ├── Matriz_Indicadores_Gestion_SGSST_2023_2026.xlsx
    ├── Matriz_Legal_SGRL_SST_Colombia.xlsx
    ├── Matriz_Peligros_Riesgos_GTC45_Vigilancia.xlsx
    ├── Matriz_Mantenimiento_Preventivo_Vigilancia_2023_2026.xlsx
    ├── Matriz_Seguimiento_EMO_2023_2026.xlsx
    ├── Plan_Anual_Trabajo_PHVA_2023_2026.xlsx
    └── Plan_Formacion_y_Capacitaciones_2023_2026.xlsx
```

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Publicación en Streamlit Community Cloud

1. Crear un repositorio en GitHub y subir todos los archivos de esta carpeta.
2. Ingresar a https://share.streamlit.io/ con la cuenta de GitHub.
3. Seleccionar el repositorio, rama y archivo principal `app.py`.
4. Deploy.

## Nota de protección de datos

Antes de subir el repositorio a GitHub o publicarlo en Streamlit Community Cloud, anonimiza datos personales, documentos de identidad, teléfonos, correos, diagnósticos, restricciones médicas y cualquier información sensible. Si el repositorio es público, los archivos de la carpeta `data/` también pueden quedar expuestos.
