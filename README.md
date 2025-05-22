[![DeepWiki](https://img.shields.io/badge/DeepWiki-junior19a2000%2FSTE-blue.svg?logo=data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACwAAAAyCAYAAAAnWDnqAAAAAXNSR0IArs4c6QAAA05JREFUaEPtmUtyEzEQhtWTQyQLHNak2AB7ZnyXZMEjXMGeK/AIi+QuHrMnbChYY7MIh8g01fJoopFb0uhhEqqcbWTp06/uv1saEDv4O3n3dV60RfP947Mm9/SQc0ICFQgzfc4CYZoTPAswgSJCCUJUnAAoRHOAUOcATwbmVLWdGoH//PB8mnKqScAhsD0kYP3j/Yt5LPQe2KvcXmGvRHcDnpxfL2zOYJ1mFwrryWTz0advv1Ut4CJgf5uhDuDj5eUcAUoahrdY/56ebRWeraTjMt/00Sh3UDtjgHtQNHwcRGOC98BJEAEymycmYcWwOprTgcB6VZ5JK5TAJ+fXGLBm3FDAmn6oPPjR4rKCAoJCal2eAiQp2x0vxTPB3ALO2CRkwmDy5WohzBDwSEFKRwPbknEggCPB/imwrycgxX2NzoMCHhPkDwqYMr9tRcP5qNrMZHkVnOjRMWwLCcr8ohBVb1OMjxLwGCvjTikrsBOiA6fNyCrm8V1rP93iVPpwaE+gO0SsWmPiXB+jikdf6SizrT5qKasx5j8ABbHpFTx+vFXp9EnYQmLx02h1QTTrl6eDqxLnGjporxl3NL3agEvXdT0WmEost648sQOYAeJS9Q7bfUVoMGnjo4AZdUMQku50McDcMWcBPvr0SzbTAFDfvJqwLzgxwATnCgnp4wDl6Aa+Ax283gghmj+vj7feE2KBBRMW3FzOpLOADl0Isb5587h/U4gGvkt5v60Z1VLG8BhYjbzRwyQZemwAd6cCR5/XFWLYZRIMpX39AR0tjaGGiGzLVyhse5C9RKC6ai42ppWPKiBagOvaYk8lO7DajerabOZP46Lby5wKjw1HCRx7p9sVMOWGzb/vA1hwiWc6jm3MvQDTogQkiqIhJV0nBQBTU+3okKCFDy9WwferkHjtxib7t3xIUQtHxnIwtx4mpg26/HfwVNVDb4oI9RHmx5WGelRVlrtiw43zboCLaxv46AZeB3IlTkwouebTr1y2NjSpHz68WNFjHvupy3q8TFn3Hos2IAk4Ju5dCo8B3wP7VPr/FGaKiG+T+v+TQqIrOqMTL1VdWV1DdmcbO8KXBz6esmYWYKPwDL5b5FA1a0hwapHiom0r/cKaoqr+27/XcrS5UwSMbQAAAABJRU5ErkJggg==)](https://deepwiki.com/junior19a2000/STE)

# Hermeticidad de Tanques de CL y OPDH - Reporte Dinámico

## Descripción
Este proyecto es una aplicación interactiva desarrollada con Marimo y Altair en Python, diseñada para generar un reporte dinámico y visual sobre la inspección periódica de hermeticidad de tanques enterrados que almacenan Combustibles Líquidos (CL) y Otros Productos Derivados de los Hidrocarburos (OPDH) a nivel nacional. La norma aplicada corresponde al Decreto Supremo N° 001-2022-MINEM-EM y sus resoluciones complementarias.

> **Nota**: La aplicación se enfoca exclusivamente en los tanques ubicados en estaciones de servicio.

## Funcionalidades

- **Carga de datos** desde fuentes oficiales:
  - Base de datos de componentes de tanques (`DATA TANQUES.xlsx`).
  - Base de datos de pruebas de hermeticidad (`DATA PRUEBAS.xlsx`).
  - Descarga de registros de agentes habilitados mediante la plataforma PVO de OSINERGMIN.
- **Limpieza y consolidación** de la información:
  - Depuración de duplicados y registros incompletos.
  - Asignación de regiones según departamento, provincia y distrito.
  - Cálculo de fecha límite de inspección, edad y estado del tanque.
- **Generación de matriz** completa con detalles de cada compartimiento y prueba.
- **Indicadores resumidos** por Oficina Regional:
  - Cumplimiento de registro de información (completo, incompleto, nulo).
  - Cumplimiento de pruebas de hermeticidad (completo, incompleto, nulo).
- **Análisis interactivo**:
  - Filtros por región, estado de registro y hermeticidad.
  - Exportación de tablas y gráficos a archivos Excel.
- **Visualizaciones** con Altair para:
  - Estado de cumplimiento por oficina regional.
  - Componentes que acreditaron fugas.

## Tecnologías y Dependencias

- **Python 3.8+**
- [Marimo](https://pypi.org/project/marimo/) (App framework)
- [pandas](https://pandas.pydata.org/)
- [Altair](https://altair-viz.github.io/)
- [requests](https://docs.python-requests.org/)
- [openpyxl](https://openpyxl.readthedocs.io/)

```bash
pip install marimo pandas altair requests openpyxl
```

## Estructura de Archivos

```plaintext
├── app.py                # Código principal de la aplicación
├── DATA TANQUES.xlsx     # Datos de fabricación e instalación de tanques
├── DATA PRUEBAS.xlsx     # Datos de pruebas de hermeticidad
├── CRONOGRAMA.png        # Imagen con cronograma de inspecciones
├── requirements.txt      # Lista de dependencias (opcional)
└── README.md             # Documentación del proyecto
```

## Uso

1. Colocar los archivos de datos (`DATA TANQUES.xlsx`, `DATA PRUEBAS.xlsx`) en el mismo directorio que `app.py`.
2. Ejecutar la aplicación:
   ```bash
   python app.py
   ```
3. Abrir el navegador en la dirección proporcionada por Marimo (por defecto `http://localhost:8000`).
4. Interactuar con los controles para filtrar resultados y generar reportes.

## Exportación de Resultados

- **Matriz general**: Botón para descargar `MATRIZ.xlsx` con todos los registros consolidados.
- **Análisis regional**: Opciones para filtrar agentes según estado y región, con descarga de Excel.
- **Análisis crítico**: Identificación de componentes con fugas y exportación de archivos.

## Contribuciones

Las contribuciones son bienvenidas. Por favor, abra un issue o envíe un pull request con mejoras, correcciones o nuevas funcionalidades.

## Licencia

Este proyecto está bajo la licencia MIT. Consulte el archivo `LICENSE` para más información.
