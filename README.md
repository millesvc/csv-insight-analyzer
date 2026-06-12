# 🔬 CSV Insight Analyzer

> Herramienta de análisis de datos empresarial que automatiza la exploración, estadística descriptiva y generación de reportes a partir de archivos CSV.

![Python](https://img.shields.io/badge/Python-3.11+-1A1A2E?style=for-the-badge&logo=python&logoColor=E94560)
![pandas](https://img.shields.io/badge/pandas-2.0+-16213E?style=for-the-badge&logo=pandas&logoColor=white)
![matplotlib](https://img.shields.io/badge/matplotlib-3.7+-0F3460?style=for-the-badge)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1+-E94560?style=for-the-badge)
![License](https://img.shields.io/badge/Licencia-MIT-9EB8D9?style=for-the-badge)

---

## 📋 Descripción

**CSV Insight Analyzer** es una herramienta de línea de comandos diseñada para automatizar el análisis exploratorio de datos (EDA) sobre cualquier archivo `.csv`. Genera estadísticas descriptivas, detecta problemas de calidad de datos, crea visualizaciones y exporta reportes listos para presentar — todo con un solo comando.

Desarrollada bajo los estándares de **MillesvcStudio** como proyecto de portafolio profesional de nivel Ingeniería Civil Informática, con foco en código limpio, arquitectura escalable y automatización de flujos empresariales.

---

## 🎯 Objetivos del Proyecto

- Reducir el tiempo de análisis exploratorio manual de datos.
- Proveer un reporte estructurado y reproducible con un solo comando.
- Servir como base escalable para integraciones con pipelines de datos, dashboards o APIs.
- Demostrar manejo profesional de librerías de análisis y visualización en Python.

---

## 🗂️ Estructura del Proyecto

```
csv-insight-analyzer/
│
├── data/
│   └── sample.csv              # Dataset de muestra (empleados ficticios)
│
├── reports/                    # Generado automáticamente al ejecutar
│   ├── reporte.txt
│   ├── reporte.xlsx
│   └── charts/
│       ├── hist_salario.png
│       ├── hist_edad.png
│       ├── boxplot_numericas.png
│       ├── correlaciones.png
│       └── valores_nulos.png
│
├── analyzer.py                 # Módulo principal de análisis
├── requirements.txt
├── README.md
└── .gitignore
```

---

## ⚙️ Tecnologías Utilizadas

| Librería      | Versión mínima | Uso principal                              |
|---------------|----------------|--------------------------------------------|
| `pandas`      | 2.0.0          | Carga, manipulación y estadísticas de datos|
| `matplotlib`  | 3.7.0          | Generación de gráficos PNG                 |
| `openpyxl`    | 3.1.0          | Exportación de reportes XLSX con estilos   |
| `pathlib`     | stdlib         | Manejo de rutas multiplataforma            |

---

## 🚀 Instalación

### 1. Clonar el repositorio

```bash
git clone https://github.com/millesvcstudio/csv-insight-analyzer.git
cd csv-insight-analyzer
```

### 2. Crear entorno virtual (recomendado)

```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# macOS / Linux
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## 💻 Uso

### Analizar el CSV de muestra incluido

```bash
python analyzer.py
```

### Analizar tu propio archivo CSV

```bash
python analyzer.py ruta/a/tu/archivo.csv
```

### Ejemplos

```bash
# Dataset de ventas
python analyzer.py data/ventas_2024.csv

# Dataset de inventario
python analyzer.py /home/usuario/inventario.csv
```

---

## 📊 Funcionalidades

### Análisis Estructural
- Cantidad de filas y columnas
- Nombres de columnas y tipos de datos detectados automáticamente

### Calidad de Datos
- Porcentaje de completitud del dataset
- Conteo de valores nulos por columna con visualización en barra
- Detección de filas duplicadas

### Estadísticas Descriptivas (por columna numérica)
| Métrica          | Descripción                          |
|------------------|--------------------------------------|
| Promedio         | Media aritmética                     |
| Mediana          | Valor central                        |
| Máximo / Mínimo  | Rango de valores                     |
| Desv. Estándar   | Dispersión de datos                  |
| Varianza         | Dispersión al cuadrado               |
| Q1 / Q3 / IQR    | Cuartiles y rango intercuartílico    |
| Asimetría        | Sesgo de la distribución             |
| Curtosis         | Forma de los colas de la distribución|

### Visualizaciones Automáticas (PNG)
- **Histogramas** por cada columna numérica
- **Boxplot** agrupado de todas las columnas numéricas
- **Mapa de calor de correlaciones** entre variables
- **Gráfico de barras** de valores nulos por columna

### Exportación de Reportes
- `reports/reporte.txt` — reporte de texto plano con todos los resultados
- `reports/reporte.xlsx` — libro Excel con 3 hojas: Resumen, Estadísticas, Datos Originales
- `reports/charts/*.png` — todos los gráficos generados

---

## 📁 Dataset de Muestra

El archivo `data/sample.csv` contiene datos ficticios de 30 empleados con columnas:

| Columna              | Tipo     | Descripción                       |
|----------------------|----------|-----------------------------------|
| `id`                 | Entero   | Identificador único               |
| `nombre`             | Texto    | Nombre completo                   |
| `edad`               | Entero   | Edad en años                      |
| `ciudad`             | Texto    | Ciudad de residencia (Chile)      |
| `salario`            | Decimal  | Salario mensual en CLP            |
| `departamento`       | Texto    | Área de la empresa                |
| `años_experiencia`   | Entero   | Años de experiencia laboral       |
| `evaluacion`         | Decimal  | Puntuación de desempeño (1.0-5.0) |
| `fecha_ingreso`      | Fecha    | Fecha de incorporación            |

> El dataset incluye intencionalmente 2 valores nulos en `salario` para demostrar la detección de calidad de datos.

---

## 🧩 Arquitectura del Código

`analyzer.py` está organizado en módulos funcionales independientes:

```
analyzer.py
├── load_csv()              # Carga con manejo de encodings
├── print_overview()        # Resumen estructural
├── analyze_data_quality()  # Nulos y duplicados
├── compute_statistics()    # Estadísticas descriptivas
├── generate_charts()       # Visualizaciones PNG
├── export_txt_report()     # Exportar TXT
├── export_xlsx_report()    # Exportar XLSX (3 hojas)
└── run_analysis()          # Orquestador del pipeline
```

Cada función tiene docstring, tipado de parámetros y manejo de excepciones, lo que facilita pruebas unitarias, reutilización y extensión futura.

---

## 🛡️ Manejo de Errores

El sistema maneja explícitamente:

| Error                    | Causa                                        |
|--------------------------|----------------------------------------------|
| `FileNotFoundError`      | Ruta al CSV inexistente                      |
| `ValueError`             | Extensión incorrecta o archivo vacío         |
| `pd.errors.ParserError`  | CSV mal formateado                           |
| `UnicodeDecodeError`     | Reintento automático con encoding `latin-1`  |

---

## 🔮 Mejoras Futuras

- [ ] **CLI interactivo** con `argparse` o `typer` (selección de columnas, umbral de nulos)
- [ ] **Detección de outliers** con método IQR y Z-score
- [ ] **Reporte HTML** con gráficos embebidos y CSS personalizado
- [ ] **Análisis de columnas categóricas** (frecuencias, valores únicos, top-N)
- [ ] **Soporte para múltiples archivos** y comparación entre datasets
- [ ] **API REST** con FastAPI para uso como microservicio
- [ ] **Dashboard web** con Streamlit o Dash
- [ ] **Integración con Google Sheets** via API
- [ ] **Tests unitarios** con `pytest`
- [ ] **CI/CD** con GitHub Actions para validación automática

---

## 👤 Autor

**MillesvcStudio** — Desarrollo web y soluciones digitales para PYMES  
📍 Viña Del Mar, Chile  
🔗 [github.com/millesvc](https://github.com/millesvc)

---

## 📄 Licencia

Este proyecto está bajo la licencia [MIT](LICENSE). Libre para uso, modificación y distribución con atribución.
