# Generador de Constancias CRM 📑

Herramienta diseñada para automatizar la creación de documentos de constancia de requerimientos CRM a partir de una fuente de datos en Excel. Transforma filas estructuradas en archivos de Word (`.docx`) con formato profesional listos para su revisión o entrega.

## ✨ Características

- **Generación Automatizada**: Crea múltiples documentos `.docx` en segundos a partir de una hoja de cálculo.
- **Formateo Dinámico**: Incluye secciones de evidencia, situación actual, alcance funcional y visualización de cambios.
- **Tratamiento de Nombres**: Limpia automáticamente caracteres inválidos para crear archivos seguros y ordenados.
- **Gestión de Saltos de Línea**: Mantiene la estructura del texto capturado en el Excel, respetando párrafos y listas.
- **Filtrado Inteligente**: Solo procesa registros válidos que contengan definiciones de módulos y submódulos.

## 📁 Estructura del Proyecto

- `main.py`: Script principal en Python que contiene la lógica de procesamiento y generación de documentos.
- `CRM - REQUERIMIENTOS.xlsx`: Archivo Excel de entrada (hoja: `Requerimiento`).
- `Constancias_Finales_CRM/`: Directorio donde se almacenan todos los documentos generados.
- `.venv/`: Entorno virtual de Python para el aislamiento de dependencias.

## 🚀 Instalación y Uso

### Requisitos Previos

- Python 3.8 o superior.
- Microsoft Word (para visualizar las constancias).

### Pasos para Ejecutar

1. **Instalar Dependencias**:
   Asegúrate de tener instaladas las librerías necesarias ejecutando:
   ```bash
   pip install pandas python-docx openpyxl
   ```

2. **Preparar los Datos**:
   Asegúrate de que el archivo `CRM - REQUERIMIENTOS.xlsx` esté en la raíz del proyecto y que la pestaña `Requerimiento` tenga las columnas correctas.

3. **Ejecutar el Generador**:
   Inicia el proceso con el siguiente comando:
   ```bash
   python main.py
   ```

4. **Verificar Resultados**:
   Revisa la carpeta `Constancias_Finales_CRM/` para encontrar los documentos `.docx` generados.

## 🛠️ Tecnologías Utilizadas

- **[Pandas](https://pandas.pydata.org/):** Para la manipulación y análisis eficiente de datos desde el Excel.
- **[python-docx](https://python-docx.readthedocs.io/):** Para la creación y el formateo programático de documentos Word.
- **[Openpyxl](https://openpyxl.readthedocs.io/):** Motor de lectura para archivos Excel modernos (.xlsx).

---
*Desarrollado para optimizar el flujo de documentación de requerimientos CRM.*
