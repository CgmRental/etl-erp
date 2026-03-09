# Extractor de Establecimientos Anexos - Guía de Uso

## 📄 Descripción

Este script está **específicamente diseñado** para extraer la información de archivos PDF de "Establecimientos Anexos" de SUNAT, convirtiéndolos automáticamente a Excel con formato profesional.

## 🎯 Características Específicas

✅ Extrae **4 columnas principales**:
   - Código
   - Tipo de Establecimiento
   - Dirección
   - Actividad Económica

✅ Genera **3 hojas en Excel**:
   - **Establecimientos**: Listado completo con todos los datos
   - **Resumen**: Información general y estadísticas
   - **Por Tipo**: Conteo de establecimientos por tipo

✅ Formato profesional:
   - Headers con colores corporativos
   - Filtros automáticos
   - Columnas ajustadas al contenido
   - Texto envuelto en direcciones largas
   - Primera fila fija para scroll

## 🚀 Instalación

```bash
pip install pdfplumber pandas openpyxl --break-system-packages
```

## 📝 Uso Básico

### Opción 1: Con el script principal
```bash
python establecimientos_pdf_to_excel.py Establecimientos_Anexos.pdf
```

Esto generará automáticamente: `Establecimientos_Anexos.xlsx`

### Opción 2: Especificar nombre de salida
```bash
python establecimientos_pdf_to_excel.py Establecimientos_Anexos.pdf mi_lista.xlsx
```

### Opción 3: Si el método principal no funciona bien
```bash
python establecimientos_alternativo.py Establecimientos_Anexos.pdf
```

## 📊 Estructura del Excel Generado

### Hoja 1: Establecimientos
| Código | Tipo de Establecimiento | Dirección | Actividad Económica |
|--------|------------------------|-----------|---------------------|
| 0251   | AG. AGENCIA           | CAL.ANTENOR ORREGO... | - |
| 0010   | AG. AGENCIA           | AV. SALAVERRY NRO. 121... | - |
| ...    | ...                   | ... | ... |

**Características**:
- Filtros automáticos en todas las columnas
- Primera fila fija (scroll independiente)
- Direcciones con wrap text para mejor lectura
- Bordes en todas las celdas

### Hoja 2: Resumen
```
Información                    | Valor
------------------------------|------------------
Archivo Origen                | Establecimientos_Anexos.pdf
Total de Establecimientos     | 100
Fecha de Extracción          | 2026-01-28 10:30:00
Tipos de Establecimiento      | 4
```

### Hoja 3: Por Tipo
```
Tipo de Establecimiento | Cantidad
-----------------------|----------
AG. AGENCIA           | 95
DE. DEPOSITO          | 3
OF. OF.ADMINIST.      | 1
LO. L. COMERCIAL      | 1
```

## 🎨 Colores del Formato

- **Headers principales**: Azul corporativo (#0066CC)
- **Hoja Resumen**: Azul profesional (#4472C4)
- **Hoja Por Tipo**: Verde (#70AD47)
- **Filas alternadas**: Gris claro (#F2F2F2)

## ⚙️ Personalización

Para modificar el formato, edita estas secciones en el script:

### Cambiar colores de headers:
```python
header_fill = PatternFill(start_color='TU_COLOR', end_color='TU_COLOR', fill_type='solid')
```

### Ajustar ancho de columnas:
```python
ws.column_dimensions['A'].width = 15  # Código
ws.column_dimensions['C'].width = 100  # Dirección más ancha
```

### Agregar nuevas hojas de análisis:
```python
# Ejemplo: Análisis por departamento
depto_stats = df['Dirección'].str.extract(r'(\w+) -')[0].value_counts()
depto_stats.to_excel(writer, sheet_name='Por Departamento')
```

## 🔧 Solución de Problemas

### Problema: Pocas filas extraídas
**Solución**: Usar el script alternativo
```bash
python establecimientos_alternativo.py Establecimientos_Anexos.pdf
```

### Problema: Direcciones cortadas
**Solución**: Aumentar el ancho de columna C en el script:
```python
ws.column_dimensions['C'].width = 120  # Más ancho
```

### Problema: Error "No module named 'pdfplumber'"
**Solución**:
```bash
pip install pdfplumber --break-system-packages
```

### Problema: Algunos registros se pierden
**Causa**: El PDF puede tener formato inconsistente en algunas páginas
**Solución**: Revisar manualmente el Excel generado y complementar datos faltantes

## 💡 Tips Avanzados

### Procesar múltiples archivos PDF:
```bash
for pdf in *.pdf; do
    python establecimientos_pdf_to_excel.py "$pdf"
done
```

### Combinar varios Excel en uno solo:
```python
import pandas as pd
from pathlib import Path

all_files = Path('.').glob('*.xlsx')
dfs = [pd.read_excel(f, sheet_name='Establecimientos') for f in all_files]
combined = pd.concat(dfs, ignore_index=True)
combined.to_excel('todos_establecimientos.xlsx', index=False)
```

### Filtrar por tipo de establecimiento:
```python
df = pd.read_excel('Establecimientos_Anexos.xlsx', sheet_name='Establecimientos')
agencias = df[df['Tipo de Establecimiento'].str.contains('AGENCIA')]
agencias.to_excel('solo_agencias.xlsx', index=False)
```

## 📞 Estructura del Código

```
establecimientos_pdf_to_excel.py
├── clean_text()              # Limpia y normaliza texto
├── extract_establecimientos_from_pdf()  # Extrae datos del PDF
├── save_to_excel_formatted() # Guarda con formato
└── process_establecimientos_pdf()  # Función principal
```

## 🎯 Casos de Uso

1. **Análisis de cobertura**: Ver distribución geográfica de establecimientos
2. **Reportes internos**: Generar listados actualizados
3. **Auditorías**: Verificar consistencia de información
4. **Planificación**: Identificar zonas sin cobertura
5. **Migración de datos**: Pasar de PDF a sistemas internos

## 📈 Ejemplo de Salida

```
Procesando: Establecimientos_Anexos.pdf
Total de páginas: 4
  Extrayendo página 1/4...
  Extrayendo página 2/4...
  Extrayendo página 3/4...
  Extrayendo página 4/4...

✓ Registros extraídos: 100

Guardando en: Establecimientos_Anexos.xlsx
✓ Archivo guardado con formato profesional

======================================================================
✓ PROCESO COMPLETADO EXITOSAMENTE
======================================================================

Archivo generado: Establecimientos_Anexos.xlsx
Total de establecimientos: 100
Columnas extraídas: Código, Tipo de Establecimiento, Dirección, Actividad Económica
```

## 🌟 Ventajas sobre otros métodos

| Característica | Este Script | Copia Manual | Otros Extractores |
|----------------|------------|--------------|-------------------|
| Velocidad | ⚡ Segundos | ⏰ Horas | 🐌 Variable |
| Precisión | ✅ Alta | ⚠️ Media | ⚠️ Variable |
| Formato | 🎨 Profesional | 📝 Básico | 📊 Básico |
| Automatización | 🤖 Total | 👤 Manual | 🔧 Parcial |
| Reproducibilidad | ✅ 100% | ❌ Baja | ⚠️ Media |
