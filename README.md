# Concentrador Excel con Merge de Datos

Esta herramienta permite consolidar archivos Excel y realizar una fusión (merge) controlada con datos desde un archivo externo.

## Funcionalidades

1. **Concentrador General**: Consolida múltiples archivos Excel en un único archivo, identificando duplicados.
2. **Actualización Incremental**: Compara y actualiza archivos de concentrado.
3. **Merge de Datos**: Integra información desde `data.xlsx` hacia el concentrado general, usando el número de expediente como identificador único.

## Requisitos

- Node.js (versión 14 o superior)
- Excel instalado para visualizar los resultados

## Instalación

```bash
npm install
```

## Configuración

1. Crear una carpeta llamada `concentrado-crk` en el Escritorio (Desktop)
2. Colocar los archivos Excel a consolidar en esta carpeta
3. Para merge: asegurarse de que `data.xlsx` esté en la carpeta `merge-general`

## Uso

### Ejecución paso a paso (scripts individuales)

```bash
# Paso 1: Generar el concentrado general
node index.js

# Paso 2: Actualizar concentrado (comparar y agregar nuevos registros)
node index2.js

# Paso 3: Realizar merge de datos
node merge.js
```

### Ejecución completa automatizada

```bash
# Ejecutar el proceso completo secuencial (todos los pasos con pausas)
node proceso-completo.js
```

Este comando ejecuta automáticamente los tres pasos en secuencia, mostrando todos los logs en tiempo real y con pausas de 5 segundos entre cada paso.

### Scripts disponibles (npm)

```bash
# Solo concentrador general
npm start

# Solo merge de datos
npm run merge

# Concentrado + merge (sin index2.js)
npm run complete

# Diagnóstico de merge
npm run diagnostico
```

## Archivos Generados

- `concentrado-general.xlsx`: Archivo con todos los registros consolidados.
- `duplicados.xlsx`: Registros identificados como duplicados.
- En carpeta merge-general:
  - `reporte-merge.xlsx`: Reporte detallado del proceso de merge.
  - Archivos de respaldo con timestamps.

## Notas Importantes

- El merge solo integra registros si el número de expediente coincide exactamente.
- Las carpetas deben existir antes de ejecutar los scripts.
- Para un proceso completo óptimo, se recomienda usar el script `proceso-completo.js`.