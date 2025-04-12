# Concentrador Excel con Merge de Datos

Esta herramienta permite consolidar archivos Excel y realizar una fusión (merge) controlada con datos desde un archivo externo.

## Funcionalidades

1. **Concentrador General**: Consolida múltiples archivos Excel en un único archivo, identificando duplicados.
2. **Merge de Datos**: Integra información desde `data.xlsx` hacia el concentrado general, usando el número de expediente como identificador único.

## Estructura del Proyecto

- `index.js`: Script principal para generar el concentrado general.
- `merge-data.js`: Script para realizar el merge de datos desde `data.xlsx`.
- `package.json`: Configuración del proyecto y scripts disponibles.

## Requisitos

- Node.js (versión 14 o superior)
- Excel instalado para visualizar los resultados

## Instalación

1. Clonar o descargar este repositorio
2. Abrir una terminal en la carpeta del proyecto
3. Instalar las dependencias:

```bash
npm install
```

## Configuración

1. Crear una carpeta llamada `concentrado-crk` en el Escritorio (Desktop)
2. Colocar los archivos Excel a consolidar en esta carpeta
3. Asegurarse de que `data.xlsx` esté en la misma carpeta

## Uso

### Generar solo el Concentrado General

```bash
npm start
```

### Realizar solo el Merge de Datos

```bash
npm run merge
```

### Proceso Completo (Concentrado + Merge)

```bash
npm run complete
```

## Proceso de Merge

El script `merge-data.js` realiza lo siguiente:

1. Lee el archivo `data.xlsx` y el concentrado general.
2. Busca coincidencias de expedientes entre ambos archivos:
   - El número de expediente se toma de la columna D (4) en `data.xlsx`.
   - Se busca coincidencia exacta en la columna A del concentrado general.
3. Para cada coincidencia, inserta los datos desde la columna AW (49) del concentrado.
4. Genera un reporte detallado de la operación.
5. Registra los expedientes procesados para evitar duplicidades en futuras ejecuciones.

## Archivos Generados

- `concentrado-general.xlsx`: Archivo con todos los registros consolidados y datos fusionados.
- `duplicados.xlsx`: Registros identificados como duplicados.
- `reporte-merge.xlsx`: Reporte detallado del proceso de merge.
- `registros-procesados.json`: Registro interno de expedientes ya procesados.

## Notas Importantes

- El merge solo integra registros si el número de expediente coincide exactamente.
- Los registros ya procesados se almacenan para evitar duplicaciones si se ejecuta el script varias veces.
- El reporte detalla qué registros se integraron y cuáles no, con el motivo correspondiente.

## Solución de Problemas

Si encuentra problemas durante la ejecución:

1. Verifique que los archivos estén en la ubicación correcta.
2. Revise que el formato de los números de expediente sea consistente entre ambos archivos.
3. Para reiniciar el proceso de merge desde cero, elimine el archivo `registros-procesados.json`.