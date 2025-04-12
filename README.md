# Concentrador de Archivos Excel

Este script permite consolidar múltiples archivos Excel (.xlsx) en un solo archivo, identificando y separando registros duplicados.

## Descripción

El script realiza las siguientes funciones:

- Lee todos los archivos Excel (.xlsx) de una carpeta especificada.
- Identifica registros duplicados (basados en la columna "Expediente").
- Consolida todos los registros únicos en un archivo Excel general.
- Guarda los registros duplicados en un archivo separado.
- Preserva los formatos de fecha y hora originales.
- Genera un reporte detallado del proceso.

## Requisitos Previos

- Node.js (versión 14 o superior)
- npm (gestor de paquetes de Node.js)

## Instalación

1. Clone o descargue este repositorio.
2. Navegue a la carpeta del proyecto en su terminal.
3. Instale las dependencias ejecutando:

```bash
npm install
```

## Configuración

Edite el archivo `index.js` para configurar las rutas y parámetros:

```javascript
// ======================================
//  CONFIGURACIÓN DE RUTAS Y ARCHIVOS
// ======================================
const carpeta = path.join(process.env.HOME, "Desktop", "concentrado-crk");
const archivoConcentrado = path.join(carpeta, "concentrado-general.xlsx");
const archivoDuplicados = path.join(carpeta, "duplicados.xlsx");

/**
 * Columna que identifica de forma única a cada expediente (columna A).
 * Ajuste este valor si su encabezado real es diferente.
 */
const COLUMNA_EXPEDIENTE = "Expediente";
```

Ajuste estas variables según sus necesidades:

- `carpeta`: Ruta donde se encuentran los archivos Excel a consolidar.
- `archivoConcentrado`: Ruta donde se guardará el archivo consolidado.
- `archivoDuplicados`: Ruta donde se guardarán los registros duplicados.
- `COLUMNA_EXPEDIENTE`: Nombre de la columna que sirve como identificador único.

## Uso

1. Asegúrese de que todos los archivos Excel (.xlsx) a consolidar estén en la carpeta configurada.
2. Ejecute el script con:

```bash
npm start
```

O alternativamente:

```bash
node index.js
```

3. El script generará dos archivos:
   - `concentrado-general.xlsx`: Contiene todos los registros únicos.
   - `duplicados.xlsx`: Contiene los registros duplicados.

## Comportamiento del Script

- Si un registro aparece en múltiples archivos, se mantiene la última versión encontrada en el concentrado general.
- Los duplicados son:
  1. Registros con el mismo identificador que aparecen más de una vez en el mismo archivo.
  2. Registros que ya estaban marcados como duplicados en una ejecución anterior.
- El script preservará los formatos de fecha (dd/mm/yyyy) y hora (hh:mm:ss) en las columnas correspondientes.

## Características Especiales

- **Formatos Preservados**: El script detecta automáticamente columnas de fecha y hora y aplica el formato adecuado.
- **Normalización de Datos**: Elimina espacios en blanco extra y estandariza los datos.
- **Detección Robusta de Fechas y Horas**: Admite múltiples formatos de entrada.
- **Reconciliación de Totales**: Verifica que la suma de registros sea correcta.
- **Persistencia de Duplicados**: Mantiene el registro de duplicados entre ejecuciones.

## Solución de Problemas

- **Error de archivo no encontrado**: Asegúrese de que la carpeta configurada existe en la ruta especificada.
- **Error al leer los archivos**: Verifique que los archivos Excel no estén abiertos en otra aplicación.
- **Diferencia en totales**: Puede ocurrir si hay registros sin el identificador único. Revise los datos originales.

## Licencia

ISC