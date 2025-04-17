// node merge.js
import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

// ======================================
//  CONFIGURACIÓN DE RUTAS Y ARCHIVOS
// ======================================
const carpetaBase = path.join(process.env.HOME, "Desktop", "concentrado-crk");
const carpetaMerge = path.join(carpetaBase, "merge-general");

// Los archivos ahora se buscan dentro de la carpeta de merge
const archivoConcentrado = path.join(carpetaMerge, "concentrado-general.xlsx");
const archivoData = path.join(carpetaMerge, "data.xlsx");
const archivoReporte = path.join(carpetaMerge, "reporte-merge.xlsx");

// El backup se guarda en la misma carpeta de merge
const archivoBackup = path.join(carpetaMerge, `concentrado-backup-${Date.now()}.xlsx`);

// ======================================
//  CREAR BACKUP DEL CONCENTRADO
// ======================================
function crearBackup() {
  try {
    if (fs.existsSync(archivoConcentrado)) {
      // Usamos el método de bajo nivel para copiar el archivo byte a byte
      // Esto asegura que se preserve absolutamente todo el contenido y formato
      const contenidoOriginal = fs.readFileSync(archivoConcentrado);
      fs.writeFileSync(archivoBackup, contenidoOriginal);
      console.log(`Backup creado: ${archivoBackup}`);
      return true;
    } else {
      console.error("No se encontró el archivo concentrado para hacer backup");
      return false;
    }
  } catch (error) {
    console.error("Error al crear backup:", error.message);
    return false;
  }
}

// ======================================
//  NORMALIZAR VALOR DE EXPEDIENTE
// ======================================
function normalizarExpediente(valor) {
  if (valor === null || valor === undefined) return "";
  
  // Convertir a string y eliminar espacios extras
  return String(valor).trim().replace(/\s+/g, "");
}

// ======================================
//  LEER ARCHIVO EXCEL CON EXCELJS
// ======================================
async function leerArchivoExcel(rutaArchivo, descripcion = "Excel") {
  try {
    console.log(`Leyendo archivo ${rutaArchivo} con ExcelJS...`);
    
    // Verificar existencia del archivo
    if (!fs.existsSync(rutaArchivo)) {
      console.error(`❌ ERROR: El archivo ${rutaArchivo} no existe.`);
      return { datos: [], encabezados: [], workbook: null, worksheet: null };
    }
    
    const stats = fs.statSync(rutaArchivo);
    console.log(`  - Tamaño del archivo: ${(stats.size / 1024).toFixed(2)} KB`);
    
    // Crear workbook y leer archivo
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(rutaArchivo);
    
    // Verificar que haya hojas
    if (workbook.worksheets.length === 0) {
      console.error(`❌ ERROR: No se encontraron hojas en el archivo ${rutaArchivo}`);
      return { datos: [], encabezados: [], workbook: null, worksheet: null };
    }
    
    // Usar la primera hoja
    const worksheet = workbook.worksheets[0];
    console.log(`  - Hoja encontrada: ${worksheet.name}`);
    console.log(`  - Filas: ${worksheet.rowCount}, Columnas: ${worksheet.columnCount}`);
    
    // Extraer encabezados (primera fila)
    const encabezados = [];
    const primeraFila = worksheet.getRow(1);
    primeraFila.eachCell((cell, colNumber) => {
      encabezados[colNumber - 1] = cell.value?.toString() || `Columna_${colNumber}`;
    });
    
    console.log(`  - Encabezados encontrados: ${encabezados.length}`);
    console.log(`  - Ejemplo de encabezados: ${encabezados.slice(0, 5).join(', ')}...`);
    
    // Convertir datos a array de objetos
    const datos = [];
    
    // Empezar desde la fila 2 (después de encabezados)
    for (let i = 2; i <= worksheet.rowCount; i++) {
      const fila = worksheet.getRow(i);
      const registro = {};
      let tieneValores = false;
      
      // Asignar valor a cada propiedad según encabezados
      encabezados.forEach((encabezado, index) => {
        if (encabezado) {
          const celda = fila.getCell(index + 1);
          const valor = celda.value;
          
          // Asignar valor al registro
          registro[encabezado] = valor;
          
          // Verificar si la celda tiene algún valor
          if (valor !== null && valor !== undefined && valor !== "") {
            tieneValores = true;
          }
        }
      });
      
      // Solo agregar filas que contengan datos
      if (tieneValores) {
        datos.push(registro);
      }
      
      // Mostrar progreso cada 1000 filas o en la última fila
      if (i % 1000 === 0 || i === worksheet.rowCount) {
        console.log(`  - Procesadas ${i} de ${worksheet.rowCount} filas...`);
      }
    }
    
    console.log(`  - Total registros extraídos: ${datos.length}`);
    
    // Mostrar algunos ejemplos de datos
    if (datos.length > 0) {
      console.log(`\nEjemplos de registros en ${descripcion}:`);
      datos.slice(0, 3).forEach((registro, i) => {
        console.log(`  - Registro #${i+1}: ${JSON.stringify(registro).substring(0, 150)}...`);
      });
    }
    
    return { datos, encabezados, workbook, worksheet };
  } catch (error) {
    console.error(`Error al leer ${rutaArchivo} con ExcelJS:`, error.message);
    console.error(error.stack);
    return { datos: [], encabezados: [], workbook: null, worksheet: null };
  }
}

// ======================================
//  GENERAR REPORTE EXCEL CON EXCELJS
// ======================================
async function generarReporteExcelJS(reporte, rutaReporte) {
  try {
    const workbook = new ExcelJS.Workbook();
    
    // Hoja de estadísticas
    const hojaEstadisticas = workbook.addWorksheet("Estadísticas");
    hojaEstadisticas.columns = [
      { header: 'Estadística', key: 'estadistica', width: 40 },
      { header: 'Valor', key: 'valor', width: 20 }
    ];
    
    // Agregar filas de estadísticas
    hojaEstadisticas.addRow({ 
      estadistica: "Total registros en data.xlsx", 
      valor: reporte.totalRegistros 
    });
    hojaEstadisticas.addRow({ 
      estadistica: "Registros integrados correctamente", 
      valor: reporte.integrados.length 
    });
    hojaEstadisticas.addRow({ 
      estadistica: "Registros no integrados", 
      valor: reporte.noIntegrados.length 
    });
    
    // Dar formato a la hoja de estadísticas
    hojaEstadisticas.getRow(1).font = { bold: true };
    hojaEstadisticas.getColumn(1).font = { bold: true };
    
    // Hoja de registros integrados
    if (reporte.integrados.length > 0) {
      const hojaIntegrados = workbook.addWorksheet("Integrados");
      
      // Configurar encabezados
      hojaIntegrados.columns = [
        { header: 'Expediente', key: 'expediente', width: 20 },
        { header: 'FilaData', key: 'filaData', width: 10 },
        { header: 'FilaConcentrado', key: 'filaConcentrado', width: 15 },
        { header: 'Estado', key: 'estado', width: 15 }
      ];
      
      // Agregar filas de datos
      hojaIntegrados.addRows(reporte.integrados);
      
      // Dar formato
      hojaIntegrados.getRow(1).font = { bold: true };
    }
    
    // Hoja de registros no integrados
    if (reporte.noIntegrados.length > 0) {
      const hojaNoIntegrados = workbook.addWorksheet("No Integrados");
      
      // Configurar encabezados
      hojaNoIntegrados.columns = [
        { header: 'Expediente', key: 'expediente', width: 20 },
        { header: 'FilaData', key: 'filaData', width: 10 },
        { header: 'Motivo', key: 'motivo', width: 40 }
      ];
      
      // Agregar filas de datos
      hojaNoIntegrados.addRows(reporte.noIntegrados);
      
      // Dar formato
      hojaNoIntegrados.getRow(1).font = { bold: true };
    }
    
    // Guardar el workbook
    await workbook.xlsx.writeFile(rutaReporte);
    console.log(`   - Reporte guardado en: ${rutaReporte}`);
    
    return true;
  } catch (error) {
    console.error("Error al generar reporte:", error.message);
    return false;
  }
}

// ======================================
//  REALIZAR MERGE DIRECTO PRESERVANDO FORMATOS
// ======================================
async function realizarMergeDirecto() {
  console.log("\n=== INICIANDO PROCESO DE MERGE DIRECTO (CONSERVANDO FORMATO) ===\n");

  // Verificar que exista la carpeta de merge
  if (!fs.existsSync(carpetaMerge)) {
    console.log(`Creando carpeta de merge: ${carpetaMerge}`);
    try {
      fs.mkdirSync(carpetaMerge, { recursive: true });
    } catch (error) {
      console.error(`Error al crear carpeta de merge: ${error.message}`);
      return;
    }
  }
  
  // Verificar archivos
  if (!fs.existsSync(archivoConcentrado)) {
    console.error(`Error: El archivo concentrado ${archivoConcentrado} no existe`);
    console.log("IMPORTANTE: Debe copiar primero el concentrado-general.xlsx a la carpeta merge-general");
    return;
  }
  
  if (!fs.existsSync(archivoData)) {
    console.error(`Error: El archivo ${archivoData} no existe`);
    console.log("IMPORTANTE: Debe colocar el archivo data.xlsx en la carpeta merge-general");
    return;
  }
  
  // Crear backup antes de modificar (usando copia byte a byte para preservar todo)
  if (!crearBackup()) {
    console.error("No se pudo crear el backup. Abortando proceso por seguridad.");
    return;
  }
  
  // Leer archivos con máxima preservación de formatos
  console.log("Leyendo archivos...");
  
  // Leer el archivo de datos con ExcelJS
  console.log("Leyendo archivo data.xlsx...");
  const resultadoData = await leerArchivoExcel(archivoData, "data.xlsx");
  
  if (!resultadoData.datos || resultadoData.datos.length === 0) {
    console.error("Error: No se pudieron extraer datos del archivo data.xlsx");
    return;
  }
  
  // Extraer valores y encabezados del archivo de datos
  const datosData = resultadoData.datos;
  const encabezadosData = resultadoData.encabezados;
  
  // Información de diagnóstico adicional
  console.log(`\nResumen de lectura de data.xlsx:
- Filas encontradas: ${datosData.length}
- Encabezados encontrados: ${encabezadosData.length}
`);

  // Ahora leemos el concentrado con ExcelJS también
  console.log("\nLeyendo archivo concentrado-general.xlsx...");
  const resultadoConcentrado = await leerArchivoExcel(archivoConcentrado, "concentrado");
  
  if (!resultadoConcentrado.datos || resultadoConcentrado.datos.length === 0) {
    console.error("Error: No se pudieron extraer datos del archivo concentrado");
    return;
  }
  
  // Extraer valores y encabezados del concentrado
  const datosConcentrado = resultadoConcentrado.datos;
  const encabezadosConcentrado = resultadoConcentrado.encabezados;
  const workbookConcentrado = resultadoConcentrado.workbook;
  const worksheetConcentrado = resultadoConcentrado.worksheet;
  
  console.log(`\nResumen de lectura de concentrado:
- Filas encontradas: ${datosConcentrado.length}
- Encabezados encontrados: ${encabezadosConcentrado.length}
`);
  
  // Identificar columna de expediente en data.xlsx
  const COLUMNA_EXPEDIENTE_DATA = "Nº de pieza";
  
  // Verificar que la columna exista en los datos
  if (!datosData || !datosData.length) {
    console.error(`Error: No se encontraron datos en el archivo data.xlsx`);
    return;
  }
  
  // Comprobar si la columna de expediente existe en los datos
  let columnaExpedienteEncontrada = false;
  const encabezadosDisponibles = [];
  
  if (datosData.length > 0) {
    // Recopilar todos los nombres de columnas disponibles para diagnóstico
    Object.keys(datosData[0]).forEach(nombre => {
      encabezadosDisponibles.push(nombre);
      
      // Comprobar si alguna columna coincide con nuestro criterio
      if (nombre === COLUMNA_EXPEDIENTE_DATA) {
        columnaExpedienteEncontrada = true;
      }
    });
  }
  
  if (!columnaExpedienteEncontrada) {
    console.error(`Error: No se encontró la columna "${COLUMNA_EXPEDIENTE_DATA}" en data.xlsx`);
    console.log("Columnas disponibles:", encabezadosDisponibles.join(", "));
    
    // Intentar identificar columnas alternativas que puedan contener números de expediente
    const posiblesColumnas = encabezadosDisponibles.filter(nombre => 
      /expediente|pieza|n[úu]mero|n[°º]|id/i.test(nombre)
    );
    
    if (posiblesColumnas.length > 0) {
      console.log("Posibles columnas alternativas de expediente:", posiblesColumnas.join(", "));
    }
    
    return;
  }
  
  // Crear un mapa de expedientes del concentrado para búsqueda eficiente
  console.log("\nIndexando expedientes del concentrado...");
  const mapaExpedientesConcentrado = new Map();
  
  // Buscar los expedientes en la primera columna
  const primeraColumnaConcentrado = encabezadosConcentrado[0];
  
  // Verificar que exista una primera columna
  if (!primeraColumnaConcentrado) {
    console.error("Error: No se pudo identificar la primera columna del concentrado");
    return;
  }
  
  // Indexar los expedientes del concentrado
  datosConcentrado.forEach((registro, indice) => {
    const expedienteRaw = registro[primeraColumnaConcentrado];
    
    // Normalizar para comparación
    const expediente = normalizarExpediente(expedienteRaw);
    
    if (expediente) {
      mapaExpedientesConcentrado.set(expediente, {
        indice: indice + 2, // +2 porque en Excel las filas empiezan en 1 y la primera es encabezado
        expediente,  // Valor normalizado
        original: expedienteRaw  // Valor original sin normalizar
      });
    }
  });
  
  console.log(`   - Indexados ${mapaExpedientesConcentrado.size} expedientes únicos del concentrado`);
  
  // Mostrar algunos ejemplos de expedientes indexados
  console.log("\nEjemplos de expedientes indexados del concentrado:");
  let count = 0;
  for (const [expediente, info] of mapaExpedientesConcentrado.entries()) {
    if (count < 5) {
      console.log(`   - [Fila ${info.indice}] "${info.original}" -> "${expediente}"`);
      count++;
    } else {
      break;
    }
  }
  
  // Preparar contadores para el reporte
  let expedientesEncontrados = 0;
  let expedientesNoEncontrados = 0;
  
  // Información para el reporte final
  const reporte = {
    totalRegistros: datosData.length,
    integrados: [],
    noIntegrados: []
  };
  
  // Calcular dónde empezar a insertar nuevos datos (columna AW / 49)
  const columnaInicio = 48; // 48 = columna índice 48, que en Excel sería AW (0-indexado)
  
  // Asegurarse de que haya suficientes columnas
  const columnasNecesarias = columnaInicio + encabezadosData.length;
  while (worksheetConcentrado.columnCount < columnasNecesarias) {
    // Agregar columnas si es necesario
    worksheetConcentrado.getColumn(worksheetConcentrado.columnCount + 1); // Esto crea la columna
  }
  
  // Insertar encabezados de data.xlsx en la fila 1 del concentrado
  encabezadosData.forEach((encabezado, indice) => {
    const colNum = columnaInicio + indice + 1; // +1 porque ExcelJS usa índices base 1
    const celda = worksheetConcentrado.getCell(1, colNum);
    celda.value = encabezado;
    celda.font = { bold: true };
  });
  
  // Procesar cada registro de data.xlsx
  console.log("\nProcesando registros de data.xlsx...");
  
  for (let i = 0; i < datosData.length; i++) {
    // Obtener el registro actual y el valor de expediente
    const registro = datosData[i];
    const expedienteDataRaw = registro[COLUMNA_EXPEDIENTE_DATA];
    const expedienteData = normalizarExpediente(expedienteDataRaw);
    
    // Buscar el expediente en el concentrado
    if (expedienteData && mapaExpedientesConcentrado.has(expedienteData)) {
      expedientesEncontrados++;
      
      // Información del expediente encontrado
      const infoExpediente = mapaExpedientesConcentrado.get(expedienteData);
      
      // Fila en el concentrado donde insertar los datos
      const filaConcentrado = infoExpediente.indice;
      
      // Mostrar algunos ejemplos de expedientes encontrados
      if (expedientesEncontrados <= 5 || i % 1000 === 0) {
        console.log(`   - [${i+1}/${datosData.length}] Expediente "${expedienteData}" encontrado en fila ${filaConcentrado} del concentrado`);
      }
      
      // Insertar cada columna de data.xlsx en el concentrado
      encabezadosData.forEach((encabezado, indiceColumna) => {
        const valor = registro[encabezado];
        
        if (valor !== "" && valor !== null && valor !== undefined) {
          // Obtener la celda donde insertar
          const colNum = columnaInicio + indiceColumna + 1; // +1 porque ExcelJS usa índices base 1
          const celda = worksheetConcentrado.getCell(filaConcentrado, colNum);
          
          // Asignar el valor manteniendo el tipo de dato apropiado
          celda.value = valor;
          
          // Si es fecha, establecer formato adecuado
          if (valor instanceof Date) {
            celda.numFmt = 'dd/mm/yyyy';
          }
        }
      });
      
      // Registrar como integrado exitosamente
      reporte.integrados.push({
        expediente: expedienteDataRaw,
        filaData: i + 2, // +2 para mostrar fila real en Excel (base 1 + encabezado)
        filaConcentrado: filaConcentrado,
        estado: "Integrado"
      });
    } else {
      expedientesNoEncontrados++;
      
      // Mostrar algunos ejemplos de expedientes no encontrados
      if (expedientesNoEncontrados <= 5 || i % 1000 === 0) {
        console.log(`   - [${i+1}/${datosData.length}] Expediente "${expedienteData}" NO encontrado en concentrado`);
      }
      
      // Registrar como no integrado
      reporte.noIntegrados.push({
        expediente: expedienteDataRaw,
        filaData: i + 2,
        motivo: expedienteData ? "Expediente no encontrado en concentrado" : "Valor de expediente vacío"
      });
    }
    
    // Mostrar progreso cada 1000 registros
    if (i > 0 && i % 1000 === 0) {
      console.log(`   - Progreso: ${i}/${datosData.length} registros procesados...`);
    }
  }
  
  // Mostrar resumen de coincidencias
  console.log(`\nRESUMEN DE COINCIDENCIAS:`);
  console.log(`   - Expedientes encontrados: ${expedientesEncontrados}`);
  console.log(`   - Expedientes NO encontrados: ${expedientesNoEncontrados}`);
  
  // Guardar el concentrado actualizado, preservando las propiedades
  if (expedientesEncontrados > 0) {
    console.log("\nGuardando concentrado actualizado (con preservación de formato)...");
    try {
      // Guardar con ExcelJS
      await workbookConcentrado.xlsx.writeFile(archivoConcentrado);
      console.log(`   - Concentrado guardado exitosamente con ${expedientesEncontrados} registros integrados`);
      console.log(`   - FORMATO ORIGINAL PRESERVADO (filtros, columnas ocultas, estilos, etc.)`);
    } catch (error) {
      console.error("Error al guardar el concentrado:", error.message);
    }
    
    // Generar reporte detallado
    await generarReporteExcelJS(reporte, archivoReporte);
  } else {
    console.log("\nNo se encontraron coincidencias. No se modificará el concentrado.");
  }
  
  // Mostrar resumen final
  console.log("\n======================================");
  console.log("         RESUMEN DEL PROCESO          ");
  console.log("======================================");
  console.log(`Total registros en data.xlsx:       ${datosData.length}`);
  console.log(`Registros integrados correctamente: ${expedientesEncontrados}`);
  console.log(`Registros no integrados:            ${expedientesNoEncontrados}`);
  console.log(`Reporte guardado en:                ${archivoReporte}`);
  console.log("======================================");
  console.log("\nIMPORTANTE: Formato original preservado - No es necesario");
  console.log("copiar manualmente el archivo, ya que se ha modificado el");
  console.log("archivo original conservando todos sus formatos.");
  console.log("\nSi desea una copia adicional, puede encontrar el backup en:");
  console.log(`${archivoBackup}`);
  console.log("======================================");
}

// Ejecutar el proceso de merge mejorado
realizarMergeDirecto().catch(error => {
  console.error("Error en el proceso principal:", error);
});