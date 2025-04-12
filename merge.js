import fs from "fs";
import path from "path";
import xlsx from "xlsx";

// ======================================
//  CONFIGURACIÓN DE RUTAS Y ARCHIVOS
// ======================================
// Cambio: Ruta base ahora apunta a la subcarpeta merge-general
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
      fs.copyFileSync(archivoConcentrado, archivoBackup);
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
//  LEER ARCHIVO EXCEL
// ======================================
function leerExcel(rutaArchivo) {
  try {
    // Opciones para preservar formatos y fechas
    const opciones = { 
      cellDates: true,
      cellNF: true,
      cellStyles: true,
      raw: false
    };
    
    const workbook = xlsx.readFile(rutaArchivo, opciones);
    const nombreHoja = workbook.SheetNames[0];
    const hoja = workbook.Sheets[nombreHoja];
    
    // Convertir a JSON para procesamiento
    const datos = xlsx.utils.sheet_to_json(hoja, { 
      defval: "",
      raw: false
    });
    
    return { workbook, nombreHoja, hoja, datos };
  } catch (error) {
    console.error(`Error al leer ${rutaArchivo}:`, error.message);
    return { workbook: null, nombreHoja: "", hoja: null, datos: [] };
  }
}

// ======================================
//  REALIZAR MERGE DIRECTO
// ======================================
function realizarMergeDirecto() {
  console.log("\n=== INICIANDO PROCESO DE MERGE DIRECTO ===\n");

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
  
  // Crear backup antes de modificar
  if (!crearBackup()) {
    console.error("No se pudo crear el backup. Abortando proceso por seguridad.");
    return;
  }
  
  // Leer archivos
  console.log("Leyendo archivos...");
  const { workbook: dataWorkbook, datos: datosData } = leerExcel(archivoData);
  const { 
    workbook: concentradoWorkbook, 
    nombreHoja: nombreHojaConcentrado, 
    hoja: hojaConcentrado, 
    datos: datosConcentrado 
  } = leerExcel(archivoConcentrado);
  
  if (!dataWorkbook || !concentradoWorkbook) {
    console.error("Error al leer los archivos Excel");
    return;
  }
  
  console.log(`   - Leídos ${datosData.length} registros de data.xlsx`);
  console.log(`   - Leídos ${datosConcentrado.length} registros de concentrado-general.xlsx`);
  
  // Identificar columna de expediente en data.xlsx
  const COLUMNA_EXPEDIENTE_DATA = "Nº de pieza";
  
  if (!datosData.length || !(COLUMNA_EXPEDIENTE_DATA in datosData[0])) {
    console.error(`Error: No se encontró la columna "${COLUMNA_EXPEDIENTE_DATA}" en data.xlsx`);
    return;
  }
  
  // Crear un mapa de expedientes del concentrado para búsqueda eficiente
  console.log("\nIndexando expedientes del concentrado...");
  const mapaExpedientesConcentrado = new Map();
  
  // Buscar los expedientes en la primera columna (A)
  datosConcentrado.forEach((registro, indice) => {
    // Tomar el primer valor (columna A) como expediente
    const primeraColumna = Object.keys(registro)[0];
    const expedienteRaw = registro[primeraColumna];
    
    // Normalizar para comparación
    const expediente = normalizarExpediente(expedienteRaw);
    
    if (expediente) {
      mapaExpedientesConcentrado.set(expediente, {
        indice,      // Posición en el array (0-based)
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
      console.log(`   - [${info.indice + 1}] "${info.original}" -> "${expediente}"`);
      count++;
    } else {
      break;
    }
  }
  
  // Extraer encabezados de data.xlsx
  const encabezadosData = Object.keys(datosData[0]);
  
  // Preparar contadores para el reporte
  let expedientesEncontrados = 0;
  let expedientesNoEncontrados = 0;
  const registrosNoIntegrados = [];
  const registrosIntegrados = [];
  
  // Información para el reporte final
  const reporte = {
    totalRegistros: datosData.length,
    integrados: [],
    noIntegrados: []
  };
  
  // Obtener el rango actual del concentrado
  const rangoConcentrado = xlsx.utils.decode_range(hojaConcentrado['!ref']);
  
  // Calcular dónde empezar a insertar nuevos datos (columna AW / 49)
  const columnaInicio = 48; // 48 = AW (base 0)
  
  // Insertar encabezados de data.xlsx en la fila 1 del concentrado
  encabezadosData.forEach((encabezado, indice) => {
    const ref = xlsx.utils.encode_cell({ r: 0, c: columnaInicio + indice });
    hojaConcentrado[ref] = { 
      t: 's', 
      v: encabezado,
      s: { font: { bold: true } } // Estilo negrita para encabezados
    };
  });
  
  // Actualizar el rango si se extendió
  if (columnaInicio + encabezadosData.length - 1 > rangoConcentrado.e.c) {
    rangoConcentrado.e.c = columnaInicio + encabezadosData.length - 1;
    hojaConcentrado['!ref'] = xlsx.utils.encode_range(rangoConcentrado);
  }
  
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
      
      // Fila en el concentrado donde insertar los datos (base 0)
      const filaConcentrado = infoExpediente.indice + 1; // +1 porque los datos empiezan en la fila 1 (después de encabezados)
      
      // Mostrar algunos ejemplos de expedientes encontrados
      if (expedientesEncontrados <= 5 || i % 1000 === 0) {
        console.log(`   - [${i+1}/${datosData.length}] Expediente "${expedienteData}" encontrado en fila ${filaConcentrado + 1} del concentrado`);
      }
      
      // Insertar cada columna de data.xlsx en el concentrado
      encabezadosData.forEach((encabezado, indiceColumna) => {
        const valor = registro[encabezado];
        
        if (valor !== "") {
          // Referencia de celda donde insertar
          const ref = xlsx.utils.encode_cell({ r: filaConcentrado, c: columnaInicio + indiceColumna });
          
          // Determinar el tipo de dato para Excel
          let tipo = 's'; // string por defecto
          let valorCelda = valor;
          
          // Detectar números
          if (typeof valor === 'number' || !isNaN(Number(valor))) {
            tipo = 'n';
            valorCelda = typeof valor === 'number' ? valor : Number(valor);
          }
          // Detectar fechas
          else if (valor instanceof Date) {
            tipo = 'd';
            valorCelda = valor;
          }
          
          // Crear la celda con el valor
          hojaConcentrado[ref] = { t: tipo, v: valorCelda };
        }
      });
      
      // Registrar como integrado exitosamente
      reporte.integrados.push({
        Expediente: expedienteDataRaw,
        FilaData: i + 2, // +2 para mostrar fila real en Excel (base 1 + encabezado)
        FilaConcentrado: filaConcentrado + 1, // +1 para mostrar fila real en Excel (base 1)
        Estado: "Integrado"
      });
    } else {
      expedientesNoEncontrados++;
      
      // Mostrar algunos ejemplos de expedientes no encontrados
      if (expedientesNoEncontrados <= 5 || i % 1000 === 0) {
        console.log(`   - [${i+1}/${datosData.length}] Expediente "${expedienteData}" NO encontrado en concentrado`);
      }
      
      // Registrar como no integrado
      reporte.noIntegrados.push({
        Expediente: expedienteDataRaw,
        FilaData: i + 2,
        Motivo: expedienteData ? "Expediente no encontrado en concentrado" : "Valor de expediente vacío"
      });
    }
  }
  
  // Mostrar resumen de coincidencias
  console.log(`\nRESUMEN DE COINCIDENCIAS:`);
  console.log(`   - Expedientes encontrados: ${expedientesEncontrados}`);
  console.log(`   - Expedientes NO encontrados: ${expedientesNoEncontrados}`);
  
  // Guardar el concentrado actualizado
  if (expedientesEncontrados > 0) {
    console.log("\nGuardando concentrado actualizado...");
    try {
      xlsx.writeFile(concentradoWorkbook, archivoConcentrado, { bookSST: true });
      console.log(`   - Concentrado guardado exitosamente con ${expedientesEncontrados} registros integrados`);
    } catch (error) {
      console.error("Error al guardar el concentrado:", error.message);
    }
    
    // Generar reporte detallado
    generarReporte(reporte);
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
  console.log("\nIMPORTANTE: Si el proceso fue exitoso, debe copiar manualmente");
  console.log(`el archivo concentrado actualizado a la carpeta principal:`);
  console.log(`De: ${archivoConcentrado}`);
  console.log(`A:  ${path.join(carpetaBase, "concentrado-general.xlsx")}`);
  console.log("======================================");
}

// ======================================
//  GENERAR REPORTE EXCEL
// ======================================
function generarReporte(reporte) {
  try {
    const workbook = xlsx.utils.book_new();
    
    // Hoja de estadísticas
    const estadisticas = [
      { Estadística: "Total registros en data.xlsx", Valor: reporte.totalRegistros },
      { Estadística: "Registros integrados correctamente", Valor: reporte.integrados.length },
      { Estadística: "Registros no integrados", Valor: reporte.noIntegrados.length }
    ];
    
    const hojaEstadisticas = xlsx.utils.json_to_sheet(estadisticas);
    xlsx.utils.book_append_sheet(workbook, hojaEstadisticas, "Estadísticas");
    
    // Hoja de registros integrados
    if (reporte.integrados.length > 0) {
      const hojaIntegrados = xlsx.utils.json_to_sheet(reporte.integrados);
      xlsx.utils.book_append_sheet(workbook, hojaIntegrados, "Integrados");
    }
    
    // Hoja de registros no integrados
    if (reporte.noIntegrados.length > 0) {
      const hojaNoIntegrados = xlsx.utils.json_to_sheet(reporte.noIntegrados);
      xlsx.utils.book_append_sheet(workbook, hojaNoIntegrados, "No Integrados");
    }
    
    // Guardar el workbook
    xlsx.writeFile(workbook, archivoReporte);
    console.log(`   - Reporte guardado en: ${archivoReporte}`);
    
    return true;
  } catch (error) {
    console.error("Error al generar reporte:", error.message);
    return false;
  }
}

// Ejecutar el proceso de merge
realizarMergeDirecto();