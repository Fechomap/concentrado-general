import fs from "fs";
import path from "path";
import ExcelJS from "exceljs";

// ======================================
//  CONFIGURACIÓN DE RUTAS Y ARCHIVOS
// ======================================
const rutaPrincipal = path.join(process.env.HOME, "Desktop", "concentrado-crk");
const rutaMerge = path.join(rutaPrincipal, "merge-general");

const archivoPrincipal = path.join(rutaPrincipal, "concentrado-general.xlsx");
const archivoMerge = path.join(rutaMerge, "concentrado-general.xlsx");

// Backup del archivo de merge (para seguridad)
const archivoBackup = path.join(rutaMerge, `concentrado-general-backup-${Date.now()}.xlsx`);

// ======================================
//  CREAR BACKUP DEL ARCHIVO DE MERGE
// ======================================
function crearBackup() {
  try {
    if (fs.existsSync(archivoMerge)) {
      // Usamos el método de bajo nivel para copiar el archivo byte a byte
      // Esto asegura que se preserve absolutamente todo el contenido y formato
      const contenidoOriginal = fs.readFileSync(archivoMerge);
      fs.writeFileSync(archivoBackup, contenidoOriginal);
      console.log(`Backup creado: ${archivoBackup}`);
      return true;
    } else {
      console.error(`Error: El archivo ${archivoMerge} no existe`);
      return false;
    }
  } catch (error) {
    console.error("Error al crear backup:", error.message);
    return false;
  }
}

// ======================================
//  NORMALIZAR VALOR DE LA CLAVE
// ======================================
function normalizarClave(valor) {
  if (valor === null || valor === undefined) return "";
  
  // Si es un objeto Date o un número, convertir a string
  if (valor instanceof Date) {
    return valor.toISOString();
  } else if (typeof valor === 'number') {
    return String(valor);
  } else if (typeof valor === 'boolean') {
    return valor ? "true" : "false";
  }
  
  // Convertir a string y eliminar espacios extras
  const valorString = String(valor).trim();
  return valorString.replace(/\s+/g, "");
}

// ======================================
//  ENCONTRAR ÚLTIMA FILA CON DATOS
// ======================================
function encontrarUltimaFilaConDatos(worksheet) {
  let ultimaFila = 0;
  
  // Recorrer filas desde el final hacia el principio
  for (let i = worksheet.rowCount; i >= 1; i--) {
    const fila = worksheet.getRow(i);
    let tieneValores = false;
    
    // Verificar si la fila tiene al menos un valor
    fila.eachCell({ includeEmpty: false }, () => {
      tieneValores = true;
    });
    
    if (tieneValores) {
      ultimaFila = i;
      break;
    }
  }
  
  return ultimaFila;
}

// ======================================
//  LEER ARCHIVO EXCEL CON EXCELJS
// ======================================
async function leerArchivoExcel(rutaArchivo, descripcion = "Excel") {
  try {
    console.log(`Leyendo archivo ${rutaArchivo}...`);
    
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
    
    // Buscar la última fila con datos reales (evitar filas vacías al final)
    const ultimaFilaReal = encontrarUltimaFilaConDatos(worksheet);
    console.log(`  - Última fila con datos reales: ${ultimaFilaReal}`);
    
    // Extraer encabezados (primera fila)
    const encabezados = [];
    const primeraFila = worksheet.getRow(1);
    primeraFila.eachCell((cell, colNumber) => {
      encabezados[colNumber - 1] = cell.value?.toString() || `Columna_${colNumber}`;
    });
    
    console.log(`  - Encabezados encontrados: ${encabezados.length}`);
    if (encabezados.length > 0) {
      console.log(`  - Primer encabezado: ${encabezados[0]}`);
    }
    
    // Extraer datos (solo necesitamos la columna A para comparar)
    const datos = [];
    const clavesUnicas = new Set();
    
    // Empezar desde la fila 2 (después de encabezados) hasta la última fila real
    for (let i = 2; i <= ultimaFilaReal; i++) {
      const fila = worksheet.getRow(i);
      
      // Obtenemos solo el valor de la columna A (índice 1)
      const celdaColumnaA = fila.getCell(1);
      const valorClave = celdaColumnaA.value;
      
      // Solo procesamos filas con valor en columna A
      if (valorClave !== null && valorClave !== undefined && valorClave !== "") {
        const claveNormalizada = normalizarClave(valorClave);
        
        // Solo procesamos claves no vacías después de normalizar
        if (claveNormalizada !== "") {
          // Guardamos el valor original y la clave normalizada
          datos.push({
            valorOriginal: valorClave,
            claveNormalizada: claveNormalizada,
            indice: i
          });
          
          // Agregamos a conjunto de claves únicas
          clavesUnicas.add(claveNormalizada);
        }
      }
      
      // Mostrar progreso cada 1000 filas o en la última fila
      if (i % 1000 === 0 || i === ultimaFilaReal) {
        console.log(`  - Procesadas ${i} de ${ultimaFilaReal} filas...`);
      }
    }
    
    console.log(`  - Total registros extraídos: ${datos.length}`);
    console.log(`  - Total claves únicas: ${clavesUnicas.size}`);
    
    // Si hay discrepancia entre registros extraídos y claves únicas, reportar duplicados
    if (datos.length !== clavesUnicas.size) {
      console.log(`  - Atención: Se detectaron ${datos.length - clavesUnicas.size} claves duplicadas`);
    }
    
    return { 
      datos, 
      encabezados, 
      workbook, 
      worksheet,
      clavesUnicas: Array.from(clavesUnicas),
      ultimaFilaReal
    };
  } catch (error) {
    console.error(`Error al leer ${rutaArchivo}:`, error.message);
    console.error(error.stack);
    return { 
      datos: [], 
      encabezados: [], 
      workbook: null, 
      worksheet: null, 
      clavesUnicas: [],
      ultimaFilaReal: 0 
    };
  }
}

// ======================================
//  COPIAR CELDA CON FORMATO COMPLETO
// ======================================
function copiarCeldaConFormato(celdaOrigen, celdaDestino) {
  try {
    // Copiar valor
    celdaDestino.value = celdaOrigen.value;
    
    // Copiar formato
    if (celdaOrigen.style) {
      // Intentar copiar estilo
      try {
        // Crear copia profunda del estilo para evitar referencias
        celdaDestino.style = JSON.parse(JSON.stringify(celdaOrigen.style));
      } catch (styleError) {
        // En caso de error, aplicar propiedades de estilo una por una
        if (celdaOrigen.style.font) celdaDestino.font = celdaOrigen.style.font;
        if (celdaOrigen.style.fill) celdaDestino.fill = celdaOrigen.style.fill;
        if (celdaOrigen.style.border) celdaDestino.border = celdaOrigen.style.border;
        if (celdaOrigen.style.alignment) celdaDestino.alignment = celdaOrigen.style.alignment;
      }
    }
    
    // Copiar formato de número si existe
    if (celdaOrigen.numFmt) {
      celdaDestino.numFmt = celdaOrigen.numFmt;
    }
    
    // Copiar fórmula si existe
    if (celdaOrigen.formula) {
      celdaDestino.formula = celdaOrigen.formula;
    }
    
    // Copiar hipervínculo si existe
    if (celdaOrigen.hyperlink) {
      celdaDestino.hyperlink = celdaOrigen.hyperlink;
    }
    
    // Copiar comentarios si existen
    if (celdaOrigen.note) {
      celdaDestino.note = celdaOrigen.note;
    }
  } catch (error) {
    // En caso de error, al menos asegurar que se copie el valor
    celdaDestino.value = celdaOrigen.value;
  }
}

// ======================================
//  PROCESO PRINCIPAL
// ======================================
async function compararYActualizarArchivos() {
  console.log("\n=== INICIANDO PROCESO DE COMPARACIÓN Y ACTUALIZACIÓN ===\n");
  
  // Verificar que existan las carpetas y archivos necesarios
  if (!fs.existsSync(rutaPrincipal)) {
    console.error(`Error: La carpeta principal ${rutaPrincipal} no existe`);
    return;
  }
  
  if (!fs.existsSync(rutaMerge)) {
    console.log(`Creando carpeta de merge: ${rutaMerge}`);
    try {
      fs.mkdirSync(rutaMerge, { recursive: true });
    } catch (error) {
      console.error(`Error al crear carpeta de merge: ${error.message}`);
      return;
    }
  }
  
  // Verificar archivos
  if (!fs.existsSync(archivoPrincipal)) {
    console.error(`Error: El archivo principal ${archivoPrincipal} no existe`);
    return;
  }
  
  if (!fs.existsSync(archivoMerge)) {
    console.error(`Error: El archivo de merge ${archivoMerge} no existe`);
    console.log("IMPORTANTE: Debe existir un archivo de merge previo para realizar la comparación.");
    return;
  }
  
  // Crear backup antes de modificar (por seguridad)
  if (!crearBackup()) {
    console.error("No se pudo crear el backup. Abortando proceso por seguridad.");
    return;
  }
  
  // Paso 1: Leer ambos archivos Excel
  console.log("\n1. Leyendo archivos Excel...");
  
  // Leer archivo principal
  const resultadoPrincipal = await leerArchivoExcel(archivoPrincipal, "principal");
  if (!resultadoPrincipal.datos || resultadoPrincipal.datos.length === 0) {
    console.error("Error: No se pudieron extraer datos del archivo principal");
    return;
  }
  
  // Leer archivo de merge
  const resultadoMerge = await leerArchivoExcel(archivoMerge, "merge");
  if (!resultadoMerge.datos || resultadoMerge.datos.length === 0) {
    console.error("Error: No se pudieron extraer datos del archivo de merge");
    return;
  }
  
  // Paso 2: Identificar registros nuevos (claves en archivo principal que no están en merge)
  console.log("\n2. Identificando registros nuevos...");
  
  // Obtenemos las claves del archivo de merge para búsqueda eficiente
  const clavesMerge = new Set(resultadoMerge.clavesUnicas);
  
  // Identificar claves nuevas (en archivo principal pero no en merge)
  const clavesNuevas = resultadoPrincipal.clavesUnicas.filter(clave => !clavesMerge.has(clave));
  
  console.log(`   - Claves únicas en archivo principal: ${resultadoPrincipal.clavesUnicas.length}`);
  console.log(`   - Claves únicas en archivo de merge: ${resultadoMerge.clavesUnicas.length}`);
  console.log(`   - Claves nuevas encontradas: ${clavesNuevas.length}`);
  
  // Si no hay claves nuevas, terminamos el proceso
  if (clavesNuevas.length === 0) {
    console.log("\n✅ No se encontraron diferencias. No se realizaron inserciones.");
    return;
  }
  
  // Paso 3: Identificar los registros completos a insertar
  console.log("\n3. Preparando registros para inserción...");
  
  // Mapa para búsqueda eficiente por clave normalizada
  const mapaRegistrosPrincipal = new Map();
  resultadoPrincipal.datos.forEach(registro => {
    mapaRegistrosPrincipal.set(registro.claveNormalizada, registro);
  });
  
  // Registros a insertar (completos, con todas sus columnas)
  const registrosAInsertar = [];
  
  // Listado de claves que se insertarán (mostrar hasta 10, luego resumir el resto)
  console.log("\nListado de claves a insertar:");
  const maxClavesAMostrar = 10;
  let clavesVisibles = 0;
  
  for (const claveNueva of clavesNuevas) {
    // Obtener registro completo del archivo principal
    const registroInfo = mapaRegistrosPrincipal.get(claveNueva);
    
    // Registrar para inserción
    if (registroInfo) {
      registrosAInsertar.push({
        indice: registroInfo.indice,
        clave: claveNueva,
        valorOriginal: registroInfo.valorOriginal
      });
      
      // Mostrar en consola (limitado para no saturar)
      if (clavesVisibles < maxClavesAMostrar) {
        console.log(`   - ${registroInfo.valorOriginal} [clave: ${claveNueva}]`);
        clavesVisibles++;
      }
    }
  }
  
  // Si hay más claves de las mostradas, indicar cuántas quedan
  if (clavesNuevas.length > maxClavesAMostrar) {
    const restantes = clavesNuevas.length - maxClavesAMostrar;
    console.log(`   - ... y ${restantes} más.`);
  }
  
  // Paso 4: Insertar registros nuevos al final del archivo de merge
  console.log(`\n4. Insertando ${registrosAInsertar.length} registros nuevos al archivo de merge...`);
  
  // Obtener workbook y worksheet del archivo de merge
  const workbookMerge = resultadoMerge.workbook;
  const worksheetMerge = resultadoMerge.worksheet;
  
  // Obtener workbook y worksheet del archivo principal
  const workbookPrincipal = resultadoPrincipal.workbook;
  const worksheetPrincipal = resultadoPrincipal.worksheet;
  
  // Determinar la última fila del archivo de merge (donde insertaremos)
  // Usamos la última fila real con datos, no la propiedad rowCount que puede incluir filas vacías
  let ultimaFilaMerge = resultadoMerge.ultimaFilaReal;
  console.log(`   - Última fila con datos en archivo de merge: ${ultimaFilaMerge}`);
  
  // Insertar registros nuevos
  let registrosInsertados = 0;
  for (const registro of registrosAInsertar) {
    try {
      // Incrementar contador de última fila
      ultimaFilaMerge++;
      
      // Obtener la fila completa del archivo principal
      const filaPrincipal = worksheetPrincipal.getRow(registro.indice);
      
      // Crear una nueva fila en el archivo de merge
      const filaMerge = worksheetMerge.getRow(ultimaFilaMerge);
      
      // Copiar todas las celdas de la fila del archivo principal a la nueva fila en merge
      filaPrincipal.eachCell((cell, colNumber) => {
        const celdaMerge = filaMerge.getCell(colNumber);
        copiarCeldaConFormato(cell, celdaMerge);
      });
      
      registrosInsertados++;
      
      // Mostrar progreso cada 100 registros
      if (registrosInsertados % 100 === 0) {
        console.log(`   - Progreso: ${registrosInsertados}/${registrosAInsertar.length} registros insertados...`);
      }
    } catch (error) {
      console.error(`Error al insertar registro con clave ${registro.clave}:`, error.message);
    }
  }
  
  // Paso 5: Guardar el archivo de merge con los nuevos registros
  console.log("\n5. Guardando archivo de merge actualizado...");
  try {
    await workbookMerge.xlsx.writeFile(archivoMerge);
    console.log(`   - Archivo de merge guardado exitosamente con ${registrosInsertados} registros nuevos insertados`);
  } catch (error) {
    console.error("Error al guardar el archivo de merge:", error.message);
    console.error("IMPORTANTE: No se pudieron guardar los cambios. Verifica permisos y que el archivo no esté abierto.");
    return;
  }
  
  // Mostrar resumen final
  console.log("\n======================================");
  console.log("         RESUMEN DEL PROCESO          ");
  console.log("======================================");
  console.log(`Total registros en archivo principal: ${resultadoPrincipal.datos.length}`);
  console.log(`Total registros en archivo de merge:  ${resultadoMerge.datos.length}`);
  console.log(`Claves únicas en archivo principal:   ${resultadoPrincipal.clavesUnicas.length}`);
  console.log(`Claves únicas en archivo de merge:    ${resultadoMerge.clavesUnicas.length}`);
  console.log(`Claves nuevas detectadas:             ${clavesNuevas.length}`);
  console.log(`Registros nuevos insertados:          ${registrosInsertados}`);
  console.log(`Nuevo total en archivo de merge:      ${resultadoMerge.datos.length + registrosInsertados}`);
  console.log(`Backup guardado en:                   ${archivoBackup}`);
  console.log("======================================");
  console.log("\nIMPORTANTE: Se ha preservado el formato original del archivo de merge");
  console.log("sin alterar su estructura. Solo se han agregado los registros nuevos al final.");
  console.log("======================================");
}

// Ejecutar el proceso principal
compararYActualizarArchivos().catch(error => {
  console.error("Error en el proceso principal:", error);
  console.error(error.stack);
});