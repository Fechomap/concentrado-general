import fs from "fs";
import path from "path";
import xlsx from "xlsx";

// ======================================
//  CONFIGURACIÓN DE RUTAS Y ARCHIVOS
// ======================================
const carpeta = path.join(process.env.HOME, "Desktop", "concentrado-crk");
const archivoConcentrado = path.join(carpeta, "concentrado-general.xlsx");
const archivoDuplicados = path.join(carpeta, "duplicados.xlsx");

/**
 * Columna que identifica de forma única a cada expediente (columna A).
 * Ajusta este valor si tu encabezado real es diferente.
 */
const COLUMNA_EXPEDIENTE = "Expediente";

// ======================================
//  FUNCIÓN PARA LEER ARCHIVO XLSX
// ======================================
export function leerExcel(rutaArchivo) {
  try {
    const workbook = xlsx.readFile(rutaArchivo, { 
      cellDates: true,
      cellNF: true,
      cellStyles: true
    });
    const hoja = workbook.Sheets[workbook.SheetNames[0]];
    return xlsx.utils.sheet_to_json(hoja, { raw: false, defval: "" });
  } catch (error) {
    console.error(`Error al leer ${rutaArchivo}:`, error.message);
    return [];
  }
}

// ======================================
//  FUNCIÓN PARA GUARDAR DATOS EN XLSX
//  PRESERVANDO FORMATOS ORIGINALES
// ======================================
export function guardarExcel(datos, rutaArchivo) {
  try {
    // Crear workbook y convertir los datos a una hoja
    const workbook = xlsx.utils.book_new();
    const hoja = xlsx.utils.json_to_sheet(datos);
    
    // Si la hoja tiene celdas, configurar formatos específicos
    if (hoja['!ref']) {
      // Obtener encabezados para identificar las columnas
      const headers = {};
      const range = xlsx.utils.decode_range(hoja['!ref']);
      
      // Mapeo de encabezados
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellRef = xlsx.utils.encode_cell({ r: 0, c: col });
        const cellHeader = hoja[cellRef];
        if (cellHeader && cellHeader.v) {
          headers[col] = cellHeader.v;
        }
      }
      
      // Aplicar formato a todas las celdas según su columna
      for (let row = range.s.r + 1; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellRef = xlsx.utils.encode_cell({ r: row, c: col });
          const cell = hoja[cellRef];
          if (!cell) continue;
          
          const headerName = headers[col];
          
          // FORMATEO: Columnas con "fecha" en el nombre (pero no "hora")
          if (/fecha/i.test(headerName) && !/hora/i.test(headerName)) {
            cell.z = 'dd/mm/yyyy';
            // Si no es ya un objeto Date, intentar convertir
            if (!(cell.v instanceof Date)) {
              const fecha = parsearFechaConSeguridad(cell.v);
              if (fecha instanceof Date) {
                cell.v = fecha;
                cell.t = 'd';
              }
            }
          }
          
          // FORMATEO: Columnas con nombres que comienzan con "t" y una letra (ta, tc, tt)
          if (/^t[cat]/i.test(headerName)) { // tiempoArribo, tiempoContacto, tiempoTermino
            cell.z = 'hh:mm:ss';
            // Si no es ya un objeto Date y tiene valor, intentar convertir
            if (!(cell.v instanceof Date) && cell.v) {
              const hora = parsearHoraConSeguridad(cell.v);
              if (hora instanceof Date) {
                cell.v = hora;
                cell.t = 'd';
              }
            }
          }
        }
      }
    }
    
    // Agregar la hoja al workbook y guardar
    xlsx.utils.book_append_sheet(workbook, hoja, "Consolidado");
    xlsx.writeFile(workbook, rutaArchivo);
    
    return true;
  } catch (error) {
    console.error(`Error al guardar ${rutaArchivo}:`, error.message);
    return false;
  }
}

// ======================================
//  PARSEAR FECHA DE MANERA ROBUSTA
// ======================================
function parsearFechaConSeguridad(valor) {
  // Si ya es una fecha, devolver tal cual
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }
  
  // Si no hay valor, retornar null
  if (!valor) return null;
  
  try {
    // Si es número, probablemente sea un número de serie de Excel
    if (!isNaN(valor)) {
      const numValue = Number(valor);
      // Fechas de Excel (número de días desde 1900-01-01)
      if (numValue > 1000) { // Filtro para números grandes que probablemente sean fechas
        // 25569 es el ajuste para 01/01/1970 (fecha UNIX epoch)
        const msDate = (numValue - 25569) * 86400 * 1000;
        const fecha = new Date(msDate);
        if (!isNaN(fecha.getTime())) return fecha;
      }
    }
    
    // Si es string, intentar varios formatos de fecha
    if (typeof valor === 'string') {
      // Primero intentar con formato DD/MM/YYYY o DD-MM-YYYY
      const regexFecha = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/;
      const match = valor.match(regexFecha);
      
      if (match) {
        const dia = parseInt(match[1], 10);
        const mes = parseInt(match[2], 10) - 1; // Meses en JS son 0-11
        const año = parseInt(match[3], 10);
        
        const fecha = new Date(año, mes, dia);
        // Validar que la fecha sea correcta (evita 31/02/2023 por ejemplo)
        if (!isNaN(fecha.getTime()) && 
            fecha.getDate() === dia && 
            fecha.getMonth() === mes && 
            fecha.getFullYear() === año) {
          return fecha;
        }
      }
      
      // Intentar con Date.parse como último recurso
      const parsedDate = new Date(Date.parse(valor));
      if (!isNaN(parsedDate.getTime())) {
        return parsedDate;
      }
    }
  } catch (error) {
    console.error("Error al parsear fecha:", error.message);
  }
  
  return null;
}

// ======================================
//  PARSEAR HORA DE MANERA ROBUSTA
// ======================================
function parsearHoraConSeguridad(valor) {
  // Si ya es una fecha, devolver tal cual
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }
  
  // Si no hay valor, retornar null
  if (!valor) return null;
  
  try {
    // Si es número, probablemente sea un decimal de Excel (fracción de día)
    if (!isNaN(valor)) {
      const numValue = Number(valor);
      if (numValue >= 0 && numValue < 1) { // Solo fracciones de día (0-1)
        const totalSeconds = Math.round(numValue * 24 * 60 * 60);
        const hours = Math.floor(totalSeconds / 3600);
        const minutes = Math.floor((totalSeconds % 3600) / 60);
        const seconds = totalSeconds % 60;
        
        // Base date: 1899-12-30 (fecha base de Excel para tiempos)
        return new Date(1899, 11, 30, hours, minutes, seconds);
      }
    }
    
    // Si es string, intentar varios formatos
    if (typeof valor === 'string') {
      // Formato HH:MM:SS
      const regexHora = /(\d{1,2}):(\d{1,2}):(\d{1,2})/;
      const match = valor.match(regexHora);
      
      if (match) {
        const horas = parseInt(match[1], 10);
        const minutos = parseInt(match[2], 10);
        const segundos = parseInt(match[3], 10);
        
        // Validar rangos
        if (horas >= 0 && horas < 24 && 
            minutos >= 0 && minutos < 60 && 
            segundos >= 0 && segundos < 60) {
          return new Date(1899, 11, 30, horas, minutos, segundos);
        }
      }
      
      // Formato HHMMSS (sin separadores)
      const regexHoraSinSep = /^(\d{2})(\d{2})(\d{2})$/;
      const matchSinSep = valor.match(regexHoraSinSep);
      
      if (matchSinSep) {
        const horas = parseInt(matchSinSep[1], 10);
        const minutos = parseInt(matchSinSep[2], 10);
        const segundos = parseInt(matchSinSep[3], 10);
        
        if (horas >= 0 && horas < 24 && 
            minutos >= 0 && minutos < 60 && 
            segundos >= 0 && segundos < 60) {
          return new Date(1899, 11, 30, horas, minutos, segundos);
        }
      }
    }
  } catch (error) {
    console.error("Error al parsear hora:", error.message);
  }
  
  return null;
}

// ======================================
//  NORMALIZAR REGISTRO
// ======================================
function normalizarRegistro(reg) {
  const nuevo = {};
  for (let key of Object.keys(reg)) {
    let val = reg[key];
    if (typeof val === "string") {
      val = val.trim().replace(/\s+/g, " ");
    }
    nuevo[key] = val;
  }
  return nuevo;
}

// ======================================
//  OBTENER EL VALOR DEL EXPEDIENTE
// ======================================
function obtenerExpediente(registro) {
  let val = registro[COLUMNA_EXPEDIENTE];
  if (!val) {
    const keys = Object.keys(registro);
    if (keys.length > 0) {
      val = registro[keys[0]];
    }
  }
  return typeof val === "string" ? val.trim().replace(/\s+/g, " ") : val;
}

// ======================================
//  PROCESO PRINCIPAL (ACUMULATIVO Y RECONCILIADO)
// ======================================
export function unirYSepararExcels() {
  console.log("\n=== INICIANDO PROCESO DE CONSOLIDACIÓN ===\n");
  
  // Paso 1: Leer los archivos existentes
  console.log("Cargando archivos existentes (si existen)...");
  
  // Mapa de duplicados existentes (si existe el archivo)
  const duplicadosExistentes = new Map();
  if (fs.existsSync(archivoDuplicados)) {
    const registrosDuplicados = leerExcel(archivoDuplicados);
    registrosDuplicados.forEach(reg => {
      const registro = normalizarRegistro(reg);
      const expediente = obtenerExpediente(registro);
      if (expediente) {
        duplicadosExistentes.set(expediente, registro);
      }
    });
    console.log(`   - Cargados ${duplicadosExistentes.size} registros de duplicados existentes.`);
  }
  
  // Paso 2: Leer todos los archivos de origen
  console.log("\nVerificando archivos de origen...");
  
  // Excluir archivos temporales y los archivos de salida
  const tempFileRegex = /^[~$].+/;
  const archivosEntrada = fs.readdirSync(carpeta).filter(f => 
    f.endsWith(".xlsx") && 
    !tempFileRegex.test(f) &&
    f !== path.basename(archivoConcentrado) &&
    f !== path.basename(archivoDuplicados)
  );
  
  console.log(`   - Encontrados ${archivosEntrada.length} archivos fuente.`);
  
  // Inicializar estructuras de datos para el procesamiento
  const registrosPorArchivo = new Map();
  const contadorExpedientes = new Map();
  const registroMasRecientePorExpediente = new Map();
  
  let totalRegistrosLeidos = 0;
  let registrosExistentesMantenidos = 0;
  
  // NUEVO: Cargar registros del concentrado existente si existe
  if (fs.existsSync(archivoConcentrado)) {
    console.log("\nCargando registros del concentrado existente...");
    const registrosConcentrado = leerExcel(archivoConcentrado);
    
    for (const reg of registrosConcentrado) {
      const registro = normalizarRegistro(reg);
      const expediente = obtenerExpediente(registro);
      
      if (expediente) {
        // Agregar al mapa de registros más recientes
        registroMasRecientePorExpediente.set(expediente, registro);
        registrosExistentesMantenidos++;
      }
    }
    
    console.log(`   - Cargados ${registrosExistentesMantenidos} registros de concentrado existente.`);
    
    // Si no hay archivos nuevos para procesar, terminamos sin cambios
    if (archivosEntrada.length === 0) {
      console.log("\nNo hay archivos fuente nuevos para procesar. El concentrado se mantiene sin cambios.");
      return;
    }
  }
  
  // Paso 3: Leer todos los archivos de origen
  console.log("\nLeyendo archivos de origen...");
  
  for (const nombreArchivo of archivosEntrada) {
    const rutaArchivo = path.join(carpeta, nombreArchivo);
    console.log(`   - Procesando ${nombreArchivo}...`);
    
    // Leer todos los registros de este archivo
    const registros = leerExcel(rutaArchivo).map(r => normalizarRegistro(r));
    registrosPorArchivo.set(nombreArchivo, registros);
    
    // Set para evitar contar el mismo expediente más de una vez por archivo
    const expedientesEnEsteArchivo = new Set();
    
    // Mapa de expedientes a sus registros en este archivo
    const mapaRegistrosDeEsteArchivo = new Map();
    
    // Procesar registros
    for (const registro of registros) {
      const expediente = obtenerExpediente(registro);
      if (!expediente) continue;
      
      // Guardar el registro más reciente por expediente (último encontrado)
      registroMasRecientePorExpediente.set(expediente, registro);
      
      // Guardar este registro en el mapa de este archivo
      mapaRegistrosDeEsteArchivo.set(expediente, registro);
      
      // Marcar que este expediente apareció en este archivo
      expedientesEnEsteArchivo.add(expediente);
    }
    
    // Incrementar contador para cada expediente único en este archivo
    for (const expediente of expedientesEnEsteArchivo) {
      contadorExpedientes.set(expediente, (contadorExpedientes.get(expediente) || 0) + 1);
    }
    
    totalRegistrosLeidos += registros.length;
  }
  
  // Total de expedientes únicos (todos los que están en el mapa de registros más recientes)
  const expedientesUnicos = new Set(registroMasRecientePorExpediente.keys());
  console.log(`   - Total registros leídos de archivos fuente: ${totalRegistrosLeidos}`);
  console.log(`   - Total expedientes únicos: ${expedientesUnicos.size}`);
  console.log(`   - Registros existentes mantenidos: ${registrosExistentesMantenidos}`);
  
  // Paso 4: Detectar duplicados dentro de los mismos archivos
  console.log("\nDetectando duplicados reales dentro de cada archivo...");
  
  // Para cada archivo, detectamos duplicados
  const expedientesDuplicados = new Set();
  
  for (const [nombreArchivo, registros] of registrosPorArchivo.entries()) {
    // Map para contar expedientes dentro de este archivo
    const contadorExpedientesEnArchivo = new Map();
    
    // Contar cada expediente en este archivo
    for (const registro of registros) {
      const expediente = obtenerExpediente(registro);
      if (!expediente) continue;
      
      contadorExpedientesEnArchivo.set(
        expediente, 
        (contadorExpedientesEnArchivo.get(expediente) || 0) + 1
      );
    }
    
    // Detectar duplicados en este archivo (aparecen más de una vez)
    let duplicadosEnEsteArchivo = 0;
    
    for (const [expediente, contador] of contadorExpedientesEnArchivo.entries()) {
      if (contador > 1) {
        expedientesDuplicados.add(expediente);
        duplicadosEnEsteArchivo++;
      }
    }
    
    console.log(`   - ${nombreArchivo}: ${duplicadosEnEsteArchivo} duplicados internos.`);
  }
  
  // Paso 5: Preparar el concentrado y los duplicados
  console.log("\nPreparando concentrado general y lista de duplicados...");
  
  // Todos los expedientes van al concentrado (versión más reciente)
  const registrosConcentrado = Array.from(registroMasRecientePorExpediente.values());
  
  // Los duplicados son:
  // 1. Los que ya estaban en duplicados existentes
  // 2. Los que aparecen duplicados dentro de un mismo archivo
  const registrosDuplicados = [];
  let duplicadosNuevos = 0;
  let duplicadosYaRegistrados = 0;
  
  // Procesar cada expediente para determinar si va a duplicados
  for (const expediente of expedientesUnicos) {
    // Si ya estaba en duplicados existentes, lo mantenemos
    if (duplicadosExistentes.has(expediente)) {
      registrosDuplicados.push(duplicadosExistentes.get(expediente));
      duplicadosYaRegistrados++;
    } 
    // Si es un duplicado nuevo, lo agregamos
    else if (expedientesDuplicados.has(expediente)) {
      registrosDuplicados.push(registroMasRecientePorExpediente.get(expediente));
      duplicadosNuevos++;
    }
  }
  
  // Paso 6: Guardar archivos finales
  console.log("\nGuardando archivos finales...");
  
  console.log(`   - Guardando ${registrosConcentrado.length} registros en concentrado...`);
  const resultadoConcentrado = guardarExcel(registrosConcentrado, archivoConcentrado);
  
  console.log(`   - Guardando ${registrosDuplicados.length} registros en duplicados...`);
  const resultadoDuplicados = guardarExcel(registrosDuplicados, archivoDuplicados);
  
  // Paso 7: Generar reporte final con reconciliación de totales
  generarReporteFinal({
    archivosEntrada,
    totalRegistrosLeidos,
    totalExpedientesUnicos: expedientesUnicos.size,
    totalConcentrado: registrosConcentrado.length,
    totalDuplicados: registrosDuplicados.length,
    duplicadosNuevos,
    duplicadosYaRegistrados,
    registrosExistentesMantenidos,
    resultadoConcentrado,
    resultadoDuplicados
  });
}

// ======================================
//  REPORTE FINAL CON RECONCILIACIÓN
// ======================================
function generarReporteFinal({ 
  archivosEntrada,
  totalRegistrosLeidos,
  totalExpedientesUnicos,
  totalConcentrado,
  totalDuplicados,
  duplicadosNuevos,
  duplicadosYaRegistrados,
  registrosExistentesMantenidos = 0,
  resultadoConcentrado,
  resultadoDuplicados
}) {
  console.log("\n======================================");
  console.log("       REPORTE FINAL DE PROCESO       ");
  console.log("======================================");
  
  // Reporte de archivos procesados
  console.log("Archivos procesados:");
  archivosEntrada.forEach(archivo => {
    console.log(`- ${archivo}`);
  });
  
  // Calcular reconciliación
  const diferenciaTotal = totalRegistrosLeidos - totalExpedientesUnicos;
  const diferenciaConcentradoDuplicados = totalExpedientesUnicos - totalConcentrado + totalDuplicados;
  
  console.log("\n--------------------------------------");
  console.log(`Archivos procesados:            ${archivosEntrada.length}`);
  console.log(`Registros totales leídos:       ${totalRegistrosLeidos}`);
  console.log(`Expedientes únicos:             ${totalExpedientesUnicos}`);
  console.log(`Duplicados en registros leídos: ${diferenciaTotal}`);
  console.log(`Expedientes en concentrado:     ${totalConcentrado}`);
  console.log(`  - Registros mantenidos:       ${registrosExistentesMantenidos}`);
  console.log(`  - Registros nuevos/actualizados: ${totalConcentrado - registrosExistentesMantenidos}`);
  console.log(`Expedientes en duplicados:      ${totalDuplicados}`);
  console.log(`  - Duplicados nuevos:          ${duplicadosNuevos}`);
  console.log(`  - Duplicados ya registrados:  ${duplicadosYaRegistrados}`);
  
  // Verificar reconciliación
  if (diferenciaTotal >= 0 && totalConcentrado === totalExpedientesUnicos) {
    console.log(`\n✅ TOTALES RECONCILIADOS CORRECTAMENTE`);
  } else {
    console.log(`\n⚠️ DIFERENCIA EN TOTALES: ${diferenciaConcentradoDuplicados}`);
  }
  
  console.log("--------------------------------------");
  
  if (resultadoConcentrado && resultadoDuplicados) {
    console.log("\n✅ PROCESO COMPLETADO EXITOSAMENTE");
  } else {
    console.log("\n⚠️ PROCESO COMPLETADO CON ADVERTENCIAS");
    if (!resultadoConcentrado) console.log("   - Error al guardar concentrado");
    if (!resultadoDuplicados) console.log("   - Error al guardar duplicados");
  }
  console.log("======================================\n");
}

// ======================================
//  EJECUTAMOS
// ======================================
unirYSepararExcels();