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
            
            // Si no es ya un objeto Date y tiene valor, intentar convertir
            if (!(cell.v instanceof Date) && cell.v) {
              // Guardar el valor original para comparación
              const valorOriginal = cell.v;
              
              // Intentar parsear con nuestro método mejorado
              const fecha = parsearFechaConSeguridad(cell.v);
              
              if (fecha instanceof Date) {
                // Verificar si la fecha resultante es significativamente diferente 
                // del valor original (posible interpretación incorrecta)
                let aplicarFecha = true;
                
                // Si el valor original era un string en formato dd/mm/yyyy, verificamos
                // que la conversión respete exactamente el día
                if (typeof valorOriginal === 'string') {
                  const matchFecha = valorOriginal.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
                  if (matchFecha) {
                    const diaOriginal = parseInt(matchFecha[1], 10);
                    // Si el día no coincide, NO aplicamos la conversión
                    if (fecha.getDate() !== diaOriginal) {
                      aplicarFecha = false;
                      // Intentar una última vez con un enfoque directo para este caso específico
                      const dia = parseInt(matchFecha[1], 10);
                      const mes = parseInt(matchFecha[2], 10) - 1;
                      const año = parseInt(matchFecha[3], 10);
                      const fechaManual = new Date(año, mes, dia);
                      
                      if (!isNaN(fechaManual.getTime()) && 
                          fechaManual.getDate() === dia && 
                          fechaManual.getMonth() === mes && 
                          fechaManual.getFullYear() === año) {
                        cell.v = fechaManual;
                        cell.t = 'd';
                      }
                    }
                  }
                }
                
                // Solo aplicamos la conversión si pasó las verificaciones
                if (aplicarFecha) {
                  cell.v = fecha;
                  cell.t = 'd';
                }
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
//  PARSEAR FECHA DE MANERA ROBUSTA Y PRECISA
// ======================================
function parsearFechaConSeguridad(valor) {
  // Si ya es una fecha, devolver tal cual - nunca modificar una fecha válida
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }
  
  // Si no hay valor, retornar null
  if (!valor) return null;
  
  try {
    // CASO 1: Valor numérico - posible serial de Excel
    if (!isNaN(valor)) {
      const numValue = Number(valor);
      
      // Verificar si es un número serial de Excel válido (fechas razonables)
      // Excel inicia en 1/1/1900 (serial 1), añadimos filtro para evitar falsos positivos
      if (numValue >= 1000 && numValue <= 50000) { // Rango de fechas razonables (~1903 hasta ~2036)
        // IMPORTANTE: El ajuste exacto para corregir el problema del año bisiesto 1900 en Excel
        // y para convertir correctamente a la epoch de JavaScript (1/1/1970)
        
        // Primero creamos una fecha con año/mes/día exactos según la numeración de Excel
        // Este enfoque evita problemas de zona horaria
        const diasDesde1900 = Math.floor(numValue);
        
        // 1. Calculamos la fecha exacta usando aritmética de fechas
        // Nota: La fecha 0 en Excel es 0/1/1900
        
        // Establecemos el 31/12/1899 como fecha base (un día antes del 1/1/1900)
        // y sumamos los días correspondientes al serial
        const fechaExcel = new Date(1899, 11, 31);
        fechaExcel.setDate(fechaExcel.getDate() + diasDesde1900);
        
        // Si el número incluye fracción de día (hora), la añadimos
        const fraccionDia = numValue - diasDesde1900;
        if (fraccionDia > 0) {
          const milisegundosDia = 24 * 60 * 60 * 1000;
          fechaExcel.setTime(fechaExcel.getTime() + Math.round(fraccionDia * milisegundosDia));
        }
        
        // 2. Validamos que la fecha sea correcta
        if (!isNaN(fechaExcel.getTime())) {
          return fechaExcel;
        }
      }
    }
    
    // CASO 2: String - intentamos varios formatos de fecha con precisión
    if (typeof valor === 'string') {
      // Limpiar el string de espacios extra y caracteres no deseados
      const valorLimpio = valor.trim().replace(/\s+/g, " ");
      
      // CASO 2.1: Formato exacto DD/MM/YYYY o DD-MM-YYYY
      // Este es el formato más común en sistemas españoles/latinos
      const regexFechaLatina = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/;
      const matchLatino = valorLimpio.match(regexFechaLatina);
      
      if (matchLatino) {
        const dia = parseInt(matchLatino[1], 10);
        const mes = parseInt(matchLatino[2], 10) - 1; // Meses en JS son 0-11
        const año = parseInt(matchLatino[3], 10);
        
        // Crear fecha usando UTC para evitar cualquier ajuste por zona horaria
        // Luego convertimos a fecha local preservando exactamente los valores
        const fechaUTC = new Date(Date.UTC(año, mes, dia));
        const fecha = new Date(año, mes, dia);
        
        // Validar fecha: verificamos que el día, mes y año sean los esperados
        // Esto protege contra fechas inválidas como 31/02/2023
        if (!isNaN(fecha.getTime()) && 
            fecha.getDate() === dia && 
            fecha.getMonth() === mes && 
            fecha.getFullYear() === año) {
          return fecha;
        }
      }
      
      // CASO 2.2: Formato YYYY-MM-DD (formato ISO)
      const regexFechaISO = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;
      const matchISO = valorLimpio.match(regexFechaISO);
      
      if (matchISO) {
        const año = parseInt(matchISO[1], 10);
        const mes = parseInt(matchISO[2], 10) - 1; // Meses en JS son 0-11
        const dia = parseInt(matchISO[3], 10);
        
        const fecha = new Date(año, mes, dia);
        
        // Validar como en el caso anterior
        if (!isNaN(fecha.getTime()) && 
            fecha.getDate() === dia && 
            fecha.getMonth() === mes && 
            fecha.getFullYear() === año) {
          return fecha;
        }
      }
      
      // CASO 2.3: Verificar otros formatos pero preservando con precisión la fecha
      // NOTA: Date.parse es peligroso para formatos ambiguos, usamos solo en casos de último recurso
      
      // EVITAMOS usar Date.parse() directamente aquí para prevenir interpretaciones
      // ambiguas que puedan alterar el día debido a zonas horarias.
      
      // En su lugar, intentamos identificar patrones comunes y garantizar precisión
      
      // Por ejemplo: dd/mm/yyyy HH:MM:SS o formatos similares
      const regexFechaHora = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?/;
      const matchFechaHora = valorLimpio.match(regexFechaHora);
      
      if (matchFechaHora) {
        const dia = parseInt(matchFechaHora[1], 10);
        const mes = parseInt(matchFechaHora[2], 10) - 1;
        const año = parseInt(matchFechaHora[3], 10);
        const hora = parseInt(matchFechaHora[4], 10);
        const minutos = parseInt(matchFechaHora[5], 10);
        const segundos = matchFechaHora[6] ? parseInt(matchFechaHora[6], 10) : 0;
        
        const fecha = new Date(año, mes, dia, hora, minutos, segundos);
        
        // Validación rigurosa para asegurar exactitud
        if (!isNaN(fecha.getTime()) && 
            fecha.getDate() === dia && 
            fecha.getMonth() === mes && 
            fecha.getFullYear() === año &&
            fecha.getHours() === hora &&
            fecha.getMinutes() === minutos) {
          return fecha;
        }
      }
    }
  } catch (error) {
    console.error("Error al parsear fecha:", error.message, "para valor:", valor);
  }
  
  // Si llegamos aquí, no pudimos parsear la fecha con seguridad.
  // Devolvemos null en lugar de intentar usar métodos menos fiables
  // para evitar posibles errores de interpretación.
  return null;
}

// ======================================
//  PARSEAR HORA DE MANERA ROBUSTA Y PRECISA
// ======================================
function parsearHoraConSeguridad(valor) {
  // Si ya es una fecha, devolver tal cual - preservar valores válidos
  if (valor instanceof Date && !isNaN(valor.getTime())) {
    return valor;
  }
  
  // Si no hay valor, retornar null
  if (!valor) return null;
  
  try {
    // CASO 1: Valor numérico - posible fracción de día de Excel
    if (!isNaN(valor)) {
      const numValue = Number(valor);
      
      // Las horas en Excel son fracciones de día entre 0 y 1
      if (numValue >= 0 && numValue < 1) {
        // Convertimos la fracción en milisegundos
        const millisInDay = 24 * 60 * 60 * 1000;
        const millisFromFraction = Math.round(numValue * millisInDay);
        
        // Para horas, usamos una fecha base fija para evitar cualquier
        // problema de zona horaria o DST
        const baseDate = new Date(1899, 11, 30, 0, 0, 0, 0);
        const timeDate = new Date(baseDate.getTime() + millisFromFraction);
        
        // Verificar que la conversión resultó en una fecha válida
        if (!isNaN(timeDate.getTime())) {
          return timeDate;
        }
      }
    }
    
    // CASO 2: String - intentamos varios formatos de hora con precisión
    if (typeof valor === 'string') {
      // Limpiar el string de espacios extra
      const valorLimpio = valor.trim().replace(/\s+/g, " ");
      
      // CASO 2.1: Formato HH:MM:SS (el más común)
      const regexHora = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?$/;
      const match = valorLimpio.match(regexHora);
      
      if (match) {
        const horas = parseInt(match[1], 10);
        const minutos = parseInt(match[2], 10);
        const segundos = match[3] ? parseInt(match[3], 10) : 0;
        
        // Validar rangos para asegurar que es una hora válida
        if (horas >= 0 && horas < 24 && 
            minutos >= 0 && minutos < 60 && 
            segundos >= 0 && segundos < 60) {
            
          // Usamos una fecha base fija para almacenar solo el tiempo
          const baseDate = new Date(1899, 11, 30);
          baseDate.setHours(horas, minutos, segundos, 0);
          
          return baseDate;
        }
      }
      
      // CASO 2.2: Formato HHMMSS (sin separadores)
      const regexHoraSinSep = /^(\d{2})(\d{2})(\d{2})$/;
      const matchSinSep = valorLimpio.match(regexHoraSinSep);
      
      if (matchSinSep) {
        const horas = parseInt(matchSinSep[1], 10);
        const minutos = parseInt(matchSinSep[2], 10);
        const segundos = parseInt(matchSinSep[3], 10);
        
        // Validar rangos
        if (horas >= 0 && horas < 24 && 
            minutos >= 0 && minutos < 60 && 
            segundos >= 0 && segundos < 60) {
            
          // Mismo enfoque de fecha base fija
          const baseDate = new Date(1899, 11, 30);
          baseDate.setHours(horas, minutos, segundos, 0);
          
          return baseDate;
        }
      }
      
      // CASO 2.3: Formato de 12 horas (HH:MM AM/PM)
      const regexHora12h = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?\s*(AM|PM|am|pm)$/i;
      const match12h = valorLimpio.match(regexHora12h);
      
      if (match12h) {
        let horas = parseInt(match12h[1], 10);
        const minutos = parseInt(match12h[2], 10);
        const segundos = match12h[3] ? parseInt(match12h[3], 10) : 0;
        const esPM = match12h[4].toLowerCase() === 'pm';
        
        // Ajustar las horas para formato 12h
        if (esPM && horas < 12) horas += 12;
        if (!esPM && horas === 12) horas = 0;
        
        // Validar rangos
        if (horas >= 0 && horas < 24 && 
            minutos >= 0 && minutos < 60 && 
            segundos >= 0 && segundos < 60) {
            
          const baseDate = new Date(1899, 11, 30);
          baseDate.setHours(horas, minutos, segundos, 0);
          
          return baseDate;
        }
      }
    }
  } catch (error) {
    console.error("Error al parsear hora:", error.message, "para valor:", valor);
  }
  
  // Si llegamos aquí, no pudimos parsear la hora con seguridad
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