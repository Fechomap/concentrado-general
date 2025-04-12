import fs from "fs";
import path from "path";
import xlsx from "xlsx";

// ======================================
//  CONFIGURACIÓN DE RUTAS Y ARCHIVOS
// ======================================
const carpeta = path.join(process.env.HOME, "Desktop", "concentrado-crk");
const archivoConcentrado = path.join(carpeta, "concentrado-general.xlsx");
const archivoBackup = path.join(carpeta, `concentrado-backup-limpieza-${Date.now()}.xlsx`);

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
//  LIMPIAR DATOS BASURA
// ======================================
function limpiarDatosBasura() {
  console.log("\n=== INICIANDO LIMPIEZA DE CONCENTRADO ===\n");
  
  // Verificar archivo
  if (!fs.existsSync(archivoConcentrado)) {
    console.error("Error: El archivo concentrado-general.xlsx no existe");
    return;
  }
  
  // Crear backup antes de modificar
  if (!crearBackup()) {
    console.error("No se pudo crear el backup. Abortando proceso por seguridad.");
    return;
  }
  
  try {
    // Leer el archivo Excel completo
    console.log("Leyendo archivo concentrado...");
    const workbook = xlsx.readFile(archivoConcentrado, { 
      cellDates: true,
      cellNF: true,
      cellStyles: true
    });
    
    const nombreHoja = workbook.SheetNames[0];
    const hoja = workbook.Sheets[nombreHoja];
    
    // Obtener el rango actual
    const rango = xlsx.utils.decode_range(hoja['!ref']);
    
    console.log(`   - Rango actual: ${hoja['!ref']}`);
    console.log(`   - Filas: ${rango.e.r + 1}, Columnas: ${rango.e.c + 1}`);
    
    // Verificar si hay contenido en las columnas BL-BP (64-67)
    const columnasProblematicas = ['BL', 'BM', 'BN', 'BO', 'BP'];
    const indicesColumnas = columnasProblematicas.map(col => xlsx.utils.decode_col(col));
    
    // Buscar contenido en las columnas problemáticas
    let celdasEncontradas = 0;
    
    // Recorrer todas las filas y columnas problemáticas
    for (let r = 0; r <= rango.e.r; r++) {
      for (const c of indicesColumnas) {
        // Solo verificar si la columna está dentro del rango
        if (c <= rango.e.c) {
          const ref = xlsx.utils.encode_cell({ r, c });
          if (hoja[ref]) {
            celdasEncontradas++;
            
            // Eliminar la celda
            delete hoja[ref];
            
            // Mostrar solo las primeras 10 celdas eliminadas
            if (celdasEncontradas <= 10) {
              console.log(`   - Eliminada celda ${ref} (fila ${r+1}, columna ${xlsx.utils.encode_col(c)})`);
            } else if (celdasEncontradas === 11) {
              console.log("   - (... más celdas eliminadas)");
            }
          }
        }
      }
    }
    
    console.log(`\nSe encontraron y eliminaron ${celdasEncontradas} celdas con datos no deseados.`);
    
    // Guardar el archivo limpio
    if (celdasEncontradas > 0) {
      console.log("\nGuardando concentrado limpio...");
      xlsx.writeFile(workbook, archivoConcentrado, { bookSST: true });
      console.log("   - Archivo guardado exitosamente.");
    } else {
      console.log("\nNo se encontraron datos para limpiar. El archivo permanece sin cambios.");
    }
    
    console.log("\n======================================");
    console.log("       LIMPIEZA DEL CONCENTRADO      ");
    console.log("======================================");
    console.log(`Celdas encontradas y eliminadas: ${celdasEncontradas}`);
    console.log(`Backup creado en: ${archivoBackup}`);
    console.log("======================================");
    
  } catch (error) {
    console.error("Error al limpiar el concentrado:", error.message);
  }
}

// Ejecutar la limpieza
limpiarDatosBasura();