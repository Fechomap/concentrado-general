// node proceso-completo.js
import { exec } from 'child_process';
import { setTimeout } from 'timers/promises';

// Función para crear una línea divisoria en la consola
function imprimirDivisor(mensaje = "") {
  const linea = "=".repeat(50);
  console.log("\n" + linea);
  if (mensaje) {
    console.log(`${mensaje}`);
    console.log(linea);
  }
}

// Función para ejecutar un comando y devolver una promesa con el resultado
function ejecutarComando(comando) {
  return new Promise((resolve, reject) => {
    console.log(`\n> Ejecutando: ${comando}\n`);
    
    const proceso = exec(comando);
    
    // Redirigir salida estándar
    proceso.stdout.on('data', (data) => {
      process.stdout.write(data);
    });
    
    // Redirigir errores
    proceso.stderr.on('data', (data) => {
      process.stderr.write(data);
    });
    
    // Manejar finalización
    proceso.on('close', (code) => {
      if (code === 0) {
        resolve();
      } else {
        reject(new Error(`El proceso falló con código de salida ${code}`));
      }
    });
    
    // Manejar errores del proceso
    proceso.on('error', (err) => {
      reject(err);
    });
  });
}

// Función principal que ejecuta los scripts en secuencia
async function ejecutarProceso() {
  try {
    // Banner inicial
    imprimirDivisor("PROCESO DE CONCENTRACIÓN Y MERGE DE DATOS");
    console.log("\nEjecutando secuencialmente los scripts:");
    console.log("1. index.js - Concentrador de archivos Excel");
    console.log("2. index2.js - Comparación y actualización");
    console.log("3. merge.js - Merge de datos");
    
    // Primera etapa: Ejecutar index.js
    imprimirDivisor("ETAPA 1: EJECUTANDO INDEX.JS");
    await ejecutarComando('node index.js');
    console.log("\n✅ index.js completado");
    
    // Pausa entre procesos (1 segundos)
    console.log("\nEsperando 1 segundos antes de continuar...");
    await setTimeout(1000);
    
    // Segunda etapa: Ejecutar index2.js
    imprimirDivisor("ETAPA 2: EJECUTANDO INDEX2.JS");
    await ejecutarComando('node index2.js');
    console.log("\n✅ index2.js completado");
    
    // Pausa entre procesos (1 segundos)
    console.log("\nEsperando 1 segundos antes de continuar...");
    await setTimeout(1000);
    
    // Tercera etapa: Ejecutar merge.js
    imprimirDivisor("ETAPA 3: EJECUTANDO MERGE.JS");
    await ejecutarComando('node merge.js');
    console.log("\n✅ merge.js completado");
    
    // Finalización
    imprimirDivisor("PROCESO COMPLETADO");
    console.log("\n✅ Todos los scripts se han ejecutado correctamente");
    
  } catch (error) {
    imprimirDivisor("ERROR EN EL PROCESO");
    console.error(`\n❌ Error: ${error.message}`);
  }
}

// Ejecutar el proceso principal
ejecutarProceso();