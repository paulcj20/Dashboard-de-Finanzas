/*
  Genera un archivo .gs con los contenidos de los JSON de la carpeta Planillas.
  Salida: PlanillasData.gs en el directorio raíz del proyecto.
*/
const fs = require('fs');
const path = require('path');

const ROOT = __dirname ? path.resolve(__dirname, '..') : process.cwd();
const SRC_DIR = path.join(ROOT, 'Planillas');
const OUT_FILE = path.join(ROOT, 'PlanillasData.gs');

function main() {
  if (!fs.existsSync(SRC_DIR)) {
    console.error(`No existe la carpeta Planillas: ${SRC_DIR}`);
    process.exit(1);
  }

  const entries = fs.readdirSync(SRC_DIR, { withFileTypes: true });
  const jsonFiles = entries
    .filter((e) => e.isFile() && e.name.toLowerCase().endsWith('.json'))
    .map((e) => e.name);

  const planillas = {};

  for (const file of jsonFiles) {
    const filePath = path.join(SRC_DIR, file);
    try {
      const content = fs.readFileSync(filePath, 'utf8');
      const data = JSON.parse(content);
      const base = path.basename(file, path.extname(file));
      planillas[base] = data;
    } catch (err) {
      console.error(`Error leyendo ${file}:`, err.message);
      process.exitCode = 1;
    }
  }

  const banner = `/**\n * Archivo generado automáticamente por scripts/generate-planillas.js\n * Fuente: carpeta Planillas\n * Fecha: ${new Date().toISOString()}\n * No editar a mano.\n */`;

  // Serializamos como JS válido asignado a una variable global
  const body = `var PLANILLAS = ${JSON.stringify(planillas, null, 2)};`;

  fs.writeFileSync(OUT_FILE, `${banner}\n\n${body}\n`);
  console.log(`Generado ${OUT_FILE} con ${Object.keys(planillas).length} archivo(s).`);
}

main();
