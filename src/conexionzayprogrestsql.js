const { Pool } = require('pg');
const dotenv = require('dotenv');
dotenv.config();

// Creamos un pool de conexiones
const pool = new Pool({
  host:     process.env.PGHOST,
  user:     process.env.PGUSER,
  password: process.env.PGPASSWORD,
  database: process.env.PGDATABASE,
  port:     process.env.PGPORT,
  // Opcionales:
  max:      20,    // número máximo de conexiones en el pool
  idleTimeoutMillis: 30000, // libera conexión tras 30s sin uso
  connectionTimeoutMillis: 2000, // espera hasta 2s para conectar
});

// Función para testear la conexión
async function testConnection() {
  try {
    const client = await pool.connect();
    console.log('✅ Conectado a PostgreSQL');
    client.release();
  } catch (err) {
    console.error('❌ Error al conectar:', err);
  }
}

// (Opcional) versión con “hoy” desde la BD
async function getReporteThyssenDolaresHoy() {
  const sql = `SELECT * FROM sp_reportethyssenhrup_dolares(current_date)`;
  const { rows } = await pool.query(sql);
  return rows;
}
module.exports = {
  testConnection,
  getReporteThyssenDolaresHoy
};
