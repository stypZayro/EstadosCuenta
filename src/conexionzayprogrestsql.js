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
/**
 * Ejecuta el SP sp_reporte_thyssenkrup y devuelve sus filas.
 * @param {number} clienteId 
 * @returns {Promise<Array<Object>>}
 */
async function getReporteThyssen(clienteId) {
  const sql = `SELECT * FROM sp_reporte_thyssenkrup($1)`;
  const params = [clienteId];
  const { rows } = await pool.query(sql, params);
  return rows;
}

module.exports = {
  testConnection,
  getReporteThyssen,
};
