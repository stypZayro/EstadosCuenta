const fs = require('fs');
const path = require('path');
const sql = require('mssql');
require('dotenv').config();

// --- Configuración de conexión (asegúrate de usar el mismo objeto que usas para tu SP)
const config = { 
  server: process.env.server,
  authentication: {
    type: 'default',
    options: {
      userName: process.env.user, 
      password: process.env.password
    }
  },
  options: {
    port: 1433,
    database: process.env.databaseData,
    trustServerCertificate: true,
  },
  requestTimeout: 1300000,
};
  sql.connect(config)
  .then(() => {
      console.log('Connected to the SQL database DATASTAGE');
  })
  .catch(error => {
      console.error('Error connecting to the database:', error);
  });

// --- Ejecuta el SP pasando el JSON armado
async function ejecutaSPconJSON(tablasJSON) {
  let pool;
  try {
    pool = await sql.connect(config);
    const jsonData = JSON.stringify(tablasJSON);
    const result = await pool
      .request()
      .input('jsonData', sql.NVarChar(sql.MAX), jsonData)
      .execute('DATASTAGE12.dbo.sp_InsertarDesdeJSON');
    console.log('SP ejecutado: RowsAffected:', result.rowsAffected);
    return { ok: true, filas: result.rowsAffected };
  } catch (err) {
    console.error('Error al ejecutar sp_InsertarDesdeJSON:', err);
    throw err;
  } finally {
    await pool?.close();
  }
}



module.exports = {
  ejecutaSPconJSON:ejecutaSPconJSON,
};


