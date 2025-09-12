const mysql = require('mysql2');
const dotenv = require('dotenv');
dotenv.config();

const config = {
  host: '192.168.0.37',
  user: 'slamm3',
  password: 'masterm3',
  database: 'turbosaai',
  port: 3306,
};

const connection = mysql.createConnection(config);


async function sp_ObtenerDatosFacturaSemanal_Thyssen(listaPedimentos, impexp) {
  return new Promise((resolve, reject) => {
    const query = `CALL sp_ObtenerDatosFacturaSemanal_Thyssen(?, ?)`;

    connection.query(query, [ listaPedimentos, impexp], (error, results) => {
      if (error) {
        console.error('Error al ejecutar el SP: sp_ObtenerDatosFacturaSemanal_Thyssen', error.message);
        return reject(`Error al ejecutar el SP: ${error.message}`);
      }

      const datos = results[0]; // Primer conjunto de resultados
      resolve(datos || []);
    });
  });
}



async function sp_ObtenerDatosFacturaexpoSemanal_Thyssen(listaPedimentos, impexp) {
  return new Promise((resolve, reject) => {
    const query = `CALL sp_ObtenerDatosFacturaexpoSemanal_Thyssen(?, ?)`;

    connection.query(query, [ listaPedimentos, impexp], (error, results) => {
      if (error) {
        console.error('Error al ejecutar el SP: sp_ObtenerDatosFacturaexpoSemanal_Thyssen', error.message);
        return reject(`Error al ejecutar el SP: ${error.message}`);
      }

      const datos = results[0]; // Primer conjunto de resultados
      resolve(datos || []);
    });
  });
}




module.exports = {
  sp_ObtenerDatosFacturaSemanal_Thyssen:sp_ObtenerDatosFacturaSemanal_Thyssen,
  sp_ObtenerDatosFacturaexpoSemanal_Thyssen:sp_ObtenerDatosFacturaexpoSemanal_Thyssen,
};