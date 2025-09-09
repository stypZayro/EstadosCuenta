const mysql = require('mysql2');
const dotenv = require('dotenv');
dotenv.config();

const config = {
  host: process.env.DB_HOST,
  user: process.env.DB_USER,
  password: process.env.DB_PASSWORD,
  database: process.env.DB_NAME,
  port: 26533,
};

const connection = mysql.createConnection(config);

connection.connect(error => {
  if (error) {
    console.error('Error connecting to the database:', error);
    return;
  }
  console.log('Connected to the MySQL database');
});

async function sp_obtener_datos_identificadores() {
  try {
    return new Promise((resolve, reject) => {
      const query = `CALL sp_obtener_datos_identificadores()`;
      connection.query(query, (error, results) => {
        if (error) {
          return reject('Error al conectar o consultar la base de datos: sp_obtener_datos_identificadores ', error.message);
        }
        const datos = results[0]; // Primer array de resultados
        if (datos.length === 0) {
          resolve([]);
        } else {
          resolve(datos);
        }
      });
    });
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
  }
}
async function sp_obtener_referencia(Referencia) {
  try {
    return new Promise((resolve, reject) => {
      const query = `CALL sp_obtener_referencia(?)`;
      connection.query(query, [Referencia], (error, results) => {
        if (error) {
          return reject('Error al conectar o consultar la base de datos: sp_obtener_referencia', error.message);
        }
        const datos = results[0]; // Primer array de resultados
        if (datos.length === 0) {
          resolve([]);
        } else {
          resolve(datos);
        }
      });
    });
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
  }
}
async function sp_obtener_datos_permisos() {
  try {
    return new Promise((resolve, reject) => {
      const query = `CALL sp_obtener_datos_permisos()`;
      connection.query(query, (error, results) => {
        if (error) {
          return reject('Error al conectar o consultar la base de datos: sp_obtener_datos_permisos ', error.message);
        }
        const datos = results[0]; // Primer array de resultados
        if (datos.length === 0) {
          resolve([]);
        } else {
          resolve(datos);
        }
      });
    });
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
  }
}

async function sp_ObtenerDatosFactura(listaPedimentos, impexp) {
  return new Promise((resolve, reject) => {
    const query = `CALL sp_ObtenerDatosFactura(?, ?)`;

    connection.query(query, [ listaPedimentos, impexp], (error, results) => {
      if (error) {
        console.error('Error al ejecutar el SP: sp_ObtenerDatosFactura', error.message);
        return reject(`Error al ejecutar el SP: ${error.message}`);
      }

      const datos = results[0]; // Primer conjunto de resultados
      resolve(datos || []);
    });
  });
}

async function sp_ObtenerDatosFacturaSemanal(listaPedimentos, impexp) {
  return new Promise((resolve, reject) => {
    const query = `CALL sp_ObtenerDatosFactura(?, ?)`;

    connection.query(query, [ listaPedimentos, impexp], (error, results) => {
      if (error) {
        console.error('Error al ejecutar el SP: sp_ObtenerDatosFacturaSemanal', error.message);
        return reject(`Error al ejecutar el SP: ${error.message}`);
      }

      const datos = results[0]; // Primer conjunto de resultados
      resolve(datos || []);
    });
  });
}
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

async function sp_ObtenerDatosFacturaexpo(listaPedimentos, impexp) {
  return new Promise((resolve, reject) => {
    const query = `CALL sp_ObtenerDatosFacturaexpo(?, ?)`;

    connection.query(query, [ listaPedimentos, impexp], (error, results) => {
      if (error) {
        console.error('Error al ejecutar el SP: sp_ObtenerDatosFacturaexpo', error.message);
        return reject(`Error al ejecutar el SP: ${error.message}`);
      }

      const datos = results[0]; // Primer conjunto de resultados
      resolve(datos || []);
    });
  });
}
async function sp_ObtenerDatosFacturaexpoSemanal(listaPedimentos, impexp) {
  return new Promise((resolve, reject) => {
    const query = `CALL sp_ObtenerDatosFacturaexpo(?, ?)`;

    connection.query(query, [ listaPedimentos, impexp], (error, results) => {
      if (error) {
        console.error('Error al ejecutar el SP: sp_ObtenerDatosFacturaexpoSemanal', error.message);
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
  sp_obtener_datos_identificadores:sp_obtener_datos_identificadores,
  sp_obtener_referencia:sp_obtener_referencia,
  sp_obtener_datos_permisos:sp_obtener_datos_permisos,
  sp_ObtenerDatosFactura:sp_ObtenerDatosFactura,
  sp_ObtenerDatosFacturaexpo:sp_ObtenerDatosFacturaexpo,
  sp_ObtenerDatosFacturaSemanal:sp_ObtenerDatosFacturaSemanal,
  sp_ObtenerDatosFacturaexpoSemanal:sp_ObtenerDatosFacturaexpoSemanal,
  sp_ObtenerDatosFacturaSemanal_Thyssen:sp_ObtenerDatosFacturaSemanal_Thyssen,
  sp_ObtenerDatosFacturaexpoSemanal_Thyssen:sp_ObtenerDatosFacturaexpoSemanal_Thyssen,
};