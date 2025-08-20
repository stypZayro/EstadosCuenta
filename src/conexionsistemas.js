const sql=require('mssql');
const dotenv=require('dotenv');
dotenv.config();
const config={ 
    server: process.env.server,
    authentication:{
        type: 'default',
        options:{
            userName: process.env.user, 
            password: process.env.password
        }
    },
    options:{
        port:1433,
        database: process.env.databasesis,
        trustServerCertificate: true,
    }
}
async function conexionsis(){
  try{
    let pool = await sql.connect(config);
  console.log("Conected...");
  }
  catch(error){
    console.log("Error conexion "+error);

  }
} 
async function sp_obtener_datos_iden(Pedimento, Aduana, Clave, Fecha, Partida, CompUno, CompDos, CompTres, Cliente, Referencia) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('Pedimento', sql.NVarChar(50), Pedimento)
      .input('Aduana', sql.NVarChar(50), Aduana)
      .input('Clave', sql.NVarChar(50), Clave)
      .input('Fecha', sql.Date, Fecha)
      .input('Partida', sql.NVarChar(50), Partida)
      .input('CompUno', sql.NVarChar(50), CompUno)
      .input('CompDos', sql.NVarChar(50), CompDos)
      .input('CompTres', sql.NVarChar(50), CompTres)
      .input('Cliente', sql.NVarChar(250), Cliente)
      .input('Referencia', sql.NVarChar(50), Referencia)
      .execute('sp_obtener_datos_iden'); // Ejecuta el SP

    return resultado.recordset;
  } catch (error) {
    console.error('Error al ejecutar el procedimiento almacenado: sp_obtener_datos_iden ', error.message);
    throw error;
  }
}
async function sp_insertar_ident(
  Pedimento, Aduana, Clave, Fecha, Partida, 
  CompUno, CompDos, CompTres, Cliente, Referencia, Correo
) {
  try {
      let pool = await sql.connect(config);

      let resultado = await pool.request()
          .input('Pedimento', sql.NVarChar(50), Pedimento)
          .input('Aduana', sql.NVarChar(50), Aduana)
          .input('Clave', sql.NVarChar(50), Clave)
          .input('Fecha', sql.Date, Fecha)
          .input('Partida', sql.NVarChar(50), Partida)
          .input('CompUno', sql.NVarChar(50), CompUno)
          .input('CompDos', sql.NVarChar(50), CompDos)
          .input('CompTres', sql.NVarChar(50), CompTres)
          .input('Cliente', sql.NVarChar(250), Cliente)
          .input('Referencia', sql.NVarChar(50), Referencia)
          .input('Correo', sql.NVarChar(255), Correo)
          .execute('sp_insertar_ident'); 

      return resultado.rowsAffected; 
  } catch (error) {
      console.error('Error al ejecutar el procedimiento almacenado: sp_insertar_ident', error.message);
      throw error;
  } finally {
      sql.close(); // Cierra la conexión
  }
}
async function sp_obtener_ident_no_enviados() {
  let pool;
  try {
      pool = await sql.connect(config);

      let resultado = await pool.request()
          .execute('sp_obtener_ident_no_enviados'); 

      return resultado.recordset; // Devuelve los datos en lugar de rowsAffected
  } catch (error) {
      console.error('Error al ejecutar el procedimiento almacenado:', error.message);
      throw error;
  } finally {
      if (pool) await pool.close(); // Cierra la conexión solo si se creó
  }
}
async function sp_obtener_ident_por_pedimento(Pedimento) {
  try {
      let pool = await sql.connect(config);

      let resultado = await pool.request()
          .input('Pedimento', sql.NVarChar(50), Pedimento)
          .execute('sp_obtener_ident_por_pedimento'); 

      return resultado.recordset || []; // 🔹 Retorna los datos obtenidos
  } catch (error) {
      console.error('Error al ejecutar sp_obtener_ident_por_pedimento:', error.message);
      throw error;
  }
}

async function sp_actualizar_enviado(Pedimento) {
  try {
      let pool = await sql.connect(config);

      let resultado = await pool.request()
          .input('Pedimento', sql.NVarChar(50), Pedimento)
          .execute('sp_actualizar_enviado'); 

      return resultado.rowsAffected[0] || 0; // 🔹 Retorna el número de filas afectadas
  } catch (error) {
      console.error('Error al ejecutar sp_actualizar_enviado:', error.message);
      throw error;
  }
}

async function sp_obtener_traImpExp(Referencia) {
  try {


    // Conectar a la base de datos
    let pool = await sql.connect(config);

    // Ejecutar el procedimiento almacenado
    let result = await pool.request()
      .input('Referencia', sql.NVarChar(255), Referencia)
      .execute('sp_obtener_traImpExp'); // Nombre del SP

    // Retornar el resultado
    return result.recordset;  // Devuelve los correos obtenidos

  } catch (error) {
    console.error('Error al ejecutar el SP:', error.message);
    throw error;
  } 
}
async function sp_obtener_email(Referencia, servicio) {
  try {


    // Conectar a la base de datos
    let pool = await sql.connect(config);

    // Ejecutar el procedimiento almacenado
    let result = await pool.request()
      .input('Referencia', sql.NVarChar(255), Referencia)
      .input('servicio', sql.NVarChar(50), servicio)
      .execute('sp_obtener_email'); // Nombre del SP

    // Retornar el resultado
    return result.recordset;  // Devuelve los correos obtenidos

  } catch (error) {
    console.error('Error al ejecutar el SP:', error.message);
    throw error;
  } 
}
async function guardafacturakmx (referencia,facturainformacion)  {
    try {
      // Establecer la conexión
      const pool = new sql.ConnectionPool(config);
      await pool.connect();
      //console.log(facturainformacion)
      const request = pool.request();
       request.input('referencia', sql.VarChar(50), referencia);
       // Agregar parámetro de entrada que contenga el JSON
       request.input('json', sql.NVarChar(sql.MAX), JSON.stringify(facturainformacion));
       

       // Ejecutar el stored procedure
      const result = await request.execute('FACTURASKMXADICIONALES');
      const respuesta = result.recordset;

      // Cerrar la conexión
      await sql.close();

      return respuesta;
    } catch (error) {
      console.error('Error al conectar o consultar la base de datos:', error.message);
    } finally {
      // Cerrar la conexión
      await sql.close();
    }
  };
  async function reportemensualsolocostos ()  {
    try {
      // Establecer la conexión
      let respuesta
      let pool =await sql.connect(config);
      let result = await pool.request().query("sp_reporte_mensual_facturas_solo_totaltes");
      // Verificar si el conjunto de resultados es null o tiene longitud cero
      if (!result.recordset || result.recordset.length === 0) {
        //console.log('No se encontraron resultados.');
        respuesta=""
        return respuesta;
      }
      else{
        return result.recordset;
      }
      
  
    } catch (error) {
      console.error('Error al conectar o consultar la base de datos:', error.message);
    }
  };
async function reportemensual (referencia)  {
    try {
      // Establecer la conexión
      let respuesta
      let pool =await sql.connect(config);
      let result = await pool.request().query("sp_reporte_mensual_facturas '"+referencia+"'");
      // Verificar si el conjunto de resultados es null o tiene longitud cero
      if (!result.recordset || result.recordset.length === 0) {
        //console.log('No se encontraron resultados.');
        respuesta=""
        return respuesta;
      }
      else{
        return result.recordset;
      }
      
  
    } catch (error) {
      console.error('Error al conectar o consultar la base de datos:', error.message);
    }
  };
  
module.exports={
    conexionsis:conexionsis,
    sp_obtener_datos_iden:sp_obtener_datos_iden,
    sp_insertar_ident:sp_insertar_ident,
    sp_obtener_ident_no_enviados:sp_obtener_ident_no_enviados,
    sp_obtener_ident_por_pedimento:sp_obtener_ident_por_pedimento,
    sp_actualizar_enviado:sp_actualizar_enviado,
    sp_obtener_traImpExp:sp_obtener_traImpExp,
    sp_obtener_email:sp_obtener_email,
    guardafacturakmx:guardafacturakmx,
    reportemensualsolocostos:reportemensualsolocostos,
    reportemensual:reportemensual,
    
}