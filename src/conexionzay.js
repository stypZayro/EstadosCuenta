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
        database: process.env.databasezay,
        trustServerCertificate: true,
    },
    requestTimeout: 1300000,
}
async function conexionzay(){
  try{
    let pool = await sql.connect(config);
  console.log("Conected...");
  }
  catch(error){
    console.log("Error conexion "+error);

  }
} 
async function getdata_ReporteThyssenhrup_dolares(fechafin){
    try{
      let pool =await sql.connect(config);
      let res = await pool.request().query("execute zayro.dbo.sp_ReporteThyssenhrup_dolares '"+fechafin+"'");
      pool.close();
      return res.recordset;
    }
    catch(error){ 
      console.log("Error de get "+error);
    }
}

async function getdata_EdoCtaThyssenkrupp_Dolares(numCliente) {
  try {
    let pool = await sql.connect(config);
    let res = await pool.request().query(`execute zayro.dbo.spEdoCtaThyssenkrup_Dolares '${numCliente}'`);
    pool.close();
    return res.recordset
  }
  catch(error) {
    console.log("Error de get " + error);
  }
}

async function getdata_EdoCuentaTkDolares(clienteID) {
  try {
    let pool = await sql.connect(config);
    let res = await pool.request().query(`execute zayro.dbo.spEdoCuentaTkDolares '${clienteID}'`);
    pool.close();
    return res.recordset;
  }
  catch(error) {
    console.log("Error de get " + error)
  }
}
async function sp_clientesestadocuenta() {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_lista_clientes");
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function sp_Rmensual_1(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("Rmensual_1 "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}

async function Rmensual_1_distinct(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("Rmensual_1_distinct "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function ultimodepositocliente(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("ultimodepositocliente "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function sp_cargossinaplicar(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_cargossinaplicar "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function antiguedadsaldos(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("antiguedadsaldos "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function datosinicialescliente(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("datosinicialescliente "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function sp_cargossinaplicar_distinct(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_cargossinaplicar_distinct "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
async function contactosestadoscuenta(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("contactosestadoscuenta "+cliente);
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
  } finally {
    // Cerrar la conexión
    await sql.close();
  }
}
module.exports={
  conexionzay:conexionzay,
  getdata_ReporteThyssenhrup_dolares:getdata_ReporteThyssenhrup_dolares,
  getdata_EdoCtaThyssenkrupp_Dolares:getdata_EdoCtaThyssenkrupp_Dolares,
  getdata_EdoCuentaTkDolares:getdata_EdoCuentaTkDolares,
  sp_clientesestadocuenta:sp_clientesestadocuenta,
  sp_Rmensual_1:sp_Rmensual_1,
  ultimodepositocliente:ultimodepositocliente,
  antiguedadsaldos:antiguedadsaldos,
  datosinicialescliente:datosinicialescliente,
  Rmensual_1_distinct:Rmensual_1_distinct,
  sp_cargossinaplicar:sp_cargossinaplicar,
  datosinicialescliente:datosinicialescliente,
  sp_cargossinaplicar_distinct:sp_cargossinaplicar_distinct,
  contactosestadoscuenta:contactosestadoscuenta,
}