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
        database: process.env.databasezam,
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
async function getdata_ReporteThyssenhrup_pesos(fechafin){
    try{
      let pool =await sql.connect(config);
      let res = await pool.request().query("execute dbo.sp_ReporteThyssenhrup_pesos '"+fechafin+"'");
      pool.close();
      return res.recordset;
    }
    catch(error){ 
      console.log("Error de get "+error);
    }
}

async function getdata_edoCtaThyssenkrupp_Pesos(numCliente) {
  try {
    let pool = await sql.connect(config);
    let res = await pool.request().query(`execute dbo.spEdoCtaThyssenkrup_Pesos '${numCliente}'`);
    pool.close();
    return res.recordset
  }
  catch(error) {
    console.log("Error de get " + error);
  }
}

async function getdata_EdoCuentaTkPesos(clienteID) {
  try {
    let pool = await sql.connect(config);
    let res = await pool.request().query(`execute dbo.spEdoCuentaTkPesos '${clienteID}'`);
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
/*************************************************** */
async function sp_Rmensual_2(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("Rmensual_2 "+cliente);
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
async function Rmensual_2_distinct(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("Rmensual_2_distinct "+cliente);
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
async function ultimodepositocliente2(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("ultimodepositocliente2 "+cliente);
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

async function antiguedadsaldos2(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("antiguedadsaldos2 "+cliente);
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
async function datosinicialescliente2(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("datosinicialescliente2 "+cliente);
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
async function sp_clientesestadocuenta2() {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_lista_clientes2");
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
async function sp_cargossinaplicar2(Cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_cargossinaplicar2 "+ Cliente);
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
async function sp_cargossinaplicar_distinct2(cliente) {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_cargossinaplicar_distinct2 "+cliente);
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
async function sp_actualizarlimitesdecreditocliente(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_actualizarlimitesdecreditocliente")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
module.exports={
  conexionzay:conexionzay,
  getdata_ReporteThyssenhrup_pesos:getdata_ReporteThyssenhrup_pesos,
  getdata_edoCtaThyssenkrupp_Pesos:getdata_edoCtaThyssenkrupp_Pesos,
  getdata_EdoCuentaTkPesos:getdata_EdoCuentaTkPesos,
  sp_clientesestadocuenta:sp_clientesestadocuenta,
  sp_Rmensual_1:sp_Rmensual_1,
  ultimodepositocliente:ultimodepositocliente,
  antiguedadsaldos:antiguedadsaldos,
  datosinicialescliente:datosinicialescliente,
  sp_Rmensual_2:sp_Rmensual_2,
  ultimodepositocliente2:ultimodepositocliente2,
  antiguedadsaldos2:antiguedadsaldos2,
  datosinicialescliente2:datosinicialescliente2,
  Rmensual_1_distinct:Rmensual_1_distinct,
  sp_cargossinaplicar:sp_cargossinaplicar,
  sp_clientesestadocuenta2:sp_clientesestadocuenta2,
  sp_cargossinaplicar2:sp_cargossinaplicar2,
  Rmensual_2_distinct:Rmensual_2_distinct,
  sp_cargossinaplicar_distinct:sp_cargossinaplicar_distinct,
  sp_cargossinaplicar_distinct2:sp_cargossinaplicar_distinct2,
  contactosestadoscuenta:contactosestadoscuenta,
  sp_actualizarlimitesdecreditocliente:sp_actualizarlimitesdecreditocliente,
}