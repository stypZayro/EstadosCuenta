const sql=require('mssql');
const dotenv=require('dotenv');
dotenv.config();
const config1={ 
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
        database: process.env.databaseram,
        trustServerCertificate: true,
    },
    requestTimeout: 1300000,
}
async function obtenercampos  ()  {
  try {
    // Establecer la conexión
    let respuesta
    let pool1 =await sql.connect(config1);
    let result = await pool1.request().query("exec SP_REPORTE_SICA_DIARIO_GRAFICA");
    // Verificar si el conjunto de resultados es null o tiene longitud cero
    if (!result.recordset || result.recordset.length === 0) {
      //console.log('No se encontraron resultados.');
      respuesta=""
      pool1.close();
      return respuesta;
    }
    else{
      pool1.close();
      return result;
    }
    

  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
  } 
};
async function sicadiario  ()  {
  try {
    // Establecer la conexión
    let respuesta
    let pool2 =await sql.connect(config1);
    let result = await pool2.request().query("exeC SP_REPORTE_SICA_DIARIO");
    // Verificar si el conjunto de resultados es null o tiene longitud cero
    if (!result.recordset || result.recordset.length === 0) {
      //console.log('No se encontraron resultados.');
      respuesta=""
      pool2.close();
      return respuesta;
    }
    else{
      pool2.close();
      return result.recordset;
    }
    

  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
  } 
};
async function sp_loginaceesostoken(usuario,password) {
  try {
    let pool = await sql.connect(config1);

    let resultado = await pool.request()
      .input('usuario', sql.VarChar(50), usuario)
      .input('password', sql.VarChar(30), password)
      .execute('ramadre.dbo.sp_loginaceesostoken'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_altaToken(usuario) {
  try {
    let pool = await sql.connect(config1);

    let resultado = await pool.request()
      .input('usuario', sql.VarChar(50), usuario)
      .execute('ramadre.dbo.sp_altaToken'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_validartoken(token) {
  try {
    let pool = await sql.connect(config1);

    let resultado = await pool.request()
      .input('token', sql.UniqueIdentifier, token)
      .execute('ramadre.dbo.sp_validartoken'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
  async function altafactura (numfactura)  {
    try {
      // Establecer la conexión
      let respuesta
      let pool =await sql.connect(config1);
      let result = await pool.request().query("altafactura '"+numfactura+"'");
      
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
  };
  async function consultarfactura (numfactura)  {
    try {
      // Establecer la conexión
      let respuesta
      let pool =await sql.connect(config1);
      let result = await pool.request().query("consultarfactura '"+numfactura+"'");
      // Verificar si el conjunto de resultados es null o tiene longitud cero
      if (!result.recordset || result.recordset.length === 0) {
        //console.log('No se encontraron resultados.');
        respuesta=""
        pool.close();
        return respuesta;
      }
      else{
        pool.close();
        return result.recordset;
      }
      
  
    } catch (error) {
      console.error('Error al conectar o consultar la base de datos:', error.message);
    } finally {
      // Cerrar la conexión
      await sql.close();
    }
  };
  async function actualizarfactura (numfactura)  {
    try {
      // Establecer la conexión
      let respuesta
      let pool =await sql.connect(config1);
      let result = await pool.request().query("ramadre.dbo.actualizafactura '"+numfactura+"'");
      // Verificar si el conjunto de resultados es null o tiene longitud cero
      if (!result.recordset || result.recordset.length === 0) {
        //console.log('No se encontraron resultados.');
        respuesta=""
        pool.close();
        return respuesta;
      }
      else{
        pool.close();
        return result.recordset;
      }
      
  
    } catch (error) {
      console.error('Error al conectar o consultar la base de datos:', error.message);
    } finally {
      // Cerrar la conexión
      await sql.close();
    }
  };
module.exports={
  obtenercampos:obtenercampos,
  sicadiario:sicadiario,
  sp_loginaceesostoken:sp_loginaceesostoken,
  sp_altaToken:sp_altaToken,
  sp_validartoken:sp_validartoken,
  altafactura:altafactura,
    consultarfactura:consultarfactura,
    actualizarfactura:actualizarfactura,
}