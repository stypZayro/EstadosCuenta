const sql=require('mssql');
const dotenv=require('dotenv');
dotenv.config();
const config={ 
    server: process.env.serveraduana,
    authentication:{
        type: 'default',
        options:{
            userName: process.env.user, 
            password: process.env.passwordaduana
        }
    },
    options:{
        port:23390,
        database: process.env.database,
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
  async function facturakmx(referencia, user, resultados) {
    try {
        // Establecer la conexión
        let respuesta;
        let pool = await sql.connect(config);
        let request = pool.request();
        request.input('Referencia', sql.NVarChar(255), referencia);
        request.input('user', sql.NVarChar(255), user);
        request.input('json', sql.NVarChar(sql.MAX), JSON.stringify( resultados));

        let result = await request.execute('aduana.dbo.sp_FACTURAS_KMX');

        // Verificar si el conjunto de resultados es null o tiene longitud cero
        if (!result.recordset || result.recordset.length === 0) {
            //console.log('No se encontraron resultados.');
            respuesta = "";
            return respuesta;
        } else {
            return result.recordset;
        }
    } catch (error) {
        console.error('Error al conectar o consultar la base de datos:', error.message);
    } finally {
        // Cerrar la conexión
        await sql.close();
    }
};
async function facturakmx_inventario(esmensual) {
    try {
        // Establecer la conexión
        let respuesta;
        let pool = await sql.connect(config);
        let request = pool.request();
        request.input('esmensual', sql.NVarChar(255), esmensual);

        let result = await request.execute('aduana.dbo.sp_FACTURAS_KMX_inventario');

        // Verificar si el conjunto de resultados es null o tiene longitud cero
        if (!result.recordset || result.recordset.length === 0) {
            //console.log('No se encontraron resultados.');
            respuesta = "";
            return respuesta;
        } else {
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
    facturakmx:facturakmx,
    facturakmx_inventario:facturakmx_inventario,

   
  }