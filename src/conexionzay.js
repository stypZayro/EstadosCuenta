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
      return res.recordset;
    }
    catch(error){ 
      console.log("Error de get "+error);
    }
}
module.exports={
  conexionzay:conexionzay,
  getdata_ReporteThyssenhrup_dolares:getdata_ReporteThyssenhrup_dolares,
}