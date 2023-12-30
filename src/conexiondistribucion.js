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
        database: process.env.databasedist,
        trustServerCertificate: true,
    }
}
async function conexiondistribucion(){
  try{
    let pool = await sql.connect(config);
  console.log("Conected...");
  }
  catch(error){
    console.log("Error conexion "+error);

  }
} 
async function getdata_reportedistribucion(){
    try{
      let pool =await sql.connect(config);
      let res = await pool.request().query("select dst_inventario.no_controlE as 'NoControl', Nom as Customer, CONVERT(varchar, dst_entrada.fecha_llegada,103) as 'ArrivalDate', "+
      "dst_entrada.hora_llegada as 'Time', dst_cat_lineas.Nombre as Carrier, dst_entrada.caja as Trailer, "+
      "dst_inventario.serie as Serial, dst_inventario.numero_referencia as Skid, dst_inventario.no_parte as Part, "+
      "dst_inventario.descripcion as 'Description', "+
      "dst_inventario.cantidad as Quantity, "+
      "Case When PAQ_DESCB.id = dst_inventario.unidad Then PAQ_DESCB.DESCB else '' End as Unit, "+
      "dst_inventario.cantidad_bulto as Qty, "+
      "Case When A.id =dst_inventario.unidad_bulto Then A.DESCB else '' End as Unit2, "+
      "dst_inventario.peso as 'Weight', "+
      "dst_cat_secciones.descripcion as Section, DATEDIFF(D, dst_entrada.fecha_llegada, getdate () )as 'DaysInWarehouse' "+
      "from distribucion.dbo.dst_inventario "+
      "left join distribucion.dbo.Clientes on clientes.cliente_id = dst_inventario.Cliente "+
      "left join distribucion.dbo.dst_entrada on dst_entrada.no_control = dst_inventario.no_controlE "+
      "left join distribucion.dbo.dst_cat_lineas on dst_cat_lineas.id = dst_entrada.linea "+
      "left join distribucion.dbo.PAQ_DESCB on PAQ_DESCB.id = dst_inventario.unidad "+
      "left Join distribucion.dbo.PAQ_DESCB A on A.id =dst_inventario.unidad_bulto "+
      "left join distribucion.dbo.dst_cat_secciones on dst_cat_secciones.ID = dst_inventario.seccion "+
      "where dst_inventario.Cliente = '10012' and dst_entrada.fecha_llegada is not null");
      return res.recordset;
    }
    catch(error){ 
      console.log("Error de get "+error);
    }
  }
module.exports={
    conexiondistribucion:conexiondistribucion,
    getdata_reportedistribucion:getdata_reportedistribucion,
  }