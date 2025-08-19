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
    },
    requestTimeout: 1300000,
}

async function getdata(){
  try{
    let pool = await sql.connect(config);
  console.log("Conected...");
  }
  catch(error){
    console.log("Error conexion "+error);

  }
}
let date=new Date();
let lunes = date.getDay();
let fechaDia    = date.getUTCDate();
var formatearFecha;
var obtenerFechaInicioDeMes;
var obtenerFechaFinDeMes;
var fechaInicio;
var fechaFin;
var fechaInicioFormateada='';
var fechaFinFormateada='';
if (fechaDia==1){
  console.log("Inicio de mes");
   formatearFecha = fecha => {
    const mes = fecha.getMonth() + 1;
    const dia = fecha.getDate();
    return `${fecha.getFullYear()}-${(mes < 10 ? '0' : '').concat(mes)}-${(dia < 10 ? '0' : '').concat(dia)}`;
  };
  obtenerFechaInicioDeMes=  obtenerFechaInicioDeMes = () => {
    const fechaInicio = new Date();
    // Iniciar en este año, este mes, en el día 1
    return new Date(fechaInicio.getFullYear(), fechaInicio.getMonth()-1, 1);
  };
  obtenerFechaFinDeMes = () => {
    const fechaFin = new Date();
    // Iniciar en este año, el siguiente mes, en el día 0 (así que así nos regresamos un día)
    return new Date(fechaFin.getFullYear(), fechaFin.getMonth() , 0);
  };
  fechaInicio = obtenerFechaInicioDeMes();
  fechaFin = obtenerFechaFinDeMes();
  fechaInicioFormateada = formatearFecha(fechaInicio);
  fechaFinFormateada = formatearFecha(fechaFin);
}
else{
  if (lunes==1){
    console.log("lunes");
    formatearFecha = fecha => {
      const mes = fecha.getMonth() + 1;
      const dia = fecha.getDate();
      return `${fecha.getFullYear()}-${(mes < 10 ? '0' : '').concat(mes)}-${(dia < 10 ? '0' : '').concat(dia)}`;
    };
    obtenerFechaInicioDeMes=  obtenerFechaInicioDeMes = () => {
      const fechaInicio = new Date();
      // Iniciar en este año, este mes, en el día 1
      return new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), 1);
    };
    obtenerFechaFinDeMes = () => {
      const fechaFin = new Date();
      return new Date(fechaFin.getFullYear(), fechaFin.getMonth() , fechaFin.getUTCDate());
    };  
    fechaInicio = obtenerFechaInicioDeMes();
    fechaFin = obtenerFechaFinDeMes();
    fechaInicioFormateada = formatearFecha(fechaInicio);
    fechaFinFormateada = formatearFecha(fechaFin);
  }
}
/************************************************************************** */
if(fechaInicioFormateada==''){
  formatearFecha = fecha => {
    const mes = fecha.getMonth() + 1;
    const dia = fecha.getDate();
    return `${fecha.getFullYear()}-${(mes < 10 ? '0' : '').concat(mes)}-${(dia < 10 ? '0' : '').concat(dia)}`;
  };
  obtenerFechaInicioDeMes=  obtenerFechaInicioDeMes = () => {
    const fechaInicio = new Date();
    // Iniciar en este año, este mes, en el día 1
    return new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), 1);
  };
  fechaInicio = obtenerFechaInicioDeMes();
  fechaInicioFormateada = formatearFecha(fechaInicio);
}
if(fechaFinFormateada==''){
  formatearFecha = fecha => {
    const mes = fecha.getMonth() + 1;
    const dia = fecha.getDate();
    return `${fecha.getFullYear()}-${(mes < 10 ? '0' : '').concat(mes)}-${(dia < 10 ? '0' : '').concat(dia)}`;
  };
  obtenerFechaFinDeMes = () => {
    const fechaFin = new Date();
    // Iniciar en este año, el siguiente mes, en el día 0 (así que así nos regresamos un día)
    return new Date(fechaFin.getFullYear(), fechaFin.getMonth() , fechaFin.getUTCDate());
  }; 
  fechaFin = obtenerFechaFinDeMes();
  fechaFinFormateada = formatearFecha(fechaFin);
}
console.log(`El inicio de mes es ${fechaInicioFormateada} y el fin es ${fechaFinFormateada}`);
async function getdata_BebidasMundiales(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("exec sp_reporte_topo_chico_bebidasmundiales'"+fechaInicioFormateada+"','"+fechaFinFormateada+"','2327'");
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_TopoChico(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("exec sp_reporte_topo_chico_bebidasmundiales'"+fechaInicioFormateada+"','"+fechaFinFormateada+"','1378'");
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
/******************************************/
/******************************************/
//Consultas Semaforo
async function getdata_SemaforoEjecutivos(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_correos_semaforo");
    
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoReporte(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_reportes_semaforo '"+ejecutivoid+"'");
    
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function sp_informacion_cumpleanios(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_informacion_cumpleanios");
    
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
/*
async function getdata_SemaforoNuevoLaredo(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom As 'Nombre', Sincruzar.Bodreferencia As 'Referencia', CONVERT(varchar,Sincruzar.Bodfecha,103) As 'Fecha', "+ 
    "Sincruzar.Trastatus As 'Estado', isnull(Sincruzar.Dias,'0') As 'Dias', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') As 'Eta', tblusua.usunombre As 'Ejecutivo'  "+ 
    "FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.*,Clientes.USUA_EXPO  "+ 
    "FROM (SELECT trafico.traImpExp,trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia, "+ 
    "(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias,  "+ 
    "Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod "+  
    "LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia  "+ 
    "WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='')  "+ 
    "AND Trafico.PREF = 'T3485240' AND (Tblbod.Bodvirtual=0) AND (TblBod.PORLLEGAR = 0)  "+ 
    "AND (Trafico.trastatus <> 'CONCLUIDO' AND Trafico.trastatus <> 'CRUZO')) AS BodegaTrafico  "+ 
    "LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua  "+ 
    "ON (Sincruzar.Usuario_id1 = Tblusua.Usuario_id and Sincruzar.traImpExp=1)  "+ 
    "or (Sincruzar.USUA_EXPO = Tblusua.Usuario_id and Sincruzar.traImpExp=0) Where tblusua.Usuario_id='"+ejecutivoid+ "' ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoVeracruz(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom as 'Nombre', Sincruzar.Bodreferencia As 'Referencia', CONVERT(varchar,Sincruzar.Bodfecha,103) As 'Fecha', Sincruzar.Trastatus as 'Estado', isnull(Sincruzar.Dias,'0') as 'Dias', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') as 'Eta', tblusua.usunombre as 'Ejecutivo' FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.* FROM (SELECT trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia,(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias, Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='') AND (Trafico.Trareferencia LIKE 'ZV%' OR Trafico.Trareferencia LIKE 'EZV%') AND (Tblbod.Bodvirtual=0) AND (TblBod.PORLLEGAR = 0) AND (Trafico.trastatus <> 'DEPENDE DEL CLIENTE' AND Trafico.trastatus <> 'CONCLUIDO' AND Trafico.trastatus <> 'CRUZO')) AS BodegaTrafico LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua ON Sincruzar.Usuario_id1 = Tblusua.Usuario_id Where tblusua.Usuario_id='"+ejecutivoid+ "'ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoCorresponsalias(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom as 'Nombre', Sincruzar.Bodreferencia as 'Referencia', CONVERT(varchar,Sincruzar.Bodfecha,103) as 'Fecha', Sincruzar.Trastatus as 'Estado', isnull(Sincruzar.Dias,'0') as 'Dias', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') as 'Eta', tblusua.usunombre as 'Ejecutivo' FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.* FROM (SELECT trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia,(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias, Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='') AND Trafico.PREF <> 'T3485240' AND Trafico.Trareferencia NOT LIKE 'ZV%' AND Trafico.Trareferencia NOT LIKE 'EZV%' AND Trafico.Trareferencia NOT LIKE 'ZA%' AND Trafico.Trareferencia NOT LIKE 'EZA%' AND (Tblbod.Bodvirtual=0) AND (TblBod.PORLLEGAR = 0) AND (Trafico.trastatus <> 'DEPENDE DEL CLIENTE' AND Trafico.trastatus <> 'CONCLUIDO' AND Trafico.trastatus <> 'CRUZO')) AS BodegaTrafico LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua ON Sincruzar.Usuario_id1 = Tblusua.Usuario_id Where tblusua.Usuario_id='"+ejecutivoid+ "'ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoAICM(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom as 'Nombre', Sincruzar.Bodreferencia as 'Referencia', CONVERT(varchar,Sincruzar.Bodfecha,103) as 'Fecha', Sincruzar.Trastatus as 'Estado', isnull(Sincruzar.Dias,'0') as 'Dias', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') as 'Eta', tblusua.usunombre as 'Ejecutivo' FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.* FROM (SELECT trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia,(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias, Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='') AND Trafico.PREF <> 'T3485240' AND (Trafico.Trareferencia LIKE 'ZA%' OR Trafico.Trareferencia LIKE 'EZA%') AND (Tblbod.Bodvirtual=0) AND (TblBod.PORLLEGAR = 0) AND (Trafico.trastatus <> 'DEPENDE DEL CLIENTE' AND Trafico.trastatus <> 'CONCLUIDO' AND Trafico.trastatus <> 'CRUZO')) AS BodegaTrafico LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua ON Sincruzar.Usuario_id1 = Tblusua.Usuario_id Where tblusua.Usuario_id='"+ejecutivoid+ "'ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoVirtuales(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom as 'Nombre', Sincruzar.Bodreferencia as 'Referencia', CONVERT(varchar,Sincruzar.Bodfecha,103) as 'Fecha', Sincruzar.Trastatus as 'Estado', isnull(Sincruzar.Dias,'0') as 'Dias', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') as 'Eta', tblusua.usunombre as 'Ejecutivo' FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.* FROM (SELECT trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia,(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias, Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='') AND (Tblbod.Bodvirtual=1)  AND (TblBod.PORLLEGAR = 0) AND (Trafico.trastatus <> 'DEPENDE DEL CLIENTE' AND Trafico.trastatus <> 'CONCLUIDO' AND Trafico.trastatus <> 'CRUZO')) AS BodegaTrafico LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua ON Sincruzar.Usuario_id1 = Tblusua.Usuario_id Where tblusua.Usuario_id='"+ejecutivoid+ "'ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoFerrocarril(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom as 'Nombre', Sincruzar.Bodreferencia as 'Referencia', CONVERT(varchar,Sincruzar.Bodfecha,103) as 'Fecha', Sincruzar.Trastatus as 'Estado', isnull(Sincruzar.Dias,'0') as 'Dias', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') as 'Eta', tblusua.usunombre as 'Ejecutivo' FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.* FROM (SELECT trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia,(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias, Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='') AND (Tblbod.bodtipemb='FERROCARRIL')   AND (TblBod.PORLLEGAR = 0) AND (Trafico.trastatus <> 'DEPENDE DEL CLIENTE' AND Trafico.trastatus <> 'CONCLUIDO' AND Trafico.trastatus <> 'CRUZO') ) AS BodegaTrafico LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua ON Sincruzar.Usuario_id1 = Tblusua.Usuario_id Where tblusua.Usuario_id='"+ejecutivoid+ "'ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_SemaforoPorLlegar(ejecutivoid){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("SELECT Sincruzar.Nom as 'Nombre', Sincruzar.Bodreferencia as 'Referencia', isnull(CONVERT(varchar,Sincruzar.TRAFETA,103),'') as 'Eta' FROM (SELECT Clientes.nom, Clientes.Cliente_ID, Clientes.Usuario_id1, BodegaTrafico.* FROM (SELECT trafico.trastatus, tblbod.bodcli, TblBod.bodfecha, Tblbod.Bodhora, Tblbod.bodreferencia,(DATEDIFF ( day , Bodfecha , GETDATE() )) as Dias, Tblbod.bodusuario, Tblbod.PORLLEGAR, Trafico.TRAFETA FROM Tblbod LEFT JOIN Trafico ON Tblbod.Bodreferencia = Trafico.Trareferencia WHERE Bodfecha <= GETDATE() AND (Trafico.Trafechacruce IS NULL OR Trafico.Trafechacruce='') AND tblBod.PORLLEGAR = 1) AS BodegaTrafico LEFT JOIN Clientes ON BodegaTrafico.bodcli = Cliente_id) AS Sincruzar LEFT JOIN Tblusua ON Sincruzar.Bodusuario = Tblusua.Usuario_id Where tblusua.Usuario_id='"+ejecutivoid+ "'ORDER BY Dias ASC");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}*/
/******************************************/
/******************************************/
async function getdata_mercanciasenbodega(cliente_id){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("select bodreferencia as 'Referencia', "+
    "isnull(CONVERT(varchar,bodfecha,103),' ') as 'FechaArribo', "+
    "isnull(bodhora,' ') as 'HoraArribo', "+
    "isnull(nom,' ') as 'Cliente', "+
    "isnull(proNom,' ') as 'Proveedor', "+
    "isnull(BODEMB,' ') as 'Embarcador', "+
    "' ' AS 'Factura', "+
    "isnull(TBLFLET.FLENOMBRE,' ') AS 'LineadeArribo', "+
    "isnull(BODCATEGORIA,' ') AS 'Categoria', "+
    "isnull(bodnopedido,' ') as 'Pedido', "+
    "isnull(bodbno,' ') as 'Guia', "+
    "isnull(bodpesolbs,'0.00') as 'PesoLbs', "+
    "isnull(format(bodpesolbs/2.2046,'.00'),'') AS 'PesoKgs', "+
    "isnull(bodbultos,' ') as 'Bultos', "+
    "isnull(bodcaja,' ') as 'Caja', "+
    "isnull(traStatus,' ') as 'Estatus', "+
    "isnull(DATEDIFF(DAY,bodfecha,SYSDATETIME()),'0') as 'Diasenbodega', "+
    "isnull(boddescmer,' ') as 'Descripcion', "+
    "isnull(TRAOBSERVACIONES,' ') as 'Observaciones', "+
    "' ' as 'Loadentrada', "+
    "' ' as 'Loadsalida', "+
    "isnull(BODCOMEN1,' ') as 'ObsTrafico'  "+
    "from aduana.dbo.tblBod "+ 
    "left join aduana.dbo.trafico  "+
    "on trareferencia = bodreferencia and tracli = bodcli "+ 
    "LEFT JOIN aduana.dbo.TBLFLET "+ 
    "ON TBLFLET.FLECLAVE = tblBod.BODFLE "+ 
    "left join aduana.dbo.procli "+ 
    "on proveedor_id = bodprocli "+ 
    "left join aduana.dbo.clientes "+ 
    "on bodcli = CLIENTE_ID "+
    "where bodcli = '"+cliente_id+"' "+  
    "and (traregimen <> 'F4' or traregimen IS NULL) "+ 
    "and traImpExp = '1' and BODVIRTUAL <> 1 "+ 
    "and tblBod.PREF = 'T3485240' "+ 
    "and bodfecha >= '2006-01-01' "+ 
    "and Trafico.traFechaCruce IS NULL "+   
    "order by bodfecha desc");
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_correos_ejecutivos_cliente(cliente_id){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_correos_ejecutivos_cliente"+" '"+cliente_id+"'")

    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_listaclientes(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("Select distinct cliente_id,clientes.NUMERO as numero,Nom as nomcli From tblBod left join trafico on trareferencia = bodreferencia and tracli = bodcli LEFT JOIN TBLFLET ON TBLFLET.FLECLAVE = TBLBOD.BODFLE left join procli on proveedor_id = bodprocli left join clientes on bodcli = CLIENTE_ID where (traregimen <> 'F4' or traregimen IS NULL) and traImpExp = '1' and BODVIRTUAL <> 1 and tblBod.PREF = 'T3485240' and bodfecha >= '2006-01-01' and Trafico.traFechaCruce IS NULL and cliente_id <> '2316' order by Nom asc");
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
/******************************************/ 
/******************************************/
async function getdata_ReporteASN(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_reporte_ASN");
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
/******************************************/
/******************************************/
async function getdata_Thyssenkrupp_pendientes(fechaini, fechafin){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_ReporteThyssenhrup_pendientes '" +fechaini+"','"+fechafin+"'");
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
/******************************************/
/******************************************/
async function getdata_correos_reporte(tiporeporte){
  try{
    let pool =await sql.connect(config);
    //let res = await pool.request().query("SELECT 1 AS Test")
    let res = await pool.request().query("execute sp_correos_reporte"+" '"+tiporeporte+"'")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
/******************************************/
/******************************************/
async function getdata_reporte_kawassaki(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_reportekawassaki")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}


async function getdata_hb101(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_hb101")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_hb102(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_hb102")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_hb101(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_hb101")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_hb102(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_hb102")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function getdata_hb103(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_hb103")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function revisarasnexisten (referencia)  {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_revisar_ASN_existen "+"'"+referencia+"'");
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
  } 
};
async function ejecutar_sp_Asn  (referencia)  {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_ejecutar_sp_Asn "+"'"+referencia+"'");
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
  } 
};

async function facturasaenviar ()  {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_selectBitacoraFacturasCorreo");
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
  } 
};
async function actualizarestadofactura (referencia)  {
  try {
    // Establecer la conexión
    let respuesta
    let pool =await sql.connect(config);
    let result = await pool.request().query("sp_actualizar_estado_factura "+"'"+referencia+"'");
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
  } 
};
async function FRACCIONTBLPARTESVS101ESTANENBODEGAIMPORTACION(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute FRACCIONTBLPARTESVS101ESTANENBODEGAIMPORTACION")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function FRACCIONTBLPARTESVS101ESTANENBODEGAEXPORTACION(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute FRACCIONTBLPARTESVS101ESTANENBODEGAEXPORTACION")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function sp_AVISO_AUTOMATICO_HB201_SIN_EDI(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute sp_AVISO_AUTOMATICO_HB201_SIN_EDI")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function Sp_kfantasma(){
  try{
    let pool =await sql.connect(config);
    let res = await pool.request().query("execute Sp_kfantasma")
    pool.close();
    return res.recordset;
  }
  catch(error){ 
    console.log("Error de get "+error);
  }
}
async function sp_ObtenerInformacionPedimento(ClienteId,ImpExp,Pedimento,renglon) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('ClienteId', sql.Int, ClienteId)
      .input('ImpExp', sql.Int, ImpExp)
      .input('Pedimento', sql.VarChar(8), Pedimento)
      .input('Partida', sql.Int, renglon)
      .execute('aduana.dbo.sp_ObtenerInformacionPedimento'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_ObtenerPedimentos(ClienteId) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('ClienteId', sql.Int, ClienteId)
      .execute('aduana.dbo.sp_ObtenerPedimentos'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_ObtenerPedimentos_Semanal(ClienteId) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('ClienteId', sql.Int, ClienteId)
      .execute('aduana.dbo.sp_ObtenerPedimentos_Semanal'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_totaldeclientesrevisarfraccionesIMMEX() {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .execute('aduana.dbo.sp_totaldeclientesrevisarfraccionesIMMEX'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_obtenerfraccionesIMMEXnoautorizadas(Clientenumero) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('Clientenumero', sql.Int, Clientenumero)
      .execute('aduana.dbo.sp_obtenerfraccionesIMMEXnoautorizadas'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_obtenerejecutivogerentesubcliente(numero) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('numero', sql.VarChar(5), numero)
      .execute('aduana.dbo.sp_obtenerejecutivogerentesubcliente'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_obtenerclientesreporteanexo24() {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .execute('aduana.dbo.sp_obtenerclientesreporteanexo24'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
async function sp_correos_ejecutivos_cliente_anexo24(cliente) {
  try {
    let pool = await sql.connect(config);

    let resultado = await pool.request()
      .input('cliente', sql.VarChar(5),cliente)
      .execute('aduana.dbo.sp_correos_ejecutivos_cliente_anexo24'); 

    return resultado.recordset;
  } catch (error) {
    console.error('Error al conectar o consultar la base de datos:', error.message);
    throw error; 
  } finally {
    sql.close(); // Cierra la conexión
  }
}
 async function altafacturasencontradas (referencia)  {
    try {
      // Establecer la conexión
      let respuesta
      let pool =await sql.connect(config);
      //console.log(referencia)
      let result = await pool.request().query("aduana.dbo.sp_agregarfacturasarevisarasn "+"'"+referencia+"'");
      //console.log(result)
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
  async function sp_existeboxid(resultados) {
    try {
        // Establecer la conexión
        let respuesta;
        let pool = await sql.connect(config);
        let request = pool.request();
        request.input('partesjson', sql.NVarChar(sql.MAX), JSON.stringify( resultados));
  
        let result = await request.execute('aduana.dbo.sp_existeboxid');
  
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
  async function sp_noexisteboxid(resultados) {
    try {
        // Establecer la conexión
        let respuesta;
        let pool = await sql.connect(config);
        let request = pool.request();
        request.input('partesjson', sql.NVarChar(sql.MAX), JSON.stringify( resultados));
  
        let result = await request.execute('aduana.dbo.sp_noexisteboxid');
  
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
  async function sp_existeboxid103(resultados) {
    try {
        // Establecer la conexión
        let respuesta;
        let pool = await sql.connect(config);
        let request = pool.request();
        request.input('partesjson', sql.NVarChar(sql.MAX), JSON.stringify( resultados));
  
        let result = await request.execute('aduana.dbo.sp_existeboxid_103');
  
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


/******************************************/
/******************************************/
module.exports={
  getdata:getdata,
  getdata_BebidasMundiales:getdata_BebidasMundiales,
  getdata_TopoChico:getdata_TopoChico,
  /*getdata_SemaforoNuevoLaredo:getdata_SemaforoNuevoLaredo,
  getdata_SemaforoVeracruz:getdata_SemaforoVeracruz,
  getdata_SemaforoCorresponsalias:getdata_SemaforoCorresponsalias,
  getdata_SemaforoAICM:getdata_SemaforoAICM,
  getdata_SemaforoVirtuales:getdata_SemaforoVirtuales,
  getdata_SemaforoFerrocarril:getdata_SemaforoFerrocarril,
  getdata_SemaforoPorLlegar:getdata_SemaforoPorLlegar,*/
  getdata_mercanciasenbodega:getdata_mercanciasenbodega,
  getdata_correos_ejecutivos_cliente:getdata_correos_ejecutivos_cliente,
  getdata_listaclientes:getdata_listaclientes,
  getdata_SemaforoEjecutivos:getdata_SemaforoEjecutivos,
  getdata_SemaforoReporte:getdata_SemaforoReporte,
  getdata_ReporteASN:getdata_ReporteASN,
  getdata_Thyssenkrupp_pendientes,getdata_Thyssenkrupp_pendientes,
  getdata_correos_reporte:getdata_correos_reporte,
  getdata_reporte_kawassaki:getdata_reporte_kawassaki,
  getdata_hb101:getdata_hb101,
  getdata_hb102:getdata_hb102,
  getdata_hb103:getdata_hb103,
  revisarasnexisten:revisarasnexisten,
  ejecutar_sp_Asn:ejecutar_sp_Asn,
  facturasaenviar:facturasaenviar,
  actualizarestadofactura:actualizarestadofactura,
  FRACCIONTBLPARTESVS101ESTANENBODEGAIMPORTACION:FRACCIONTBLPARTESVS101ESTANENBODEGAIMPORTACION,
  FRACCIONTBLPARTESVS101ESTANENBODEGAEXPORTACION:FRACCIONTBLPARTESVS101ESTANENBODEGAEXPORTACION,
  sp_AVISO_AUTOMATICO_HB201_SIN_EDI:sp_AVISO_AUTOMATICO_HB201_SIN_EDI,
  Sp_kfantasma:Sp_kfantasma,
  sp_ObtenerInformacionPedimento:sp_ObtenerInformacionPedimento,
  sp_ObtenerPedimentos:sp_ObtenerPedimentos,
  sp_informacion_cumpleanios:sp_informacion_cumpleanios,
  sp_totaldeclientesrevisarfraccionesIMMEX:sp_totaldeclientesrevisarfraccionesIMMEX,
  sp_obtenerfraccionesIMMEXnoautorizadas:sp_obtenerfraccionesIMMEXnoautorizadas,
  sp_obtenerejecutivogerentesubcliente:sp_obtenerejecutivogerentesubcliente,
  sp_ObtenerPedimentos_Semanal:sp_ObtenerPedimentos_Semanal,
  sp_obtenerclientesreporteanexo24:sp_obtenerclientesreporteanexo24,
  sp_correos_ejecutivos_cliente_anexo24:sp_correos_ejecutivos_cliente_anexo24,
  altafacturasencontradas:altafacturasencontradas,
  sp_existeboxid:sp_existeboxid,
  sp_existeboxid103:sp_existeboxid103,
  sp_noexisteboxid:sp_noexisteboxid,
}