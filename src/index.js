
const express=require('express');
const morgan =require("morgan");
const xl=require('excel4node');
const path = require('path');
const sql=require('./conexionaduana');
const sqldist=require('./conexiondistribucion');
const sqlzay=require('./conexionzay');
const sqlzam=require('./conexionzamudio');
const sqlsis=require('./conexionsis');
const fs=require('fs');
const nodemailer=require('nodemailer');
const dotenv=require('dotenv');
dotenv.config();
const app=express();
app.set("port",4000);
app.listen(app.get("port"));
console.log("con señal "+app.get("port"));
app.use(express.json());
app.use(express.urlencoded({extended:true}));
app.use(morgan("dev"));
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Proceso para reportes de Bebidas Mundiales y Topo Chico 
//Solo se ejecutan el dia primero de mes o los dias lunes
app.get('/getdata_BebidasMundiales',function(req,res,next){
   

   sql.getdata_BebidasMundiales().then((result)=>{
      //res.json(result);
      //console.log(result) 
      var wb= new xl.Workbook();
      let nombreArchivo="Reporte Bebidas Mundiales";
      var ws=wb.addWorksheet("BebidasMundi");
      //Estilo Columnas
      var estiloTitulo=wb.createStyle({
         font:{
         name: 'Arial',
         color: '#FFFFFF',
         size:10,
         bold: true,
         } ,
         fill:{
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid',
            fgColor: '#008000',
         },
      });
      var estilocontenido=wb.createStyle({
         font:{
            name: 'Arial',
            color: '#000000',
            size:10,
         }
      });
         //Nombre de las columnas
      ws.cell(1,1).string("REFERENCIA").style(estiloTitulo);
      ws.cell(1,2).string("FECHA APERTURA PEDIMENTO").style(estiloTitulo);
      ws.cell(1,3).string("CLIENTE").style(estiloTitulo);
      ws.cell(1,4).string("PROVEEDOR").style(estiloTitulo);
      ws.cell(1,5).string("FACTURA").style(estiloTitulo);
      ws.cell(1,6).string("CFDI").style(estiloTitulo);
      ws.cell(1,7).string("FECHA FACTURA").style(estiloTitulo);
      ws.cell(1,8).string("TIPO OPERACION").style(estiloTitulo);
      ws.cell(1,9).string("FECHA CRUCE").style(estiloTitulo);
      ws.cell(1,10).string("C001CAAT").style(estiloTitulo);
      ws.cell(1,11).string("CAJA").style(estiloTitulo);
      ws.cell(1,12).string("PLACAS").style(estiloTitulo);
      ws.cell(1,13).string("PEDIMENTO").style(estiloTitulo);
      let numfila=2;
      result.forEach(reglonactual => {
         ws.cell(numfila,1).string(reglonactual.REFERENCIA).style(estilocontenido);
         ws.cell(numfila,2).string(reglonactual.FECHAAPERTURAPEDIMENTO).style(estilocontenido);
         ws.cell(numfila,3).string(reglonactual.CLIENTE).style(estilocontenido);
         ws.cell(numfila,4).string(reglonactual.PROVEEDOR).style(estilocontenido);
         ws.cell(numfila,5).string(reglonactual.FACTURA).style(estilocontenido);
         ws.cell(numfila,6).string(reglonactual.CFDI).style(estilocontenido);
         ws.cell(numfila,7).string(reglonactual.FECHAFACTURA).style(estilocontenido);
         ws.cell(numfila,8).string(reglonactual.TIPOOPERACION).style(estilocontenido);
         ws.cell(numfila,9).string(reglonactual.FECHACRUCE).style(estilocontenido);
         ws.cell(numfila,10).string("").style(estilocontenido);
         ws.cell(numfila,11).string(reglonactual.CAJA).style(estilocontenido);
         ws.cell(numfila,12).string(reglonactual.PLACAS).style(estilocontenido);
         ws.cell(numfila,13).string(reglonactual.PEDIMENTO).style(estilocontenido);
         numfila=numfila+1;
      });
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
         function downloadFile(){res.download(pathExcel);}
         downloadFile();

         /*fs.rm(pathExcel,function(err){
            if(err)console.log(err);*/
            /*else*/ console.log("Archivo descargado exitoso");
            
         /*});*/
        
      }
   });
   })
  
});
app.get('/getdata_TopoChico',function(req,res,next){
   

   sql.getdata_TopoChico().then((result)=>{
      //res.json(result);
      //console.log(result) 
      var wb= new xl.Workbook();
      let nombreArchivo="Reporte Topo Chico";
      var ws=wb.addWorksheet("Topo Chico");
      //Estilo Columnas
      var estiloTitulo=wb.createStyle({
         font:{
         name: 'Arial',
         color: '#FFFFFF',
         size:10,
         bold: true,
         } ,
         fill:{
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid',
            fgColor: '#008000',
         },
      });
      var estilocontenido=wb.createStyle({
         font:{
            name: 'Arial',
            color: '#000000',
            size:10,
         }
      });
         //Nombre de las columnas
      ws.cell(1,1).string("REFERENCIA").style(estiloTitulo);
      ws.cell(1,2).string("FECHA APERTURA PEDIMENTO").style(estiloTitulo);
      ws.cell(1,3).string("CLIENTE").style(estiloTitulo);
      ws.cell(1,4).string("PROVEEDOR").style(estiloTitulo);
      ws.cell(1,5).string("FACTURA").style(estiloTitulo);
      ws.cell(1,6).string("CFDI").style(estiloTitulo);
      ws.cell(1,7).string("FECHA FACTURA").style(estiloTitulo);
      ws.cell(1,8).string("TIPO OPERACION").style(estiloTitulo);
      ws.cell(1,9).string("FECHA CRUCE").style(estiloTitulo);
      ws.cell(1,10).string("C001CAAT").style(estiloTitulo);
      ws.cell(1,11).string("CAJA").style(estiloTitulo);
      ws.cell(1,12).string("PLACAS").style(estiloTitulo);
      ws.cell(1,13).string("PEDIMENTO").style(estiloTitulo);
      let numfila=2;
      result.forEach(reglonactual => {
         ws.cell(numfila,1).string(reglonactual.REFERENCIA).style(estilocontenido);
         ws.cell(numfila,2).string(reglonactual.FECHAAPERTURAPEDIMENTO).style(estilocontenido);
         ws.cell(numfila,3).string(reglonactual.CLIENTE).style(estilocontenido);
         ws.cell(numfila,4).string(reglonactual.PROVEEDOR).style(estilocontenido);
         ws.cell(numfila,5).string(reglonactual.FACTURA).style(estilocontenido);
         ws.cell(numfila,6).string(reglonactual.CFDI).style(estilocontenido);
         ws.cell(numfila,7).string(reglonactual.FECHAFACTURA).style(estilocontenido);
         ws.cell(numfila,8).string(reglonactual.TIPOOPERACION).style(estilocontenido);
         ws.cell(numfila,9).string(reglonactual.FECHACRUCE).style(estilocontenido);
         ws.cell(numfila,10).string("").style(estilocontenido);
         ws.cell(numfila,11).string(reglonactual.CAJA).style(estilocontenido);
         ws.cell(numfila,12).string(reglonactual.PLACAS).style(estilocontenido);
         ws.cell(numfila,13).string(reglonactual.PEDIMENTO).style(estilocontenido);
         numfila=numfila+1;
      });
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
         function downloadFile(){res.download(pathExcel);}
         downloadFile();

         /*fs.rm(pathExcel,function(err){
            if(err)console.log(err);*/
            /*else*/ console.log("Archivo descargado exitoso");
            
         /*});*/
        
      }
   });
   })
});
app.get('/getdata_enviarcorreoBebMunTopChic',function(req,res,next){
   //Se hizo de esta manera porque primero se tienen que generarlos dos reportes, ya que ambos se mandan en un solo correo pero cada metodo es un reporte
   sql.getdata_correos_reporte('1').then((result)=>{
      result.forEach(renglonactual=>{
         enviarMailBebMunTopChic(renglonactual.correos);
      })
   })
   //enviarMailBebMunTopChic();
});
enviarMailBebMunTopChic=async(correos)=>{
   const config ={
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth:{
         user:process.env.useremail, 
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false,
       }, 
   } 

   const mensaje ={
      
      from:'it.sistemas@zayro.com',
      //to:'aby.zamora@arcacontal.com,valentin.garza@arcacontal.com,avazquez@zayro.com,exportacion203@zayro.com,gerenciati@zayro.com,sistemas@zayro.com',
      to: correos, 
      //to: 'oswal15do@gmail.com',
      subject:'Envio de reportes',
      attachments:[
         {filename:'Reporte Bebidas Mundiales.xlsx',
         path:'./src/excel/Reporte Bebidas Mundiales.xlsx'},
         {filenname:'Reporte Topo Chico.xlsx',
         path:'./src/excel/Reporte Topo Chico.xlsx'}],
      text:'Por medio de este conducto nos permitimos enviarles este reporte. Atentamente: Zamudio y Rodríguez. ',
   }
   const transport = nodemailer.createTransport(config);
   //transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   const info=await transport.sendMail(mensaje);
   console.log(info); 
   //console.log(correos);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Proceso para reportes de distribucion 
//Ya se manda el correo cambiar los detinatarios en la funcion del correo
app.get('/getdata_reportedistribucion',function(req,res,next){
   

   sqldist.getdata_reportedistribucion().then((result)=>{
      //res.json(result);
      //console.log(result) 
      var wb= new xl.Workbook();
      let nombreArchivo="Reporte Distribucion";
      var ws=wb.addWorksheet("Reporte");
      //Estilo Columnas
      var estiloTitulo=wb.createStyle({
         font:{
         name: 'Arial',
         color: '#FFFFFF',
         size:10,
         bold: true,
         } ,
         fill:{
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid',
            fgColor: '#288BA8',
         },
      });
      var estilocontenido=wb.createStyle({
         font:{
            name: 'Arial',
            color: '#000000',
            size:10,
         }
      });
         //Nombre de las columnas
      ws.cell(1,1).string("No Control").style(estiloTitulo);
      ws.cell(1,2).string("Customer").style(estiloTitulo);
      ws.cell(1,3).string("Arrival Date").style(estiloTitulo);
      ws.cell(1,4).string("Time").style(estiloTitulo);
      ws.cell(1,5).string("Carrier").style(estiloTitulo);
      ws.cell(1,6).string("Trailer").style(estiloTitulo);
      ws.cell(1,7).string("Serial").style(estiloTitulo);
      ws.cell(1,8).string("Skid").style(estiloTitulo);
      ws.cell(1,9).string("Part").style(estiloTitulo);
      ws.cell(1,10).string("Description").style(estiloTitulo);
      ws.cell(1,11).string("Quantity").style(estiloTitulo);
      ws.cell(1,12).string("Unit").style(estiloTitulo);
      ws.cell(1,13).string("Qty").style(estiloTitulo);
      ws.cell(1,14).string("Unit2").style(estiloTitulo);
      ws.cell(1,15).string("Weight").style(estiloTitulo);
      ws.cell(1,16).string("Section").style(estiloTitulo);
      ws.cell(1,17).string("Days In Warehouse").style(estiloTitulo);
      let numfila=2;
      result.forEach(reglonactual => {
         ws.cell(numfila,1).string(reglonactual.NoControl).style(estilocontenido);//A
         ws.cell(numfila,2).string(reglonactual.Customer).style(estilocontenido);//B
         ws.cell(numfila,3).string(reglonactual.ArrivalDate).style(estilocontenido);//C
         ws.cell(numfila,4).string(reglonactual.Time).style(estilocontenido);//D
         ws.cell(numfila,5).string(reglonactual.Carrier).style(estilocontenido);//E
         ws.cell(numfila,6).string(reglonactual.Trailer).style(estilocontenido);//F
         ws.cell(numfila,7).string(reglonactual.Serial).style(estilocontenido);//G
         ws.cell(numfila,8).string(reglonactual.Skid).style(estilocontenido);//H
         ws.cell(numfila,9).string(reglonactual.Part).style(estilocontenido);//I
         ws.cell(numfila,10).string(reglonactual.Description).style(estilocontenido);//J
         ws.cell(numfila,11).number(reglonactual.Quantity).style(estilocontenido);//K
         ws.cell(numfila,12).string(reglonactual.Unit).style(estilocontenido);//L
         ws.cell(numfila,13).number(reglonactual.Qty).style(estilocontenido);//M
         ws.cell(numfila,14).string(reglonactual.Unit2).style(estilocontenido);//N
         ws.cell(numfila,15).number(reglonactual.Weight).style(estilocontenido);//O
         ws.cell(numfila,16).string(reglonactual.Section).style(estilocontenido);//P
         ws.cell(numfila,17).number(reglonactual.DaysInWarehouse).style(estilocontenido);//Q
         numfila=numfila+1;
      });
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
         function downloadFile(){res.download(pathExcel);}
         downloadFile();
         /*fs.rm(pathExcel,function(err){
            if(err)console.log(err);*/
            /*else*/ console.log("Archivo descargado exitoso");
            
         /*});*/
      }
   });
   })
   sql.getdata_correos_reporte('2').then((result)=>{
      result.forEach(renglonactual=>{
         enviarMailreportedistribucion(renglonactual.correos);
      })
   })
   //enviarMailreportedistribucion()
});
enviarMailreportedistribucion=async(correos)=>{
   const config ={
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth:{
         user:process.env.useremail, 
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false,
       }, 
   } 
   const mensaje ={
      from:'it.sistemas@zayro.com',
      //to:'ca.she@logisteed-america.com, ja.diaz@logisteed-america.com, lfzamudio@zayro.com, lzamudio@zayro.com, distribution1@zayro.com, alule@zayro.com, avazquez@zayro.com',
      //to:'oswal15do@gmail.com',
      to:correos,
      subject:'Envio de reporte de distribucion',
      attachments:[
         {filename:'Reporte Distribucion.xlsx',
         path:'./src/excel/Reporte Distribucion.xlsx'}],
      text:'Por medio de este conducto nos permitimos enviarles este reporte. Atentamente: Zamudio y Rodríguez. ',
   }
   const transport = nodemailer.createTransport(config);
   //transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   const info=await transport.sendMail(mensaje);
   console.log(info); 
   //console.log(correos); 
}   
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Proceso para semaforo
app.get('/getdata_semaforo',function(req,res,next){
   let config ={
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth:{
         user:process.env.useremail, 
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false,
       }, 
   } 
   let transport = nodemailer.createTransport(config);
   sql.getdata_SemaforoEjecutivos().then((resultado)=> { 
      resultado.forEach(ren=>{
   var wb= new xl.Workbook();
   let nombreArchivo="Reporte Semaforo";
      //Estilo Columnas
   var estiloTitulo=wb.createStyle({
         font:{
         name: 'Arial',
         color: '#FFFFFF',
         size:10,
         bold: true,
         } ,
         fill:{
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid',
            fgColor: '#000000',
      },
   });
   var estilocontenidorojo=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#ff0000',
      },
   });
   var estilocontenidoamarillo=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#f7ff00',
      },
   });
   var estilocontenidoverde=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#00ff00',
      },
   });
   var estilocontenidoporelcliente=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#6f00ff',
      },
   });
   var estilocontenidoporllegar=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#ff7c00',
      },
   });
   var estilocontenido=estiloTitulo;
   var variablews;
   let ws=wb.addWorksheet("NuevoLaredo");
   ws.cell(1,1).string("Cliente").style(estiloTitulo);
   ws.cell(1,2).string("Referencia").style(estiloTitulo);
   ws.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   ws.cell(1,4).string("Estado").style(estiloTitulo);
   ws.cell(1,5).string("Dias").style(estiloTitulo);
   ws.cell(1,6).string("ETA").style(estiloTitulo);
   ws.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   let wsver=wb.addWorksheet("Veracruz");
   wsver.cell(1,1).string("Cliente").style(estiloTitulo);
   wsver.cell(1,2).string("Referencia").style(estiloTitulo);
   wsver.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   wsver.cell(1,4).string("Estado").style(estiloTitulo);
   wsver.cell(1,5).string("Dias").style(estiloTitulo);
   wsver.cell(1,6).string("ETA").style(estiloTitulo);
   wsver.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   let wsCorresponsalias=wb.addWorksheet("Corresponsalias");
   wsCorresponsalias.cell(1,1).string("Cliente").style(estiloTitulo);
   wsCorresponsalias.cell(1,2).string("Referencia").style(estiloTitulo);
   wsCorresponsalias.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   wsCorresponsalias.cell(1,4).string("Estado").style(estiloTitulo);
   wsCorresponsalias.cell(1,5).string("Dias").style(estiloTitulo);
   wsCorresponsalias.cell(1,6).string("ETA").style(estiloTitulo);
   wsCorresponsalias.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   let wsAICM=wb.addWorksheet("AICM");
   wsAICM.cell(1,1).string("Cliente").style(estiloTitulo);
   wsAICM.cell(1,2).string("Referencia").style(estiloTitulo);
   wsAICM.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   wsAICM.cell(1,4).string("Estado").style(estiloTitulo);
   wsAICM.cell(1,5).string("Dias").style(estiloTitulo);
   wsAICM.cell(1,6).string("ETA").style(estiloTitulo);
   wsAICM.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   let wsVirtuales=wb.addWorksheet("Virtuales");
   wsVirtuales.cell(1,1).string("Cliente").style(estiloTitulo);
   wsVirtuales.cell(1,2).string("Referencia").style(estiloTitulo);
   wsVirtuales.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   wsVirtuales.cell(1,4).string("Estado").style(estiloTitulo);
   wsVirtuales.cell(1,5).string("Dias").style(estiloTitulo);
   wsVirtuales.cell(1,6).string("ETA").style(estiloTitulo);
   wsVirtuales.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   let wsFerr=wb.addWorksheet("Ferrocarril");
   wsFerr.cell(1,1).string("Cliente").style(estiloTitulo);
   wsFerr.cell(1,2).string("Referencia").style(estiloTitulo);
   wsFerr.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   wsFerr.cell(1,4).string("Estado").style(estiloTitulo);
   wsFerr.cell(1,5).string("Dias").style(estiloTitulo);
   wsFerr.cell(1,6).string("ETA").style(estiloTitulo);
   wsFerr.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   let wsporllegar=wb.addWorksheet("Por llegar");
   wsporllegar.cell(1,1).string("Cliente").style(estiloTitulo);
   wsporllegar.cell(1,2).string("Referencia").style(estiloTitulo);
   wsporllegar.cell(1,3).string("Fecha de Entrada").style(estiloTitulo);
   wsporllegar.cell(1,4).string("Estado").style(estiloTitulo);
   wsporllegar.cell(1,5).string("Dias").style(estiloTitulo);
   wsporllegar.cell(1,6).string("ETA").style(estiloTitulo);
   wsporllegar.cell(1,7).string("Ejecutivo").style(estiloTitulo);
   //ws.cell(2,1).string("1").style(estilocontenido);//A
   //ws1.cell(2,1).string("8").style(estilocontenido);//A
   //console.log(ws)
      //Nombre de las columnas
   
         //console.log(ren.id)
         sql.getdata_SemaforoReporte(ren.id).then((result)=> {
            let numfila;
            let numfilanld=2;
            let numfilaver=2;
            let numfilacorr=2;
            let numfilaaicm=2;
            let numfilavir=2;
            let numfilaferr=2;
            let numfilaporlle=2;
            //console.log(result)
             result.forEach(reglonactual => {
               if (reglonactual.Dias==0){
                  estilocontenido=estilocontenidoverde;
               }else{
                  if (reglonactual.Dias==1){
                     estilocontenido=estilocontenidoamarillo;
                  }else{
                     if(reglonactual.Dias>1){
                        estilocontenido=estilocontenidorojo;
                     }
                  }
               }
               if(reglonactual.Estado=='DEPENDE DEL CLIENTE'){
                  estilocontenido=estilocontenidoporelcliente;
               }
               switch(reglonactual.tipo)
               {
                  case '1':
                     variablews=ws;
                     numfila=numfilanld;
                     break;
                  case'2':
                     variablews=wsver;
                     numfila=numfilaver;
                   break;
                  case'3':
                     variablews=wsCorresponsalias;
                     numfila=numfilacorr;
                   break;
                  case'4':
                     variablews=wsAICM;
                     numfila=numfilaaicm;
                   break;
                  case'5':
                     variablews=wsVirtuales;
                     numfila=numfilavir;
                   break;
                  case'6':
                     variablews=wsFerr;
                     numfila=numfilaferr;
                   break;
                  case'7':
                     variablews=wsporllegar;
                     numfila=numfilaporlle;
                   break;
               }
               //console.log(variablews)
               if (reglonactual.Nombre==''){
                  variablews.cell(numfila,1).string("").style(estilocontenido);//A
               }
               else{
                  variablews.cell(numfila,1).string(reglonactual.Nombre).style(estilocontenido);//A
               }
               if(reglonactual.Referencia==''){
                  variablews.cell(numfila,2).string("").style(estilocontenido);//B
               }
               else{
                  variablews.cell(numfila,2).string(reglonactual.Referencia).style(estilocontenido);//B
               }
               if(reglonactual.Fecha==''){
                  variablews.cell(numfila,3).string("").style(estilocontenido);//C
               }
               else{
                  variablews.cell(numfila,3).string(reglonactual.Fecha).style(estilocontenido);//C
               }
               if(reglonactual.Estado==''){
                  variablews.cell(numfila,4).string("").style(estilocontenido);//D
               }
               else{
                  variablews.cell(numfila,4).string(reglonactual.Estado).style(estilocontenido);//D
               }
               if(reglonactual.Dias==''){
                  variablews.cell(numfila,5).number(0).style(estilocontenido);//E
               }
               else{
                  variablews.cell(numfila,5).number(reglonactual.Dias).style(estilocontenido);//E
               }
               if(reglonactual.Eta==''){
                  variablews.cell(numfila,6).string("").style(estilocontenido);//F
               }
               else{
                  variablews.cell(numfila,6).string(reglonactual.Eta).style(estilocontenido);//F
               }
               if(reglonactual.Ejecutivo==''){
                  variablews.cell(numfila,7).string("").style(estilocontenido);//G
               }
               else{
                  variablews.cell(numfila,7).string(reglonactual.Ejecutivo).style(estilocontenido);//G
               }
               switch(reglonactual.tipo)
               {
                  case '1':
                     numfilanld=numfilanld+1;
                     break;
                  case'2':
                     numfilaver=numfilaver+1;
                     break;
                  case'3':
                     numfilanumfilacorr=numfilacorr+1;
                     break;
                  case'4':
                     numfilaaicm=numfilaaicm+1;
                     break;
                  case'5':
                     numfilavir=numfilavir+1;
                     break;
                  case'6':
                     numfilafer=numfilaferr+1;
                     break;
                  case'7':
                     numfilaporlle=numfilaporlle+1;
                     break;
               }
               //numfila=numfila+1;
               //console.log(ws)
            });
            const pathExcel=path.join(__dirname,'excel',nombreArchivo+' '+ren.nombre+'.xlsx');
            wb.write(pathExcel,function(err,stats){
               if(err) console.log(err);
               else{
                  //console.log("Archivo descargado exitoso");
               }
            });
            //Enviar correo
            enviarMailsemaforo(ren.email,nombreArchivo,ren.nombre,transport);
         });

      })
   })   
   res.json("Archivos generados") 
   /*const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
            wb.write(pathExcel,function(err,stats){
               if(err) console.log(err);
               else{
                  res.json("Archivos generados")
                  //console.log("Archivo descargado exitoso");
               }
            });*/
   //res.json("Archivos generados")
});
enviarMailsemaforo=async(correo, nombreArchivo,nombre,transport)=>{
   const mensaje ={
      from:'it.sistemas@zayro.com',
      to: correo+',avazquez@zayro.com,sistemas@zayro.com,gerenciati@zayro.com',
      //to: 'Oswal15do@gmail.com',
      subject:'Reporte: '+nombre,
      attachments:[
         {filename:nombreArchivo+' '+nombre+'.xlsx',
         path:'./src/excel/'+nombreArchivo+' '+nombre+'.xlsx'}],
         text:'Por medio de este conducto nos permitimos enviarles este reporte. Atentamente: Zamudio y Rodríguez. ',
      }
      console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      //transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
      const info=await transport.sendMail(mensaje);
      //console.log(info);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Proceso Mercancias en Bodega
//Ya se manda el correo cambiar los detinatarios en la funcion del correo
app.get('/getdata_reportemercanciasenbodega',function(req,res,next){
   
   let config ={
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth:{
         user:process.env.useremail, 
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false,
       }, 
   } 
   let transport = nodemailer.createTransport(config);
   /*sqlsis.getdata_listaclientes().then((result)=>{
          //res.json(result);
      //console.log(result) 
      return result
   });*/
 
 /*********************************************** sqlsis.getdata_listaclientes().then((result)=>*/
   sql.getdata_listaclientes().then((result)=>{
      //console.log(result)
      /*********************************************** result.forEach((renglonactual*/
      result.forEach((renglonactual=>{ setTimeout(()=>{
         var wb= new xl.Workbook();
         let date=new Date();
         let fechaDia    = date.getUTCDate();
                  let fechaMes=date.getUTCMonth();
                  let fechaAnio=date.getUTCFullYear();
         //let nombreArchivo="Mercancias en Bodega "+fechaDia +"-"+fechaMes+"-"+fechaAnio;
         let nombreArchivo="Mercancias en Bodega ";
          //console.log(result1)
          var ws=wb.addWorksheet(nombreArchivo);
          //Estilo Columnas
       var estiloTitulo=wb.createStyle({
          font:{
          name: 'Arial',
          color: '#FFFFFF',
          size:10,
          bold: true,
       } ,
       fill:{
          type: 'pattern', // the only one implemented so far.
          patternType: 'solid',
          fgColor: '#288BA8',
       },
       });
       var estilocontenido=wb.createStyle({
          font:{
             name: 'Arial',
             color: '#000000',
             size:10,
          }
       });
             //Nombre de las columnas
       ws.cell(1,1).string("Referencia").style(estiloTitulo);
       ws.cell(1,2).string("Fecha Arribo").style(estiloTitulo);
       ws.cell(1,3).string("Hora Arribo").style(estiloTitulo);
       ws.cell(1,4).string("Cliente").style(estiloTitulo);
       ws.cell(1,5).string("Proveedor").style(estiloTitulo);
       ws.cell(1,6).string("Embarcador").style(estiloTitulo);
       ws.cell(1,7).string("Factura").style(estiloTitulo);
       ws.cell(1,8).string("Linea de Arribo").style(estiloTitulo);
       ws.cell(1,9).string("Categoria").style(estiloTitulo);
       ws.cell(1,10).string("Pedido").style(estiloTitulo);
       ws.cell(1,11).string("Guia").style(estiloTitulo);
       ws.cell(1,12).string("Peso Lbs").style(estiloTitulo);
       ws.cell(1,13).string("Peso Kgs").style(estiloTitulo);
       ws.cell(1,14).string("Bultos").style(estiloTitulo);
       ws.cell(1,15).string("Caja").style(estiloTitulo);
       ws.cell(1,16).string("Estatus").style(estiloTitulo);
       ws.cell(1,17).string("Dias en Bodega").style(estiloTitulo);
       ws.cell(1,18).string("Descripcion").style(estiloTitulo);
       ws.cell(1,19).string("Observaciones").style(estiloTitulo);
       ws.cell(1,20).string("Load entrada").style(estiloTitulo);
       ws.cell(1,21).string("Load salida").style(estiloTitulo);
       ws.cell(1,22).string("Obs. Trafico").style(estiloTitulo);
      /*********************************************** */
      sql.getdata_mercanciasenbodega(renglonactual.cliente_id).then((result1)=>{
            //console.log(result1)
           /*if (result1.json() == "")
           {
               console.log("Vacio")
           }*/
            let numfila=2;
            /*********************************************** */2
            result1.forEach((reglonactual1=>{
               if (reglonactual1.Referencia==''){
                  ws.cell(numfila,1).string("").style(estilocontenido);//A
               }else{
                  ws.cell(numfila,1).string(reglonactual1.Referencia).style(estilocontenido);//A
               }
               if (reglonactual1.FechaArribo==''){
                  ws.cell(numfila,2).string("").style(estilocontenido);//B
               }else{
                  ws.cell(numfila,2).string(reglonactual1.FechaArribo).style(estilocontenido);//B
               }
               if (reglonactual1.HoraArribo==''){
                  ws.cell(numfila,3).string("").style(estilocontenido);//C
               }else{
                  ws.cell(numfila,3).string(reglonactual1.HoraArribo).style(estilocontenido);//C
               }
               if (reglonactual1.Cliente==''){
                  ws.cell(numfila,4).string("").style(estilocontenido);//D
               }else{
                  ws.cell(numfila,4).string(reglonactual1.Cliente).style(estilocontenido);//D
               }
               if (reglonactual1.Proveedor==''){
                  ws.cell(numfila,5).string("").style(estilocontenido);//E
               }else{
                  ws.cell(numfila,5).string(reglonactual1.Proveedor).style(estilocontenido);//E
               }
               if (reglonactual1.Embarcador==''){
                  ws.cell(numfila,6).string("").style(estilocontenido);//F
               }else{
                  ws.cell(numfila,6).string(reglonactual1.Embarcador).style(estilocontenido);//F
               }
               if (reglonactual1.Factura==''){
                  ws.cell(numfila,7).string("").style(estilocontenido);//G
               
               }else{
                  ws.cell(numfila,7).string(reglonactual1.Factura).style(estilocontenido);//G
               }
               if (reglonactual1.LineadeArribo==''){
                  ws.cell(numfila,8).string("").style(estilocontenido);//H
               
               }else{
                  ws.cell(numfila,8).string(reglonactual1.LineadeArribo).style(estilocontenido);//H
               }
               if (reglonactual1.Categoria==''){
                  ws.cell(numfila,9).string("").style(estilocontenido);//I
               
               }else{
                  ws.cell(numfila,9).string(reglonactual1.Categoria).style(estilocontenido);//I
               }
               if (reglonactual1.Pedido==''){
                  ws.cell(numfila,10).string("").style(estilocontenido);//J
               
               }else{
                  ws.cell(numfila,10).string(reglonactual1.Pedido).style(estilocontenido);//J
               }
               if (reglonactual1.Guia==''){
                  ws.cell(numfila,11).string("").style(estilocontenido);//K
               
               }else{
                  ws.cell(numfila,11).string(reglonactual1.Guia).style(estilocontenido);//K
               }
               if (reglonactual1.PesoLbs==''){
                  ws.cell(numfila,12).number(0.00).style(estilocontenido);//L
               
               }else{
                  ws.cell(numfila,12).number(reglonactual1.PesoLbs).style(estilocontenido);//L
               }
               if (reglonactual1.PesoKgs==''){
                  ws.cell(numfila,13).number(0.00).style(estilocontenido);//M
               }else{
                  ws.cell(numfila,13).string(reglonactual1.PesoKgs).style(estilocontenido);//M
               }
               if (reglonactual1.Bultos==''){
                  ws.cell(numfila,14).number(0).style(estilocontenido);//N
               }else{
                  ws.cell(numfila,14).number(reglonactual1.Bultos).style(estilocontenido);//N
               }
               if (reglonactual1.Caja==''){
                  ws.cell(numfila,15).string("").style(estilocontenido);//O
               }else{
                  ws.cell(numfila,15).string(reglonactual1.Caja).style(estilocontenido);//O
               }
               if (reglonactual1.Estatus==''){
                  ws.cell(numfila,16).string("").style(estilocontenido);//P
               }else{
                  ws.cell(numfila,16).string(reglonactual1.Estatus).style(estilocontenido);//P
               }
               if (reglonactual1.DiasenBodega==''){
                  ws.cell(numfila,17).number(0).style(estilocontenido);//Q
               }else{
                  ws.cell(numfila,17).number(reglonactual1.Diasenbodega).style(estilocontenido);//Q
               }
               if (reglonactual1.Descripcion==''){
                  ws.cell(numfila,18).string("").style(estilocontenido);//R
              
               }else{
                  ws.cell(numfila,18).string(reglonactual1.Descripcion).style(estilocontenido);//R
               }
               if (reglonactual1.Observaciones==''){
                  ws.cell(numfila,19).string("").style(estilocontenido);//S
                  
               }else{
                  ws.cell(numfila,19).string(reglonactual1.Observaciones).style(estilocontenido);//S
               }
               if (reglonactual1.Loadentrada==''){
                  ws.cell(numfila,20).string("").style(estilocontenido);//T
                  
               }else{
                  ws.cell(numfila,20).string(reglonactual1.Loadentrada).style(estilocontenido);//T
               }
               if (reglonactual1.Loadsalida==''){
                  ws.cell(numfila,21).string("").style(estilocontenido);//U
                  
               }else{
                  ws.cell(numfila,21).string(reglonactual1.Loadsalida).style(estilocontenido);//U
               }
               if (reglonactual1.ObsTrafico==''){
                  ws.cell(numfila,22).string("").style(estilocontenido);//V
               }else{
                  ws.cell(numfila,22).string(reglonactual1.ObsTrafico).style(estilocontenido);//V
               }
               
               numfila=numfila+1;
             
            }))//Fin renglon actual 1
            /*********************************************** */
            /************************************************* */       
            const pathExcel=path.join(__dirname,'excel',nombreArchivo+' '+renglonactual.numero+'.xlsx');
            //Guardar
            wb.write(pathExcel,function(err,stats){
               if(err) console.log(err);
               else{
               /*fs.rm(pathExcel,function(err){
                     if(err)console.log(err);
                     else {
                        
                        console.log("Archivo descargado exitoso");
                     }
                     
                  });*/

               //function downloadFile(){res.download(pathExcel);}
               //downloadFile();
               console.log("Archivo descargado exitoso");   
               }
               
            });
            var correoscliente;
            var correosclientecc;
            sql.getdata_correos_ejecutivos_cliente(renglonactual.cliente_id).then((resultc)=>{
               //res.json(resultc)
               resultc.forEach((reglonactual2=>{
                  correoscliente=reglonactual2.correos
                  correosclientecc=reglonactual2.correoscc
               }))
               console.log(correoscliente,correosclientecc)
               //Despues de guardar tiene que mandar el correo
               enviarMailreportebodega(correoscliente,correosclientecc,nombreArchivo,renglonactual.nomcli,renglonactual.numero,transport);
            }); 
            
            
         });
         /*********************************************** */  1 
      },10000)
        
      }))
      
      /*********************************************** result.forEach((renglonactual*/  
   });
   /*********************************************** finsqlsis.getdata_listaclientes().then((result)=>*/
   res.json("Archivos generados");
});
//Los correos a quienes va se tienen que configurar
enviarMailreportebodega=async(correos, correoscc,nombreArchivo,nomcli,numero,transport)=>{
   const mensaje ={
      from:'it.sistemas@zayro.com',
      to: correos+', '+correoscc+',sistemas@zayro.com',
      //to: 'Oswal15do@gmail.com',
      subject:'Mercancia en bodega: '+nomcli+' - '+numero,
      attachments:[
         {filename:nombreArchivo+'.xlsx',
         path:'./src/excel/'+nombreArchivo+' '+numero+'.xlsx'}],
         text:'Por medio de este conducto nos permitimos enviarles este reporte. Atentamente: Zamudio y Rodríguez. ',
      }
      console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      //transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
      const info=await transport.sendMail(mensaje);
      console.log(info);
} 

/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Reporte ASN
app.get('/getdata_reporteASN',function async (req,res,next){
   let config ={
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth:{
         user:process.env.useremail, 
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false,
       }, 
   } 
   let transport = nodemailer.createTransport(config);
   sql.getdata_ReporteASN().then((resultado)=> { 
      var wb= new xl.Workbook();
      let nombreArchivo="Reporte ASN";
      var estiloTitulo=wb.createStyle({
         font:{
         name: 'Arial',
         color: '#FFFFFF',
         size:10,
         bold: true,
         } ,
         fill:{
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid',
            fgColor: '#000000',
      },
      });
      var estilocontenidoletraroja=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#ff0000',
         size:10,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#f7ff00',
      },
      });
      var estilocontenidonormal=wb.createStyle({
         font:{
            name: 'Arial',
            color: '#000000',
            size:10,
         },
         fill:{
            type: 'pattern', // the only one implemented so far.
            patternType: 'solid',
            fgColor: '#f7ff00',
         },
      });
      var estilocontenido;
      var variablews;
      let ws=wb.addWorksheet("Targets Bodega");
      ws.cell(1,1).string("Fecha").style(estiloTitulo);
      ws.cell(1,2).string("PO LCN").style(estiloTitulo);
      ws.cell(1,3).string("POD LCN").style(estiloTitulo);
      ws.cell(1,4).string("PO").style(estiloTitulo);
      ws.cell(1,5).string("BOX ID").style(estiloTitulo);
      ws.cell(1,6).string("NUM PARTE").style(estiloTitulo);
      ws.cell(1,7).string("PIEZAS LONG BOX").style(estiloTitulo);
      ws.cell(1,8).string("PIEZAS PO").style(estiloTitulo);
      ws.cell(1,9).string("FECHA APROX LAREDO").style(estiloTitulo);
      ws.cell(1,10).string("PROVEEDOR").style(estiloTitulo);
      ws.cell(1,11).string("FECHA ESCANEADO").style(estiloTitulo);
      ws.cell(1,12).string("PALLET").style(estiloTitulo);
      ws.cell(1,13).string("ESCANEADO").style(estiloTitulo);
      ws.cell(1,14).string("FECHA CIERRE").style(estiloTitulo);
      ws.cell(1,15).string("HB203").style(estiloTitulo);
      let wscan=wb.addWorksheet("Cancelados");
      wscan.cell(1,1).string("Fecha").style(estiloTitulo);
      wscan.cell(1,2).string("PO LCN").style(estiloTitulo);
      wscan.cell(1,3).string("POD LCN").style(estiloTitulo);
      wscan.cell(1,4).string("PO").style(estiloTitulo);
      wscan.cell(1,5).string("BOX ID").style(estiloTitulo);
      wscan.cell(1,6).string("NUM PARTE").style(estiloTitulo);
      wscan.cell(1,7).string("PIEZAS LONG BOX").style(estiloTitulo);
      wscan.cell(1,8).string("PIEZAS PO").style(estiloTitulo);
      wscan.cell(1,9).string("FECHA APROX LAREDO").style(estiloTitulo);
      wscan.cell(1,10).string("PROVEEDOR").style(estiloTitulo);
      wscan.cell(1,11).string("FECHA ESCANEADO").style(estiloTitulo);
      wscan.cell(1,12).string("PALLET").style(estiloTitulo);
      wscan.cell(1,13).string("ESCANEADO").style(estiloTitulo);
      wscan.cell(1,14).string("FECHA CIERRE").style(estiloTitulo);
      wscan.cell(1,15).string("HB203").style(estiloTitulo);
      let numfilabod=2;
      let numfilacan=2;
      var numfila;
      resultado.forEach(ren=>{
         if (ren.fecha_scan==''){
            estilocontenido=estilocontenidoletraroja;
         }else{
            estilocontenido=estilocontenidonormal;
         }
         switch(ren.tipo){
            case 1:
               variablews=ws;
               numfila=numfilabod;
               break;
            case 2:
               variablews=wscan;
               numfila=numfilacan;
               break;
         }

         if (ren.fecha==''){
            variablews.cell(numfila,1).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,1).string(ren.fecha).style(estilocontenido);//A
         }
         if (ren.po_lcn==''){
            variablews.cell(numfila,2).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,2).string(ren.po_lcn).style(estilocontenido);//A
         }
         if (ren.pod_lcn==''){
            variablews.cell(numfila,3).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,3).string(ren.pod_lcn).style(estilocontenido);//A
         }
         if (ren.po==''){
            variablews.cell(numfila,4).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,4).string(ren.po).style(estilocontenido);//A
         }
         if (ren.box_id==''){
            variablews.cell(numfila,5).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,5).string(ren.box_id).style(estilocontenido);//A
         }
         if (ren.numparte==''){
            variablews.cell(numfila,6).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,6).string(ren.numparte).style(estilocontenido);//A
         }
         if (ren.piezas_longbox==''){
            variablews.cell(numfila,7).number(0).style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,7).number(ren.piezas_longbox).style(estilocontenido);//A
         }
         if (ren.piezas_po==''){
            variablews.cell(numfila,8).number(0).style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,8).number(ren.piezas_po).style(estilocontenido);//A
         }
         if (ren.fecha_aprox_laredo==''){
            variablews.cell(numfila,9).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,9).string(ren.fecha_aprox_laredo).style(estilocontenido);//A
         }
         if (ren.proveedor==''){
            variablews.cell(numfila,10).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,10).string(ren.proveedor).style(estilocontenido);//A
         }
         if (ren.fecha_scan==''){
            variablews.cell(numfila,11).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,11).string(ren.fecha_scan).style(estilocontenido);//A
         }
         if (ren.pallet==''){
            variablews.cell(numfila,12).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,12).string(ren.pallet).style(estilocontenido);//A
         }
         if (ren.escaneado==''){
            variablews.cell(numfila,13).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,13).string(ren.escaneado).style(estilocontenido);//A
         }
         if (ren.fecha_cierre==''){
            variablews.cell(numfila,14).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,14).string(ren.fecha_cierre).style(estilocontenido);//A
         }
         if (ren.hb203==''){
            variablews.cell(numfila,15).string("").style(estilocontenido);//A
         }
         else{
            variablews.cell(numfila,15).string(ren.hb203).style(estilocontenido);//A
         }

         switch(ren.tipo){
            case 1:
               numfilabod=numfilabod+1;
               break;
            case 2:
               numfilacan=numfilacan+1;
               break;
         }

      });
      const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
      wb.write(pathExcel,function(err,stats){
         if(err) console.log(err);
         else{
            res.json("Archivo generado");
            console.log("Archivo descargado exitoso");
         }
      });
      //enviarMailASN(nombreArchivo,transport)
      sql.getdata_correos_reporte('3').then((result)=>{
         result.forEach(renglonactual=>{
            enviarMailASN(nombreArchivo,transport,renglonactual.correos)
         })
      })
   });

});
enviarMailASN=async(nombreArchivo,transport,correos)=>{
   const mensaje ={
      from:'it.sistemas@zayro.com',
      //to: 'lzamudio@zayro.com,soportetecnico@zayro.com,sistemas@zayro.com',
      to:correos,
      //to: 'Oswal15do@gmail.com',
      subject:'Reporte ASN',
      attachments:[
         {filename:nombreArchivo+'.xlsx',
         path:'./src/excel/'+nombreArchivo+'.xlsx'}],
         text:'Reporte ASN ',
      }
      console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      //transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
      const info=await transport.sendMail(mensaje);
      console.log(info);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//REPORTE Thyssenkrupp 
app.get('/getdata_Thyssenkrupp/:fechaini/:fechafin',function(req,res,next){
   let config ={
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth:{
         user:process.env.useremail, 
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false,
       }, 
   } 
   let transport = nodemailer.createTransport(config);
   var fechaini=req.params.fechaini;
   var fechafin=req.params.fechafin;
   /*console.log(req.params.fechaini);
   console.log(req.params.fechafin);*/
   var wb= new xl.Workbook();
   let nombreArchivo="Estado de cuenta Thyssenkrupp";
   var wsUSD=wb.addWorksheet("Facturadas DLLS");
   var wsMXN=wb.addWorksheet("Facturadas MXN");
   var wsPEN=wb.addWorksheet("Pendientes por facturar");
   //Estilo Columnas
   var estiloTitulo=wb.createStyle({
      font:{
      name: 'Arial',
      color: '#FFFFFF',
      size:10,
      bold: true,
      } ,
      fill:{
      type: 'pattern', // the only one implemented so far.
      patternType: 'solid',
      fgColor: '#00BCF3',
      },
   });
   var estilocontenido=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      }
   });
   var estilozamprov=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#FFFFFF',
         size:14,
         bold: true,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#00448D',
      }
   });
   var estilototal=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
         bold: true,
      },
      fill:{
         type: 'pattern', // the only one implemented so far.
         patternType: 'solid',
         fgColor: '#FFF300',
      }
   });
   wsUSD.cell(1,1).string("No. Proveedor").style(estiloTitulo);
   wsUSD.cell(1,2,1,12,true).string("Razon Social").style(estiloTitulo);
   wsUSD.cell(2,1).string("971556").style(estilozamprov);
   wsUSD.cell(2,2,2,12,true).string("ZAMUDIO Y RODRIGUEZ").style(estilozamprov);
   
   //
   wsMXN.cell(1,1).string("No. Proveedor").style(estiloTitulo);
   wsMXN.cell(1,2,1,12,true).string("Razon Social").style(estiloTitulo);
   wsMXN.cell(2,1).string("971556").style(estilozamprov);
   wsMXN.cell(2,2,2,12,true).string("ZAMUDIO Y RODRIGUEZ").style(estilozamprov);
   //
   wsPEN.cell(1,1).string("No. Proveedor").style(estiloTitulo);
   wsPEN.cell(1,2,1,12,true).string("Razon Social").style(estiloTitulo);
   wsPEN.cell(2,1).string("971556").style(estilozamprov);
   wsPEN.cell(2,2,2,12,true).string("ZAMUDIO Y RODRIGUEZ").style(estilozamprov);
   sqlzay.getdata_ReporteThyssenhrup_dolares(fechafin).then((result)=>{
      //res.json(result);
      //console.log(result) 
         //Nombre de las columnas
      wsUSD.cell(4,1).string("No. Proveedor").style(estiloTitulo);
      wsUSD.cell(4,2).string("Razon Social").style(estiloTitulo);
      wsUSD.cell(4,3).string("No. Factura").style(estiloTitulo);
      wsUSD.cell(4,4).string("Fecha").style(estiloTitulo);
      wsUSD.cell(4,5).string("Credito").style(estiloTitulo);
      wsUSD.cell(4,6).string("Vencimiento").style(estiloTitulo);
      wsUSD.cell(4,7).string("IMP/EXP").style(estiloTitulo);
      wsUSD.cell(4,8).string("PO").style(estiloTitulo);
      wsUSD.cell(4,9).string("Cuenta Contable").style(estiloTitulo);
      wsUSD.cell(4,10).string("SubTotal").style(estiloTitulo);
      wsUSD.cell(4,11).string("IVA").style(estiloTitulo);
      wsUSD.cell(4,12).string("Retencion").style(estiloTitulo);
      wsUSD.cell(4,13).string("Total").style(estiloTitulo);
      wsUSD.cell(4,14).string("Moneda").style(estiloTitulo);
      wsUSD.cell(4,15).string("Comentarios").style(estiloTitulo);
      let numfila=5;
      let total=0;
      result.forEach(reglonactual => {
         wsUSD.cell(numfila,1).string(reglonactual.NoProveedor).style(estilocontenido);
         wsUSD.cell(numfila,2).string(reglonactual.RazonSocial).style(estilocontenido);
         wsUSD.cell(numfila,3).string(reglonactual.NoFactura).style(estilocontenido);
         wsUSD.cell(numfila,4).string(reglonactual.Fecha).style(estilocontenido);
         wsUSD.cell(numfila,5).number(reglonactual.Credito).style(estilocontenido);
         wsUSD.cell(numfila,6).string(reglonactual.Vencimiento).style(estilocontenido);
         wsUSD.cell(numfila,7).string(reglonactual.IMPEXP).style(estilocontenido);
         if (reglonactual.PO==""){
            wsUSD.cell(numfila,8).string("").style(estilocontenido);
         }else{
            wsUSD.cell(numfila,8).string(reglonactual.PO).style(estilocontenido);
         }
         wsUSD.cell(numfila,9).string(reglonactual.CuentaContable).style(estilocontenido);
         if(reglonactual.SubTotal==""){
            wsUSD.cell(numfila,10).number(0).style(estilocontenido);
         }else{
            wsUSD.cell(numfila,10).number(reglonactual.Subtotal).style(estilocontenido);
         }
         wsUSD.cell(numfila,11).number(reglonactual.IVA).style(estilocontenido);
         wsUSD.cell(numfila,12).number(reglonactual.Retencion).style(estilocontenido);
         wsUSD.cell(numfila,13).number(reglonactual.Total).style(estilocontenido);
         total=total+reglonactual.Total
         wsUSD.cell(numfila,14).string(reglonactual.Moneda).style(estilocontenido);
         wsUSD.cell(numfila,15).string("").style(estilocontenido);
         numfila=numfila+1;
      });
      wsUSD.cell(1,13).string("Total").style(estilototal);
      wsUSD.cell(2,13).number(total).style(estilototal);
      wsUSD.cell(numfila,12).string("Total").style(estilototal);
      wsUSD.cell(numfila,13).number(total).style(estilototal);
      
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
      }
   });
   })
   sqlzam.getdata_ReporteThyssenhrup_pesos(fechafin).then((result)=>{
      //res.json(result);
      //console.log(result) 
         //Nombre de las columnas
      wsMXN.cell(4,1).string("No. Proveedor").style(estiloTitulo);
      wsMXN.cell(4,2).string("Razon Social").style(estiloTitulo);
      wsMXN.cell(4,3).string("No. Factura").style(estiloTitulo);
      wsMXN.cell(4,4).string("Fecha").style(estiloTitulo);
      wsMXN.cell(4,5).string("Credito").style(estiloTitulo);
      wsMXN.cell(4,6).string("Vencimiento").style(estiloTitulo);
      wsMXN.cell(4,7).string("IMP/EXP").style(estiloTitulo);
      wsMXN.cell(4,8).string("PO").style(estiloTitulo);
      wsMXN.cell(4,9).string("Cuenta Contable").style(estiloTitulo);
      wsMXN.cell(4,10).string("SubTotal").style(estiloTitulo);
      wsMXN.cell(4,11).string("IVA").style(estiloTitulo);
      wsMXN.cell(4,12).string("Retencion").style(estiloTitulo);
      wsMXN.cell(4,13).string("Total").style(estiloTitulo);
      wsMXN.cell(4,14).string("Moneda").style(estiloTitulo);
      wsMXN.cell(4,15).string("Comentarios").style(estiloTitulo);
      let numfilamxn=5;
      let total=0;
      result.forEach(reglonactual => {
         wsMXN.cell(numfilamxn,1).string(reglonactual.NoProveedor).style(estilocontenido);
         wsMXN.cell(numfilamxn,2).string(reglonactual.RazonSocial).style(estilocontenido);
         wsMXN.cell(numfilamxn,3).string(reglonactual.NoFactura).style(estilocontenido);
         wsMXN.cell(numfilamxn,4).string(reglonactual.Fecha).style(estilocontenido);
         wsMXN.cell(numfilamxn,5).number(reglonactual.Credito).style(estilocontenido);
         wsMXN.cell(numfilamxn,6).string(reglonactual.Vencimiento).style(estilocontenido);
         wsMXN.cell(numfilamxn,7).string(reglonactual.IMPEXP).style(estilocontenido);
         if (reglonactual.PO==""){
            wsMXN.cell(numfilamxn,8).string("").style(estilocontenido);
         }else{
            wsMXN.cell(numfilamxn,8).string(reglonactual.PO).style(estilocontenido);
         }
         wsMXN.cell(numfilamxn,9).string(reglonactual.CuentaContable).style(estilocontenido);
         if(reglonactual.SubTotal==""){
            wsMXN.cell(numfilamxn,10).number(0).style(estilocontenido);
         }else{
            wsMXN.cell(numfilamxn,10).number(reglonactual.Subtotal).style(estilocontenido);
         }
         wsMXN.cell(numfilamxn,11).number(reglonactual.IVA).style(estilocontenido);
         wsMXN.cell(numfilamxn,12).number(reglonactual.Retencion).style(estilocontenido);
         wsMXN.cell(numfilamxn,13).number(reglonactual.Total).style(estilocontenido);
         total=total+reglonactual.Total
         wsMXN.cell(numfilamxn,14).string(reglonactual.Moneda).style(estilocontenido);
         wsMXN.cell(numfilamxn,15).string("").style(estilocontenido);
         numfilamxn=numfilamxn+1;
      });
      wsMXN.cell(1,13).string("Total").style(estilototal);
      wsMXN.cell(2,13).number(total).style(estilototal);
      wsMXN.cell(numfilamxn,12).string("Total").style(estilototal);
      wsMXN.cell(numfilamxn,13).number(total).style(estilototal);
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
      }
   });
   })
   sql.getdata_Thyssenkrupp_pendientes(fechaini,fechafin).then((result)=>{
      //res.json(result);
      //console.log(result) 
         //Nombre de las columnas
      wsPEN.cell(4,1).string("No. Proveedor").style(estiloTitulo);
      wsPEN.cell(4,2).string("Razon Social").style(estiloTitulo);
      wsPEN.cell(4,3).string("No. Factura").style(estiloTitulo);
      wsPEN.cell(4,4).string("Fecha").style(estiloTitulo);
      wsPEN.cell(4,5).string("Credito").style(estiloTitulo);
      wsPEN.cell(4,6).string("Vencimiento").style(estiloTitulo);
      wsPEN.cell(4,7).string("IMP/EXP").style(estiloTitulo);
      wsPEN.cell(4,8).string("PO").style(estiloTitulo);
      wsPEN.cell(4,9).string("Cuenta Contable").style(estiloTitulo);
      wsPEN.cell(4,10).string("SubTotal").style(estiloTitulo);
      wsPEN.cell(4,11).string("IVA").style(estiloTitulo);
      wsPEN.cell(4,12).string("Retencion").style(estiloTitulo);
      wsPEN.cell(4,13).string("Total").style(estiloTitulo);
      wsPEN.cell(4,14).string("Moneda").style(estiloTitulo);
      wsPEN.cell(4,15).string("Comentarios").style(estiloTitulo);
      let numfila=5;
      result.forEach(reglonactual => {
         wsPEN.cell(numfila,1).string(reglonactual.NoProveedor).style(estilocontenido);
         wsPEN.cell(numfila,2).string(reglonactual.RazonSocial).style(estilocontenido);
         wsPEN.cell(numfila,3).string(reglonactual.NoFactura).style(estilocontenido);
         wsPEN.cell(numfila,4).string("").style(estilocontenido);
         wsPEN.cell(numfila,5).number(reglonactual.Credito).style(estilocontenido);
         wsPEN.cell(numfila,6).string("").style(estilocontenido);
         wsPEN.cell(numfila,7).string(reglonactual.IMPEXP).style(estilocontenido);
         wsPEN.cell(numfila,8).string("").style(estilocontenido);
         wsPEN.cell(numfila,9).string(reglonactual.CuentaContable).style(estilocontenido);
         wsPEN.cell(numfila,13).string("$").style(estilocontenido);
         wsPEN.cell(numfila,14).string("").style(estilocontenido);
         wsPEN.cell(numfila,15).string(reglonactual.Comentarios).style(estilocontenido);
         numfila=numfila+1;
      });
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
      }
   });
   })
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
         res.json("Archivo generado")
         console.log("Archivo generado")
      }
   });
   setTimeout(()=>{
      sql.getdata_correos_reporte('4').then((result)=>{
         result.forEach(renglonactual=>{
            enviarMailEstadoCuentaThyn(nombreArchivo,transport,renglonactual.correos);
         })
      })
      //enviarMailEstadoCuentaThyn(nombreArchivo,transport);
   },10000);
});
enviarMailEstadoCuentaThyn=async(nombreArchivo,transport,correos)=>{
   const mensaje ={
      from:'it.sistemas@zayro.com',
      to: correos,
      //to: 'Oswal15do@gmail.com',
      subject:'Estado de cuenta Thyssenkrupp',
      attachments:[
         {filename:nombreArchivo+'.xlsx',
         path:'./src/excel/'+nombreArchivo+'.xlsx'}],
         text:'Estado de cuenta Thyssenkrupp',
      }
      console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      //transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
      const info=await transport.sendMail(mensaje);
      console.log(info);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Referencias pendiente de entregar a facturacion

