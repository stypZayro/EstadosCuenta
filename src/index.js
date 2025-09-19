const express = require('express');
const morgan = require('morgan');
const xl = require('excel4node');
const path = require('path');
const fs = require('fs');
const { exec } = require('child_process');

const sql = require('./conexionaduana');
const sqldist = require('./conexiondistribucion');
const sqlzay = require('./conexionzay');
const sqlzam = require('./conexionzamudio');
const sqlSIS = require('./conexionsis');
const sqlSISTEMAS = require('./conexionsistemas');
const sqlram = require('./conexionRam');
const mysql = require('./conexionmysql');
const pg = require('./conexionzayprogrestsql');
////
const socketIO = require('socket.io');
const http = require('http');
const nodemailer = require('nodemailer');
const dotenv = require('dotenv');
const ExcelJS = require('exceljs');
const helmet = require('helmet');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const jwt = require('jsonwebtoken');
const xlsx = require('xlsx-style');
const archiver = require('archiver');
const sqlData=require('./insertAscToTable');

const { v4: uuidv4 } = require('uuid'); // ✅ faltaba
const { validate, authBearer, errorHandler, z } = require('./middlewares');

dotenv.config();

const app = express();
const server = http.createServer(app);
// const io = socketIO(server); // (si lo vas a usar)

// ===== Config =====
const PORT = process.env.PORT || 3001;
const CORS_ORIGINS = (process.env.CORS_ORIGINS || 'https://www.zayro.com')
  .split(',')
  .map(s => s.trim());

// ===== Seguridad / transporte base =====
app.set('trust proxy', 1);
app.use(helmet());
app.disable('x-powered-by');
app.use(cors({ origin: CORS_ORIGINS, credentials: true }));

// Logs
app.use(morgan('dev'));

// JSON parser con límite y manejo de errores JSON
app.use(express.json({ limit: '512kb' }));
app.use((err, req, res, next) => {
  if (err?.type === 'entity.too.large') {
    return res.status(413).json({ error: 'payload_too_large' });
  }
  if (err instanceof SyntaxError && 'body' in err) {
    return res.status(400).json({ error: 'invalid_json' });
  }
  next(err);
});

// Rate limit GLOBAL suave
app.use(rateLimit({
  windowMs: 60 * 1000,
  max: 100,
  standardHeaders: true,
  legacyHeaders: false
}));

// Request-Id para trazabilidad
app.use((req, res, next) => {
  req.id = req.get('X-Request-Id') || uuidv4();
  res.setHeader('X-Request-Id', req.id);
  next();
});

// Rechazo de métodos y Content-Type inesperados
const methodsPermitidos = ['GET','POST','PUT','PATCH','DELETE','OPTIONS'];
app.use((req, res, next) => {
  if (!methodsPermitidos.includes(req.method)) {
    return res.status(405).json({ error: 'method_not_allowed' });
  }
  if (['POST','PUT','PATCH'].includes(req.method)) {
    const ct = (req.get('Content-Type') || '').toLowerCase();
    if (!ct.startsWith('application/json')) {
      return res.status(415).json({ error: 'unsupported_media_type' });
    }
  }
  next();
});

// Normalización GLOBAL de strings (trim; '' -> undefined)
function normalizeStrings(obj) {
  if (obj && typeof obj === 'object') {
    for (const k of Object.keys(obj)) {
      const v = obj[k];
      if (typeof v === 'string') {
        const trimmed = v.trim();
        obj[k] = trimmed === '' ? undefined : trimmed;
      } else if (v && typeof v === 'object') {
        normalizeStrings(v);
      }
    }
  }
}
app.use((req, res, next) => {
  normalizeStrings(req.query);
  normalizeStrings(req.body);
  next();
});

// VALIDACIÓN GLOBAL de query comunes si aparecen
const GlobalQuerySchema = z.object({
  page: z.coerce.number().int().min(1).max(1000).optional(),
  limit: z.coerce.number().int().min(1).max(1000).optional(),
  sort: z.string().max(64).optional(),
  order: z.enum(['asc','desc']).optional(),
  desde: z.string().optional()
    .refine(v => !v || /^\d{4}-\d{2}-\d{2}$/.test(v) || !Number.isNaN(Date.parse(v)), { message: 'Fecha inválida: use YYYY-MM-DD o ISO' }),
  hasta: z.string().optional()
    .refine(v => !v || /^\d{4}-\d{2}-\d{2}$/.test(v) || !Number.isNaN(Date.parse(v)), { message: 'Fecha inválida: use YYYY-MM-DD o ISO' }),
  usuario: z.string().max(64).optional(),
  sucursal: z.string().max(64).optional(),
}).passthrough()
  .superRefine((obj, ctx) => {
    if (obj.desde && obj.hasta) {
      const d1 = new Date(obj.desde);
      const d2 = new Date(obj.hasta);
      if (d1 > d2) ctx.addIssue({ code: z.ZodIssueCode.custom, message: 'desde no puede ser mayor que hasta', path: ['hasta'] });
    }
  });

app.use((req, res, next) => {
  const q = GlobalQuerySchema.safeParse(req.query);
  if (!q.success) {
    return res.status(400).json({ error: 'validation', where: 'query', details: q.error.issues });
  }
  req.query = q.data;
  next();
});

// Hardening de body
const denyBodyKeys = new Set(['$where', '$expr', '__proto__', 'constructor']);
app.use((req, res, next) => {
  if (req.body && typeof req.body === 'object') {
    for (const k of Object.keys(req.body)) {
      if (denyBodyKeys.has(k)) {
        return res.status(400).json({ error: 'invalid_field', field: k });
      }
      const stack = [req.body[k]];
      let depth = 0, maxDepth = 20;
      while (stack.length) {
        const cur = stack.pop();
        if (cur && typeof cur === 'object') {
          depth++;
          if (depth > maxDepth) return res.status(400).json({ error: 'object_too_deep' });
          for (const kk of Object.keys(cur)) stack.push(cur[kk]);
        }
      }
    }
  }
  next();
});

// ====================== Rutas públicas si las hay ======================
// app.get('/health', (req,res)=>res.json({ok:true}));

// ====================== Rutas privadas /api ============================
const api = express.Router();
api.use(authBearer); // ✅ toda /api requiere token (usa tu verificador real)
app.use('/api', api);

// ====== Rate limit fuerte para descargas/reportes ======
const heavyLimiter = rateLimit({
  windowMs: 60 * 1000,
  max: 10,
  standardHeaders: true,
  legacyHeaders: false
});

// ====== Asegurar carpeta excel/ ======
const excelDir = path.join(__dirname, 'excel');
if (!fs.existsSync(excelDir)) fs.mkdirSync(excelDir, { recursive: true });



app.get('/api/getdata_BebidasMundiales', async function(req, res, next) {
   try {
       const result = await sql.getdata_BebidasMundiales();

       const wb = new xl.Workbook();
       const nombreArchivo = "Reporte Bebidas Mundiales";
       const ws = wb.addWorksheet("BebidasMundi");

       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#008000',
           },
       });
       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       const columnas = [
           "REFERENCIA", "FECHA APERTURA PEDIMENTO", "CLIENTE", "PROVEEDOR",
           "FACTURA", "CFDI", "FECHA FACTURA", "TIPO OPERACION", "FECHA CRUCE",
           "C001CAAT", "CAJA", "PLACAS", "PEDIMENTO"
       ];
       columnas.forEach((columna, index) => {
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
           Object.keys(reglonactual).forEach((columna, idx) => {
               ws.cell(numfila, idx + 1).string(reglonactual[columna]).style(estilocontenido);
           });
           numfila++;
       });

       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');

       wb.write(pathExcel, async function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               try {
                   await fs.promises.access(pathExcel, fs.constants.F_OK);
                   res.download(pathExcel, () => {
                       //fs.unlink(pathExcel, (err) => {
                           if (err) console.error(err);
                           else console.log("Archivo descargado y eliminado exitosamente.");
                       //});
                   });
               } catch (err) {
                   console.error(err);
                   res.status(500).send("Error al acceder al archivo Excel generado.");
               }
           }
       });
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
});
app.get('/api/getdata_TopoChico', function(req, res, next) {
   sql.getdata_TopoChico().then((result) => {
       // Crear un nuevo libro de Excel y una nueva hoja de cálculo
       var wb = new xl.Workbook();
       let nombreArchivo = "Reporte Topo Chico";
       var ws = wb.addWorksheet("Topo Chico");

       // Definir estilos para títulos y contenido
       var estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#008000',
           },
       });
       var estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       // Definir el encabezado de las columnas
       const columnas = [
           "REFERENCIA", "FECHA APERTURA PEDIMENTO", "CLIENTE", "PROVEEDOR",
           "FACTURA", "CFDI", "FECHA FACTURA", "TIPO OPERACION", "FECHA CRUCE",
           "C001CAAT", "CAJA", "PLACAS", "PEDIMENTO"
       ];
       columnas.forEach((columna, index) => {
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       // Llenar la hoja de cálculo con los datos
       result.forEach((reglonactual, index) => {
           const numfila = index + 2;
           Object.keys(reglonactual).forEach((columna, idx) => {
               ws.cell(numfila, idx + 1).string(reglonactual[columna]).style(estilocontenido);
           });
       });

       // Guardar el archivo Excel y enviarlo como descarga al cliente
       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');
       wb.write(pathExcel, function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               res.download(pathExcel, () => {
                   // Eliminar el archivo después de que se haya descargado
                   //fs.unlink(pathExcel, (err) => {
                       if (err) console.error(err);
                       else console.log("Archivo descargado y eliminado exitosamente.");
                   //});*
               });
           }
       });
   }).catch((err) => {
       console.error(err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   });
});
app.get('/api/getdata_enviarcorreoBebMunTopChic',function(req,res,next){
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
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
app.get('/api/getdata_reportedistribucion', async function(req, res, next) {
   try {
       const result = await sqldist.getdata_reportedistribucion();

       const wb = new xl.Workbook();
       const nombreArchivo = "Distribution Report";
       const ws = wb.addWorksheet("Reporte");

       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#288BA8',
           },
       });
       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       const columnas = [
           "No Control", "Customer", "Arrival Date", "Time", "Carrier",
           "Trailer", "Serial", "Skid", "Part", "Description", "Quantity",
           "Unit", "Qty", "Unit2", "Weight", "Section", "Days In Warehouse"
       ];
       columnas.forEach((columna, index) => {
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
           Object.keys(reglonactual).forEach((columna, idx) => {
               const valor = reglonactual[columna] !== null && reglonactual[columna] !== undefined ? reglonactual[columna].toString() : '';
               ws.cell(numfila, idx + 1).string(valor).style(estilocontenido);
           });
           numfila++;
       });

       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');

       wb.write(pathExcel, async function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               try {
                   await fs.promises.access(pathExcel, fs.constants.F_OK);
                   res.download(pathExcel, () => {
                       
                           if (err) console.error(err);
                           else console.log("Archivo descargado y eliminado exitosamente.");

                   });
               } catch (err) {
                   console.error(err);
                   res.status(500).send("Error al acceder al archivo Excel generado.");
               }
           }
       });

       const correosResult = await sql.getdata_correos_reporte('2');
       correosResult.forEach(renglonactual => {
           enviarMailreportedistribucion(renglonactual.correos);
       });
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
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
      //to:'programacion@zayro.com',
      to:correos,
      subject:'Distribution Report',
      attachments:[
         {filename:'Distribution Report.xlsx',
         path:'./src/excel/Distribution Report.xlsx'}],
      text:'Please find attached the report',
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos); 
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
app.get('/api/getdata_reportemeow', async function(req, res, next) {
   try {
       const config = {
           host: process.env.hostemail,
           port: process.env.portemail,
           secure: true,
           auth: {
               user: process.env.useremail,
               pass: process.env.passemail
           },
           tls: {
               rejectUnauthorized: false,
           },
       };

       const result = await sqldist.getdata_reportemeow();

       const wb = new xl.Workbook();
       const ws = wb.addWorksheet("Report");

       // Estilo de columnas
       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern', // the only one implemented so far.
               patternType: 'solid',
               fgColor: '#288BA8',
           },
       });

       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       // Nombre de las columnas
       const columnas = ["No Control", "Customer", "Arrival Date", "Time", "Carrier", "Trailer", "Serial", "Skid", "Part", "Description", "Quantity", "Unit", "Qty", "Unit2", "Weight", "Section", "Days In Warehouse"];
       columnas.forEach((titulo, index) => {
           ws.cell(1, index + 1).string(titulo).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
           columnas.forEach((columna, index) => {
               const contenido = reglonactual[columna.replace(/ /g, '')];
               if (contenido !== undefined) {
                   if (typeof contenido === 'string') {
                       ws.cell(numfila, index + 1).string(contenido).style(estilocontenido);
                   } else if (typeof contenido === 'number') {
                       ws.cell(numfila, index + 1).number(contenido).style(estilocontenido);
                   }
               }
           });
           numfila++;
       });

       // Ruta
       const nombreArchivo = "Meow Products Report";
       const pathExcel = path.join(__dirname, 'excel', `${nombreArchivo}.xlsx`);

       // Guardar y descargar
       wb.write(pathExcel, function(err, stats) {
           if (err) {
               console.log(err);
               res.status(500).send("Error al guardar el archivo.");
           } else {
               console.log("Archivo descargado exitoso");
               res.download(pathExcel);
           }
       });

       const correosResult = await sql.getdata_correos_reporte('6');
       correosResult.forEach(renglonactual => {
           setTimeout(() => {
               enviarMailreportemeow(renglonactual.correos);
           }, 1000);
       });

   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
});
enviarMailreportemeow=async(correos)=>{
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
      //to:'programacion@zayro.com',
      to:correos,
      subject:'Meow Products Report',
      attachments:[
         {filename:'Meow Products Report.xlsx',
         path:'./src/excel/Meow Products Report.xlsx'}],
      text:'Through this channel we allow ourselves to send you this report. Sincerely: Zayro International. ',
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos); 
}   
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Proceso para semaforo
app.get('/api/getdata_semaforo',function(req,res,next){
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
         sql.getdata_SemaforoReporte(ren.id).then((result)=> {setTimeout(()=>{
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
               //console.log(reglonactual.Dias);
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
               if(reglonactual.Dias==0){
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
            setTimeout(()=>{enviarMailsemaforo(ren.email,nombreArchivo,ren.nombre,config)},1000)
         });
      },30000)

      })
   })   
   res.json("Archivos generados") 
   //res.json("Archivos generados")
});
enviarMailsemaforo=async(correo, nombreArchivo,nombre,config)=>{
   let transport = nodemailer.createTransport(config);
   const mensaje ={
      from:'it.sistemas@zayro.com',
      to: correo+',avazquez@zayro.com,programacion@zayro.com',
      //to:'programacion@zayro.com',
      subject:'Reporte: '+nombre,
      attachments:[
         {filename:nombreArchivo+' '+nombre+'.xlsx',
         path:'./src/excel/'+nombreArchivo+' '+nombre+'.xlsx'}],
         text:'Por medio de este conducto nos permitimos enviarles este reporte. Atentamente: Zamudio y Rodríguez. ',
      }
      //console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
      transport.sendMail(mensaje,(error, info) => {
         if (error) {
           console.error('Error al enviar el correo:', error);
         } else {
           console.log('Correo enviado:', info.response);
         }
         
         // Cierra el transporte después de enviar el correo
         transport.close()

      });
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Proceso Mercancias en Bodega
//Ya se manda el correo cambiar los detinatarios en la funcion del correo
app.get('/api/getdata_reportemercanciasenbodega',function(req,res,next){
   
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
      pool: true,
   } 


 

   sql.getdata_listaclientes().then((result)=>{
      //console.log(result)

      
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
       ws.cell(1,23).string("Link de consulta                                          ").style(estiloTitulo);
      sql.getdata_mercanciasenbodega(renglonactual.cliente_id).then((result1)=>{
            let numfila=2;
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
               if (reglonactual1.Referencia==''){
                  ws.cell(numfila,23).string("").style(estilocontenido);//A
               }else{
                  ws.cell(numfila,23).link("https://slamm3.zayro.com/SLAMDIGITAL4/VISORREFGEN.ASPX?ref="+reglonactual1.Referencia).style(estilocontenido);//A
               }
               numfila=numfila+1;
             
            }))//Fin renglon actual 1
            const columnWidths = [];
            for (let col = 1; col <= 23; col++) { // Asumiendo que tienes 23 columnas
               let maxLength = 0;
               for (let row = 1; row <= numfila; row++) { // Asumiendo que numfila es el número total de filas
                  const cell = ws.cell(row, col);
                  const cellLength = cell ? cell.toString().length : 0;
                  maxLength = Math.max(maxLength, cellLength);
               }
               columnWidths.push(maxLength);
            }

            // Ajustar el ancho de las columnas
            columnWidths.forEach((width, colIndex) => {
               ws.column(colIndex + 1).setWidth(width * 1.2); // Ajusta el ancho según necesites
            });
     
            const pathExcel=path.join(__dirname,'excel',nombreArchivo+' '+renglonactual.numero+'.xlsx');
            //Guardar
            wb.write(pathExcel,function(err,stats){
               if(err) console.log(err);
               else{
               

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
               //console.log(correoscliente,correosclientecc)
               //Despues de guardar tiene que mandar el correo
               setTimeout(()=>{enviarMailreportebodega(correoscliente,correosclientecc,nombreArchivo,renglonactual.nomcli,renglonactual.numero,config)},3000);
            }); 
            
            
         });

        
         
      },5000)}))
      
      

   });

   res.json("Archivos generados");
});

//Los correos a quienes va se tienen que configurar
enviarMailreportebodega=async(correos, correoscc,nombreArchivo,nomcli,numero,config)=>{
   let transport = nodemailer.createTransport(config);
   const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));
   const mensaje ={
      from:'it.sistemas@zayro.com',
      to: correos+', '+correoscc+',programacion@zayro.com,gerenciati@zayro.com',
      //to:'programacion@zayro.com',
      subject:'Mercancia en bodega: '+nomcli+' - '+numero,
      attachments:[
         {filename:nombreArchivo+' '+numero+'.xlsx',
         path:'./src/excel/'+nombreArchivo+' '+numero+'.xlsx'}],
         text:'Por medio de este conducto nos permitimos enviarles este reporte. Atentamente: Zamudio y Rodríguez. ',
      }
      //console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      try{
         await delay(5000)
         await transport.sendMail(mensaje,(error, info) => {
            if (error) {
              console.error('Error al enviar el correo:', error);
            } else {
              console.log('Correo enviado:', info.response);
            }
            // Cierra el transporte después de enviar el correo
            transport.close()
         }
         );
      }catch(error){
         console.error('Error al enviar el correo:', error.message);
         let intentos=3;
         while (intentos>0){
         await delay(5000)
         try{
            
            await transport.sendMail(mensaje,(error, info) => {
               if (error) {
                 console.error('Error al enviar el correo:', error);
               } else {
                 console.log('Correo enviado:', info.response);
               }
               // Cierra el transporte después de enviar el correo
               transport.close()
            });
            break;
         }catch (retryError){
            console.error('Error al enviar el correo:', error.message);
            await delay(5000); // Pausa de 5 segundos antes de reintentar
            intentos--; // Reintenta el mismo correo
         }

      }
      if (intentos === 0) {
         console.error('Se agotaron los reintentos. No se pudo enviar el correo:', correo);
       }
      }
      
} 

/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//Reporte ASN
app.get('/api/getdata_reporteASN',async function (req,res,next){
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
   resultados=await sql.getdata_ReporteASN();
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
      let ws=wb.addWorksheet("TargetsUltimos10Dias");
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
      let wscan=wb.addWorksheet("Targets Reporte");
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
      resultados.forEach(ren => {
         estilocontenido = ren.fecha_scan === '' ? estilocontenidoletraroja : estilocontenidonormal;

         switch(ren.tipo) {
            case 1:
               variablews = ws;
               numfila = numfilabod;
               break;
            case 2:
               variablews = wscan;
               numfila = numfilacan;
               break;
         }

         // Rellenar las celdas
         variablews.cell(numfila, 1).string(ren.fecha || "").style(estilocontenido);
         variablews.cell(numfila, 2).string(ren.po_lcn || "").style(estilocontenido);
         variablews.cell(numfila, 3).string(ren.pod_lcn || "").style(estilocontenido);
         variablews.cell(numfila, 4).string(ren.po || "").style(estilocontenido);
         variablews.cell(numfila, 5).string(ren.box_id || "").style(estilocontenido);
         variablews.cell(numfila, 6).string(ren.numparte || "").style(estilocontenido);
         variablews.cell(numfila, 7).number(ren.piezas_longbox || 0).style(estilocontenido);
         variablews.cell(numfila, 8).number(ren.piezas_po || 0).style(estilocontenido);
         variablews.cell(numfila, 9).string(ren.fecha_aprox_laredo || "").style(estilocontenido);
         variablews.cell(numfila, 10).string(ren.proveedor || "").style(estilocontenido);
         variablews.cell(numfila, 11).string(ren.fecha_scan || "").style(estilocontenido);
         variablews.cell(numfila, 12).string(ren.pallet || "").style(estilocontenido);
         variablews.cell(numfila, 13).string(ren.escaneado || "").style(estilocontenido);
         variablews.cell(numfila, 14).string(ren.fecha_cierre || "").style(estilocontenido);
         variablews.cell(numfila, 15).string(ren.hb203 || "").style(estilocontenido);

         switch(ren.tipo) {
            case 1:
               numfilabod++;
               break;
            case 2:
               numfilacan++;
               break;
         }
      });
   
      const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');
         wb.write(pathExcel, async function(err, stats) {
            if(err) {
               console.log(err);
               return res.status(500).json("Error al generar el archivo");
            } else {
               res.json("Archivo generado");

               let correos = await sql.getdata_correos_reporte('3');
               correos.forEach(renglonactual => {
                  enviarMailASN(nombreArchivo, transport, renglonactual.correos);
               });
            }
         });

});
enviarMailASN=async(nombreArchivo,transport,correos)=>{
   const mensaje ={
      from:'it.sistemas@zayro.com',
      //to: 'lzamudio@zayro.com,soportetecnico@zayro.com,sistemas@zayro.com',
      to:correos,
      //to: 'programacion@zayro.com',
      subject:'Reporte ASN',
      attachments:[
         {filename:nombreArchivo+'.xlsx',
         path:'./src/excel/'+nombreArchivo+'.xlsx'}],
         text:'Reporte ASN ',
      }
     // console.log(mensaje)
      //const transport = nodemailer.createTransport(config);
      transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
     //const info=await transport.sendMail(mensaje);
     transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   });
      //console.log(info);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//REPORTE Thyssenkrupp 
app.get('/api/getdata_Thyssenkrupp/:fechaini/:fechafin', async function(req,res,next){
   
   var fechaini=req.params.fechaini;
   var fechafin=req.params.fechafin;

   var wb= new xl.Workbook();
   
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
  //const pgArr=await pgconect.getReporteThyssen(595);
   const zayArr=await pg.getReporteThyssenDolaresHoy();
   console.log(zayArr);
   


      //res.json(result);
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
      const asText = v => (v == null ? '' : String(v));
const asNum  = v => {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return v;
  if (typeof v === 'string') {
    // quita $, comas, espacios, etc. (ajusta si usas coma decimal)
    const n = Number(v.replace(/[^0-9.-]/g, ''));
    return Number.isFinite(n) ? n : 0;
  }
  const n = Number(v);
  return Number.isFinite(n) ? n : 0;
};

// (opcional) estilo de moneda
const estiloMoneda = wb.createStyle({ numberFormat: '#,##0.00' });

zayArr.forEach(reglonactual => {
  wsUSD.cell(numfila,1).string(asText(reglonactual.NoProveedor)).style(estilocontenido);
  wsUSD.cell(numfila,2).string(asText(reglonactual.RazonSocial)).style(estilocontenido);
  wsUSD.cell(numfila,3).string(asText(reglonactual.NoFactura)).style(estilocontenido);
  wsUSD.cell(numfila,4).string(asText(reglonactual.Fecha)).style(estilocontenido);
  wsUSD.cell(numfila,5).number(asNum(reglonactual.Credito)).style(estilocontenido);
  wsUSD.cell(numfila,6).string(asText(reglonactual.Vencimiento)).style(estilocontenido);
  wsUSD.cell(numfila,7).string(asText(reglonactual.IMPEXP)).style(estilocontenido);

  wsUSD.cell(numfila,8).string(asText(reglonactual.PO)).style(estilocontenido);

  wsUSD.cell(numfila,9).string(asText(reglonactual.CuentaContable)).style(estilocontenido);

  // OJO: usa 'Subtotal' (no 'SubTotal')
  wsUSD.cell(numfila,10).number(asNum(reglonactual.Subtotal)).style(estiloMoneda || estilocontenido);
  wsUSD.cell(numfila,11).number(asNum(reglonactual.IVA)).style(estiloMoneda || estilocontenido);
  wsUSD.cell(numfila,12).number(asNum(reglonactual.Retencion)).style(estiloMoneda || estilocontenido);
  wsUSD.cell(numfila,13).number(asNum(reglonactual.Total)).style(estiloMoneda || estilocontenido);

  total += asNum(reglonactual.Total); // evita concatenación de strings

  wsUSD.cell(numfila,14).string(asText(reglonactual.Moneda)).style(estilocontenido);
  wsUSD.cell(numfila,15).string(asText(reglonactual.Comentarios || '')).style(estilocontenido);

  numfila += 1;
});
      //-------------------------------------------------------------------------------------
      //console.log(pgArr)
   /*  pgArr.forEach(reglon => {
      wsUSD.cell(numfila, 1).string(reglon.noproveedor || '').style(estilocontenido);
      wsUSD.cell(numfila, 2).string(reglon.razonsocial || '').style(estilocontenido);
      wsUSD.cell(numfila, 3).string(reglon.nofactura || '').style(estilocontenido);
      wsUSD.cell(numfila, 4).string(reglon.fecha || '').style(estilocontenido);

      // Crédito viene vacío ('') en tu JSON, así que lo dejamos como texto
      wsUSD.cell(numfila, 5)
            .string(reglon.credito || '')
            .style(estilocontenido);

      wsUSD.cell(numfila, 6).string(reglon.vencimiento || '').style(estilocontenido);
      wsUSD.cell(numfila, 7).string(reglon.impexp   || '').style(estilocontenido);
      wsUSD.cell(numfila, 8).string(reglon.po       || '').style(estilocontenido);
      wsUSD.cell(numfila, 9).string(reglon.cuentacontable || '').style(estilocontenido);

      // Para los campos numéricos, parseamos y comprobamos
      const subtotal = parseFloat(reglon.subtotal);
      if (!isNaN(subtotal)) {
         wsUSD.cell(numfila, 10).number(subtotal).style(estilocontenido);
      } else {
         wsUSD.cell(numfila,10).number(0).style(estilocontenido);
      }

      const iva = parseFloat(reglon.iva);
      wsUSD.cell(numfila, 11)
            .number(!isNaN(iva) ? iva : 0)
            .style(estilocontenido);

      const ret = parseFloat(reglon.retencion);
      wsUSD.cell(numfila, 12)
            .number(!isNaN(ret) ? ret : 0)
            .style(estilocontenido);

      const tot = parseFloat(reglon.total);
      if (!isNaN(tot)) {
         wsUSD.cell(numfila, 13).number(tot).style(estilocontenido);
         total += tot;
      } else {
         wsUSD.cell(numfila,13).number(0).style(estilocontenido);
      }

      wsUSD.cell(numfila, 14).string(reglon.moneda || '').style(estilocontenido);
      wsUSD.cell(numfila, 15).string(reglon.comentarios || '').style(estilocontenido);

      numfila++;
      });

*/

      wsUSD.cell(1,13).string("Total").style(estilototal);
      wsUSD.cell(2,13).number(total).style(estilototal);
      wsUSD.cell(numfila,12).string("Total").style(estilototal);
      wsUSD.cell(numfila,13).number(total).style(estilototal);
   /*pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
      }
   });*/

   const resultadopesos=await sqlzam.getdata_ReporteThyssenhrup_pesos(fechafin);
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
      total=0;
      resultadopesos.forEach(reglonactual => {
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
       /*
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   wb.write(pathExcel,function(err,stats){
      if(err) console.log(err);
      else{
      }
   });*/
   const pendientes=await sql.getdata_Thyssenkrupp_pendientes(fechaini,fechafin);
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
      numfila=5;
      pendientes.forEach(reglonactual => {
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
      let nombreArchivo="Estado de cuenta Thyssenkrupp";

      const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
      //Guardar
      await wb.write(pathExcel,function(err,stats){
         if(err) console.log(err);
         else{
            //res.json("Archivo generado")
            console.log("Archivo generado")
         }
      });
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
   
   const resp= await enviarMailEstadoCuentaThynss(nombreArchivo,transport,'');
   //const resp='enviado'
   if (resp=='enviado'){
         res.json('enviado')
   }else{
         res.json('error')
   }

   
});
enviarMailEstadoCuentaThynss=async(nombreArchivo,transport,correos)=>{
   const mensaje ={
      from:'sistemas@zayro.com',
      //to: correos,
      to:'cobranza@zayro.com',
      cc:'contador@zayro.com;hvargas@zayro.com;programacion@zayro.com',
      subject:'Estado de cuenta Thyssenkrupp',
      attachments:[ 
         {filename:nombreArchivo+'.xlsx',
         path:'./src/excel/'+nombreArchivo+'.xlsx'}],
         text:'Estado de cuenta Thyssenkrupp',
      }
      //console.log(mensaje)
      //console.log(transport)
      //const transport = nodemailer.createTransport(config);
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
         return 'enviado'
      }

      transport.close()
   });
      //console.log(info);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//REPORTE KAWASSAKI
app.get('/api/getdata_kawassaki',async function(req,res,next){
      //res.json(result);
      //console.log(result) 
      var wb= new xl.Workbook();
      let nombreArchivo="Kawasaki Report";
      var ws=wb.addWorksheet("KawasakiReport");
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
      ws.cell(1,1).string("CREACION").style(estiloTitulo);
      ws.cell(1,2).string("PO LCN").style(estiloTitulo);
      ws.cell(1,3).string("POD LCN").style(estiloTitulo);
      ws.cell(1,4).string("LONG BOX ID").style(estiloTitulo);
      ws.cell(1,5).string("MASTER LABEL").style(estiloTitulo);
      ws.cell(1,6).string("#PARTE").style(estiloTitulo);
      ws.cell(1,7).string("PIEZAS LONG BOX").style(estiloTitulo);
      ws.cell(1,8).string("PIEZAS PO").style(estiloTitulo);
      ws.cell(1,9).string("FECHA APROX LAREDO").style(estiloTitulo);
      ws.cell(1,10).string("PROVEEDOR").style(estiloTitulo);
      ws.cell(1,11).string("FECHA ESCANEO").style(estiloTitulo);
      ws.cell(1,12).string("NOMBRE").style(estiloTitulo);
      let numfila=2;
      const result =await sql.getdata_reporte_kawassaki();

      result.forEach(reglonactual => {
         if (reglonactual.Creacion==''){
            ws.cell(numfila,1).string("").style(estilocontenido);//A
         }else{
            ws.cell(numfila,1).string(reglonactual.Creacion).style(estilocontenido);//A
         }
         if (reglonactual.PO_LCN==""){
            ws.cell(numfila,2).string("").style(estilocontenido);//B
         }else{
            ws.cell(numfila,2).string(reglonactual.PO_LCN).style(estilocontenido);//B
         }
         if(reglonactual.POD_LCN==''){
            ws.cell(numfila,3).string("").style(estilocontenido);//C
         }else{
            ws.cell(numfila,3).string(reglonactual.POD_LCN).style(estilocontenido);//C
         }
         if(reglonactual.Long_Box_ID==''){
            ws.cell(numfila,4).string("").style(estilocontenido);//D
         }else{
            ws.cell(numfila,4).string(reglonactual.Long_Box_ID).style(estilocontenido);//D
         }
         if (reglonactual.MasterLabel==''){
            ws.cell(numfila,5).string("").style(estilocontenido);//E
         }else{
            ws.cell(numfila,5).string(reglonactual.MasterLabel).style(estilocontenido);//E
         }
         if (reglonactual.NumParte==''){
            ws.cell(numfila,6).string("").style(estilocontenido);//F
         }else{
            ws.cell(numfila,6).string(reglonactual.NumParte).style(estilocontenido);//F
         }
         if(reglonactual.Piezas_Long_Box==''){
            ws.cell(numfila,7).number(0).style(estilocontenido);//G
         }
         else{
            ws.cell(numfila,7).number(reglonactual.Piezas_Long_Box).style(estilocontenido);//G
         }
         if (reglonactual.Piezas_PO==""){
            ws.cell(numfila,8).number(0).style(estilocontenido);//H
         }
         else{
            ws.cell(numfila,8).number(reglonactual.Piezas_PO).style(estilocontenido);//H
         }
         if(reglonactual.Fecha_Aprox_Laredo==''){
            ws.cell(numfila,9).string("").style(estilocontenido);//I
         }else{
            ws.cell(numfila,9).string(reglonactual.Fecha_Aprox_Laredo).style(estilocontenido);//I
         }
         if(reglonactual.Proveedor==''){
            ws.cell(numfila,10).string(reglonactual.Proveedor).style(estilocontenido);//J
         }else{
            ws.cell(numfila,10).string(reglonactual.Proveedor).style(estilocontenido);//J
         }
         if(reglonactual.FechaEscaneo==''){
            ws.cell(numfila,11).string("").style(estilocontenido);//K
         }else{
            ws.cell(numfila,11).string(reglonactual.FechaEscaneo).style(estilocontenido);//K
         }
         if (reglonactual.Nombre==''){
            ws.cell(numfila,12).string("").style(estilocontenido);//L
         }else{
            ws.cell(numfila,12).string(reglonactual.Nombre).style(estilocontenido);//L
         }
         numfila=numfila+1;

       //Ruta
   })
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   await wb.write(pathExcel,function(err,stats){
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
   
  
   await sql.getdata_correos_reporte('5').then((result)=>{
      result.forEach(renglonactual=>{
         enviarMailreportekawassaki(renglonactual.correos);
      })
      })
   //enviarMailreportedistribucion()
});
enviarMailreportekawassaki=async(correos)=>{
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
      //to:'gerenciati@zayro.com',
      to:correos,
      subject:'KAWASAKI  INVENTORY REPORT',
      attachments:[
         {filename:'Kawasaki Report.xlsx',
         path:'./src/excel/Kawasaki Report.xlsx'}],
      text:'Through this channel we allow ourselves to send you this report. Sincerely: Zayro International. ',
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   });
   //console.log(correos); 
}   
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
//HB203
const documentsFolder = '/slamsuite/tools/ZyroK/HB203/BackUp'
app.get('/api/getdata_hb203',async function(req,res,next){
   const hb101=await sql.getdata_hb101();
   const hb102=await sql.getdata_hb102();
   const hb103=await sql.getdata_hb103();
   var wb= new xl.Workbook();
   let nombreArchivo="Catalogo HB1";
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
      alignment: {
         horizontal: 'center', // Centrar horizontalmente
         vertical: 'center',   // Centrar verticalmente
     }
   });
   var estilocontenido=wb.createStyle({
      font:{
         name: 'Arial',
         color: '#000000',
         size:10,
      },
      alignment: {
         horizontal: 'center', // Centrar horizontalmente
         vertical: 'center',   // Centrar verticalmente
     }
   });
   var ws=wb.addWorksheet("HB101");
   ws.cell(1,1).string("FECHA").style(estiloTitulo);
   ws.cell(1,2).string("HORA").style(estiloTitulo);
   ws.cell(1,3).string("DESTINO CODIGO").style(estiloTitulo);
   ws.cell(1,4).string("ORIGEN CODIGO").style(estiloTitulo);
   ws.cell(1,5).string("NUM. PARTE").style(estiloTitulo);
   ws.cell(1,6).string("DESCRIPCION").style(estiloTitulo);
   ws.cell(1,7).string("VALOR IM").style(estiloTitulo);
   ws.cell(1,8).string("MONEDA IM").style(estiloTitulo);
   ws.cell(1,9).string("TERMINO IM").style(estiloTitulo);
   ws.cell(1,10).string("VALOR EX").style(estiloTitulo);
   ws.cell(1,11).string("MONEDA EX").style(estiloTitulo);
   ws.cell(1,12).string("TERMINO EX").style(estiloTitulo);
   ws.cell(1,13).string("PAIS ORIGEN").style(estiloTitulo);
   ws.cell(1,14).string("HS IM").style(estiloTitulo);
   ws.cell(1,15).string("HS EX").style(estiloTitulo);
   ws.cell(1,16).string("ES IMEX").style(estiloTitulo);
   ws.cell(1,17).string("ESPROSEC").style(estiloTitulo);
   ws.cell(1,18).string("PROGRAMA MX").style(estiloTitulo);
   ws.cell(1,19).string("CODE LINEA EXP").style(estiloTitulo);
   ws.cell(1,20).string("ECC").style(estiloTitulo);
   ws.cell(1,21).string("CODIGO HAZMAT").style(estiloTitulo);
   ws.cell(1,22).string("UNIDAD PESO").style(estiloTitulo);
   const NULL="NULL"
   var numfila=2;
   hb101.forEach(reglonactual => {
      if (reglonactual.Fecha==''){
         ws.cell(numfila,1).string(NULL).style(estilocontenido);//A
      }else{
         ws.cell(numfila,1).string(reglonactual.Fecha).style(estilocontenido);//A
      }
      if (reglonactual.Hora==''){
         ws.cell(numfila,2).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,2).string(reglonactual.Hora).style(estilocontenido);//B
      }
      if (reglonactual.DestinoCodigo==''){
         ws.cell(numfila,3).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,3).string(reglonactual.DestinoCodigo).style(estilocontenido);//B
      }
      if (reglonactual.OrigenCodigo==''){
         ws.cell(numfila,4).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,4).string(reglonactual.OrigenCodigo).style(estilocontenido);//B
      }
      if (reglonactual.NumParte==''){
         ws.cell(numfila,5).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,5).string(reglonactual.NumParte).style(estilocontenido);//B
      }
      if (reglonactual.Descripcion==''){
         ws.cell(numfila,6).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,6).string(reglonactual.Descripcion).style(estilocontenido);//B
      }
      if (reglonactual.ValorIM==''){
         ws.cell(numfila,7).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,7).number(reglonactual.ValorIM).style(estilocontenido);//B
      }
      if (reglonactual.MonedaIM==''){
         ws.cell(numfila,8).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,8).string(reglonactual.MonedaIM).style(estilocontenido);//B
      }
      if (reglonactual.TerminoIM==''){
         ws.cell(numfila,9).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,9).string(reglonactual.TerminoIM).style(estilocontenido);//B
      }
      if (reglonactual.ValorEX==''){
         ws.cell(numfila,10).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,10).number(reglonactual.ValorEX).style(estilocontenido);//B
      }
      if (reglonactual.MonedaEX==''){
         ws.cell(numfila,11).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,11).string(reglonactual.MonedaEX).style(estilocontenido);//B
      }
      if (reglonactual.TerminoEX==''){
         ws.cell(numfila,12).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,12).string(reglonactual.TerminoEX).style(estilocontenido);//B
      }
      if (reglonactual.PaisOrigen==''){
         ws.cell(numfila,13).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,13).string(reglonactual.PaisOrigen).style(estilocontenido);//B
      }
      if (reglonactual.HsIM==''){
         ws.cell(numfila,14).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,14).string(reglonactual.HsIM).style(estilocontenido);//B
      }
      if (reglonactual.HsEx==''){
         ws.cell(numfila,15).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,15).string(reglonactual.HsEx).style(estilocontenido);//B
      }
      if (reglonactual.EsIMEX==''){
         ws.cell(numfila,16).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,16).number(reglonactual.EsIMEX).style(estilocontenido);//B
      }
      if (reglonactual.EsProsec==''){
         ws.cell(numfila,17).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,17).number(reglonactual.EsProsec).style(estilocontenido);//B
      }
      if (reglonactual.ProgramaMx==''){
         ws.cell(numfila,18).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,18).string(reglonactual.ProgramaMx).style(estilocontenido);//B
      }
      if (reglonactual.CodeLineaExp==''){
         ws.cell(numfila,19).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,19).string(reglonactual.CodeLineaExp).style(estilocontenido);//B
      }
      if (reglonactual.ECC==''){
         ws.cell(numfila,20).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,20).string(reglonactual.ECC).style(estilocontenido);//B
      }
      if (reglonactual.CodigoHazmat==''){
         ws.cell(numfila,21).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,21).string(reglonactual.CodigoHazmat).style(estilocontenido);//B
      }
      if (reglonactual.Unidad_Peso==''){
         ws.cell(numfila,22).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,22).number(reglonactual.Unidad_Peso).style(estilocontenido);//B
      } 
      numfila=numfila+1;
   })

   var ws=wb.addWorksheet("HB102");
   ws.cell(1,1).string("FECHA").style(estiloTitulo);
   ws.cell(1,2).string("HORA").style(estiloTitulo);
   ws.cell(1,3).string("DESTINO CODIGO").style(estiloTitulo);
   ws.cell(1,4).string("ORIGEN CODIGO").style(estiloTitulo);
   ws.cell(1,5).string("LCN PO NO").style(estiloTitulo);
   ws.cell(1,6).string("LCN PO DETAIL NO").style(estiloTitulo);
   ws.cell(1,7).string("KMC PO NO").style(estiloTitulo);
   ws.cell(1,8).string("KMC PO DETAL NO").style(estiloTitulo);
   ws.cell(1,9).string("KMX PO NO").style(estiloTitulo);
   ws.cell(1,10).string("KMX PO DETAIL NO").style(estiloTitulo);
   ws.cell(1,11).string("NUM PARTE").style(estiloTitulo);
   ws.cell(1,12).string("PIEZAS").style(estiloTitulo);
   ws.cell(1,13).string("FECHA APROX LAREDO").style(estiloTitulo);
   ws.cell(1,14).string("FECHA APROX KMX").style(estiloTitulo);
   ws.cell(1,15).string("RAZON PO").style(estiloTitulo);
   ws.cell(1,16).string("SHIP CODIGO").style(estiloTitulo);
   ws.cell(1,17).string("CAT PRODUCCION").style(estiloTitulo);
   numfila=2;
   hb102.forEach(reglonactual1 => {
      if (reglonactual1.Fecha==''){
         ws.cell(numfila,1).string(NULL).style(estilocontenido);//A
      }else{
         ws.cell(numfila,1).string(reglonactual1.Fecha).style(estilocontenido);//A
      }
      if (reglonactual1.Hora==''){
         ws.cell(numfila,2).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,2).string(reglonactual1.Hora).style(estilocontenido);//B
      }
      if (reglonactual1.DestinoCodigo==''){
         ws.cell(numfila,3).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,3).string(reglonactual1.DestinoCodigo).style(estilocontenido);//B
      }
      if (reglonactual1.OrigenCodigo==''){
         ws.cell(numfila,4).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,4).string(reglonactual1.OrigenCodigo).style(estilocontenido);//B
      }
      if (reglonactual1.LCN_PO_NO==''){
         ws.cell(numfila,5).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,5).string(reglonactual1.LCN_PO_NO).style(estilocontenido);//B
      }
      if (reglonactual1.LCN_PO_DETAIL_NO==''){
         ws.cell(numfila,6).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,6).string(reglonactual1.LCN_PO_DETAIL_NO).style(estilocontenido);//B
      }
      if (reglonactual1.KMC_PO_NO==''){
         ws.cell(numfila,7).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,7).string(reglonactual1.KMC_PO_NO).style(estilocontenido);//B
      }
      if (reglonactual1.KMC_PO_DETAL_NO==''){
         ws.cell(numfila,8).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,8).string(reglonactual1.KMC_PO_DETAL_NO).style(estilocontenido);//B
      }
      if (reglonactual1.KMX_PO_NO==''){
         ws.cell(numfila,9).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,9).string(reglonactual1.KMX_PO_NO).style(estilocontenido);//B
      }
      if (reglonactual1.KMX_PO_DETAIL_NO==''){
         ws.cell(numfila,10).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,10).string(reglonactual1.KMX_PO_DETAIL_NO).style(estilocontenido);//B
      }
      if (reglonactual1.NumParte==''){
         ws.cell(numfila,11).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,11).string(reglonactual1.NumParte).style(estilocontenido);//B
      }
      if (reglonactual1.Piezas==''){
         ws.cell(numfila,12).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,12).number(reglonactual1.Piezas).style(estilocontenido);//B
      }
      if (reglonactual1.Fecha_Aprox_Laredo==''){
         ws.cell(numfila,13).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,13).string(reglonactual1.Fecha_Aprox_Laredo).style(estilocontenido);//B
      }
      if (reglonactual1.Fecha_Aprox_KMX==''){
         ws.cell(numfila,14).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,14).string(reglonactual1.Fecha_Aprox_KMX).style(estilocontenido);//B
      }
      if (reglonactual1.RazonPO==''){
         ws.cell(numfila,15).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,15).string(reglonactual1.RazonPO).style(estilocontenido);//B
      }
      if (reglonactual1.Ship_Codigo==''){
         ws.cell(numfila,16).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,16).string(reglonactual1.Ship_Codigo).style(estilocontenido);//B
      }
      if (reglonactual1.Cat_Produccion==''){
         ws.cell(numfila,17).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,17).string(reglonactual1.Cat_Produccion).style(estilocontenido);//B
      }
      numfila=numfila+1;
   })

   var ws=wb.addWorksheet("HB103");
   ws.cell(1,1).string("FECHA").style(estiloTitulo);
   ws.cell(1,2).string("HORA").style(estiloTitulo);
   ws.cell(1,3).string("DESTINO CODIGO").style(estiloTitulo);
   ws.cell(1,4).string("ORIGEN CODIGO").style(estiloTitulo);
   ws.cell(1,5).string("BOX ID").style(estiloTitulo);
   ws.cell(1,6).string("PO LCN").style(estiloTitulo);
   ws.cell(1,7).string("POD LCN").style(estiloTitulo);
   ws.cell(1,8).string("NUM PARTE").style(estiloTitulo);
   ws.cell(1,9).string("FECHA APROX LAREDO").style(estiloTitulo);
   ws.cell(1,10).string("PIEZAS LONG BOX").style(estiloTitulo);
   ws.cell(1,11).string("PIEZAS PO").style(estiloTitulo);
   ws.cell(1,12).string("CAT PRODUCCION").style(estiloTitulo);
   ws.cell(1,13).string("PROVEEDOR").style(estiloTitulo);
   numfila=2;
   hb103.forEach(reglonactual2 => {
      if (reglonactual2.Fecha==''){
         ws.cell(numfila,1).string(NULL).style(estilocontenido);//A
      }else{
         ws.cell(numfila,1).string(reglonactual2.Fecha).style(estilocontenido);//A
      }
      if (reglonactual2.Hora==''){
         ws.cell(numfila,2).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,2).string(reglonactual2.Hora).style(estilocontenido);//B
      }
      if (reglonactual2.DestinoCodigo==''){
         ws.cell(numfila,3).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,3).string(reglonactual2.DestinoCodigo).style(estilocontenido);//B
      }
      if (reglonactual2.OrigenCodigo==''){
         ws.cell(numfila,4).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,4).string(reglonactual2.OrigenCodigo).style(estilocontenido);//B
      }
      if (reglonactual2.Box_ID==''){
         ws.cell(numfila,5).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,5).string(reglonactual2.Box_ID).style(estilocontenido);//B
      }
      if (reglonactual2.PO_LCN==''){
         ws.cell(numfila,6).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,6).string(reglonactual2.PO_LCN).style(estilocontenido);//B
      }
      if (reglonactual2.POD_LCN==''){
         ws.cell(numfila,7).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,7).string(reglonactual2.POD_LCN).style(estilocontenido);//B
      }
      if (reglonactual2.NumParte==''){
         ws.cell(numfila,8).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,8).string(reglonactual2.NumParte).style(estilocontenido);//B
      }
      if (reglonactual2.Fecha_Aprox_Laredo==''){
         ws.cell(numfila,9).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,9).string(reglonactual2.Fecha_Aprox_Laredo).style(estilocontenido);//B
      }
      if (reglonactual2.Piezas_LongBox==''){
         ws.cell(numfila,10).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,10).number(reglonactual2.Piezas_LongBox).style(estilocontenido);//B
      }
      if (reglonactual2.Piezas_PO==''){
         ws.cell(numfila,11).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,11).number(reglonactual2.Piezas_PO).style(estilocontenido);//B
      }
      if (reglonactual2.Cat_Produccion==''){
         ws.cell(numfila,12).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,12).string(reglonactual2.Cat_Produccion).style(estilocontenido);//B
      }
      if (reglonactual2.Proveedor==''){
         ws.cell(numfila,13).string(NULL).style(estilocontenido);//B
      }else{
         ws.cell(numfila,13).string(reglonactual2.Proveedor).style(estilocontenido);//B
      }
      numfila=numfila+1;
   })
   
       //Ruta
   const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
   //Guardar
   await wb.write(pathExcel, (err, stats) => {
      if (err) {
          console.error('Error al guardar el archivo de Excel:', err);
          
      } else {
          console.log('Archivo de Excel guardado exitosamente en:', pathExcel);
          // Descargar el archivo de Excel
          res.download(pathExcel, nombreArchivo+'.xlsx', (err) => {
              if (err) {
                  console.error('Error al descargar el archivo:', err);
                  // Manejar el error
              } else {
                  console.log('Archivo descargado exitoso');
              }
          });
      }
  });
   const correos=await sql.getdata_correos_reporte('7')
 
   correos.forEach(async renglonactual3=>{
      //enviarMailHB1(renglonactual.correos);
      getFilesFromPreviousDay(async (filesFromPreviousDay) => {
         await enviarMailHB1(renglonactual3.correos, filesFromPreviousDay.map(file => path.join(documentsFolder, file)));
      });
   })
})
function getFilesFromPreviousDay(callback) {
   const today = new Date();
   const dayOfWeek = today.getDay(); // Obtener el día de la semana (0 para domingo, 1 para lunes, ..., 6 para sábado)
   let previousDay = new Date(today); // Por defecto, tomará los archivos del día anterior

   if (dayOfWeek === 1) {
       // Si es lunes, retroceder 3 días para obtener archivos del viernes
       previousDay.setDate(today.getDate() - 3);

       // Obtener archivos del sábado
       getFilesFromSaturday(saturdayFiles => {
           // Obtener archivos del domingo
           getFilesFromSunday(sundayFiles => {
               // Leer archivos del día anterior
               fs.readdir(documentsFolder, (err, files) => {
                   if (err) {
                       console.error('Error al leer la carpeta de documentos:', err);
                       return;
                   }

                   const filesFromPreviousDay = files.filter(file => {
                       const filePath = path.join(documentsFolder, file);
                       const stats = fs.statSync(filePath);
                       const fileDate = new Date(stats.mtime); // Obtener la fecha de modificación del archivo

                       // Verificar si el archivo fue modificado el día anterior
                       const isPreviousDay = fileDate.toDateString() === previousDay.toDateString();
                       // Excluir los archivos que terminan con "-p"
                       const fileName = path.basename(file);
                       const notEndsWithP = !fileName.endsWith('-p');

                       return isPreviousDay && notEndsWithP;
                   });

                   // Agregar archivos del sábado y del domingo a los archivos del día anterior
                   let combinedFiles = [...filesFromPreviousDay, ...saturdayFiles, ...sundayFiles];

                   // Si es después de las 12 PM, agregar archivos del día actual
                   const isAfterNoon = today.getHours() >= 12;
                   if (isAfterNoon) {
                       getFilesFromCurrentDay(currentDayFiles => {
                           // Agregar archivos del día actual
                           combinedFiles = [...combinedFiles, ...currentDayFiles];
                           // Llamar al callback con todos los archivos
                           callback(combinedFiles);
                       });
                   } else {
                       // Llamar al callback con todos los archivos
                       callback(combinedFiles);
                   }
               });
           });
       });
   } else {
       // De lo contrario, retroceder un día para obtener archivos del día anterior
       previousDay.setDate(today.getDate() - 1);

       fs.readdir(documentsFolder, (err, files) => {
           if (err) {
               console.error('Error al leer la carpeta de documentos:', err);
               return;
           }

           const filesFromPreviousDay = files.filter(file => {
               const filePath = path.join(documentsFolder, file);
               const stats = fs.statSync(filePath);
               const fileDate = new Date(stats.mtime); // Obtener la fecha de modificación del archivo

               // Verificar si el archivo fue modificado el día anterior
               const isPreviousDay = fileDate.toDateString() === previousDay.toDateString();
               // Excluir los archivos que terminan con "-p"
               const fileName = path.basename(file);
               const notEndsWithP = !fileName.endsWith('-p');

               return isPreviousDay && notEndsWithP;
           });

           // Si es después de las 12 PM, agregar archivos del día actual
           const isAfterNoon = today.getHours() >= 12;
           if (isAfterNoon) {
               getFilesFromCurrentDay(currentDayFiles => {
                   // Agregar archivos del día actual
                   const combinedFiles = [...filesFromPreviousDay, ...currentDayFiles];
                   // Llamar al callback con todos los archivos
                   callback(combinedFiles);
               });
           } else {
               getFilesFromCurrentDay(currentDayFiles => {
               // Agregar archivos del día actual
               const combinedFiles = [...filesFromPreviousDay, ...currentDayFiles];
               // Llamar al callback con todos los archivos
               callback(combinedFiles);
            });
           }
       });
   }
}

function getFilesFromSunday(callback) {
   const today = new Date();
   const dayOfWeek = today.getDay();

   // Verificar si hoy es lunes (día 1)
   if (dayOfWeek === 1) {
       const previousSunday = new Date(today);

       // Retroceder un día para obtener el domingo anterior
       previousSunday.setDate(today.getDate() - 1);

       fs.readdir(documentsFolder, (err, files) => {
           if (err) {
               console.error('Error al leer la carpeta de documentos:', err);
               return;
           }

           const filesFromSunday = files.filter(file => {
               const filePath = path.join(documentsFolder, file);
               const stats = fs.statSync(filePath);
               const fileDate = new Date(stats.mtime);

               // Verificar si el archivo fue modificado el domingo anterior
               const isSunday = fileDate.getDay() === 0; // 0 para domingo
               const isSameSunday = fileDate.getDate() === previousSunday.getDate() &&
                                   fileDate.getMonth() === previousSunday.getMonth() &&
                                   fileDate.getFullYear() === previousSunday.getFullYear();

               // Excluir los archivos que terminan con "-p"
               const fileName = path.basename(file);
               const notEndsWithP = !fileName.endsWith('-p');

               return isSunday && isSameSunday && notEndsWithP;
           });

           callback(filesFromSunday);
       });
   } else {
       // Si no es lunes, devolver una lista vacía
       callback([]);
   }
}
function getFilesFromSaturday(callback) {
   const today = new Date();
   const dayOfWeek = today.getDay();

   // Verificar si hoy es lunes (día 1)
   if (dayOfWeek === 1) {
       const previousSaturday = new Date(today);

       // Retroceder dos días para obtener el sábado anterior
       previousSaturday.setDate(today.getDate() - 2);


       fs.readdir(documentsFolder, (err, files) => {
           if (err) {
               console.error('Error al leer la carpeta de documentos:', err);
               return;
           }

           const filesFromSaturday = files.filter(file => {
               const filePath = path.join(documentsFolder, file);
               const stats = fs.statSync(filePath);
               const fileDate = new Date(stats.mtime);

               // Verificar si el archivo fue modificado el sábado anterior
               const isSaturday = fileDate.getDay() === 6; // 6 para sábado
               const isSameSaturday = fileDate.getDate() === previousSaturday.getDate() &&
                                      fileDate.getMonth() === previousSaturday.getMonth() &&
                                      fileDate.getFullYear() === previousSaturday.getFullYear();

               // Excluir los archivos que terminan con "-p"
               const fileName = path.basename(file);
               const notEndsWithP = !fileName.endsWith('-p');

               return isSaturday && isSameSaturday && notEndsWithP;
           });

           callback(filesFromSaturday);
       });
   } else {
       // Si no es lunes, devolver una lista vacía
       callback([]);
   }
}
// Función para obtener la fecha de modificación de una carpeta
function getFolderModificationDate(folderPath) {
   try {
       const stats = fs.statSync(folderPath);
       const modificationDate = stats.mtime.toLocaleDateString(); // Devuelve la fecha de modificación en formato de cadena de texto
       return formatDate(modificationDate); // Formatea la fecha
   } catch (error) {
       console.error('Error al obtener la fecha de modificación de la carpeta:', error);
       return null; // Devuelve null en caso de error
   }
}
// Función para dar formato a la fecha "M/DD/YYYY"
function formatDate(dateString) {
   // Dividir la cadena en mes, día y año
   const [month, day, year] = dateString.split('/');

   // Crear una instancia de fecha con el año, mes y día proporcionados
   const currentDate = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));

   // Obtener el nombre abreviado del mes
   const monthAbbreviation = currentDate.toLocaleString('default', { month: 'short' });

   // Formatear la fecha como "Mmm DD"
   const formattedDate = monthAbbreviation + ' ' + parseInt(day);

   return formattedDate;
}
function getFilesFromCurrentDay(callback) {
   const today = new Date();

   fs.readdir(documentsFolder, (err, files) => {
       if (err) {
           console.error('Error al leer la carpeta de documentos:', err);
           return;
       }

       // Filtrar archivos que no terminan con "-p" y cuya fecha de modificación coincide con hoy
       const currentDayFiles = files.filter(file => {
           const filePath = path.join(documentsFolder, file);
           const stats = fs.statSync(filePath);
           const fileDate = new Date(stats.mtime);
           return fileDate.toDateString() === today.toDateString() && !file.endsWith('-p');
       });

       // Llamar al callback con los archivos del día actual
       callback(currentDayFiles);
   });
}
/*getFilesFromPreviousDay(files => {
   console.log('Archivos del día anterior:', files);
});*/
enviarMailHB1=async(correos,filesFromPreviousDay)=>{
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
   const attachments = [];
   const folder101 = '/slamsuite/tools/ZyroK/HB101';
   const folder102 = '/slamsuite/tools/ZyroK/HB102';
   const folder103 = '/slamsuite/tools/ZyroK/HB103';
   const folder104 = '/slamsuite/tools/ZyroK/HB104';
   const folder201 = '/slamsuite/tools/ZyroK/HB201';
   const folder202 = '/slamsuite/tools/ZyroK/HB202';
   const folder203 = '/slamsuite/tools/ZyroK/HB203';
   attachments.push({
      filename: 'Catalogo HB1.xlsx',
      path: './src/excel/Catalogo HB1.xlsx'
    });
    attachments.push(...filesFromPreviousDay.map(file => ({
      filename: path.basename(file),
      path: file
    })));
    const fecha101 = getFolderModificationDate(folder101);
    const fecha102 = getFolderModificationDate(folder102);
    const fecha103 = getFolderModificationDate(folder103);
    const fecha104 = getFolderModificationDate(folder104);
    const fecha201 = getFolderModificationDate(folder201);
    const fecha202 = getFolderModificationDate(folder202);
    const fecha203 = getFolderModificationDate(folder203);
    
    const tablaDatos = `
        <table border="1">
            <tr>
                <th> </th>
                <th>File with the latest date</th>
                <th> </th>
                <th>File with the latest date</th>
            </tr>
            <tr>
                <td>HB101</td>
                <td>${fecha101}</td>
                <td>HB201</td>
                <td>${fecha201}</td>
            </tr>
            <tr>
                <td>HB102</td>
                <td>${fecha102}</td>
                <td>HB202</td>
                <td>${fecha202}</td>
            </tr>
            <tr>
                <td>HB103</td>
                <td>${fecha103}</td>
                <td>HB203</td>
                <td>${fecha203}</td>
            </tr>
            <tr>
                <td>HB104</td>
                <td>${fecha104}</td>
                <td></td>
                <td></td>
            </tr>
        </table>
    `;

   
   const mensaje ={
      from:'sistemas@zayro.com',
      //to:'ca.she@logisteed-america.com, ja.diaz@logisteed-america.com, lfzamudio@zayro.com, lzamudio@zayro.com, distribution1@zayro.com, alule@zayro.com, avazquez@zayro.com',
      to:correos,
      //to:'programacion@zayro.com',
      subject:'KAWASAKI / SFTP DATES /',
      attachments: attachments,
      html: `<p>Good morning, annex SFTP file date status and you will find the DB catalog and the latest HB203 file attached.</p>${tablaDatos}`,
   }
   const transport = await nodemailer.createTransport(config);
   await transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
  
   await transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   });
   //console.log(correos); 
}   
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
app.get('/api/leermailfacturaskmc', async function(req, res) {
   try {
       const archivos = await sql.facturasaenviar();
       if (archivos.length > 0) {
           let procesados = 0;
           for (let i = 0; i < archivos.length; i++) {
               const nombreArchivo = archivos[i];
               const resultado = await sql.ejecutar_sp_Asn(nombreArchivo.NumFactura);
               const ASNExisten = await sql.revisarasnexisten(nombreArchivo.NumFactura);
               if (ASNExisten.length > 0) {
                   
                   if (resultado && resultado.length > 0) {
                     console.log(nombreArchivo.NumFactura)
                       var wb = new xl.Workbook();
                       // Resto de tu código para la creación del archivo Excel
                       var estiloTitulo = wb.createStyle({
                           font: {
                               name: 'Arial',
                               color: '#FFFFFF',
                               size: 10,
                               bold: true,
                           },
                           fill: {
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
                        var estilocontenidoencontrado=wb.createStyle({
                          font:{
                            name: 'Arial',
                            color: 'FFFFFF',
                            size:10,
                          },
                          fill:{
                              type: 'pattern', // the only one implemented so far.
                              patternType: 'solid',
                              fgColor: '#572364',
                          },
                          });
                          var estilocontenidon=wb.createStyle({
                            font:{
                               name: 'Arial',
                               color: '#000000',
                               size:10,
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
                        ws.cell(1,17).string(nombreArchivo.NumFactura).style(estiloTitulo);
                  
                        let numfilabod=2;
                        var numfila;
                        let estilocontenidoboxid=estilocontenido
                        //console.log(resultado)
                        resultado.forEach(ren=>{
                           if (ren.fecha_scan==''){
                              estilocontenido=estilocontenidoletraroja;
                           }else{
                              estilocontenido=estilocontenidonormal;
                           }
                           variablews=ws;
                           numfila=numfilabod;
                           
                  
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
                           if(ren.existen==1){
                             estilocontenidoboxid= estilocontenidoencontrado
                           }
                           else{
                             estilocontenidoboxid=estilocontenido
                           }
                           if (ren.box_id==''){
                             variablews.cell(numfila,5).string("").style(estilocontenidoboxid);//A
                             }
                            else{
                                 variablews.cell(numfila,5).string(ren.box_id).style(estilocontenidoboxid);//A
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
                          
                           numfilabod=numfilabod+1;
                                 
                  
                  
                        });
                        /********************************************************************************** */
                        var numfila1=2;
                        let estiloCelda = estilocontenidon; 
                        ASNExisten.forEach(encontrado => {
                           //console.log(encontrado.LBID+', '+encontrado.existenasn)
                          if (encontrado.existenasn==1) {
                              // Si hay coincidencia y ambos valores no están vacíos, se aplica el estilo para coincidencias
                              estiloCelda = estilocontenidoencontrado;
                          }
                          else{
                             estiloCelda = estilocontenidon; 
                          }
                          
                          //ws.cell(numfila, 5).string(ren.box_id).style(estiloCelda);
                          
                          ws.cell(numfila1, 17).string(encontrado.LBID).style(estiloCelda);
                          numfila1 = numfila1 + 1;
                      }); 
                        /********************************************************************************** */
              
                       const pathExcel = path.join(__dirname, 'excel', 'Reporte ASN VS FACTURA.xlsx');
                       wb.write(pathExcel, function(err, stats) {
                           if (err) {
                               console.error(err);
                           } else {
                               console.log("Archivo Excel generado exitosamente");
                               procesados++;
                               if (procesados === archivos.length) {
                                   res.json("Proceso completado correctamente");
                               }
                           }
                       });
                       setTimeout(()=>{
                        sql.getdata_correos_reporte('8').then((result)=>{
                           result.forEach(renglonactual=>{
                              enviarMailfacturasmail(renglonactual.correos,nombreArchivo.NumFactura);
                           })
                        })},30000)
                   } else {
                       console.log(`No hay datos válidos para el archivo ${nombreArchivo.NumFactura}`);
                   }
               } else {
                   console.log(`No se encontraron ASN para el archivo ${nombreArchivo.NumFactura}`);
               }
           }
       } else {
           res.send("No hay archivos para procesar");
       }
   } catch (error) {
       console.error("Error:", error);
       res.status(500).send("Error interno del servidor");
   }
});
enviarMailfacturasmail=async(correos, numfactura)=>{
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
      from:'sistemas@zayro.com',
      //to:'ca.she@logisteed-america.com, ja.diaz@logisteed-america.com, lfzamudio@zayro.com, lzamudio@zayro.com, distribution1@zayro.com, alule@zayro.com, avazquez@zayro.com',
      //to:'gerenciati@zayro.com',
      to:correos, 
      //to:'programacion@zayro.com',
      subject:'ASN VS FACTURA '+numfactura,
      attachments:[
         {filename:'Reporte ASN VS FACTURA.xlsx',
         path:'./src/excel/Reporte ASN VS FACTURA.xlsx'}],
      
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
        sql.actualizarestadofactura(numfactura)
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   });
   //console.log(correos); 
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
app.get('/api/getdata_SICADIARIO',async function(req, res, next) {
   try{
      const result=await sqlram.obtenercampos()
      //console.log(result)
      const sicadiario=await sqlram.sicadiario()
      const columnNames = Object.keys(result.recordset[0]);
      //console.log(columnNames)
      const data = result.recordset;
      
      // Crear un nuevo libro de Excel y una nueva hoja de cálculo
      var wb = new xl.Workbook();
      
      
      var estiloTitulo = wb.createStyle({
         font: {
            name: 'Arial',
            color: '#FFFFFF',
            size: 10,
            bold: true,
         },
         fill: {
            type: 'pattern',
            patternType: 'solid',
            fgColor: '#008000',
            },
            alignment: {
              horizontal: 'center', // Centrar horizontalmente el contenido
              wrapText: true // Ajustar automáticamente el tamaño de la celda
            }
            
      });
      var estilocontenido = wb.createStyle({
         font: {
            name: 'Arial',
            color: '#000000',
            size: 10,
         },
         alignment: {
            horizontal: 'center' // Centrar horizontalmente el contenido
          }
      });
      var ws = wb.addWorksheet("Grafica"); 
      /*data.forEach( renglon=>{
         var wscambio=wb.addWorksheet(renglon.Usuario);
         wscambio.cell(1,1).string("Poliza").style(estiloTitulo);
         wscambio.cell(1,2).string("Movimiento").style(estiloTitulo);
         wscambio.cell(1,3).string("Fecha Modificación").style(estiloTitulo);
         wscambio.cell(1,4).string("Usuario").style(estiloTitulo);
         wscambio.cell(1,5).string("Estatus").style(estiloTitulo);
         wscambio.cell(1,6).string("Sucursal").style(estiloTitulo);
         let numfila=2
         sicadiario.forEach(async ren=>{
            
            if (renglon.Usuario==ren.UsuarioID){
               wscambio.cell(numfila,1).string(ren.PolizaID).style(estilocontenido);
               wscambio.cell(numfila,2).string(ren.MovimientoID).style(estilocontenido);
               wscambio.cell(numfila,3).string(ren.FechaModificacion).style(estilocontenido);
               wscambio.cell(numfila,4).string(ren.UsuarioID).style(estilocontenido);
               wscambio.cell(numfila,5).string(ren.Estatus).style(estilocontenido);
               wscambio.cell(numfila,6).string(ren.Sucursal).style(estilocontenido);
               numfila++
            }
         })
      })*/
      for (let renglon of data) {
         var wscambio = wb.addWorksheet(renglon.Usuario);
         wscambio.cell(1, 1).string("Poliza").style(estiloTitulo);
         wscambio.cell(1, 2).string("Movimiento").style(estiloTitulo);
         wscambio.cell(1, 3).string("Fecha Modificación").style(estiloTitulo);
         wscambio.cell(1, 4).string("Usuario").style(estiloTitulo);
         wscambio.cell(1, 5).string("Estatus").style(estiloTitulo);
         wscambio.cell(1, 6).string("Sucursal").style(estiloTitulo);
      
         let numfila = 2;
         // Usar for...of para procesar las filas de sicadiario
         for (let ren of sicadiario) {
            if (renglon.Usuario == ren.UsuarioID) {
               wscambio.cell(numfila, 1).string(ren.PolizaID).style(estilocontenido);
               wscambio.cell(numfila, 2).string(ren.MovimientoID).style(estilocontenido);
               wscambio.cell(numfila, 3).string(ren.FechaModificacion).style(estilocontenido);
               wscambio.cell(numfila, 4).string(ren.UsuarioID).style(estilocontenido);
               wscambio.cell(numfila, 5).string(ren.Estatus).style(estilocontenido);
               wscambio.cell(numfila, 6).string(ren.Sucursal).style(estilocontenido);
               numfila++;
            }
         }
      }

      columnNames.forEach( (columnName, index) => {
         ws.cell(1, index + 1).string(columnName).style(estiloTitulo);;
      });
       // Agregar los datos de la consulta a la hoja de cálculo
       data.forEach(async (row, rowIndex) => {
         columnNames.forEach((columnName, index) => {
           let value = row[columnName];
       
           // Validar el tipo de dato y convertirlo si es necesario
           if (value === null || value === undefined) {
             value = 0; // Si el valor es null o undefined, reemplazarlo con 0
                        // Escribir el valor en la celda de Excel
            ws.cell(rowIndex + 2, index + 1).number(value).style(estilocontenido);
           } else if (isNaN(value)) {
             value = value.toString(); // Convertir el valor a cadena si no es un número
             ws.cell(rowIndex + 2, index + 1).string(value).style(estilocontenido);
           } else {
             value = Number(value); // Convertir el valor a número si es un número
             ws.cell(rowIndex + 2, index + 1).number(value).style(estilocontenido);
           }
       

         });
       });

    let nombreArchivo = "Reporte SICA";
    // Guardar el archivo Excel y enviarlo como descarga al cliente
    const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');
    
      await wb.write(pathExcel);
     //console.log(pathExcel)
     
        // Función para ejecutar el script de Python con reintento
        async function ejecutarScriptPython() {
         return new Promise((resolve, reject) => {
             exec(`python agregar_grafico.py "${pathExcel}"`, async (error, stdout, stderr) => {
                 if (error) {
                     console.error(`Error al ejecutar el script de Python: ${error}`);
                     reject(error);
                     return;
                 }

                 console.log(`stdout: ${stdout}`);
                 console.error(`stderr: ${stderr}`);

                 resolve();
             });
         });
     }

     const maxIntentos = 3; // Número máximo de intentos
     let intento = 1;

     // Función para ejecutar el script con reintento
     async function ejecutarConReintentos() {
         while (intento <= maxIntentos) {
             console.log(`Intento ${intento}`);
             try {
                 await ejecutarScriptPython();
                 console.log("El script de Python se ejecutó con éxito.");
                 break;
             } catch (error) {
                 console.error(`Error en el intento ${intento}: ${error}`);
                 if (intento === maxIntentos) {
                     console.error("Se alcanzó el número máximo de intentos. No se pudo ejecutar el script de Python.");
                     res.status(500).send('Error al ejecutar el script de Python');
                     return;
                 }
                 intento++;
             }
         }
     }

     // Ejecutar script de Python con reintento
     await ejecutarConReintentos();

     // Descargar el archivo Excel con el gráfico
     res.download(pathExcel, () => {
         console.log("Archivo descargado exitosamente.");
         sql.getdata_correos_reporte('9').then((result) => {
             result.forEach(renglonactual => {
                 enviarsica(renglonactual.correos);
             });
         });
     });
     
 } catch (err) {
     console.error('Error al generar el archivo Excel:', err);
     res.status(500).send('Error al generar el archivo Excel');
 }

   


});
enviarsica=async(correos)=>{
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
   const currentDate = new Date();
   const day = currentDate.getDate();
   const month = currentDate.getMonth(); // Los meses empiezan desde 0 (enero es 0)
   let subject = 'REPORTE SICA DIARIO';

   if ((month === 1 || month === 4 || month === 7 || month === 11) && day === 5) {
      subject = 'REPORTE SICA TRIMESTRAL';
   } else if (day === 5 && !(month === 4 || month === 4 || month === 7 || month === 11)) {
      subject = 'REPORTE SICA MENSUAL';
   }
   
   const mensaje ={
      
      from:'it.sistemas@zayro.com',
      //to:'aby.zamora@arcacontal.com,valentin.garza@arcacontal.com,avazquez@zayro.com,exportacion203@zayro.com,gerenciati@zayro.com,sistemas@zayro.com',
      to: correos, 
      //to:'programacion@zayro.com',
      //to: 'oswal15do@gmail.com',
      subject: subject,
      attachments:[
         {filename:'Reporte SICA.xlsx',
         path:'./src/excel/Reporte SICA.xlsx'}],
      text:'Hola buen dia, se anexa reporte de SICA',
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos);
} 
/*************************************************************************************** */
app.get('/api/getdata_partesvshb101', async function(req, res, next) {
   try {
       const result = await sql.FRACCIONTBLPARTESVS101ESTANENBODEGAIMPORTACION();
       const result2 = await sql.FRACCIONTBLPARTESVS101ESTANENBODEGAEXPORTACION();

       const wb = new xl.Workbook();
       const nombreArchivo = "Catalogo de partes vs hb101";
       const ws = wb.addWorksheet("Importacion");
       const ws2 = wb.addWorksheet("Exportacion");
       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#008000',
           },
       });
       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       const columnas = [
           "NUMERO_DE_PARTE", "FRACCION_CATALOGO_DE_PARTES", "FRACCION_HB101", "PALLET"
       ];
       columnas.forEach((columna, index) => {
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });
       columnas.forEach((columna, index) => {
         ws2.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
           Object.keys(reglonactual).forEach((columna, idx) => {
               ws.cell(numfila, idx + 1).string(reglonactual[columna]).style(estilocontenido);
           });
           numfila++;
       });
       numfila = 2;
       result2.forEach(reglonactual => {
         Object.keys(reglonactual).forEach((columna, idx) => {
             ws2.cell(numfila, idx + 1).string(reglonactual[columna]).style(estilocontenido);
         });
         numfila++;
     });

       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');

       wb.write(pathExcel, async function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               try {
                   await fs.promises.access(pathExcel, fs.constants.F_OK);
                   res.download(pathExcel, () => {
                       //fs.unlink(pathExcel, (err) => {
                           if (err) console.error(err);
                           else console.log("Archivo descargado y eliminado exitosamente.");
                           sql.getdata_correos_reporte('10').then((result)=>{
                              result.forEach(renglonactual=>{
                                 enviarMaiLPartesvsHB101(renglonactual.correos,nombreArchivo,pathExcel);
                              })
                           })
                           
                       //});
                   });
               } catch (err) {
                   console.error(err);
                   res.status(500).send("Error al acceder al archivo Excel generado.");
               }
           }
       });
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
});
enviarMaiLPartesvsHB101=async(correos,nombreArchivo,pathExcel)=>{
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
      subject:'Catalogo de partes vs HB101',
      attachments:[
         {filename:nombreArchivo + '.xlsx',
         path: './src/excel/Catalogo de partes vs hb101.xlsx'},],
      text:'Hola buen día se anexa reporte de la comparativa del catalogo de partes vs el HB101 ',
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos);
} 
/*****************************************************************/
/*****************************************************************/
/*****************************************************************/
app.get('/api/inventario', async (req, res) => {
   try{
  
   let esmensual='1';

   const tipo='mensual'
   const workbook = new ExcelJS.Workbook();
   //console.log(esmensual,tipo)
   const ws  = workbook.addWorksheet("Almacenaje_"+tipo);
   
   // Establecer los títulos de las columnas
   const columnTitles = [
     '     NUMERO DE PALLET     ','   BOXID   ','   POKMC   ','   PODKMC   ',
     '   Parts Code   ','   PCS / Box ID   ','   ACTUAL IN   ','   ACTUAL STORAGE DATE  ',
     '   MONTH COUNT IN  ','   ACTUAL OUT  ','   MONTH COUNT OUT  ','  COUNTABLE STORAGE DATE  ', '  WEEK COUNT  ', '     TOTAL     '
   ];
     ws.getRow(1).values = columnTitles;
     ws.getRow(1).font = {
       name: 'Arial',
       color: '#000000',
       size: 10,
       bold: true,
     };
     // Centrar el contenido de las columnas y autoajustar el ancho de las columnas
     for (let i = 1; i <= columnTitles.length; i++) {
       ws.getColumn(i).alignment = { horizontal: 'center' };
     }
 
     ws.columns.forEach(column => {
       let maxLength = 0;
       column.eachCell({ includeEmpty: true }, cell => {
         const columnLength = cell.value ? cell.value.toString().length : 10;
         if (columnLength > maxLength) {
           maxLength = columnLength;
         }
       });
       column.width = maxLength < 10 ? 10 : maxLength + 2; 
     });
   const result=await sqlSIS.facturakmx_inventario(esmensual);
   
   let numfila = 2;
   const dollarFormat = '$#,##0.00'; // Formato de moneda
   let pallet='';
   let banderaprimerboxid='1';
   let anterior=2;
   let actual=0;
   result.forEach(renglon=>{
     if (numfila==2){
       pallet=renglon.Pallet;
       ws.mergeCells('B'+numfila+':'+'B'+(numfila+1));
       ws.mergeCells('C'+numfila+':'+'C'+(numfila+1));
       ws.mergeCells('D'+numfila+':'+'D'+(numfila+1));
       ws.mergeCells('E'+numfila+':'+'E'+(numfila+1));
       ws.mergeCells('F'+numfila+':'+'F'+(numfila+1));
       /*actual_in=renglon.FechaEntrada;
       actual_storage_date=renglon.Storage_Days 
       april_count_in=renglon.aprilcountin;
       actual_out=renglon.Actual_Out;
       april_count_out=renglon.Abril_Count_Out;
       week_count=renglon.Semanas ;
       total=renglon.Total;*/
 
     }
     if (renglon.Pallet === 'ZZZZZZZ'){
       if (pallet != renglon.Pallet)
       {
         actual=numfila-1
         ws.mergeCells('A'+anterior+':'+'A'+actual);
         ws.mergeCells('G'+anterior+':'+'G'+actual);
         ws.mergeCells('H'+anterior+':'+'H'+actual);
         ws.mergeCells('I'+anterior+':'+'I'+actual);
         ws.mergeCells('J'+anterior+':'+'J'+actual);
         ws.mergeCells('K'+anterior+':'+'K'+actual);
         ws.mergeCells('L'+anterior+':'+'L'+actual);
         ws.mergeCells('M'+anterior+':'+'M'+actual);
         ws.mergeCells('N'+anterior+':'+'N'+actual);
 
         ws.mergeCells('B'+numfila+':'+'B'+(numfila+1));
         ws.mergeCells('C'+numfila+':'+'C'+(numfila+1));
         ws.mergeCells('D'+numfila+':'+'D'+(numfila+1));
         ws.mergeCells('E'+numfila+':'+'E'+(numfila+1));
         ws.mergeCells('F'+numfila+':'+'F'+(numfila+1));
         anterior=numfila
         ws.getCell(actual, 1).value = pallet  || '';
         pallet=renglon.Pallet;
       }
       
                
       ws.getCell(numfila, 1).value = 'Total:';
       ws.getCell(numfila, 14).value = { formula: 'SUM(' +  'N2:' + 'N' + (numfila - 1) + ')', result: 7 };
       //ws.getCell(numfila, 14).value = renglon.Total !== 0 ? renglon.Total : '-';
       ws.getCell(numfila, 14).numFmt = dollarFormat;
     }
     else{
       if (pallet != renglon.Pallet)
       {
         actual=numfila-1
         ws.mergeCells('A'+anterior+':'+'A'+actual);
         ws.mergeCells('G'+anterior+':'+'G'+actual);
         ws.mergeCells('H'+anterior+':'+'H'+actual);
         ws.mergeCells('I'+anterior+':'+'I'+actual);
         ws.mergeCells('J'+anterior+':'+'J'+actual);
         ws.mergeCells('K'+anterior+':'+'K'+actual);
         ws.mergeCells('L'+anterior+':'+'L'+actual);
         ws.mergeCells('M'+anterior+':'+'M'+actual);
         ws.mergeCells('N'+anterior+':'+'N'+actual);
 
         ws.mergeCells('B'+numfila+':'+'B'+(numfila+1));
         ws.mergeCells('C'+numfila+':'+'C'+(numfila+1));
         ws.mergeCells('D'+numfila+':'+'D'+(numfila+1));
         ws.mergeCells('E'+numfila+':'+'E'+(numfila+1));
         ws.mergeCells('F'+numfila+':'+'F'+(numfila+1));
         anterior=numfila
         ws.getCell(actual, 1).value = pallet  || '';
         pallet=renglon.Pallet;
       }
       ws.getCell(numfila, 1).value = renglon.Pallet || '';
 
       ws.getCell(numfila, 2).value = renglon.BoxID || '';
       ws.getCell(numfila, 3).value = renglon.PO || '';
       ws.getCell(numfila, 4).value = renglon.POD || '';
       ws.getCell(numfila, 5).value = renglon.NumParte || '';
       ws.getCell(numfila, 6).value = renglon.Piezas || '';
       ws.getCell(numfila, 7).value = renglon.FechaEntrada|| '';
       ws.getCell(numfila, 8).value = renglon.Storage_Days || '';
       ws.getCell(numfila, 9).value = renglon.aprilcountin || '';
       ws.getCell(numfila, 10).value = renglon.Actual_Out || '';
       ws.getCell(numfila, 11).value = renglon.Abril_Count_Out|| '';
       ws.getCell(numfila, 12).value = renglon.countstoragedate || '';
       ws.getCell(numfila, 13).value = renglon.Semanas || '';
       ws.getCell(numfila, 14).value = renglon.Total !== 0 ? renglon.Total : '-';
       ws.getCell(numfila, 14).numFmt = dollarFormat;
     }
     numfila++;
   })
 
   const pathExcel = path.join(__dirname, `Almacenaje_${tipo}.xlsx`);
   //await workbook.xlsx.writeFile(excelFilePath);
   workbook.xlsx.writeFile(pathExcel).then(async function() {
      try {
          // Verifica que el archivo existe
          await fs.promises.access(pathExcel, fs.constants.F_OK);
          const nombreDescarga = `Almacenaje_${tipo}.xlsx` || path.basename(pathExcel);
          // Enviar el archivo para descarga
          res.download(pathExcel, nombreDescarga, (err) => {
            if (err) {
               console.error("Error durante la descarga:", err);
               return; // ya se manejó el error de envío
            }

            console.log("Archivo descargado exitosamente.");

            // Ejecuta tareas asíncronas sin bloquear la respuesta
            (async () => {
               try {
                  await enviarMaiLAlmacenajes('sistemas@zayro.com', nombreDescarga, pathExcel);
               } catch (e) {
                  console.error("Error enviando correo:", e);
               }

               try {
                  await fs.promises.unlink(pathExcel);
                  console.log("Archivo eliminado exitosamente.");
               } catch (e) {
                  console.error("Error al eliminar el archivo:", e);
               }
            })();
            });
      } catch (err) {
          console.error("Error al acceder al archivo Excel generado:", err);
          res.status(500).send("Error al acceder al archivo Excel generado.");
      }
  }).catch(function(err) {
      console.error("Error al generar el archivo Excel:", err);
      res.status(500).send("Error al generar el archivo Excel.");
  });
   
 
   //await res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
   //await res.setHeader('Content-Disposition', `attachment; filename=Almacenaje_${tipo}.xlsx`);
 /*
   const fileStream = fs.createReadStream(excelFilePath);
   await fileStream.on('error', (err) => {
     console.error('Error al leer el archivo:', err);
     res.status(500).send('Error al descargar el archivo.');
   });
 
   await fileStream.on('finish', () => {
     console.log('Archivo enviado con éxito');

     fs.unlinkSync(excelFilePath);
   });
 
   await fileStream.pipe(res);
 */
   }catch (error) {
     console.error(error);
     
   }
});
// Requiere: const nodemailer = require('nodemailer');

enviarMaiLAlmacenajes = async (correos, nombreArchivo, pathExcel) => {
  const config = {
    host: process.env.hostemail,
    port: Number(process.env.portemail),
    secure: true, // deja tu configuración actual
    auth: {
      user: process.env.useremail,
      pass: process.env.passemail
    },
    tls: { rejectUnauthorized: false },
  };

  const mensaje = {
    from: 'sistemas@zayro.com',
    to: 'sistemas@zayro.com',
    subject: 'Almacenajes Mensual',
    text: 'Almacenajes Mensual',
    attachments: [
      {
        filename: `${nombreArchivo}.xlsx`,
        path: `${pathExcel}`
      }
    ]
  };

  const transport = nodemailer.createTransport(config);

  try {
    await transport.verify();
    const info = await transport.sendMail(mensaje);
    console.log('Correo enviado:', info.messageId || info.response);
    return info;
  } catch (error) {
    console.error('Error al enviar el correo:', error);
    throw error;
  } finally {
    transport.close();
  }
};

/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
//Reporte Semanal de Thyssenkrupp (NUEVO FORMATO)
app.get('/api/getdata_Thyssenkrupp', async function(req,res,next){ // Se removió /:fechaini/:fechafin
   let config = {
     host:process.env.hostemail,
     port:process.env.portemail,
     secure: true,
     auth: {
      user:process.env.useremail,
      pass:process.env.passemail
     },
      tls: {
      rejectUnauthorized: false
      }
   }
    let transport = nodemailer.createTransport(config);
    //var fechaini = req.params.fechaini;
    //var fechafin = req.params.fechafin;
   
    var wb = new xl.Workbook();
    var wb2 = new xl.Workbook();


    let nombreArchivo  = "Reporte Semanal Thyssenkrupp USD (4133)";
    let nombreArchivo2 = "Reporte Semanal Thyssenkrupp MXN (4133)";

    var wsUSD = wb.addWorksheet("Cuenta Zayro 4133");
    var wsMXN = wb2.addWorksheet("Cuenta Zamudio 4133");

   
    var estiloTitulo = wb.createStyle({
        font: {
       name:   'Arial',
       color:  '#FFFFFF',
       size:   10,
       bold:   true,
       },
        fill: {
       type:   'pattern',
       patternType:'solid',
       fgColor:'#000000',
       },
    });
    var estiloSubTitulo = wb.createStyle({
      font: {
         name: 'Arial',
         color: '1F4E78',
         size: 10,
         bold: true,
      },
      fill: {
         type: 'pattern',
         patternType: 'solid',
         fgColor: '#BDD7EE'
      }
    })
    var estiloContenido = wb.createStyle({
        font:{
         name: 'Arial',
         color: '#000000',
         size: 10,
        }
    });
    var estilozamprov = wb.createStyle({
      font: {
         name:    'Arial',
         color:   '#FFFFFF',
         size:    14,
         bold:    true,
      },
      fill: {
         type:    'pattern',
         patternType:'solid',
         fgColor: '#00448D',
      }
    });
    var estiloTotal = wb.createStyle({
      font:{
         name:    'Arial',
         color:   '#000000',
         size:    10,
         bold:    true,
      },
      fill:{
         type:    'pattern',
         patternType:'solid',
         fgColor: '#C6E0B4',
      }
    });

    
    const resultado =  await sqlzay.getdata_EdoCtaThyssenkrupp_Dolares();

      //Cabecilla
      wsUSD.cell(1,1).string("Cliente").style(estiloTitulo);
      wsUSD.cell(1,2,1,7,true).string("Razon Social").style(estiloTitulo);
      wsUSD.cell(2,1).string("4133").style(estilozamprov);
      wsUSD.cell(2,2,2,7,true).string("ZAMUDIO Y RODRIGUEZ").style(estilozamprov);    
      
      // Inicio. Últimos Cambios
      let total = 0;
      let numRenglon = 4;
      wsUSD.column(1).setWidth(14);
      wsUSD.column(2).setWidth(13);
      wsUSD.column(3).setWidth(17);
      wsUSD.column(4).setWidth(15);
      wsUSD.column(6).setWidth(18);

      // Nombre de columnas
      //console.log(resultado);
      resultado.forEach(renglonactual => {
         wsUSD.cell(numRenglon,1).string("PROYECTO").style(estiloTitulo);
         wsUSD.cell(numRenglon += 1,1).string("REFERENCIA").style(estiloTitulo);
         wsUSD.cell(numRenglon += 1,1).string("PEDIMENTO").style(estiloTitulo);
         wsUSD.cell(numRenglon += 1,1).string("TIPO OP.").style(estiloTitulo);
         wsUSD.cell(numRenglon += 1,1).string("PEDIDO").style(estiloTitulo);
         wsUSD.cell(numRenglon += 1,1).string("F. FISCAL - UUID").style(estiloTitulo);
         wsUSD.cell(numRenglon += 1,1).string("FOLIO").style(estiloSubTitulo);
         wsUSD.cell(numRenglon,2).string("SUBTOTAL").style(estiloSubTitulo);
         wsUSD.cell(numRenglon,3).string("I.TRASLADADOS 8%").style(estiloSubTitulo);
         wsUSD.cell(numRenglon,4).string("RETENCIÓN IVA").style(estiloSubTitulo);
         wsUSD.cell(numRenglon,5).string("TOTAL").style(estiloSubTitulo);
         wsUSD.cell(numRenglon,6).string("CUENTA DE GASTOS").style(estiloSubTitulo);
         numRenglon -= 6

         // Inserción de datos
         wsUSD.cell(numRenglon += 1,2).string(renglonactual.Referencia).style(estiloContenido);
         wsUSD.cell(numRenglon += 1,2).string(renglonactual.Pedimento).style(estiloContenido);
         wsUSD.cell(numRenglon += 1,2).string(renglonactual.Tipo_Operacion).style(estiloContenido);
         wsUSD.cell(numRenglon += 1,2).string(renglonactual.Pedido).style(estiloContenido);
         wsUSD.cell(numRenglon += 1,2).string(renglonactual.Folio_UUID).style(estiloContenido);
         wsUSD.cell(numRenglon += 2,1).string(renglonactual.Folio).style(estiloContenido);
         wsUSD.cell(numRenglon,2).number(Number(renglonactual.Subtotal)).style({numberFormat: '$###0.00'});
         wsUSD.cell(numRenglon,3).number(Number(renglonactual.IVA)).style({numberFormat: '$###0.00'});
         wsUSD.cell(numRenglon,4).number(Number(renglonactual.Retencion_IVA)).style({numberFormat: '$###0.00'});
         wsUSD.cell(numRenglon,5).number(Number(renglonactual.Total)).style({numberFormat: '$###0.00'});
         //wsUSD.cell(numRenglon,6).string(renglonactual.Cuenta_Gastos).style(estiloContenido);
         
         // Fin Últimos Cambios
         total = total + Number(renglonactual.Total);
         numRenglon += 2;
      })
      console.log(total);
      wsUSD.cell(numRenglon,5).string("TOTAL").style(estiloTotal)
      wsUSD.cell(numRenglon + 1,5).number(total).style(estiloContenido);

      const pathExcel1 = path.join(__dirname,'excel',nombreArchivo+'.xlsx');
      // Guardado
      wb.write(pathExcel1,function(err,stats){
         if(err) console.log(err);
         else {
            console.log("Archivo 1 Generado");
         }
      });
    

   const resultado2 = await sqlzam.getdata_edoCtaThyssenkrupp_Pesos(4133);
      
      wsMXN.cell(1,1).string("Cliente").style(estiloTitulo);
      wsMXN.cell(1,2,1,7,true).string("Razon Social").style(estiloTitulo);
      wsMXN.cell(2,1).string("4133").style(estilozamprov);
      wsMXN.cell(2,2,2,7,true).string("ZAMUDIO Y RODRIGUEZ").style(estilozamprov);

      // Inicio Últimos Cambios
      total = 0;
      numRenglon = 4;
      wsMXN.column(1).setWidth(14);
      wsMXN.column(2).setWidth(13);
      wsMXN.column(3).setWidth(17);
      wsMXN.column(4).setWidth(15);
      wsMXN.column(6).setWidth(18);
      //wsMXN.column(7).setWidth(18);

      //console.log(resultado2);
      resultado2.forEach(renglonactual2 => {
         wsMXN.cell(numRenglon,1).string("PROYECTO").style(estiloTitulo);
         wsMXN.cell(numRenglon += 1,1).string("REFERENCIA").style(estiloTitulo);
         wsMXN.cell(numRenglon += 1,1).string("PEDIMENTO").style(estiloTitulo);
         wsMXN.cell(numRenglon += 1,1).string("TIPO OP.").style(estiloTitulo);
         wsMXN.cell(numRenglon += 1,1).string("PEDIDO").style(estiloTitulo);
         wsMXN.cell(numRenglon += 1,1).string("F. FISCAL - UUID").style(estiloTitulo);
         wsMXN.cell(numRenglon += 1,1).string("FOLIO").style(estiloSubTitulo);
         wsMXN.cell(numRenglon,2).string("SUBTOTAL").style(estiloSubTitulo);
         wsMXN.cell(numRenglon,3).string("I.TRASLADADOS 8%").style(estiloSubTitulo);
         wsMXN.cell(numRenglon,4).string("RETENCIÓN IVA").style(estiloSubTitulo);
         wsMXN.cell(numRenglon,5).string("TOTAL").style(estiloSubTitulo);
         wsMXN.cell(numRenglon,6).string("CUENTA DE GASTOS").style(estiloSubTitulo);
         numRenglon -= 6
         // Inserción de datos
         
         wsMXN.cell(numRenglon += 1,2).string(renglonactual2.Referencia).style(estiloContenido);
         wsMXN.cell(numRenglon += 1,2).string(renglonactual2.Pedimento).style(estiloContenido);
         wsMXN.cell(numRenglon += 1,2).string(renglonactual2.Tipo_Operacion).style(estiloContenido);
         wsMXN.cell(numRenglon += 1,2).string(renglonactual2.Pedido).style(estiloContenido);
         wsMXN.cell(numRenglon += 1,2).string(renglonactual2.Folio_UUID).style(estiloContenido);
         wsMXN.cell(numRenglon += 2,1).string(renglonactual2.Folio).style(estiloContenido);
         wsMXN.cell(numRenglon,2).number(Number(renglonactual2.Subtotal)).style({numberFormat: '$###0.00'});
         wsMXN.cell(numRenglon,3).number(Number(renglonactual2.IVA)).style({numberFormat: '$###0.00'});
         wsMXN.cell(numRenglon,4).number(Number(renglonactual2.Retencion_IVA)).style({numberFormat: '$###0.00'});
         wsMXN.cell(numRenglon,5).number(Number(renglonactual2.Total)).style({numberFormat: '$###0.00'});
         //wsMXN.cell(numRenglon,6).string(renglonactual2.Cuenta_Gastos).style(estiloContenido);
         
         // Fin Últimos cambios
         total = total + Number(renglonactual2.Total);
         numRenglon += 2;
      });
      //console.log(total);
      wsMXN.cell(numRenglon,5).string("TOTAL").style(estiloTotal)
      wsMXN.cell(numRenglon + 1,5).number(total).style(estiloContenido);
   
      const pathExcel2 = path.join(__dirname,'excel',nombreArchivo2 + '.xlsx');
      await wb2.write(pathExcel2,function(err,stats){
         if(err) console.log(err);
         else {
            res.json("Archivos generados")
            console.log("Archivo 2 Generado");
         }
      }) 

      sql.getdata_correos_reporte('4').then((resultado) => {
         resultado.forEach(renglonactual => {
            enviarMailEstadoCuentaThyn(nombreArchivo, nombreArchivo2,  transport, renglonactual.correos);
         })
      })
}); 
enviarMailEstadoCuentaThyn = async(nombreArchivo, nombreArchivo2, transport, correos) => {
   const mensaje = {
      from:'it.sistemas@zayro.com',
      to: 'cobranza@zayro.com;sistemas@zayro.com',//correos,
      subject: 'Estado de cuenta Thyssenkrupp',
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }, {
            filename: nombreArchivo2 + '.xlsx',
            path: './src/excel/' + nombreArchivo2 + '.xlsx',
         }],
      text: 'Estado de Cuenta Thyssenkrupp',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });

}
app.get('/api/getdata_Thyssenkrupp_Regular',async function(req,res,next){
   let config = {
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth: {
         user:process.env.useremail,
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false
      }
   }
   let transport = nodemailer.createTransport(config);

   var wb = new xl.Workbook();
   var wb2 = new xl.Workbook();
   var wb3 = new xl.Workbook();
   var wb4 = new xl.Workbook();

   let nombreArchivo  = "Reporte Semanal Thyssenkrupp USD (4133)";
   let nombreArchivo2 = "Reporte Semanal Thyssenkrupp MXN (4133)";
   let nombreArchivo3 = "Reporte Semanal Thyssenkrupp USD (3157)";
   let nombreArchivo4 = "Reporte Semanal Thyssenkrupp MXN (3157)"
   var wsUSD = wb.addWorksheet("Cuenta Zayro 4133");
   var wsMXN = wb2.addWorksheet("Cuenta Zamudio 4133");
   var wsUSD2 = wb3.addWorksheet("Cuenta Zayro 3157");
   var wsMXN2 = wb4.addWorksheet("Cuenta Zamudio 3157");

   var estiloTitulo = wb.createStyle({
      font: {
      name:   'Calibri',
      color:  '#000000',
      size:   18,
      bold:   true,
      },
   });
   var estiloSubTitulo = wb.createStyle({
      font: {
      name: 'Calibri',
      color: '#000000',
      size: 11,
      bold: true,
   },
   })
   var estiloContenido = wb.createStyle({
      font:{
      name: 'Calibri',
      color: '#000000',
      size: 11,
      }
   });
   var estiloNomColumna = wb.createStyle({
   font: {
      name:    'Arial',
      color:   '#FFFFFF',
      size:    11,
      bold:    true,
   },
   fill: {
      type:    'pattern',
      patternType:'solid',
      fgColor: '#002060',
   }
   });

   const meses = [
      "ene","feb","mzo","abr","may","jun",
      "jul","ago","sept","oct","nov","dic"
   ]

   const mesesEng = [
      "jan","feb","mar","apr","may","jun",
      "jul","aug","sep","oct","nov","dec"
   ]

   let fecha = new Date;
   let diaLunes = fecha.getDate();
   let mes = fecha.getMonth();
   let anual = fecha.getFullYear();

   let fechaFormatEng = `${diaLunes}-${mesesEng[mes]}-${anual}`
   let fechaFormat = `${diaLunes}-${meses[mes]}-${anual}`

   
   const columnas = [
      "COMPANY","ACCOUNT FULL NAME","INVOICE","DATE","DUE DATE","FILE / REF NUMBER","CURRENT","USD AMOUNT","PSMA",
      "AGING","DAYS","BU","KN COMMENTS","PLANTA","THYSSEN COMMENTS","STATUS","OWNER"
   ]

   const anchoColumnas = [
      10,52,8,8,10,17,10,13,
      10,7,7,11,15,9,20,10,10
   ]

   const resultado = await sqlzay.getdata_EdoCuentaTkDolares(4133);

   for (let i = 0; i < 17; i++) {
      wsUSD.column(i+1).setWidth(anchoColumnas[i]);
   }

   wsUSD.cell(2,6).string("STATEMENT OF ACCOUNT").style(estiloTitulo);

   wsUSD.cell(4,1).string("Account N°")            .style(estiloSubTitulo);
   wsUSD.cell(4,3).string("T000239")               .style(estiloSubTitulo);
   wsUSD.cell(4,5).string("Date Report")           .style(estiloSubTitulo);
   wsUSD.cell(4,6).string(fechaFormatEng)          .style(estiloContenido);

   wsUSD.cell(5,1).string("Name")                  .style(estiloSubTitulo);
   wsUSD.cell(5,3).string("THYSSENKRUPP COMPONENTS TECHNOLOGY DE MEXICO S.A. DE C.V.").style(estiloSubTitulo);
   wsUSD.cell(6,1).string("Plant")                 .style(estiloSubTitulo);
   wsUSD.cell(6,3).string("DAMPER")                .style(estiloSubTitulo);

   columnas.forEach((columna,index) => {
      wsUSD.cell(8,index + 1).string(columna).style(estiloNomColumna);
   });

   let total = 0;

   const columnasMXN = [
      "COMPANY","ACCOUNT FULL NAME","INVOICE","DATE","DUE DATE","FILE / REF NUMBER","CURRENT","MXN AMOUNT","PSMA",
      "AGING","DAYS","BU","KN COMMENTS","PLANTA","THYSSEN COMMENTS","STATUS","OWNER"
   ]
   columnas.forEach((columna,index) => {
      wsUSD.cell(8,index + 1).string(columna).style(estiloNomColumna);
   });

   /* 
   El ingreso de datos revisa el nombre de una columna. Cualquier ajuste
   debe reflejarse en el array 'columnas'
   */
   let numRenglon = 9;
   resultado.forEach(renglonactual => {
      Object.keys(renglonactual).forEach((columna,idx) => {
         if (columna ==='DAYS') {
            wsUSD.cell(numRenglon,idx + 1).number(renglonactual[columna]).style(estiloContenido);
         } else if (typeof(renglonactual[columna]) === 'number' && columna != 'DAYS') {
            wsUSD.cell(numRenglon,idx + 1).number(renglonactual[columna]).style({numberFormat: '#,##0.00'});
            total += renglonactual[columna]
         } else {
            wsUSD.cell(numRenglon,idx + 1).string(renglonactual[columna]).style(estiloContenido)
         }
      });
      numRenglon ++
   })
   console.log(total);

   wsUSD.cell(numRenglon + 3, 8).number(total).style({numberFormat: '$#,##0.00'});

   const pathExcel = path.join(__dirname, 'excel',nombreArchivo + '.xlsx');

   wb.write(pathExcel, async function(err) {
      if (err) console.error(err);
      else {
         console.log("Reporte 1 Generado");
      }
   })

   const resultado2 = await sqlzam.getdata_EdoCuentaTkPesos(4133);

   for (let i = 0; i < 17; i++) {
      wsMXN.column(i+1).setWidth(anchoColumnas[i]);
   }

   wsMXN.cell(2,6).string("STATEMENT OF ACCOUNT").style(estiloTitulo);

   wsMXN.cell(4,1).string("Account N°")            .style(estiloSubTitulo);
   wsMXN.cell(4,3).string("T000239")               .style(estiloSubTitulo);
   wsMXN.cell(4,5).string("Date Report")           .style(estiloSubTitulo);
   wsMXN.cell(4,6).string(fechaFormat)             .style(estiloContenido);

   wsMXN.cell(5,1).string("Name")                  .style(estiloSubTitulo);
   wsMXN.cell(5,3).string("THYSSENKRUPP COMPONENTS TECHNOLOGY DE MEXICO S.A. DE C.V.").style(estiloSubTitulo);
   wsMXN.cell(6,1).string("Plant")                 .style(estiloSubTitulo);
   wsMXN.cell(6,3).string("DAMPER")                .style(estiloSubTitulo);

   columnasMXN.forEach((columna,index) => {
      wsMXN.cell(8,index + 1).string(columna).style(estiloNomColumna);
   });

   total = 0;

   numRenglon = 9;
   resultado2.forEach(renglonactual => {
      Object.keys(renglonactual).forEach((columna,idx) => {
         if (columna ==='DAYS') {
            wsMXN.cell(numRenglon,idx + 1).number(renglonactual[columna]).style(estiloContenido);
         } else if (typeof(renglonactual[columna]) === 'number' && columna != 'DAYS') {
            wsMXN.cell(numRenglon,idx + 1).number(renglonactual[columna]).style({numberFormat: '#,##0.00'});
            total += renglonactual[columna]
            //console.log(columna)
         } else {
            wsMXN.cell(numRenglon,idx + 1).string(renglonactual[columna]).style(estiloContenido)
         }
      });
      numRenglon ++
   })
   //console.log(total);

   wsMXN.cell(numRenglon + 3, 8).number(total).style({numberFormat: '$#,##0.00'});

   const pathExcel2 = path.join(__dirname, 'excel',nombreArchivo2 + '.xlsx');
   
   wb2.write(pathExcel2, async function(err) {
      if (err) console.error(err);
      else {
         console.log("Reporte 2 Generado");
      }
   })

   const resultado3 = await sqlzay.getdata_EdoCuentaTkDolares(3157);

   for (let i = 0; i < 17; i++) {
      wsUSD2.column(i+1).setWidth(anchoColumnas[i]);
   }
   
   wsUSD2.cell(2,6).string("STATEMENT OF ACCOUNT").style(estiloTitulo);

   wsUSD2.cell(4,1).string("Account N°")           .style(estiloSubTitulo);
   wsUSD2.cell(4,3).string("T000239")              .style(estiloSubTitulo);
   wsUSD2.cell(4,5).string("Date Report")          .style(estiloSubTitulo);
   wsUSD2.cell(4,6).string(fechaFormatEng)         .style(estiloContenido);

   wsUSD2.cell(5,1).string("Name")                 .style(estiloSubTitulo);
   wsUSD2.cell(5,3).string("THYSSENKRUPP COMPONENTS TECHNOLOGY DE MEXICO S.A. DE C.V.").style(estiloSubTitulo);
   wsUSD2.cell(6,1).string("Plant")                .style(estiloSubTitulo);
   wsUSD2.cell(6,3).string("DAMPER")               .style(estiloSubTitulo);

   columnas.forEach((columna,index) => {
      wsUSD2.cell(8,index + 1).string(columna).style(estiloNomColumna);
   });

   total = 0

   numRenglon = 9;
   resultado3.forEach(renglonactual => {
      Object.keys(renglonactual).forEach((columna,idx) => {
         if (columna ==='DAYS') {
            wsUSD2.cell(numRenglon,idx + 1).number(renglonactual[columna]).style(estiloContenido);
         } else if (typeof(renglonactual[columna]) === 'number' && columna != 'DAYS') {
            wsUSD2.cell(numRenglon,idx + 1).number(renglonactual[columna]).style({numberFormat: '#,##0.00'});
            total += renglonactual[columna]
         } else {
            wsUSD2.cell(numRenglon,idx + 1).string(renglonactual[columna]).style(estiloContenido)
         }
      });
      numRenglon ++
   })
   console.log(total);

   wsUSD2.cell(numRenglon + 3, 8).number(total).style({numberFormat: '$#,##0.00'});

   const pathExcel3 = path.join(__dirname, 'excel',nombreArchivo3 + '.xlsx');

   wb3.write(pathExcel3, async function(err) {
      if (err) console.error(err);
      else {
         console.log("Reporte 3 Generado");
      }
   })

   const resultado4 = await sqlzam.getdata_EdoCuentaTkPesos(3157);

   for (let i = 0; i < 17; i++) {
      wsMXN2.column(i+1).setWidth(anchoColumnas[i]);
   }

   wsMXN.cell(2,6).string("STATEMENT OF ACCOUNT").style(estiloTitulo);

   wsMXN2.cell(4,1).string("Account N°")           .style(estiloSubTitulo);
   wsMXN2.cell(4,3).string("T000239")              .style(estiloSubTitulo);
   wsMXN2.cell(4,5).string("Date Report")          .style(estiloSubTitulo);
   wsMXN2.cell(4,6).string(fechaFormat)            .style(estiloContenido);

   wsMXN2.cell(5,1).string("Name")                 .style(estiloSubTitulo);
   wsMXN2.cell(5,3).string("THYSSENKRUPP COMPONENTS TECHNOLOGY DE MEXICO S.A. DE C.V.").style(estiloSubTitulo);
   wsMXN2.cell(6,1).string("Plant")                .style(estiloSubTitulo);
   wsMXN2.cell(6,3).string("DAMPER")               .style(estiloSubTitulo);

   columnasMXN.forEach((columna,index) => {
      wsMXN2.cell(8,index + 1).string(columna).style(estiloNomColumna);
   });

   total = 0

   numRenglon = 9;
   resultado4.forEach(renglonactual => {
      Object.keys(renglonactual).forEach((columna,idx) => {
         if (columna ==='DAYS') {
            wsMXN2.cell(numRenglon,idx + 1).number(renglonactual[columna]).style(estiloContenido);
         } else if (typeof(renglonactual[columna]) === 'number' && columna != 'DAYS') {
            wsMXN2.cell(numRenglon,idx + 1).number(renglonactual[columna]).style({numberFormat: '#,##0.00'});
            total += renglonactual[columna]
            //console.log(columna)
         } else {
            wsMXN2.cell(numRenglon,idx + 1).string(renglonactual[columna]).style(estiloContenido)
         }
      });
      numRenglon ++
   })
   console.log(total);

   wsMXN2.cell(numRenglon + 3, 8).number(total).style({numberFormat: '$#,##0.00'});

   const pathExcel4 = path.join(__dirname, 'excel',nombreArchivo4 + '.xlsx');

   await wb4.write(pathExcel4, async function(err) {
      if (err) console.error(err);
      else {
         res.json("Todos los reportes generados")
         console.log("Reporte 4 Generado");
      }
   });

   setTimeout(() => {
      sql.getdata_correos_reporte('4').then((resultado) => {
         resultado.forEach(renglonactual => {
            enviarMailEstadoCuentaThyn(nombreArchivo, nombreArchivo2, nombreArchivo3, nombreArchivo4, transport, renglonactual.correos);
         })
      })
    },10000) 
})
enviarMailEstadoCuentaThyn = async(nombreArchivo, nombreArchivo2, nombreArchivo3, nombreArchivo4,transport, correos) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'cobranza@zayro.com;sistemas@zayro.com',//correos,
      subject: 'Estado de cuenta Thyssenkrupp',
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }, {
            filename: nombreArchivo2 + '.xlsx',
            path: './src/excel/' + nombreArchivo2 + '.xlsx',
         }, {
            filename: nombreArchivo3 + '.xlsx',
            path: './src/excel/' + nombreArchivo3 + '.xlsx',
         }, {
            filename: nombreArchivo4 + '.xlsx',
            path: './src/excel/' + nombreArchivo4 + '.xlsx'
         }],
      text: 'Estado de Cuenta Thyssenkrupp',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_estadosdecuentanld', async function(req, res, next) {
   const clientes=await sqlzam.sp_clientesestadocuenta();
   let config = {
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth: {
         user:process.env.useremail,
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false
      }
   }
   let transport = nodemailer.createTransport(config);
   let totalclientes=0
   for (let i = 0; i <clientes.length; i++) {
      totalclientes=totalclientes+1
      const cliente = clientes[i];
      

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Estado de Cuenta');

      // Estilo general de la hoja
      worksheet.pageSetup = {
         margins: {
            left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3
         }
      };
      // Configurar el rango con bordes blancos
      const range = { startRow: 1, endRow: 200, startCol: 1, endCol: 26 }; // 26 columnas (A-Z)

      // Aplicar bordes blancos a cada celda en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.border = {
                  top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  right: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            };
         }
      }

      // Establecer un fondo blanco para todas las celdas en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFFFF' }, // Fondo blanco
            };
         }
      }

      // Configuración de ancho de columnas
      worksheet.columns = [
         { width: 12 },//A
         { width: 6.1 },//B
         { width: 22 },//C
         { width: 13.78 },//D
         { width: 3 },//E
         { width: 9.7 },//F
         { width: 13.5 },//G
         { width: 15 },//H
         { width: 10 },//I
         { width: 5 },//J
         { width: 15 }//K
      ];
      const inicial=await sqlzam.datosinicialescliente(cliente.ClienteID);
      console.log(cliente.ClienteID)
      if(inicial.length>0){
         inicial.forEach(ini=>{

         
         // Fila 1: Encabezado con fondo amarillo
         worksheet.mergeCells('A1:K1');
         const headerCell = worksheet.getCell('A1');
         headerCell.value = ini.Sucursal.trim(); 
         headerCell.font = { bold: true, size: 14, color: { argb: 'FF0000FF' } }; // Azul
         headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
         headerCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '9AA5D4' }, // Amarillo
         };

         // Fila 2: Estado de cuenta
         worksheet.mergeCells('A2:K2');
         const estadoCell = worksheet.getCell('A2');
         const today = new Date();
         const formattedDate = `${today.getDate().toString().padStart(2, '0')}/${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getFullYear()}`;
         estadoCell.value = `Estado de Cuenta del 01/01/2006 hasta ${formattedDate}`;
         estadoCell.font = { size: 12, color: { argb: 'FF0000FF' } }; // Azul
         estadoCell.alignment = { horizontal: 'center', vertical: 'middle' };

         // Inserción de la imagen en la esquina superior izquierda
         const imagePath = path.resolve(__dirname, 'zayro.png'); // Cambia esto por la ruta de tu imagen
         const imageId = workbook.addImage({
            filename: imagePath,
            extension: 'png', // Cambia a jpg si es necesario
         });

         worksheet.addImage(imageId, {
            tl: { col: 0.5, row: 1 }, // Top-left corner: columna 0, fila 0
            ext: { width: 120, height: 100 } // Tamaño de la imagen en píxeles
         });

         // Fila 3: Dirección
         worksheet.mergeCells('A3:K3');
         const direccionCell = worksheet.getCell('A3');
         direccionCell.value = 'Hidalgo 3331 0 Col. Sector Centro Nuevo Laredo, TAMAULIPAS';
         direccionCell.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         direccionCell.alignment = { horizontal: 'center', vertical: 'middle' };

         // Fila 4: Teléfono
         worksheet.mergeCells('A4:K4');
         const telefonoCell = worksheet.getCell('A4');
         telefonoCell.value = 'Tel (867) 712-7048, Fax (867) 712-7064';
         telefonoCell.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         telefonoCell.alignment = { horizontal: 'center', vertical: 'middle' };

         // Fila 5: RFC
         worksheet.mergeCells('A5:K5');
         const rfcCell = worksheet.getCell('A5');
         rfcCell.value = 'RFC ZRS950417E15';
         rfcCell.font = { bold: true, size: 12, color: { argb: 'FF0000FF' } }; // Azul
         rfcCell.alignment = { horizontal: 'center', vertical: 'middle' };
         const formattedDatemesdiaanio = `${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getDate().toString().padStart(2, '0')}/${today.getFullYear()}`;
         // Fila 6: Fecha Reporte
         worksheet.mergeCells('I6:K6');
         const fechaReporteCell = worksheet.getCell('I6');
         fechaReporteCell.value = `Fecha Reporte ${formattedDatemesdiaanio}`;
         fechaReporteCell.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };

         // Fila 7: Cliente 
         worksheet.getCell('A8').value = 'Cliente: '+cliente.ClienteID+' '+cliente.Nombre;
         worksheet.getCell('A8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A8').alignment = { horizontal: 'left', vertical: 'middle' };
         //RFC CLIENTE Y DIRECCION
         worksheet.mergeCells('A9:E9');
         worksheet.getCell('A9').value = ini.RFC+' - '+ini.Direccion;
         worksheet.getCell('A9').font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A9').alignment = { horizontal: 'left', vertical: 'middle' };
         //Pagina
         worksheet.mergeCells('I7:K7');
         const fechaReporteCellpag = worksheet.getCell('I7');
         fechaReporteCellpag.value = `Página 1`;
         fechaReporteCellpag.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCellpag.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
         //SALDO INICIAL
         worksheet.mergeCells('G8:I8');
         worksheet.getCell('G8').value = 'SALDO INICIAL:';
         worksheet.getCell('G8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('G8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheet.mergeCells('J8:K8');
         worksheet.getCell('K8').value = ini.SaldoInicial;
         worksheet.getCell('K8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('K8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheet.getCell('K8').numFmt = '"$"#,##0.00'; 

          //Numero de Cuenta
          worksheet.mergeCells('G9:K9');
          worksheet.getCell('G9').value = 'Número de cuenta: BBVA 0447278063';
          worksheet.getCell('G9').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
          worksheet.getCell('G9').alignment = { horizontal: 'right', vertical: 'middle' };
         
         })
         // Fila para "Último Depósito"
         
         worksheet.mergeCells('A10:K10');
         const ultimoDepositoHeader = worksheet.getCell('A10');
         ultimoDepositoHeader.value = 'ULTIMO DEPOSITO';
         ultimoDepositoHeader.font = { bold: true, size: 12 };
         ultimoDepositoHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         let numfila=11;
         // Encabezados para la primera sección (último depósito) con texto negro y fondo verde fuerte
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'FECHA';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         worksheet.getCell('C'+numfila).value = 'POLIZA';
         worksheet.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('E'+numfila+':G'+numfila);
         worksheet.getCell('E'+numfila).value = 'BANCO';
         worksheet.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('H'+numfila).value = 'TIPO MOVIMIENTO';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('I'+numfila+':K'+numfila);
         worksheet.getCell('I'+numfila).value = 'IMPORTE';
         worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         numfila=numfila+1;
         // Datos de último depósito con texto negro y fondo verde fuerte
         const ultimodeposito=await sqlzam.ultimodepositocliente(cliente.ClienteID);
         if (ultimodeposito.length>0)
         {

        
            ultimodeposito.forEach(renglon=>{ 
               worksheet.mergeCells('A'+numfila+':B'+numfila);
               worksheet.getCell('A'+numfila).value = renglon.Fecha;
               worksheet.getCell('A'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('C'+numfila+':D'+numfila);
               worksheet.getCell('C'+numfila).value = renglon.PolizaAlone;
               worksheet.getCell('C'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('E'+numfila+':G'+numfila);
               worksheet.getCell('E'+numfila).value = renglon.Banco;
               worksheet.getCell('E'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('H'+numfila).value = renglon.TipoPol;
               worksheet.getCell('H'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('I'+numfila+':K'+numfila);
               worksheet.getCell('I'+numfila).value = renglon.Importe;
               worksheet.getCell('I'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('I'+numfila).numFmt = '"$"#,##0.00';
               numfila=numfila+1
            })
            const rangeverde = [];
            for (let fila = 11; fila < numfila; fila++) {
               for (let col of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) {
                  rangeverde.push(`${col}${fila}`);
               }
            }
            rangeverde.forEach(cell => {
               const currentCell = worksheet.getCell(cell);
               currentCell.border = {
                  top: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  left: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  bottom: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  right: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
               };
            });
      }
         let numfilainicial=numfila;
         worksheet.getCell('A'+numfila).value = 'FECHA';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('B'+numfila).value = 'POLIZA';
         worksheet.getCell('B'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('B'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         
         worksheet.getCell('C'+numfila).value = 'BANCO';
         worksheet.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheet.getCell('D'+numfila).value = 'TIPO MOVIMIENTO';
         worksheet.getCell('D'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheet.getCell('E'+numfila).value = 'IE';
         worksheet.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('F'+numfila).value = 'FACTURA';
         worksheet.getCell('F'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('F'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('G'+numfila).value = 'FOLIO INTERNO';
         worksheet.getCell('G'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('G'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('H'+numfila).value = 'PEDIMENTO';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('I'+numfila).value = 'PEDIDO';
         worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('J'+numfila).value = 'ANT';
         worksheet.getCell('J'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('J'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('K'+numfila).value = 'IMPORTE';
         worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('K'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
        
         numfila=numfila+1
         let referenciaanterior='';
         let saldotrafico=0;
         const result=await sqlzam.Rmensual_1_distinct(cliente.ClienteID);
         const re=await sqlzam.sp_Rmensual_1(cliente.ClienteID);
         if (re.length>0){
            result.forEach(renglonactual=>{
               
               worksheet.mergeCells('A'+numfila+':B'+numfila);
               worksheet.getCell('A'+numfila).value = 'TRAFICO: '+renglonactual.Referencia;
               worksheet.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('C'+numfila+':E'+numfila);
               worksheet.getCell('C'+numfila).value = 'Proveedor: '+renglonactual.Proveedor;
               worksheet.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('F'+numfila+':K'+numfila);
               worksheet.getCell('F'+numfila).value = 'Facturas: '+renglonactual.RefFactura;
               worksheet.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
               numfila=numfila+1
               
               re.forEach(x=>{
                  if (renglonactual.Referencia==x.Referencia){
                     worksheet.getCell('A'+numfila).value = x.Fecha//'30/08/2024  01:14:00 p. m.';//
                     worksheet.getCell('A'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
   
                     worksheet.getCell('B'+numfila).value = x.MovimientoID//'CXC';
                     worksheet.getCell('B'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                     if (x.MovimientoID.trim()=='CXC'){
                        worksheet.getCell('C'+numfila).value = x.UUID;
                        worksheet.getCell('C'+numfila).font = {  size: 6, color: { argb: 'FF000000' } };
                     }else{
                        worksheet.getCell('C'+numfila).value = x.Banco;
                        worksheet.getCell('C'+numfila).font = {  size: 9, color: { argb: 'FF000000' } }; // Negro
                     }
                     worksheet.getCell('D'+numfila).value = x.TipoPoliza;
                     worksheet.getCell('D'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                     worksheet.getCell('E'+numfila).value = x.IE//'I';
                     worksheet.getCell('E'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('F'+numfila).value = x.PolizaID//'167049';
                     worksheet.getCell('F'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('G'+numfila).value = x.FolioInterno;
                     worksheet.getCell('G'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('H'+numfila).value = x.Pedimento//'4005417';
                     worksheet.getCell('H'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('I'+numfila).value = x.Pedido//'MEX1268';
                     worksheet.getCell('I'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                     if (x.MovimientoID.trim()=='CXC'){
                        if(x.Antiguedad==90){
                           worksheet.getCell('J'+numfila).value = x.Antiguedad;
                           worksheet.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };
                           worksheet.getCell('J' + numfila).numFmt = '"+"#,##0'; 
                        }else{
                           worksheet.getCell('J'+numfila).value = x.Antiguedad;
                           worksheet.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };

                        }
                     }
                     
                     if(x.Saldo<0){
                        worksheet.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                        worksheet.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                        
                     }
                     else{
                        worksheet.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                        worksheet.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                     }
                     
                     saldotrafico=saldotrafico+x.Saldo
                     numfila=numfila+1
                  } 
               })
              
               worksheet.getCell('I'+numfila).value = 'Saldo del Tráfico';
               worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
               
               if(saldotrafico<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotrafico)//'1179.9';
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  
               }
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotrafico)//'1179.9';
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
               
               saldotrafico=0;
            numfila=numfila+1
         }) 
         }
  
         numfila=numfila+1 
         const resumen=await sqlzam.antiguedadsaldos(cliente.ClienteID);
         resumen.forEach(r=>{
            if (re.length>0){
               if(r.Saldo<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  
               }
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
            }
            
            numfila=numfila+3
            worksheet.mergeCells('C'+numfila+':I'+numfila);
            worksheet.getCell('D'+numfila).value = 'RESUMEN';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheet.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            numfila=numfila+1
            worksheet.mergeCells('D'+numfila+':F'+numfila);
            worksheet.getCell('D'+numfila).value = 'TOTAL A SU CARGO EN M.N.';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('G'+numfila+':H'+numfila);
            if (re.length>0){
               if(r.Saldo<0){
                  worksheet.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('G'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('G' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
               }
               else{
                  worksheet.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('G'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('G' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
            }
            numfila=numfila+1
            worksheet.mergeCells('D'+numfila+':H'+numfila);
            worksheet.getCell('D'+numfila).value = 'ANTIGUEDAD DE SALDOS';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheet.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            numfila=numfila+2
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = '$';
            worksheet.getCell('F'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('H'+numfila).value = '%';
            worksheet.getCell('H'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            numfila=numfila+1
            //30 dias
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo a 30 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo30);
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            let porcentaje = (r.Saldo30 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            numfila=numfila+1
            //31 a 60
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo de 31 días a 60 días';
            worksheet.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo60);
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            porcentaje = (r.Saldo60 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            numfila=numfila+1
            //61 a 90
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo de 61 días a 90 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo90);
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            porcentaje = (r.Saldo90 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00'; 
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            numfila=numfila+1
            //mas de 90
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Más de 90 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Mayor90);
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            porcentaje = (r.Mayor90 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
              
            numfila=numfila+1
            //ANTICIPOS
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Anticipos';
            worksheet.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            if(r.Deposito<0){
               worksheet.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheet.getCell('F' + numfila).numFmt = '"($"#,##0.00")"';
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
                

            }
            else{
               worksheet.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
                
            }
            
            numfila=numfila+2
         })
         
         inicial.forEach(async ini2=>{
         //SALDO A FAVOR 
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'SALDO A FAVOR';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         
         if(ini2.SaldoAFavor<0){
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheet.getCell('C' + numfila).numFmt = '"($"#,##0.00")"'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
         else{
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheet.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
 
         numfila=numfila+1
         //SALDO DEUDOR
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'SALDO DEUDOR';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
         worksheet.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
         worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         
      })
         //CARGOS ANTICIPOS SIN APLICAR
         const sinaplicar=await sqlzam.sp_cargossinaplicar(cliente.ClienteID);
         if (sinaplicar.length>0){
            numfila=numfila+2
            worksheet.mergeCells('A'+numfila+':D'+numfila);
            worksheet.getCell('A'+numfila).value = 'CARGOS/ANTICIPOS SIN APLICAR';
            worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; 
            numfila=numfila+1
            const sinaplicardistinct=await sqlzam.sp_cargossinaplicar_distinct(cliente.ClienteID)
            let saldotraf=0;
            sinaplicardistinct.forEach(sad=>{
            worksheet.mergeCells('A'+numfila+':B'+numfila);
            worksheet.getCell('A'+numfila).value = 'TRAFICO: '+sad.Referencia;
            worksheet.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('C'+numfila+':E'+numfila);
            worksheet.getCell('C'+numfila).value = 'Proveedor: '+sad.Proveedor;
            worksheet.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':K'+numfila);
            worksheet.getCell('F'+numfila).value = 'Facturas: '+sad.RefFactura;
            worksheet.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            numfila=numfila+1
            
               sinaplicar.forEach(sa=>{
                  worksheet.getCell('A'+numfila).value = sa.Fecha;
                  worksheet.getCell('A'+numfila).font = {  size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('B'+numfila).value = sa.PolizaID;
                  worksheet.getCell('B'+numfila).font = { bsize: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('C'+numfila).value = sa.Banco;
                  worksheet.getCell('C'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('D'+numfila).value = sa.TipoMovimiento;
                  worksheet.getCell('D'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
                  
                  worksheet.getCell('E'+numfila).value = sa.IE;
                  worksheet.getCell('E'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('F'+numfila).value = sa.Factura;
                  worksheet.getCell('F'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('G'+numfila).value = sa.FolioInterno;
                  worksheet.getCell('G'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('H'+numfila).value = sa.Pedimento;
                  worksheet.getCell('H'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('I'+numfila).value = sa.Pedido;
                  worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         
               
                  if(sa.Saldo<0){
                     worksheet.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheet.getCell('K' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
                     worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  else{
                     worksheet.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00""'; // Formato de moneda
                     worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  saldotraf=saldotraf+sa.Saldo
                  numfila=numfila+1
               })
               worksheet.getCell('I'+numfila).value = 'Saldo del Tráfico';
               worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
               
               if(saldotraf<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotraf)//'1179.9';
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  
                  
               }
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotraf)//'1179.9';
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                  
               }
            })

         } 
        numfila=numfila+2
         //PIE DE PAGINA
         worksheet.mergeCells('H'+numfila+':I'+numfila);
         worksheet.getCell('H'+numfila).value = 'FR-02-02-02';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         numfila=numfila+1
         worksheet.mergeCells('H'+numfila+':I'+numfila);
         worksheet.getCell('H'+numfila).value = 'Rev. 01';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('J'+numfila+':K'+numfila);
         const today = new Date();
         const formattedDate2 = `${today.getDate().toString().padStart(2, '0')}/${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getFullYear()}`;
         worksheet.getCell('J'+numfila).value = formattedDate2;
         worksheet.getCell('J'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
      
         let nombreOriginal =cliente.Nombre;
         let nombreLimpio = limpiarNombreArchivo(nombreOriginal);
         /******************************************************************************************* */
         // Guardar archivo
         let nombreArchivo='Estado_de_Cuenta_'+cliente.ClienteID+' '+nombreLimpio
         await workbook.xlsx.writeFile('Estado_de_Cuenta_'+cliente.ClienteID+' '+nombreLimpio+'.xlsx');


         const correos=await sqlzam.contactosestadoscuenta(cliente.ClienteID);
         correos.forEach(async co=>{
            await enviarMailNLD(nombreArchivo,transport,co.correos,nombreLimpio)
         })
         //
         console.log('Archivo creado exitosamente.');
         
      }
   }
   if(totalclientes==(clientes.length)){
      await res.json('Reportes Enviados')
   }
}); 
enviarMailNLD = async(nombreArchivo,transport, correos,nombreLimpio) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'cobranza@zayro.com;sistemas@zayro.com;'+correos,
      //to: 'programacion@zayro.com;',
      subject: 'Estado de cuenta '+nombreLimpio,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './' + nombreArchivo + '.xlsx',
         }],
      text: 'Estado de Cuenta Nuevo Laredo',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_estadosdecuentamxn', async function(req, res, next) {
   const clientes=await sqlzam.sp_clientesestadocuenta2();
   let config = {
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth: {
         user:process.env.useremail,
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false
      }
   }
   let transport = nodemailer.createTransport(config);
   let totalclientes=0
   for (let i = 0; i <clientes.length; i++) {
      totalclientes=totalclientes+1
      const cliente = clientes[i];
      

      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Estado de Cuenta');

      // Estilo general de la hoja
      worksheet.pageSetup = {
         margins: {
            left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3
         }
      };
      // Configurar el rango con bordes blancos
      const range = { startRow: 1, endRow: 200, startCol: 1, endCol: 26 }; // 26 columnas (A-Z)

      // Aplicar bordes blancos a cada celda en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.border = {
                  top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  right: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            };
         }
      }

      // Establecer un fondo blanco para todas las celdas en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFFFF' }, // Fondo blanco
            };
         }
      }

      // Configuración de ancho de columnas
      worksheet.columns = [
         { width: 12 },//A
         { width: 6.1 },//B
         { width: 22 },//C
         { width: 13.78 },//D
         { width: 3 },//E
         { width: 9.7 },//F
         { width: 13.5 },//G
         { width: 15 },//H
         { width: 10 },//I
         { width: 5 },//J
         { width: 15 }//K
      ];
      const inicial=await sqlzam.datosinicialescliente2(cliente.ClienteID);
      console.log(cliente.ClienteID)
      if(inicial.length>0){
         inicial.forEach(ini=>{

         
         // Fila 1: Encabezado con fondo amarillo
         worksheet.mergeCells('A1:K1');
         const headerCell = worksheet.getCell('A1');
         headerCell.value = ini.Sucursal.trim(); 
         headerCell.font = { bold: true, size: 14, color: { argb: 'FF0000FF' } }; // Azul
         headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
         headerCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '9AA5D4' }, // Amarillo
         };

         // Fila 2: Estado de cuenta
         worksheet.mergeCells('A2:K2');
         const estadoCell = worksheet.getCell('A2');
         const today = new Date();
         const formattedDate = `${today.getDate().toString().padStart(2, '0')}/${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getFullYear()}`;
         estadoCell.value = `Estado de Cuenta del 01/01/2006 hasta ${formattedDate}`;
         estadoCell.font = { size: 12, color: { argb: 'FF0000FF' } }; // Azul
         estadoCell.alignment = { horizontal: 'center', vertical: 'middle' };

         // Inserción de la imagen en la esquina superior izquierda
         const imagePath = path.resolve(__dirname, 'zayro.png'); // Cambia esto por la ruta de tu imagen
         const imageId = workbook.addImage({
            filename: imagePath,
            extension: 'png', // Cambia a jpg si es necesario
         });

         worksheet.addImage(imageId, {
            tl: { col: 0.5, row: 1 }, // Top-left corner: columna 0, fila 0
            ext: { width: 120, height: 100 } // Tamaño de la imagen en píxeles
         });

         // Fila 3: Dirección
         worksheet.mergeCells('A3:K3');
         const direccionCell = worksheet.getCell('A3');
         direccionCell.value = 'Hidalgo 3331 0 Col. Sector Centro Nuevo Laredo, TAMAULIPAS';
         direccionCell.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         direccionCell.alignment = { horizontal: 'center', vertical: 'middle' };

         // Fila 4: Teléfono
         worksheet.mergeCells('A4:K4');
         const telefonoCell = worksheet.getCell('A4');
         telefonoCell.value = 'Tel (867) 712-7048, Fax (867) 712-7064';
         telefonoCell.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         telefonoCell.alignment = { horizontal: 'center', vertical: 'middle' };

         // Fila 5: RFC
         worksheet.mergeCells('A5:K5');
         const rfcCell = worksheet.getCell('A5');
         rfcCell.value = 'RFC ZRS950417E15';
         rfcCell.font = { bold: true, size: 12, color: { argb: 'FF0000FF' } }; // Azul
         rfcCell.alignment = { horizontal: 'center', vertical: 'middle' };
         const formattedDatemesdiaanio = `${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getDate().toString().padStart(2, '0')}/${today.getFullYear()}`;
         // Fila 6: Fecha Reporte
         worksheet.mergeCells('I6:K6');
         const fechaReporteCell = worksheet.getCell('I6');
         fechaReporteCell.value = `Fecha Reporte ${formattedDatemesdiaanio}`;
         fechaReporteCell.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };

         // Fila 7: Cliente 
         worksheet.getCell('A8').value = 'Cliente: '+cliente.ClienteID+' '+cliente.Nombre;
         worksheet.getCell('A8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A8').alignment = { horizontal: 'left', vertical: 'middle' };
         //RFC CLIENTE Y DIRECCION
         worksheet.mergeCells('A9:E9');
         worksheet.getCell('A9').value = ini.RFC+' - '+ini.Direccion;
         worksheet.getCell('A9').font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A9').alignment = { horizontal: 'left', vertical: 'middle' };
         //Pagina
         worksheet.mergeCells('I7:K7');
         const fechaReporteCellpag = worksheet.getCell('I7');
         fechaReporteCellpag.value = `Página 1`;
         fechaReporteCellpag.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCellpag.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
         //SALDO INICIAL
         worksheet.mergeCells('G8:I8');
         worksheet.getCell('G8').value = 'SALDO INICIAL:';
         worksheet.getCell('G8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('G8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheet.mergeCells('J8:K8');
         worksheet.getCell('K8').value = ini.SaldoInicial;
         worksheet.getCell('K8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('K8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheet.getCell('K8').numFmt = '"$"#,##0.00'; 

          //Numero de Cuenta
          worksheet.mergeCells('G9:K9');
          worksheet.getCell('G9').value = 'Número de cuenta: Banamex 9892847077';
          worksheet.getCell('G9').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
          worksheet.getCell('G9').alignment = { horizontal: 'right', vertical: 'middle' };
         
         })
         // Fila para "Último Depósito"
         
         worksheet.mergeCells('A10:K10');
         const ultimoDepositoHeader = worksheet.getCell('A10');
         ultimoDepositoHeader.value = 'ULTIMO DEPOSITO';
         ultimoDepositoHeader.font = { bold: true, size: 12 };
         ultimoDepositoHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         let numfila=11;
         // Encabezados para la primera sección (último depósito) con texto negro y fondo verde fuerte
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'FECHA';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         worksheet.getCell('C'+numfila).value = 'POLIZA';
         worksheet.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('E'+numfila+':G'+numfila);
         worksheet.getCell('E'+numfila).value = 'BANCO';
         worksheet.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('H'+numfila).value = 'TIPO MOVIMIENTO';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('I'+numfila+':K'+numfila);
         worksheet.getCell('I'+numfila).value = 'IMPORTE';
         worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         numfila=numfila+1;
         // Datos de último depósito con texto negro y fondo verde fuerte
         const ultimodeposito=await sqlzam.ultimodepositocliente2(cliente.ClienteID);
         if (ultimodeposito.length>0)
         {

        
            ultimodeposito.forEach(renglon=>{ 
               worksheet.mergeCells('A'+numfila+':B'+numfila);
               worksheet.getCell('A'+numfila).value = renglon.Fecha;
               worksheet.getCell('A'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('C'+numfila+':D'+numfila);
               worksheet.getCell('C'+numfila).value = renglon.PolizaAlone;
               worksheet.getCell('C'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('E'+numfila+':G'+numfila);
               worksheet.getCell('E'+numfila).value = renglon.Banco;
               worksheet.getCell('E'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('H'+numfila).value = renglon.TipoPol;
               worksheet.getCell('H'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('I'+numfila+':K'+numfila);
               worksheet.getCell('I'+numfila).value = renglon.Importe;
               worksheet.getCell('I'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('I'+numfila).numFmt = '"$"#,##0.00';
               numfila=numfila+1
            })
            const rangeverde = [];
            for (let fila = 11; fila < numfila; fila++) {
               for (let col of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) {
                  rangeverde.push(`${col}${fila}`);
               }
            }
            rangeverde.forEach(cell => {
               const currentCell = worksheet.getCell(cell);
               currentCell.border = {
                  top: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  left: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  bottom: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  right: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
               };
            });
      }
         let numfilainicial=numfila;
         worksheet.getCell('A'+numfila).value = 'FECHA';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('B'+numfila).value = 'POLIZA';
         worksheet.getCell('B'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('B'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         
         worksheet.getCell('C'+numfila).value = 'BANCO';
         worksheet.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheet.getCell('D'+numfila).value = 'TIPO MOVIMIENTO';
         worksheet.getCell('D'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheet.getCell('E'+numfila).value = 'IE';
         worksheet.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('F'+numfila).value = 'FACTURA';
         worksheet.getCell('F'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('F'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('G'+numfila).value = 'FOLIO INTERNO';
         worksheet.getCell('G'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('G'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('H'+numfila).value = 'PEDIMENTO';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('I'+numfila).value = 'PEDIDO';
         worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('J'+numfila).value = 'ANT';
         worksheet.getCell('J'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('J'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('K'+numfila).value = 'IMPORTE';
         worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('K'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
        
         numfila=numfila+1
         let referenciaanterior='';
         let saldotrafico=0;
         const result=await sqlzam.Rmensual_2_distinct(cliente.ClienteID);
         const re=await sqlzam.sp_Rmensual_2(cliente.ClienteID);
         if (re.length>0){
            result.forEach(renglonactual=>{
               
               worksheet.mergeCells('A'+numfila+':B'+numfila);
               worksheet.getCell('A'+numfila).value = 'TRAFICO: '+renglonactual.Referencia;
               worksheet.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('C'+numfila+':E'+numfila);
               worksheet.getCell('C'+numfila).value = 'Proveedor: '+renglonactual.Proveedor;
               worksheet.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('F'+numfila+':K'+numfila);
               worksheet.getCell('F'+numfila).value = 'Facturas: '+renglonactual.RefFactura;
               worksheet.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
               numfila=numfila+1
               
               re.forEach(x=>{
                  if (renglonactual.Referencia==x.Referencia){
                     worksheet.getCell('A'+numfila).value = x.Fecha//'30/08/2024  01:14:00 p. m.';//
                     worksheet.getCell('A'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
   
                     worksheet.getCell('B'+numfila).value = x.MovimientoID//'CXC';
                     worksheet.getCell('B'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                     if (x.MovimientoID.trim()=='CXC'){
                        worksheet.getCell('C'+numfila).value = x.UUID;
                        worksheet.getCell('C'+numfila).font = {  size: 6, color: { argb: 'FF000000' } };
                     }else{
                        worksheet.getCell('C'+numfila).value = x.Banco;
                        worksheet.getCell('C'+numfila).font = {  size: 9, color: { argb: 'FF000000' } }; // Negro
                     }
                     worksheet.getCell('D'+numfila).value = x.TipoPoliza;
                     worksheet.getCell('D'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                     worksheet.getCell('E'+numfila).value = x.IE//'I';
                     worksheet.getCell('E'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('F'+numfila).value = x.PolizaID//'167049';
                     worksheet.getCell('F'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('G'+numfila).value = x.FolioInterno;
                     worksheet.getCell('G'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('H'+numfila).value = x.Pedimento//'4005417';
                     worksheet.getCell('H'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                     worksheet.getCell('I'+numfila).value = x.Pedido//'MEX1268';
                     worksheet.getCell('I'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                     if (x.MovimientoID.trim()=='CXC'){
                        if(x.Antiguedad==90){
                           worksheet.getCell('J'+numfila).value = x.Antiguedad;
                           worksheet.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };
                           worksheet.getCell('J' + numfila).numFmt = '"+"#,##0'; 
                        }else{
                           worksheet.getCell('J'+numfila).value = x.Antiguedad;
                           worksheet.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };

                        }
                     }
                     
                     if(x.Saldo<0){
                        worksheet.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                        worksheet.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                        
                     }
                     else{
                        worksheet.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                        worksheet.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                     }
                     
                     saldotrafico=saldotrafico+x.Saldo
                     numfila=numfila+1
                  } 
               })
              
               worksheet.getCell('I'+numfila).value = 'Saldo del Tráfico';
               worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
               
               if(saldotrafico<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotrafico)//'1179.9';
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  
               }
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotrafico)//'1179.9';
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
               
               saldotrafico=0;
            numfila=numfila+1
         }) 
         }
  
         numfila=numfila+1 
         const resumen=await sqlzam.antiguedadsaldos2(cliente.ClienteID);
         resumen.forEach(r=>{
            if (re.length>0){
               if(r.Saldo<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  
               }
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
            }
            
            numfila=numfila+3
            worksheet.mergeCells('C'+numfila+':I'+numfila);
            worksheet.getCell('D'+numfila).value = 'RESUMEN';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheet.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            numfila=numfila+1
            worksheet.mergeCells('D'+numfila+':F'+numfila);
            worksheet.getCell('D'+numfila).value = 'TOTAL A SU CARGO EN M.N.';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('G'+numfila+':H'+numfila);
            if (re.length>0){
               if(r.Saldo<0){
                  worksheet.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('G'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('G' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
               }
               else{
                  worksheet.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('G'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('G' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
            }
            numfila=numfila+1
            worksheet.mergeCells('D'+numfila+':H'+numfila);
            worksheet.getCell('D'+numfila).value = 'ANTIGUEDAD DE SALDOS';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheet.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            numfila=numfila+2
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = '$';
            worksheet.getCell('F'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('H'+numfila).value = '%';
            worksheet.getCell('H'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            numfila=numfila+1
            //30 dias
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo a 30 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo30);
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            let porcentaje = (r.Saldo30 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            numfila=numfila+1
            //31 a 60
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo de 31 días a 60 días';
            worksheet.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo60);
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            porcentaje = (r.Saldo60 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            numfila=numfila+1
            //61 a 90
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo de 61 días a 90 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo90);
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            porcentaje = (r.Saldo90 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00'; 
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            numfila=numfila+1
            //mas de 90
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Más de 90 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = Math.abs(r.Mayor90);
            worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            porcentaje = (r.Mayor90 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
              
            numfila=numfila+1
            //ANTICIPOS
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Anticipos';
            worksheet.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            if(r.Deposito<0){
               worksheet.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheet.getCell('F' + numfila).numFmt = '"($"#,##0.00")"';
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
                

            }
            else{
               worksheet.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
                
            }
            
            numfila=numfila+2
         })
         
         inicial.forEach(async ini2=>{
         //SALDO A FAVOR 
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'SALDO A FAVOR';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         
         if(ini2.SaldoAFavor<0){
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheet.getCell('C' + numfila).numFmt = '"($"#,##0.00")"'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
         else{
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheet.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
 
         numfila=numfila+1
         //SALDO DEUDOR
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'SALDO DEUDOR';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
         worksheet.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
         worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         
      })
         //CARGOS ANTICIPOS SIN APLICAR
         const sinaplicar=await sqlzam.sp_cargossinaplicar2(cliente.ClienteID);
         if (sinaplicar.length>0){
            numfila=numfila+2
            worksheet.mergeCells('A'+numfila+':D'+numfila);
            worksheet.getCell('A'+numfila).value = 'CARGOS/ANTICIPOS SIN APLICAR';
            worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; 
            numfila=numfila+1
            const sinaplicardistinct=await sqlzam.sp_cargossinaplicar_distinct2(cliente.ClienteID)
            let saldotraf=0;
            if (sinaplicardistinct.length>0)
            {
            sinaplicardistinct.forEach(sad=>{
            worksheet.mergeCells('A'+numfila+':B'+numfila);
            worksheet.getCell('A'+numfila).value = 'TRAFICO: '+sad.Referencia;
            worksheet.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('C'+numfila+':E'+numfila);
            worksheet.getCell('C'+numfila).value = 'Proveedor: '+sad.Proveedor;
            worksheet.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':K'+numfila);
            worksheet.getCell('F'+numfila).value = 'Facturas: '+sad.RefFactura;
            worksheet.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            numfila=numfila+1
            
               sinaplicar.forEach(sa=>{
                  worksheet.getCell('A'+numfila).value = sa.Fecha;
                  worksheet.getCell('A'+numfila).font = {  size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('B'+numfila).value = sa.PolizaID;
                  worksheet.getCell('B'+numfila).font = { bsize: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('C'+numfila).value = sa.Banco;
                  worksheet.getCell('C'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('D'+numfila).value = sa.TipoMovimiento;
                  worksheet.getCell('D'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
                  
                  worksheet.getCell('E'+numfila).value = sa.IE;
                  worksheet.getCell('E'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('F'+numfila).value = sa.Factura;
                  worksheet.getCell('F'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('G'+numfila).value = sa.FolioInterno;
                  worksheet.getCell('G'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('H'+numfila).value = sa.Pedimento;
                  worksheet.getCell('H'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('I'+numfila).value = sa.Pedido;
                  worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         
               
                  if(sa.Saldo<0){
                     worksheet.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheet.getCell('K' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
                     worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  else{
                     worksheet.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00""'; // Formato de moneda
                     worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  saldotraf=saldotraf+sa.Saldo
                  numfila=numfila+1
               })
               worksheet.getCell('I'+numfila).value = 'Saldo del Tráfico';
               worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
               
               if(saldotraf<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotraf)//'1179.9';
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  
                  
               }
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(saldotraf)//'1179.9';
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                  
               }
            })
         }

         } 
        numfila=numfila+2
         //PIE DE PAGINA
         worksheet.mergeCells('H'+numfila+':I'+numfila);
         worksheet.getCell('H'+numfila).value = 'FR-02-02-02';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         numfila=numfila+1
         worksheet.mergeCells('H'+numfila+':I'+numfila);
         worksheet.getCell('H'+numfila).value = 'Rev. 01';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('J'+numfila+':K'+numfila);
         const today = new Date();
         const formattedDate2 = `${today.getDate().toString().padStart(2, '0')}/${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getFullYear()}`;
         worksheet.getCell('J'+numfila).value = formattedDate2;
         worksheet.getCell('J'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
      
         let nombreOriginal =cliente.Nombre;
         let nombreLimpio = limpiarNombreArchivo(nombreOriginal);
         /******************************************************************************************* */
         // Guardar archivo
         let nombreArchivo='Estado_de_Cuenta_MXN_'+cliente.ClienteID+' '+nombreLimpio
         await workbook.xlsx.writeFile('Estado_de_Cuenta_MXN_'+cliente.ClienteID+' '+nombreLimpio+'.xlsx');

         //
         const correos=await sqlzam.contactosestadoscuenta(cliente.ClienteID);
         correos.forEach(async co=>{
            await enviarMailMXN(nombreArchivo,transport,co.correos,nombreLimpio)
         })
         console.log('Archivo creado exitosamente.');
         
      }
   }
   if(totalclientes==(clientes.length)){
      await res.json('Reportes Enviados')
   }
}); 
enviarMailMXN = async(nombreArchivo,transport, correos,nombreLimpio) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'cobranza@zayro.com;sistemas@zayro.com;'+correos,
      //to: 'programacion@zayro.com;',
      subject: 'Estado de cuenta Sucursal Aeropuerto '+nombreLimpio,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './' + nombreArchivo + '.xlsx',
         }],
      text: 'Estado de Cuenta Mensual Sucursal Aeropuerto',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_estadosdecuentadll', async function(req, res, next) {
   const clientes=await sqlzay.sp_clientesestadocuenta();
   let config = {
      host:process.env.hostemail,
      port:process.env.portemail,
      secure: true,
      auth: {
         user:process.env.useremail,
         pass:process.env.passemail
      },
      tls: {
         rejectUnauthorized: false
      }
   }
   let transport = nodemailer.createTransport(config); 
   let totalclientes=0
   for (let i = 0; i <clientes.length; i++) {
      totalclientes=totalclientes+1
      const cliente = clientes[i];
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Estado de Cuenta');
      const worksheetingles = workbook.addWorksheet('Account Statement');

      // Estilo general de la hoja
      worksheet.pageSetup = {
         margins: {
            left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3
         }
      };
      worksheetingles.pageSetup = {
         margins: {
            left: 0.5, right: 0.5, top: 0.5, bottom: 0.5, header: 0.3, footer: 0.3
         }
      };
      // Configurar el rango con bordes blancos
      const range = { startRow: 1, endRow: 200, startCol: 1, endCol: 26 }; // 26 columnas (A-Z)

      // Aplicar bordes blancos a cada celda en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.border = {
                  top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  right: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            };
         }
      }
      // Aplicar bordes blancos a cada celda en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheetingles.getCell(row, col);
            cell.border = {
                  top: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  left: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  bottom: { style: 'thin', color: { argb: 'FFFFFFFF' } },
                  right: { style: 'thin', color: { argb: 'FFFFFFFF' } },
            };
         }
      }

      // Establecer un fondo blanco para todas las celdas en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheet.getCell(row, col);
            cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFFFF' }, // Fondo blanco
            };
         }
      }
      // Establecer un fondo blanco para todas las celdas en el rango
      for (let row = range.startRow; row <= range.endRow; row++) {
         for (let col = range.startCol; col <= range.endCol; col++) {
            const cell = worksheetingles.getCell(row, col);
            cell.fill = {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFFFFFFF' }, // Fondo blanco
            };
         }
      }

      // Configuración de ancho de columnas
      worksheet.columns = [
         { width: 12 },//A
         { width: 6.1 },//B
         { width: 22 },//C
         { width: 13.78 },//D
         { width: 3 },//E
         { width: 9.7 },//F
         { width: 13.5 },//G
         { width: 15 },//H
         { width: 10 },//I
         { width: 5 },//J
         { width: 15 }//K
      ];
      // Configuración de ancho de columnas
      worksheetingles.columns = [
         { width: 12 },//A
         { width: 6.1 },//B
         { width: 22 },//C
         { width: 13.78 },//D
         { width: 3 },//E
         { width: 9.7 },//F
         { width: 13.5 },//G
         { width: 15 },//H
         { width: 10 },//I
         { width: 5 },//J
         { width: 15 }//K
      ];
      const inicial=await sqlzay.datosinicialescliente(cliente.ClienteID);
      console.log(cliente.ClienteID)
      if(inicial.length>0){
         inicial.forEach(ini=>{

         
         // Fila 1: Encabezado con fondo amarillo
         worksheet.mergeCells('A1:K1');
         const headerCell = worksheet.getCell('A1');
         headerCell.value = ini.Sucursal.trim(); 
         headerCell.font = { bold: true, size: 14, color: { argb: 'FF0000FF' } }; // Azul
         headerCell.alignment = { horizontal: 'center', vertical: 'middle' };
         headerCell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '9AA5D4' }, // Amarillo
         };
         worksheetingles.mergeCells('A1:K1');
         const headerCellingles = worksheetingles.getCell('A1');
         headerCellingles.value = ini.Sucursal.trim(); 
         headerCellingles.font = { bold: true, size: 14, color: { argb: 'FF0000FF' } }; // Azul
         headerCellingles.alignment = { horizontal: 'center', vertical: 'middle' };
         headerCellingles.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: '9AA5D4' }, // Amarillo
         };

         // Fila 2: Estado de cuenta
         worksheet.mergeCells('A2:K2');
         const estadoCell = worksheet.getCell('A2');
         const today = new Date();
         const formattedDate = `${today.getDate().toString().padStart(2, '0')}/${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getFullYear()}`;
         estadoCell.value = `Estado de Cuenta del 01/01/2006 hasta ${formattedDate}`;
         estadoCell.font = { size: 12, color: { argb: 'FF0000FF' } }; // Azul
         estadoCell.alignment = { horizontal: 'center', vertical: 'middle' };
         // Fila 2: Estado de cuenta
         worksheetingles.mergeCells('A2:K2');
         const estadoCellingles = worksheetingles.getCell('A2');
         const todayingles = new Date();
         const formattedDateingles = `${todayingles.getDate().toString().padStart(2, '0')}/${(todayingles.getMonth() + 1).toString().padStart(2, '0')}/${todayingles.getFullYear()}`;
         estadoCellingles.value = `Account Statement from 01/01/2006 to ${formattedDateingles}`;
         estadoCellingles.font = { size: 12, color: { argb: 'FF0000FF' } }; // Azul
         estadoCellingles.alignment = { horizontal: 'center', vertical: 'middle' };

         // Inserción de la imagen en la esquina superior izquierda
         const imagePath = path.resolve(__dirname, 'zayro.png'); // Cambia esto por la ruta de tu imagen
         const imageId = workbook.addImage({
            filename: imagePath,
            extension: 'png', // Cambia a jpg si es necesario
         });

         worksheet.addImage(imageId, {
            tl: { col: 0.5, row: 1 }, // Top-left corner: columna 0, fila 0
            ext: { width: 120, height: 100 } // Tamaño de la imagen en píxeles
         });
         worksheetingles.addImage(imageId, {
            tl: { col: 0.5, row: 1 }, // Top-left corner: columna 0, fila 0
            ext: { width: 120, height: 100 } // Tamaño de la imagen en píxeles
         });


         // Fila 3: Dirección
         worksheet.mergeCells('A3:K3');
         const direccionCell = worksheet.getCell('A3');
         direccionCell.value = '4113 FREE TRADE ST LAREDO TEXAS';
         direccionCell.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         direccionCell.alignment = { horizontal: 'center', vertical: 'middle' };
         // Fila 3: Dirección
         worksheetingles.mergeCells('A3:K3');
         const direccionCellingles = worksheetingles.getCell('A3');
         direccionCellingles.value = '4113 FREE TRADE ST LAREDO TEXAS';
         direccionCellingles.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         direccionCellingles.alignment = { horizontal: 'center', vertical: 'middle' };

         // Fila 4: Teléfono
         worksheet.mergeCells('A4:K4');
         const telefonoCell = worksheet.getCell('A4');
         telefonoCell.value = 'Tel (956) 717-5044, Fax (956) 717-5040';
         telefonoCell.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         telefonoCell.alignment = { horizontal: 'center', vertical: 'middle' };
         // Fila 4: Teléfono
         worksheetingles.mergeCells('A4:K4');
         const telefonoCellingles = worksheetingles.getCell('A4');
         telefonoCellingles.value = 'Tel (956) 717-5044, Fax (956) 717-5040';
         telefonoCellingles.font = { size: 11, color: { argb: 'FF0000FF' } }; // Azul
         telefonoCellingles.alignment = { horizontal: 'center', vertical: 'middle' };

        
         
         const formattedDatemesdiaanio = `${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getDate().toString().padStart(2, '0')}/${today.getFullYear()}`;
         // Fila 6: Fecha Reporte
         worksheet.mergeCells('I6:K6');
         const fechaReporteCell = worksheet.getCell('I6');
         fechaReporteCell.value = `Fecha Reporte ${formattedDatemesdiaanio}`;
         fechaReporteCell.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCell.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
         worksheetingles.mergeCells('I6:K6');
         const fechaReporteCellingles = worksheetingles.getCell('I6');
         fechaReporteCellingles.value = `Report Date ${formattedDatemesdiaanio}`;
         fechaReporteCellingles.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCellingles.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };

         // Fila 7: Cliente 
         worksheet.getCell('A8').value = 'Cliente: '+cliente.ClienteID+' '+cliente.Nombre;
         worksheet.getCell('A8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A8').alignment = { horizontal: 'left', vertical: 'middle' };
         // Fila 7: Cliente 
         worksheetingles.getCell('A8').value = 'Client: '+cliente.ClienteID+' '+cliente.Nombre;
         worksheetingles.getCell('A8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('A8').alignment = { horizontal: 'left', vertical: 'middle' }
         //RFC CLIENTE Y DIRECCION
         worksheet.mergeCells('A9:E9');
         worksheet.getCell('A9').value = ini.RFC+' - '+ini.Direccion;
         worksheet.getCell('A9').font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A9').alignment = { horizontal: 'left', vertical: 'middle' };
         //RFC CLIENTE Y DIRECCION
         worksheetingles.mergeCells('A9:E9');
         worksheetingles.getCell('A9').value = ini.RFC+' - '+ini.Direccion;
         worksheetingles.getCell('A9').font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('A9').alignment = { horizontal: 'left', vertical: 'middle' };
         //Pagina
         worksheet.mergeCells('I7:K7');
         const fechaReporteCellpag = worksheet.getCell('I7');
         fechaReporteCellpag.value = `Página 1`;
         fechaReporteCellpag.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCellpag.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
         //Pagina
         worksheetingles.mergeCells('I7:K7');
         const fechaReporteCellpagingles = worksheetingles.getCell('I7');
         fechaReporteCellpagingles.value = `Page 1`;
         fechaReporteCellpagingles.font = { size: 10, color: { argb: 'FF000000' } }; // Negro
         fechaReporteCellpagingles.alignment = { horizontal: 'right', vertical: 'middle', wrapText: true };
         //SALDO INICIAL
         worksheet.mergeCells('G8:I8');
         worksheet.getCell('G8').value = 'SALDO INICIAL:';
         worksheet.getCell('G8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('G8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheet.mergeCells('J8:K8');
         worksheet.getCell('K8').value = ini.SaldoInicial;
         worksheet.getCell('K8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('K8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheet.getCell('K8').numFmt = '"$"#,##0.00'; 
         //SALDO INICIAL
         worksheetingles.mergeCells('G8:I8');
         worksheetingles.getCell('G8').value = 'INITIAL BALANCE:';
         worksheetingles.getCell('G8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('G8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheetingles.mergeCells('J8:K8');
         worksheetingles.getCell('K8').value = ini.SaldoInicial;
         worksheetingles.getCell('K8').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('K8').alignment = { horizontal: 'right', vertical: 'middle' };
         worksheetingles.getCell('K8').numFmt = '"$"#,##0.00'; 

          //Numero de Cuenta
          worksheet.mergeCells('G9:K9');
          worksheet.getCell('G9').value = 'Número de cuenta: IBC 2115650042';
          worksheet.getCell('G9').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
          worksheet.getCell('G9').alignment = { horizontal: 'right', vertical: 'middle' };
          //Numero de Cuenta
          worksheetingles.mergeCells('G9:K9');
          worksheetingles.getCell('G9').value = 'Account Number: IBC 2115650042';
          worksheetingles.getCell('G9').font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
          worksheetingles.getCell('G9').alignment = { horizontal: 'right', vertical: 'middle' };
         
         })
         
         // Fila para "Último Depósito"
         
         worksheet.mergeCells('A10:K10');
         const ultimoDepositoHeader = worksheet.getCell('A10');
         ultimoDepositoHeader.value = 'ULTIMO DEPOSITO';
         ultimoDepositoHeader.font = { bold: true, size: 12 };
         ultimoDepositoHeader.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         // Fila para "Último Depósito"
         worksheetingles.mergeCells('A10:K10');
         const ultimoDepositoHeaderingles = worksheetingles.getCell('A10');
         ultimoDepositoHeaderingles.value = 'LAST DEPOSIT';
         ultimoDepositoHeaderingles.font = { bold: true, size: 12 };
         ultimoDepositoHeaderingles.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         let numfila=11;
         // Encabezados para la primera sección (último depósito) con texto negro y fondo verde fuerte
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'FECHA';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         worksheet.getCell('C'+numfila).value = 'POLIZA';
         worksheet.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('E'+numfila+':G'+numfila);
         worksheet.getCell('E'+numfila).value = 'BANCO';
         worksheet.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('H'+numfila).value = 'TIPO MOVIMIENTO';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheet.mergeCells('I'+numfila+':K'+numfila);
         worksheet.getCell('I'+numfila).value = 'IMPORTE';
         worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         
         worksheetingles.mergeCells('A'+numfila+':B'+numfila);
         worksheetingles.getCell('A'+numfila).value = 'DATE';
         worksheetingles.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheetingles.mergeCells('C'+numfila+':D'+numfila);
         worksheetingles.getCell('C'+numfila).value = 'VOUCHER';
         worksheetingles.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheetingles.mergeCells('E'+numfila+':G'+numfila);
         worksheetingles.getCell('E'+numfila).value = 'BANK';
         worksheetingles.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('H'+numfila).value = 'TRANSACTION TYPE';
         worksheetingles.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         worksheetingles.mergeCells('I'+numfila+':K'+numfila);
         worksheetingles.getCell('I'+numfila).value = 'AMOUNT';
         worksheetingles.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         numfila=numfila+1;
         // Datos de último depósito con texto negro y fondo verde fuerte
         const ultimodeposito=await sqlzay.ultimodepositocliente(cliente.ClienteID);
         if (ultimodeposito.length>0)
         {
            ultimodeposito.forEach(renglon=>{  
               worksheet.mergeCells('A'+numfila+':B'+numfila);
               worksheet.getCell('A'+numfila).value = renglon.Fecha;
               worksheet.getCell('A'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('C'+numfila+':D'+numfila);
               worksheet.getCell('C'+numfila).value = renglon.PolizaAlone;
               worksheet.getCell('C'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('E'+numfila+':G'+numfila);
               worksheet.getCell('E'+numfila).value = renglon.Banco;
               worksheet.getCell('E'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('H'+numfila).value = renglon.TipoPol;
               worksheet.getCell('H'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.mergeCells('I'+numfila+':K'+numfila);
               worksheet.getCell('I'+numfila).value = renglon.Importe;
               worksheet.getCell('I'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('I'+numfila).numFmt = '"$"#,##0.00';
               
               worksheetingles.mergeCells('A'+numfila+':B'+numfila);
               worksheetingles.getCell('A'+numfila).value = renglon.Fecha;
               worksheetingles.getCell('A'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.mergeCells('C'+numfila+':D'+numfila);
               worksheetingles.getCell('C'+numfila).value = renglon.PolizaAlone;
               worksheetingles.getCell('C'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.mergeCells('E'+numfila+':G'+numfila);
               worksheetingles.getCell('E'+numfila).value = renglon.Banco;
               worksheetingles.getCell('E'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('H'+numfila).value = renglon.TipoPol;
               worksheetingles.getCell('H'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.mergeCells('I'+numfila+':K'+numfila);
               worksheetingles.getCell('I'+numfila).value = renglon.Importe;
               worksheetingles.getCell('I'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('I'+numfila).numFmt = '"$"#,##0.00';
               numfila=numfila+1
            })
            const rangeverde = [];
            for (let fila = 11; fila < numfila; fila++) {
               for (let col of ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K']) {
                  rangeverde.push(`${col}${fila}`);
               }
            }
            rangeverde.forEach(cell => {
               const currentCell = worksheet.getCell(cell);
               currentCell.border = {
                  top: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  left: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  bottom: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  right: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
               };
            });
            rangeverde.forEach(cell => {
               const currentCell = worksheetingles.getCell(cell);
               currentCell.border = {
                  top: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  left: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  bottom: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
                  right: { style: 'thin', color: { argb: 'FF008000' } }, // Verde fuerte
               };
            });
      }
         numfila=numfila+1
         worksheet.getCell('A'+numfila).value = 'FECHA';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('B'+numfila).value = 'POLIZA';
         worksheet.getCell('B'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('B'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         
         worksheet.getCell('C'+numfila).value = 'BANCO';
         worksheet.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheet.getCell('D'+numfila).value = 'TIPO MOVIMIENTO';
         worksheet.getCell('D'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheet.getCell('E'+numfila).value = 'IE';
         worksheet.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('F'+numfila).value = 'FACTURA';
         worksheet.getCell('F'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('F'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('G'+numfila).value = 'FOLIO INTERNO';
         worksheet.getCell('G'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('G'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('H'+numfila).value = 'PEDIMENTO';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('I'+numfila).value = 'PEDIDO';
         worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('J'+numfila).value = 'ANT';
         worksheet.getCell('J'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('J'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheet.getCell('K'+numfila).value = 'IMPORTE';
         worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheet.getCell('K'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         worksheetingles.getCell('A'+numfila).value = 'DATE';
         worksheetingles.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('B'+numfila).value = 'VOUCHER';
         worksheetingles.getCell('B'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('B'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         
         worksheetingles.getCell('C'+numfila).value = 'BANK';
         worksheetingles.getCell('C'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('C'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheetingles.getCell('D'+numfila).value = 'TRANSACTION TYPE';
         worksheetingles.getCell('D'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte


         worksheetingles.getCell('E'+numfila).value = 'IE';
         worksheetingles.getCell('E'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('E'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('F'+numfila).value = 'INVOICE';
         worksheetingles.getCell('F'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('F'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('G'+numfila).value = 'INTERNAL FOLIO';
         worksheetingles.getCell('G'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('G'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('H'+numfila).value = 'PEDIMENTO';
         worksheetingles.getCell('H'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('H'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('I'+numfila).value = 'ORDER';
         worksheetingles.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('I'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
         
         worksheetingles.getCell('J'+numfila).value = 'ANT';
         worksheetingles.getCell('J'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('J'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         worksheetingles.getCell('K'+numfila).value = 'AMOUNT';
         worksheetingles.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('K'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte

         numfila=numfila+1

         let referenciaanterior='';
         let saldotrafico=0;
         const result=await sqlzay.Rmensual_1_distinct(cliente.ClienteID);
         const re=await sqlzay.sp_Rmensual_1(cliente.ClienteID);
         if (result.length>0){
            if (re.length>0){
               
               result.forEach(renglonactual=>{
               
                  worksheet.mergeCells('A'+numfila+':B'+numfila);
                  worksheet.getCell('A'+numfila).value = 'TRAFICO: '+renglonactual.Referencia;
                  worksheet.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.mergeCells('C'+numfila+':E'+numfila);
                  worksheet.getCell('C'+numfila).value = 'Proveedor: '+renglonactual.Proveedor;
                  worksheet.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
                  worksheet.mergeCells('F'+numfila+':K'+numfila);
                  worksheet.getCell('F'+numfila).value = 'Facturas: '+renglonactual.RefFactura;
                  worksheet.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.mergeCells('A'+numfila+':B'+numfila);
                  worksheetingles.getCell('A'+numfila).value = 'TRAFFIC: '+renglonactual.Referencia;
                  worksheetingles.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.mergeCells('C'+numfila+':E'+numfila);
                  worksheetingles.getCell('C'+numfila).value = 'Supplier: '+renglonactual.Proveedor;
                  worksheetingles.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.mergeCells('F'+numfila+':K'+numfila);
                  worksheetingles.getCell('F'+numfila).value = 'Invoices: '+renglonactual.RefFactura;
                  worksheetingles.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
               
                  numfila=numfila+1
                  
                  re.forEach(x=>{
                     if (renglonactual.Referencia==x.Referencia){
                        worksheet.getCell('A'+numfila).value = x.Fecha//'30/08/2024  01:14:00 p. m.';//
                        worksheet.getCell('A'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
      
                        worksheet.getCell('B'+numfila).value = x.MovimientoID//'CXC';
                        worksheet.getCell('B'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        
                        if (x.MovimientoID.trim()=='CXC'){
                           worksheet.getCell('C'+numfila).value = x.UUID;
                           worksheet.getCell('C'+numfila).font = {  size: 6, color: { argb: 'FF000000' } };
                        }else{
                           worksheet.getCell('C'+numfila).value = x.Banco;
                           worksheet.getCell('C'+numfila).font = {  size: 9, color: { argb: 'FF000000' } }; // Negro
                        }
                        worksheet.getCell('D'+numfila).value = x.TipoPoliza;
                        worksheet.getCell('D'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
                        
                        worksheet.getCell('E'+numfila).value = x.IE//'I';
                        worksheet.getCell('E'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheet.getCell('F'+numfila).value = x.PolizaID//'167049';
                        worksheet.getCell('F'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheet.getCell('G'+numfila).value = x.FolioInterno;
                        worksheet.getCell('G'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheet.getCell('H'+numfila).value = x.Pedimento//'4005417';
                        worksheet.getCell('H'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheet.getCell('I'+numfila).value = x.Pedido//'MEX1268';
                        worksheet.getCell('I'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        if (x.MovimientoID.trim()=='CXC'){
                           if(x.Antiguedad==90){
                              worksheet.getCell('J'+numfila).value = x.Antiguedad;
                              worksheet.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };
                              worksheet.getCell('J' + numfila).numFmt = '"+"#,##0'; 
                           }else{
                              worksheet.getCell('J'+numfila).value = x.Antiguedad;
                              worksheet.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };

                           }
                        }
                        
                        if(x.Saldo<0){
                           worksheet.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                           worksheet.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                           worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                           
                        }
                        else{
                           worksheet.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                           worksheet.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                           worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                        }

                        worksheetingles.getCell('A'+numfila).value = x.Fecha//'30/08/2024  01:14:00 p. m.';//
                        worksheetingles.getCell('A'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
      
                        worksheetingles.getCell('B'+numfila).value = x.MovimientoID//'CXC';
                        worksheetingles.getCell('B'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        
                        if (x.MovimientoID.trim()=='CXC'){
                           worksheetingles.getCell('C'+numfila).value = x.UUID;
                           worksheetingles.getCell('C'+numfila).font = {  size: 6, color: { argb: 'FF000000' } };
                        }else{
                           worksheetingles.getCell('C'+numfila).value = x.Banco;
                           worksheetingles.getCell('C'+numfila).font = {  size: 9, color: { argb: 'FF000000' } }; // Negro
                        }
                        worksheetingles.getCell('D'+numfila).value = x.TipoPoliza;
                        worksheetingles.getCell('D'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro
                        
                        worksheetingles.getCell('E'+numfila).value = x.IE//'I';
                        worksheetingles.getCell('E'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheetingles.getCell('F'+numfila).value = x.PolizaID//'167049';
                        worksheetingles.getCell('F'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheetingles.getCell('G'+numfila).value = x.FolioInterno;
                        worksheetingles.getCell('G'+numfila).font = { size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheetingles.getCell('H'+numfila).value = x.Pedimento//'4005417';
                        worksheetingles.getCell('H'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro

                        worksheetingles.getCell('I'+numfila).value = x.Pedido//'MEX1268';
                        worksheetingles.getCell('I'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                        if (x.MovimientoID.trim()=='CXC'){
                           if(x.Antiguedad==90){
                              worksheetingles.getCell('J'+numfila).value = x.Antiguedad;
                              worksheetingles.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };
                              worksheetingles.getCell('J' + numfila).numFmt = '"+"#,##0'; 
                           }else{
                              worksheetingles.getCell('J'+numfila).value = x.Antiguedad;
                              worksheetingles.getCell('J'+numfila).font = {  size: 10, color: { argb: 'FF000000' } };

                           }
                        }
                        
                        if(x.Saldo<0){
                           worksheetingles.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                           worksheetingles.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                           worksheetingles.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                           
                        }
                        else{
                           worksheetingles.getCell('K'+numfila).value = Math.abs(x.Saldo)//'1179.9';
                           worksheetingles.getCell('K'+numfila).font = {  size: 10, color: { argb: 'FF000000' } }; // Negro
                           worksheetingles.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                        }
                        
                        saldotrafico=saldotrafico+x.Saldo
                        numfila=numfila+1
                     } 
                  })
               
                  worksheetingles.getCell('I'+numfila).value = 'Traffic Balance';
                  worksheetingles.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                  
                  if(saldotrafico<0){
                     worksheetingles.getCell('K'+numfila).value = Math.abs(saldotrafico)//'1179.9';
                     worksheetingles.getCell('K'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                     worksheetingles.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                     
                  }
                  else{
                     worksheetingles.getCell('K'+numfila).value = Math.abs(saldotrafico)//'1179.9';
                     worksheetingles.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     worksheetingles.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                  }
                  
                  saldotrafico=0;
                  numfila=numfila+1
            }) 
            }
         }
         numfila=numfila+1 
         const resumen=await sqlzay.antiguedadsaldos(cliente.ClienteID);
         resumen.forEach(r=>{
            if (re.length>0){
               if(r.Saldo<0){
                  worksheet.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('K'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  worksheetingles.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheetingles.getCell('K'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  
               } 
               else{
                  worksheet.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('K'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                  worksheetingles.getCell('K'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheetingles.getCell('K'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
               
            }
            
            numfila=numfila+3
            worksheet.mergeCells('C'+numfila+':I'+numfila);
            worksheet.getCell('D'+numfila).value = 'RESUMEN';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheet.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheetingles.mergeCells('C'+numfila+':I'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'SUMMARY';
            worksheetingles.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheetingles.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            numfila=numfila+1
            worksheet.mergeCells('D'+numfila+':F'+numfila);
            worksheet.getCell('D'+numfila).value = 'TOTAL A SU CARGO EN USD';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('G'+numfila+':H'+numfila);
            worksheetingles.mergeCells('D'+numfila+':F'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'TOTAL CHARGED TO YOU IN USD';
            worksheetingles.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('G'+numfila+':H'+numfila);
            if (re.length>0){
               if(r.Saldo<0){
                  worksheet.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('G'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('G' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
                  worksheetingles.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheetingles.getCell('G'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.getCell('G' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
               }
               else{
                  worksheet.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheet.getCell('G'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheet.getCell('G' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                  worksheetingles.getCell('G'+numfila).value = Math.abs(r.Saldo)//'1179.9';
                  worksheetingles.getCell('G'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  worksheetingles.getCell('G' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
               }
            }
            numfila=numfila+1
            worksheet.mergeCells('D'+numfila+':H'+numfila);
            worksheet.getCell('D'+numfila).value = 'ANTIGUEDAD DE SALDOS';
            worksheet.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheet.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            worksheetingles.mergeCells('D'+numfila+':H'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'AGE OF BALANCES';
            worksheetingles.getCell('D'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.getCell('D'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; // Fondo verde fuerte
            worksheetingles.getCell('D'+numfila).alignment = { horizontal: 'center', vertical: 'middle' };
            numfila=numfila+2
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            worksheet.getCell('F'+numfila).value = '$';
            worksheet.getCell('F'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('H'+numfila).value = '%';
            worksheet.getCell('H'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
            
            worksheetingles.mergeCells('F'+numfila+':G'+numfila);
            worksheetingles.getCell('F'+numfila).value = '$';
            worksheetingles.getCell('F'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.getCell('H'+numfila).value = '%';
            worksheetingles.getCell('H'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            numfila=numfila+1
            //30 dias
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo a 30 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            if (r.Saldo30<0){
               worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo30)//'1179.9';
               worksheet.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheet.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo30);
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            worksheetingles.mergeCells('D'+numfila+':E'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'Balance at 30 days';
            worksheetingles.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('F'+numfila+':G'+numfila);

            if (r.Saldo30<0){
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Saldo30)//'1179.9';
               worksheetingles.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheetingles.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Saldo30);
               worksheetingles.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }

            let porcentaje = (r.Saldo30 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            
            worksheetingles.getCell('H' + numfila).value = porcentaje;
            worksheetingles.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            numfila=numfila+1
            //31 a 60
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo de 31 días a 60 días';
            worksheet.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            if (r.Saldo60<0){
               worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo60)//'1179.9';
               worksheet.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheet.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo60);
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            worksheetingles.mergeCells('D'+numfila+':E'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'Balance from 31 days to 60 days';
            worksheetingles.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('F'+numfila+':G'+numfila);
            if (r.Saldo60<0){
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Saldo60)//'1179.9';
               worksheetingles.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheetingles.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Saldo60);
               worksheetingles.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            porcentaje = (r.Saldo60 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            worksheetingles.getCell('H' + numfila).value = porcentaje;
            worksheetingles.getCell('H' + numfila).numFmt = '"%"#,##0.00';  
            worksheetingles.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            numfila=numfila+1
            //61 a 90
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Saldo de 61 días a 90 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            if (r.Saldo90<0){
               worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo90)//'1179.9';
               worksheet.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheet.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheet.getCell('F'+numfila).value = Math.abs(r.Saldo90);
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            
            worksheetingles.mergeCells('D'+numfila+':E'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'Balance from 61 days to 90 days';
            worksheetingles.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('F'+numfila+':G'+numfila);
            if (r.Saldo90<0){
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Saldo90)//'1179.9';
               worksheetingles.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheetingles.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Saldo90);
               worksheetingles.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            porcentaje = (r.Saldo90 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00'; 
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            worksheetingles.getCell('H' + numfila).value = porcentaje;
            worksheetingles.getCell('H' + numfila).numFmt = '"%"#,##0.00'; 
            worksheetingles.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            numfila=numfila+1
            //mas de 90
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Más de 90 días';
            worksheet.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);
            if (r.Mayor90<0){
               worksheet.getCell('F'+numfila).value = Math.abs(r.Mayor90)//'1179.9';
               worksheet.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheet.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheet.getCell('F'+numfila).value = Math.abs(r.Mayor90);
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            worksheetingles.mergeCells('D'+numfila+':E'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'More than 90 days';
            worksheetingles.getCell('D'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('F'+numfila+':G'+numfila);
            if (r.Mayor90<0){
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Mayor90)//'1179.9';
               worksheetingles.getCell('F' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
               worksheetingles.getCell('F'+numfila).font = {  bold: true,size: 12, color: { argb: 'FF000000' } };
            }else{
               worksheetingles.getCell('F'+numfila).value = Math.abs(r.Mayor90);
               worksheetingles.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('F' + numfila).numFmt = '"$"#,##0.00';
            }
            porcentaje = (r.Mayor90 * 100) / r.Saldo;
            // Validar si es NaN o Infinity
            if (isNaN(porcentaje) || !isFinite(porcentaje)) {
               porcentaje = null; // O '' si prefieres una celda vacía
            } else {
               porcentaje = Math.abs(porcentaje); // Asegurar que siempre sea positivo
            } 
            worksheet.getCell('H' + numfila).value = porcentaje;
            worksheet.getCell('H' + numfila).numFmt = '"%"#,##0.00';
            worksheet.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
             
            worksheetingles.getCell('H' + numfila).value = porcentaje;
            worksheetingles.getCell('H' + numfila).numFmt = '"%"#,##0.00';
            worksheetingles.getCell('H' + numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
            
            numfila=numfila+1
            //ANTICIPOS
            worksheet.mergeCells('D'+numfila+':E'+numfila);
            worksheet.getCell('D'+numfila).value = 'Anticipos';
            worksheet.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':G'+numfila);

            worksheetingles.mergeCells('D'+numfila+':E'+numfila);
            worksheetingles.getCell('D'+numfila).value = 'Advances';
            worksheetingles.getCell('D'+numfila).font = {  size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('F'+numfila+':G'+numfila);
            if(r.Deposito<0){
               worksheet.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheet.getCell('F' + numfila).numFmt = '"($"#,##0.00")"';
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               worksheetingles.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheetingles.getCell('F' + numfila).numFmt = '"($"#,##0.00")"';
               worksheetingles.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro

            }
            else{
               worksheet.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheet.getCell('F' + numfila).numFmt = '"$"#,##0.00';
               worksheet.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
               
               worksheetingles.getCell('F'+numfila).value =Math.abs(r.Deposito);
               worksheetingles.getCell('F' + numfila).numFmt = '"$"#,##0.00';
               worksheetingles.getCell('F'+numfila).font = { size: 12, color: { argb: 'FF000000' } }; // Negro
                
            }
            
            numfila=numfila+2
         })
         
         inicial.forEach(async ini2=>{
         //SALDO A FAVOR 
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'SALDO A FAVOR';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         
         worksheetingles.mergeCells('A'+numfila+':B'+numfila);
         worksheetingles.getCell('A'+numfila).value = 'CREDIT BALANCE';
         worksheetingles.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.mergeCells('C'+numfila+':D'+numfila);
         if(ini2.SaldoAFavor<0){
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheet.getCell('C' + numfila).numFmt = '"($"#,##0.00")"'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
           
            worksheetingles.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheetingles.getCell('C' + numfila).numFmt = '"($"#,##0.00")"'; 
            worksheetingles.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } };
         }
         else{
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheet.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.getCell('C'+numfila).value = Math.abs(ini2.SaldoAFavor);
            worksheetingles.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
            worksheetingles.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
 
         numfila=numfila+1
         //SALDO DEUDOR
         worksheet.mergeCells('A'+numfila+':B'+numfila);
         worksheet.getCell('A'+numfila).value = 'SALDO DEUDOR';
         worksheet.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('C'+numfila+':D'+numfila);
         if (ini2.SaldoPendiente<0){
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
            worksheet.getCell('C' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         
         }else{
            worksheet.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
            worksheet.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
            worksheet.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
         
         worksheetingles.mergeCells('A'+numfila+':B'+numfila);
         worksheetingles.getCell('A'+numfila).value = 'DEBTOR BALANCE';
         worksheetingles.getCell('A'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.mergeCells('C'+numfila+':D'+numfila);
         if (ini2.SaldoPendiente<0){
            worksheetingles.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
            worksheetingles.getCell('C' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
            worksheetingles.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         
         }else{
            worksheetingles.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
            worksheetingles.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
            worksheetingles.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
            
         }
         
         worksheetingles.getCell('C'+numfila).value = Math.abs(ini2.SaldoPendiente);
         worksheetingles.getCell('C' + numfila).numFmt = '"$"#,##0.00'; 
         worksheetingles.getCell('C'+numfila).font = { bold: true, size: 12, color: { argb: 'FF000000' } }; // Negro
         
      })
         //CARGOS ANTICIPOS SIN APLICAR
         const sinaplicar=await sqlzay.sp_cargossinaplicar(cliente.ClienteID);
         if (sinaplicar.length>0){
            numfila=numfila+2
            worksheet.mergeCells('A'+numfila+':D'+numfila);
            worksheet.getCell('A'+numfila).value = 'CARGOS/ANTICIPOS SIN APLICAR';
            worksheet.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheet.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; 
            worksheetingles.mergeCells('A'+numfila+':D'+numfila);
            worksheetingles.getCell('A'+numfila).value = 'CHARGES/ADVANCES NOT APPLIED';
            worksheetingles.getCell('A'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.getCell('A'+numfila).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF008000' } }; 
            numfila=numfila+1
            const sinaplicardistinct=await sqlzay.sp_cargossinaplicar_distinct(cliente.ClienteID)
            let saldotraf=0;
            sinaplicardistinct.forEach(sad=>{
            worksheet.mergeCells('A'+numfila+':B'+numfila);
            worksheet.getCell('A'+numfila).value = 'TRAFICO: '+sad.Referencia;
            worksheet.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('C'+numfila+':E'+numfila);
            worksheet.getCell('C'+numfila).value = 'Proveedor: '+sad.Proveedor;
            worksheet.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            worksheet.mergeCells('F'+numfila+':K'+numfila);
            worksheet.getCell('F'+numfila).value = 'Facturas: '+sad.RefFactura;
            worksheet.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            
            worksheetingles.mergeCells('A'+numfila+':B'+numfila);
            worksheetingles.getCell('A'+numfila).value = 'TRAFFIC: '+sad.Referencia;
            worksheetingles.getCell('A'+numfila).font = { bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('C'+numfila+':E'+numfila);
            worksheetingles.getCell('C'+numfila).value = 'Supplier: '+sad.Proveedor;
            worksheetingles.getCell('C'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            worksheetingles.mergeCells('F'+numfila+':K'+numfila);
            worksheetingles.getCell('F'+numfila).value = 'Invoices: '+sad.RefFactura;
            worksheetingles.getCell('F'+numfila).font = { bold: true,size: 8, color: { argb: 'FF000000' } }; // Negro
            
            numfila=numfila+1
            
               sinaplicar.forEach(sa=>{
                  worksheet.getCell('A'+numfila).value = sa.Fecha;
                  worksheet.getCell('A'+numfila).font = {  size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('B'+numfila).value = sa.PolizaID;
                  worksheet.getCell('B'+numfila).font = { bsize: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('C'+numfila).value = sa.Banco;
                  worksheet.getCell('C'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('D'+numfila).value = sa.TipoMovimiento;
                  worksheet.getCell('D'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
                  
                  worksheet.getCell('E'+numfila).value = sa.IE;
                  worksheet.getCell('E'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('F'+numfila).value = sa.Factura;
                  worksheet.getCell('F'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('G'+numfila).value = sa.FolioInterno;
                  worksheet.getCell('G'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('H'+numfila).value = sa.Pedimento;
                  worksheet.getCell('H'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheet.getCell('I'+numfila).value = sa.Pedido;
                  worksheet.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         
               
                  if(sa.Saldo<0){
                     worksheet.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheet.getCell('K' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
                     worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  else{
                     worksheet.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheet.getCell('K' + numfila).numFmt = '"$"#,##0.00""'; // Formato de moneda
                     worksheet.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }

                  worksheetingles.getCell('A'+numfila).value = sa.Fecha;
                  worksheetingles.getCell('A'+numfila).font = {  size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('B'+numfila).value = sa.PolizaID;
                  worksheetingles.getCell('B'+numfila).font = { bsize: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('C'+numfila).value = sa.Banco;
                  worksheetingles.getCell('C'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('D'+numfila).value = sa.TipoMovimiento;
                  worksheetingles.getCell('D'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
                  
                  worksheetingles.getCell('E'+numfila).value = sa.IE;
                  worksheetingles.getCell('E'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('F'+numfila).value = sa.Factura;
                  worksheetingles.getCell('F'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('G'+numfila).value = sa.FolioInterno;
                  worksheetingles.getCell('G'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('H'+numfila).value = sa.Pedimento;
                  worksheetingles.getCell('H'+numfila).font = { size: 8, color: { argb: 'FF000000' } }; // Negro
         
                  worksheetingles.getCell('I'+numfila).value = sa.Pedido;
                  worksheetingles.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
         
               
                  if(sa.Saldo<0){
                     worksheetingles.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheetingles.getCell('K' + numfila).numFmt = '"($"#,##0.00")"'; // Formato de moneda
                     worksheetingles.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  else{
                     worksheetingles.getCell('K'+numfila).value = Math.abs(sa.Saldo);
                     worksheetingles.getCell('K' + numfila).numFmt = '"$"#,##0.00""'; // Formato de moneda
                     worksheetingles.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                     
                  }
                  saldotraf=saldotraf+sa.Saldo
                  numfila=numfila+1
               })
               worksheetingles.getCell('I'+numfila).value = 'Saldo del Tráfico';
               worksheetingles.getCell('I'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
               
               if(saldotraf<0){
                  worksheetingles.getCell('K'+numfila).value = Math.abs(saldotraf)//'1179.9';
                  worksheetingles.getCell('K' + numfila).numFmt = '"("#,##0.00")"'; // Formato de moneda
                  worksheetingles.getCell('K'+numfila).font = {  bold: true,size: 10, color: { argb: 'FF000000' } }; // Negro
                  
                  
               }
               else{
                  worksheetingles.getCell('K'+numfila).value = Math.abs(saldotraf)//'1179.9';
                  worksheetingles.getCell('K' + numfila).numFmt = '"$"#,##0.00'; // Formato de moneda
                  worksheetingles.getCell('K'+numfila).font = { bold: true, size: 10, color: { argb: 'FF000000' } }; // Negro
                  
               }
            })

         } 
        numfila=numfila+2
         //PIE DE PAGINA
         worksheet.mergeCells('H'+numfila+':I'+numfila);
         worksheet.getCell('H'+numfila).value = 'FR-02-02-02';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.mergeCells('H'+numfila+':I'+numfila);
         worksheetingles.getCell('H'+numfila).value = 'FR-02-02-02';
         worksheetingles.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         numfila=numfila+1
         worksheet.mergeCells('H'+numfila+':I'+numfila);
         worksheet.getCell('H'+numfila).value = 'Rev. 01';
         worksheet.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheet.mergeCells('J'+numfila+':K'+numfila);
         worksheetingles.mergeCells('H'+numfila+':I'+numfila);
         worksheetingles.getCell('H'+numfila).value = 'Rev. 01';
         worksheetingles.getCell('H'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.mergeCells('J'+numfila+':K'+numfila);
         const today = new Date();
         const formattedDate2 = `${today.getDate().toString().padStart(2, '0')}/${(today.getMonth() + 1).toString().padStart(2, '0')}/${today.getFullYear()}`;
         worksheet.getCell('J'+numfila).value = formattedDate2;
         worksheet.getCell('J'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
         worksheetingles.getCell('J'+numfila).value = formattedDate2;
         worksheetingles.getCell('J'+numfila).font = { bold: true, size: 8, color: { argb: 'FF000000' } }; // Negro
      
         
         let nombreOriginal =cliente.Nombre;
         let nombreLimpio = limpiarNombreArchivo(nombreOriginal);
         /******************************************************************************************* */
         // Guardar archivo
         let nombreArchivo='Estado_de_Cuenta_USD '+cliente.ClienteID+' '+nombreLimpio
         await workbook.xlsx.writeFile('Estado_de_Cuenta_USD '+cliente.ClienteID+' '+nombreLimpio+'.xlsx');

         //await enviarMailDLL(nombreArchivo,transport,'',nombreLimpio)
         const correos=await sqlzay.contactosestadoscuenta(cliente.ClienteID);
         correos.forEach(async co=>{
            await enviarMailDLL(nombreArchivo,transport,co.correos,nombreLimpio)
         })
         console.log('Archivo creado exitosamente.');
         
      }
   }
   
   if(totalclientes==(clientes.length)){
      await res.json('Reportes Enviados')
   }
}); 
enviarMailDLL= async(nombreArchivo,transport, correos,nombreLimpio) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      //to: 'cobranza@zayro.com;sistemas@zayro.com;'+correos,
      to: 'programacion@zayro.com;',
      subject: 'Estado de cuenta '+nombreLimpio,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './' + nombreArchivo + '.xlsx',
         }],
      text: 'Estado de Cuenta Mensual USD',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
function limpiarNombreArchivo(nombre) {
   return nombre 
       .normalize("NFD") // Normaliza para separar caracteres acentuados
       .replace(/[\u0300-\u036f]/g, "") // Elimina los acentos
       .replace(/[^a-zA-Z0-9._-]/g, "_") // Reemplaza caracteres no permitidos con guion bajo
       .replace(/_+/g, "_") // Reemplaza múltiples guiones bajos por uno solo
       .trim(); // Elimina espacios en blanco al inicio y al final
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_avisoautomaticohb201sinedi', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
       const result = await sql.sp_AVISO_AUTOMATICO_HB201_SIN_EDI();

       const wb = new xl.Workbook();
       const nombreArchivo = "Reporte HB201 SIN EDI";
       const ws = wb.addWorksheet("HB201");

       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#008000',
           },
       });
       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       const columnas = [
           "BOX_ID", "PALLET", "FECHA_SCAN", "ENTRADA",
           "CUARENTENA", "HB201"
       ];
       columnas.forEach((columna, index) => {
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
           Object.keys(reglonactual).forEach((columna, idx) => {
               ws.cell(numfila, idx + 1).string(reglonactual[columna]).style(estilocontenido);
           });
           numfila++;
       });

       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');

       await wb.write(pathExcel, async function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               try {
                   await fs.promises.access(pathExcel, fs.constants.F_OK);
                   res.download(pathExcel, () => {
                       //fs.unlink(pathExcel, (err) => {
                           if (err) console.error(err);
                           else console.log("Archivo descargado y eliminado exitosamente.");
                       //});
                   });
               } catch (err) {
                   console.error(err);
                   res.status(500).send("Error al acceder al archivo Excel generado.");
               }
           }
       });

       await enviarMailavisoautomaticohb201sinedi(nombreArchivo,transport);
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
   

});
enviarMailavisoautomaticohb201sinedi= async(nombreArchivo,transport) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'kwteamleader@zayro.com;lfzamudio@zayro.com;whmanager@zayro.com;sistemas@zayro.com;',
      subject: 'Reporte HB201 SIN EDI',
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }],
      text: 'HOLA BUEN DÍA SE ANEXA ARCHIVO REPORTE HB201 SIN EDI',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_envioavisosidentificadoresautomaticos', async function(req, res, next) {
   try {
         let config = {
            host:process.env.hostemail,
            port:process.env.portemail,
            secure: true,
            auth: {
               user:process.env.useremail,
               pass:process.env.passemail
            },
            tls: {
               rejectUnauthorized: false
            }
         }
       let transport = nodemailer.createTransport(config); 
       const result = await mysql.sp_obtener_datos_identificadores();
       //console.log(result)
      if (result.length > 0) {
         for (let row of result) {
            
            //console.log(row)
            let { Pedimento, Aduana, Identificador, Fecha, Partida, Complemento1, Complemento2, Complemento3, Cliente, Referencia } = row;
           // console.log(row)
            //console.log(Pedimento, Aduana, Identificador, Fecha, Partida, Complemento1, Complemento2, Complemento3, Cliente, Referencia)
            let resultCheck = await sqlSISTEMAS.sp_obtener_datos_iden(Pedimento, Aduana, Identificador, Fecha,  Partida ? Partida.toString() : null, Complemento1, Complemento2, Complemento3, Cliente, Referencia)
            
            if (resultCheck.length === 0) {
               let listaIdentificadores = ["EA", "EB", "EC", "EF", "EN", "ES", "EX", "MA", "MB", "MC", "MM", "MR", "MV", "NE", "NS", "NZ", "OV", "PA", "PB", "PG", "PS", "PT", "SC", "UM", "XP"];
     
               if (listaIdentificadores.includes(Identificador)) {
                  
                  let resultCliente = await mysql.sp_obtener_referencia(Pedimento);
                  let correoEjecutivo = "";
                 
                  if (resultCliente.length > 0) {
                     let resultTrafico = await sqlSISTEMAS.sp_obtener_traImpExp(Referencia);
                  
                     let servicio = resultTrafico.length > 0 && resultTrafico[0].traImpExp ? "IMPORTACION" : "EXPORTACION";
                     let resultCorreo = await sqlSISTEMAS.sp_obtener_email(Referencia,servicio);
                     console.log(resultCorreo)
                     if (resultCorreo.length > 0) {
                        correoEjecutivo = resultCorreo[0].usuEmail;
                        console.log(resultCorreo[0].usuEmail)
                        await sqlSISTEMAS.sp_insertar_ident(Pedimento,Aduana,Identificador, Fecha, Partida ? Partida.toString() : null, Complemento1, Complemento2, Complemento3, Cliente, Referencia, correoEjecutivo)
                        //console.log(`Registro insertado: ${Pedimento} - ${Aduana}`);
                     }

                  }
                  
               }

            }
         
         }
      }
      let registros=await sqlSISTEMAS.sp_obtener_ident_no_enviados();
      //console.log(registros);
      if (registros.length === 0) {
         console.log("No hay registros pendientes de envío.");
         res.json('No hay registros pendientes de envío.')
      }else{
         for (const registro of registros) {
            const { pedimento, referencia, cliente, aduana, correo } = registro;
    
           const destinatarios = `cgonzalez@zayro.com;avazquez@zayro.com;gerenciati@zayro.com`;
           const cc = `programacion@zayro.com;${correo}`;
           //const destinatarios = `programacion@zayro.com`;
           //const cc = `programacion@zayro.com`;

    
           let texto = `<!DOCTYPE html>
           <html>
           <body>
           <div style="text-align:center;">
           <table border="1">
           <tr><td><b>Cliente</b></td><td style="color:#CC6600">${cliente}</td></tr>
           <tr><td><b>Referencia</b></td><td style="color:#0099CC">${referencia}</td></tr>
           <tr><td><b>Pedimento</b></td><td style="color:#009900">${pedimento}</td></tr>
           <tr><td><b>Aduana</b></td><td style="color:#0000CC">${aduana}</td></tr>
           <tr><td><b>Origen</b></td><td style="color:#0000CC">SLAMM3</td></tr>`;
           
           let identificadores = await sqlSISTEMAS.sp_obtener_ident_por_pedimento(pedimento) || [];
           
           // **Siempre envía una fila para cada identificador, aunque sus valores sean vacíos**
           // **Siempre envía una fila para cada identificador, aunque sus valores sean vacíos**
            identificadores.forEach(id => {
               texto += `
                  <tr><td><b>Identificador Partida ${id.partida || ''}</b></td><td style="color:#0000CC">${id.clave || ''}</td></tr>
                  <tr><td>Permiso:</td><td style="color:#0000CC">${id.permiso || ''}</td></tr>  <!-- Campo de Permiso -->
                  <tr><td>Complemento1:</td><td style="color:#0000CC">${id.compuno ?? ''}</td></tr>
                  <tr><td>Complemento2:</td><td style="color:#0000CC">${id.compdos ?? ''}</td></tr>
                  <tr><td>Complemento3:</td><td style="color:#0000CC">${id.comptres ?? ''}</td></tr>
               `;
            });

            // **Si no hay identificadores, al menos enviamos una fila con valores vacíos**
            if (identificadores.length === 0) {
               texto += `
                  <tr><td><b>Identificador Partida </b></td><td style="color:#0000CC"></td></tr>
                  <tr><td>Permiso:</td><td style="color:#0000CC"></td></tr> <!-- Campo de Permiso vacío -->
                  <tr><td>Complemento1:</td><td style="color:#0000CC"></td></tr>
                  <tr><td>Complemento2:</td><td style="color:#0000CC"></td></tr>
                  <tr><td>Complemento3:</td><td style="color:#0000CC"></td></tr>
               `;
            }

           
           texto += `</table></div></body></html>`;
           
            console.log(`Enviando correo a ${destinatarios} con CC a ${cc}`);
            await enviarMailAvusisidentificadores(destinatarios,cc,transport,texto);
            await sqlSISTEMAS.sp_actualizar_enviado(pedimento)

         }
         res.json('Proceso terminado')
      }
      
      
   } 
   catch (err) {
      console.error('EL ERROR ES ' + err);
      res.status(500).send("Error al obtener los datos de la base de datos.");
   }
});
enviarMailAvusisidentificadores= async(destinatarios,cc,transport,texto) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      to: destinatarios,
      cc:cc,
      subject: 'Asignación de Identificadores y Permisos - Agencia Aduanal Zamudio y Rodriguez S.C.',
      html: texto,
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         //console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_kfantasma', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
       const result = await sql.Sp_kfantasma();

       const wb = new xl.Workbook();
       const nombreArchivo = "Reporte K fantasma";
       const ws = wb.addWorksheet("K fantasma");

       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#008000',
           },
       });
       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });
       const estiloRojo = wb.createStyle({
         font: {
            name: 'Arial',
            color: '#FF0000',  // Color rojo
            size: 10,
            bold: true
         }
      });

       const columnas = [
           "CERRADA", "HB201", "HB201_GEN", "HB201F",
           "NO_PALLET"
       ];
       columnas.forEach((columna, index) => {
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
         const esRojo = reglonactual.colorrojo === '1';
           Object.keys(reglonactual).forEach((columna, idx) => {
            if (columna !== "colorrojo") { // No incluir el campo 'colorrojo'
               ws.cell(numfila, idx + 1)
                 .string(reglonactual[columna])
                 .style(esRojo ? estiloRojo : estilocontenido);
            }
           });
           numfila++;
       });

       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');

       await wb.write(pathExcel, async function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               try {
                   await fs.promises.access(pathExcel, fs.constants.F_OK);
                   res.download(pathExcel, () => {
                       //fs.unlink(pathExcel, (err) => {
                           if (err) console.error(err);
                           else console.log("Archivo descargado y eliminado exitosamente.");
                       //});
                   });
               } catch (err) {
                   console.error(err);
                   res.status(500).send("Error al acceder al archivo Excel generado.");
               }
           }
       });

       await enviarMailkfantasma(nombreArchivo,transport);
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
   

});
enviarMailkfantasma= async(nombreArchivo,transport) => {
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'kwteamleader@zayro.com;lfzamudio@zayro.com;whmanager@zayro.com;sistemas@zayro.com;',
      subject: 'Reporte K fantasma',
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }],
      text: 'HOLA BUEN DÍA SE ANEXA ARCHIVO DEL REPORTE',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_actualizarcreditoclientes', async function(req, res, next) {
   try {
     
      const result = await sqlzam.sp_actualizarlimitesdecreditocliente();
      if(result.length > 0)
      {
         let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
      await enviarMailactualizacioncreditos(transport);
         res.status(200).send("Se actualizaron los limites de creditos de los clientes");
      } else
      {
         res.status(500).send("No se actualizó ningún límite de crédito de los clientes");
      }
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
   

});
enviarMailactualizacioncreditos= async(transport) => {

   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'programacion@zayro.com',
      subject : `Credito de clientes Actualizados`,
      text: 'Se actualizaron los Creditos de los clientes',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_enviaranexo24kawassaki', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
      const wb = new xl.Workbook();
      const nombreArchivo = "Reporte Anexo 24";
      const ws = wb.addWorksheet("Importación");
      const wsExpo = wb.addWorksheet("Exportación");

      const estiloTitulo = wb.createStyle({
          font: {
              name: 'Arial',
              color: '#FFFFFF',
              size: 10,
              bold: true,
          },
          fill: {
              type: 'pattern',
              patternType: 'solid',
              fgColor: '#630b57',
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
      });
      const estilocontenido = wb.createStyle({
          font: {
              name: 'Arial',
              color: '#000000',
              size: 10,
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
      });
      const estiloMoneda = wb.createStyle({
         numberFormat: '"$"#,##0.00', // formato con símbolo de dólar y dos decimales
         alignment: {
             horizontal: 'right'
         }
     });
      const columnas = [
         "Pedimento","Aduana","Clave","Fecha de Pago","Proveedor","Factura",	
         "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Origen","Tratado"
         ,"Cantidad UMComercial","UMComercial",	"Valor Aduana",	"Valor Comercial",	
         "TIGI","FP IGI","FP IVA","FP IEPS",	"Tipo de cambio"
     ];
      columnas.forEach((columna, index) => {
         ws.cell(1, index + 1).string(columna).style(estiloTitulo);
      });

     
      const result = await sql.sp_ObtenerPedimentos(2316);
      //console.log('result'); 
      //PARA IMPORTASCION
      const facturas = await mysql.sp_ObtenerDatosFactura(result[0].Pedimentos,1)

      //console.log(facturas); 
      var numfila=2;
      for (const factura of facturas) {

         //console.log('Pedimento:', factura.Pedimento);
         const informacionreporte=await sql.sp_ObtenerInformacionPedimento(2316,1,factura.Pedimento,0) 
         //console.log(informacionreporte) 
         if (informacionreporte.length>0){
            
            for (const reporte of informacionreporte)
            {
               ws.cell(numfila, 1).string(reporte.Pedimento.toString()).style(estilocontenido);
               ws.cell(numfila, 2).string(reporte.Aduana.toString()).style(estilocontenido);
               ws.cell(numfila, 3).string(reporte.Clave.toString()).style(estilocontenido);

               ws.cell(numfila,4).string(factura.FechadePago).style(estilocontenido);
               ws.cell(numfila, 5).string(factura.Proveedor.toString()).style(estilocontenido);
               ws.cell(numfila, 6).string(reporte.Factura.toString()).style(estilocontenido);
               ws.cell(numfila,7).string(factura.FechaFactura).style(estilocontenido);

               ws.cell(numfila, 8).string(reporte.Producto.toString()).style(estilocontenido);
               ws.cell(numfila, 9).string(reporte.Fraccion.toString().substring(0, 8)).style(estilocontenido)
               ws.cell(numfila,10).string(reporte.OrigenDestino.toString()).style(estilocontenido);
               ws.cell(numfila,11).string(reporte.Tratado.toString()).style(estilocontenido);

               ws.cell(numfila,12).number(Number(reporte.CantidadUMComercial)).style(estilocontenido);
               ws.cell(numfila,13).string(factura.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
               ws.cell(numfila, 14).number(Number(factura.ValorAduana)).style(estiloMoneda);
               ws.cell(numfila, 15).number(Number(reporte.ValorComercial)).style(estiloMoneda);

               ws.cell(numfila,16).number(Number(reporte.TIGI)).style(estilocontenido);
               ws.cell(numfila,17).number(Number(reporte.FPIGI)).style(estilocontenido);
               ws.cell(numfila,18).number(Number(reporte.FPIVA)).style(estilocontenido);
               ws.cell(numfila,19).number(Number(reporte.FPIEPS)).style(estilocontenido);

               ws.cell(numfila,20).number(Number(factura.Tipodecambio)).style(estilocontenido);

               numfila++;
            }
            

         }

     }

     const columnasExpo = [
      "Pedimento","Aduana","Clave","Fecha de Pago","Cliente","Factura",	
      "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Destino"
      ,"Cantidad UMComercial","UMComercial",	"Valor Comercial",	
      "Valor USD",	"Tipo de cambio"
  ];
   columnasExpo.forEach((columna, index) => {
      wsExpo.cell(1, index + 1).string(columna).style(estiloTitulo);
   });
     const facturasExpo = await mysql.sp_ObtenerDatosFacturaexpo(result[0].Pedimentos,2)

     //console.log(facturas); 
     var numfilaExpo=2;
     for (const facturaexpo  of facturasExpo ) {

        console.log('Pedimento:', facturaexpo.Pedimento);
        const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(2316,0,facturaexpo.Pedimento,facturaexpo.Partida) 
        //console.log(informacionreporteExpo ) 
        if (informacionreporteExpo.length>0){
         for (const reporteexpo of informacionreporteExpo)
            {
               wsExpo.cell(numfilaExpo, 1).string(reporteexpo.Pedimento.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 2).string(reporteexpo.Aduana.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 3).string(reporteexpo.Clave.toString()).style(estilocontenido);

               wsExpo.cell(numfilaExpo,4).string(facturaexpo.FechadePago).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 5).string(facturaexpo.Proveedor.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 6).string(reporteexpo.Factura.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo,7).string(facturaexpo.FechaFactura).style(estilocontenido);

               wsExpo.cell(numfilaExpo, 8).string(reporteexpo.Producto.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 9).string(reporteexpo.Fraccion.toString().substring(0, 8)).style(estilocontenido);

               wsExpo.cell(numfilaExpo,10).string(reporteexpo.OrigenDestino.toString()).style(estilocontenido);
            
               wsExpo.cell(numfilaExpo,11).number(Number(reporteexpo.CantidadUMComercial)).style(estilocontenido);
               wsExpo.cell(numfilaExpo,12).string(facturaexpo.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
               let valorComercial = Number(reporteexpo.ValorComercial);
               let tipoCambio = Number(facturaexpo.Tipodecambio);
               let resultado = valorComercial * tipoCambio;

               wsExpo.cell(numfilaExpo, 13).number(resultado).style(estiloMoneda);
               //wsExpo.cell(numfilaExpo,13).number(Number(reporteexpo.ValorComercial)).style(estiloMoneda);
               wsExpo.cell(numfilaExpo,14).number(Number(facturaexpo.ValorDolares)).style(estiloMoneda);

               wsExpo.cell(numfilaExpo,15).number(Number(facturaexpo.Tipodecambio)).style(estilocontenido);

               numfilaExpo++;
            }
           
        }
    }
     const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
      //Guardar
      await wb.write(pathExcel, (err, stats) => {
         if (err) {
            console.error('Error al guardar el archivo de Excel:', err);
            
         } else {
            console.log('Archivo de Excel guardado exitosamente en:', pathExcel);
            // Descargar el archivo de Excel
            res.download(pathExcel, nombreArchivo+'.xlsx', (err) => {
               if (err) {
                     console.error('Error al descargar el archivo:', err);
                     // Manejar el error
               } else {
                     console.log('Archivo descargado exitoso');
                   
               }
            });
         }
   });
   /*
     //PARA EXPORTACION
     const facturasExpo = await mysql.sp_ObtenerDatosFactura(result[0].Pedimentos,2)

      //console.log(facturas); 
      for (const facturaExpo  of facturasExpo ) {

         console.log('Pedimento:', facturaExpo.Pedimento);
         const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(2316,0,facturaExpo.Pedimento) 
         console.log(informacionreporteExpo ) 
         if (informacionreporteExpo.length>0){
            
         }
     }
     */
      await enviarMailAnexo24Kawassaki(nombreArchivo,transport);
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");  
   }
   

});
enviarMailAnexo24Kawassaki= async(nombreArchivo,transport) => {
      const meses = [
         "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
      ];
      
      const fechaActual = new Date();
      fechaActual.setMonth(fechaActual.getMonth() - 1); // Retrocede un mes
      
      const mesAnterior = meses[fechaActual.getMonth()];
      const año = fechaActual.getFullYear();
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'anahi.valle@kawasakimotores.mx; gerenciati@zayro.com; cgonzalez@zayro.com; auxiliarimportacion108@zayro.com',
      cc: 'lfzamudio@zayro.com; lzamudio@zayro.com; kwteamleader@zayro.com; avazquez@zayro.com; importacion103@zayro.com;programacion@zayro.com',
      subject : `REPORTES PARA ANEXO 24 || ZAYRO || ${mesAnterior} ${año}`,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }],
      text: 'HOLA BUEN DÍA SE ANEXA ARCHIVO DE LOS REPORTES',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/  
app.get('/api/getdata_cumpleanios', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
      let datoscumple=await sql.sp_informacion_cumpleanios()  
      datoscumple.forEach(async item=>{
         await enviarMailPORTALCUMPLE(transport,item.us_NombreCompleto);
      })
      res.status(200).send(true);
       
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
   

});
enviarMailPORTALCUMPLE= async(transport,nombre) => {
const mensaje = {
  from: '"Cumpleaños Zayro" <sistemas@zayro.com>',
  to: 'zayro.nld@zayro.com; zayro.ltx@zayro.com; zayro.saz@zayro.com; transpuentes@t3polos.com; asistenterh@zayro.com',
  subject: `🎉 ¡Feliz cumpleaños, ${nombre}! 🎂✨`,
  html: `
   <style>
      body {
        margin: 0; padding: 0;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        background: #c7e4c7; /* verde pastel claro para todo el fondo */
        color: #144d14; /* verde oscuro para texto */
      }
      .container {
        max-width: 600px;
        margin: 40px auto; /* espacio arriba y abajo para que respire */
        padding: 35px 30px;
        border-radius: 22px;
        background: transparent; /* fondo transparente para que no se vea cuadro */
        box-shadow: 0 0 20px rgba(20, 77, 20, 0.3); /* sombra suave para delimitar */
        text-align: center;
      }
      h1 {
        color: #2e7d32;
        font-size: 3rem;
        margin-bottom: 18px;
        text-shadow: 1.5px 1.5px 6px #81c784;
      }
      p.lead {
        font-size: 1.3rem;
        margin: 0 0 30px;
        color: #1b5e20;
        font-weight: 700;
        letter-spacing: 0.03em;
      }
      ul {
        list-style: none;
        padding: 0;
        margin: 0 0 35px;
      }
      ul li {
        font-size: 1.15rem;
        margin: 14px 0;
        padding-left: 32px;
        position: relative;
        color: #27632a;
        font-weight: 600;
      }
      ul li::before {
        content: '🎈';
        position: absolute;
        left: 0;
        top: 0;
        font-size: 1.4rem;
      }
      .btn-cumple {
        display: inline-block;
        background: linear-gradient(45deg, #43a047, #1b5e20);
        color: red !important;
        padding: 16px 34px;
        border-radius: 45px;
        text-decoration: none;
        font-weight: 800;
        font-size: 1.25rem;
        box-shadow: 0 6px 18px rgba(27, 94, 32, 0.7);
        transition: background-color 0.3s ease, box-shadow 0.3s ease;
      }
      .btn-cumple:hover,
      .btn-cumple:focus {
        background: linear-gradient(45deg, #1b5e20, #43a047);
        box-shadow: 0 8px 26px rgba(67, 160, 71, 0.85);
      }
      cumpleGif {
         max-width: 190px;
         width: 190px;
         height: 190px;
         border-radius: 50%; 
         margin-bottom: 28px;
         box-shadow: 0 0 30px 6px rgba(67, 160, 71, 0.9);
         object-fit: cover; 
         border: 4px solid #43a047; 
      }

      .footer {
        font-size: 0.9rem;
        color: #2e7d32;
        font-style: italic;
        margin-top: 38px;
      }
        .message-list p {
         font-size: 1.15rem;
         margin: 14px 0;
         font-weight: 600;
         color: #2e7d32; 
         text-shadow: 0 0 4px rgba(0,0,0,0.1);
         }

    </style>
    <div class="container">
    <h1>🎂🎉🎈 ¡Feliz cumpleaños, ${nombre}! 🎉🎈🎂</h1>
      
      <img class="cumpleGif" src="https://gifs.org.es/gifs/2022/10/feliz-cumple-tarta-animada.gif" alt="Gif cumpleaños" />
      
     <div class="message-list">
      <p>🌞 Que cada día te regale un motivo para sonreír</p>
      <p>❤️ Que la paz y el amor te acompañen siempre</p>
      <p>🌳 Que tus metas crezcan tan fuertes como un árbol</p>
      <p>😊 Que disfrutes cada momento con alegría y gratitud</p>
      </div>

      <p style="font-weight: 700; font-size: 1.1rem; color: #2e7d32; margin-top: 20px; display: flex; align-items: center; gap: 8px; max-width: 400px;">
       <span style="display:inline-block; width: 30px; height: 30px; overflow: hidden;">
         <img 
            src="https://net.zayro.com/zayrologistics/Imagenes/zayrologo.png" 
            alt="Logo Zayro" 
            width="30" 
            height="30" 
            style="object-fit: contain; display: block; max-width: 100%; height: auto;" 
         />
      </span>
      ¡Muchas felicidades de parte de toda la familia Zayro!
      <span style="display:inline-block; width: 30px; height: 30px; overflow: hidden;">
         <img 
            src="https://net.zayro.com/zayrologistics/Imagenes/zayrologo.png" 
            alt="Logo Zayro" 
            width="30" 
            height="30" 
            style="object-fit: contain; display: block; max-width: 100%; height: auto;" 
         />
      </span>
      </p>




      <a href="https://net.zayro.com/zayrologistics/cumpleanios" target="_blank" class="btn-cumple">🎁 Felicita al cumpleañero(a) 🎁</a>
    </div>
  `,
};





   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_reportedist', async function(req, res, next) {
   try {
       const result = await sqldist.sp_reporte_distribucion(2);

       const wb = new xl.Workbook();
       const nombreArchivo = "Distribution Report";
       const ws = wb.addWorksheet("Reporte");

       const estiloTitulo = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
           },
           fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#288BA8',
           },
       });
       const estilocontenido = wb.createStyle({
           font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
           }
       });

       const columnas = [
            "Part", "Description", "Quantity"
       ];
       columnas.forEach((columna, index) => {     
           ws.cell(1, index + 1).string(columna).style(estiloTitulo);
       });

       let numfila = 2;
       result.forEach(reglonactual => {
           Object.keys(reglonactual).forEach((columna, idx) => {
               const valor = reglonactual[columna] !== null && reglonactual[columna] !== undefined ? reglonactual[columna].toString() : '';
               ws.cell(numfila, idx + 1).string(valor).style(estilocontenido);
           });
           numfila++;
       });

       const pathExcel = path.join(__dirname, 'excel', nombreArchivo + '.xlsx');

       wb.write(pathExcel, async function(err) {
           if (err) {
               console.error(err);
               res.status(500).send("Error al generar el archivo Excel.");
           } else {
               try {
                   await fs.promises.access(pathExcel, fs.constants.F_OK);
                   res.download(pathExcel, () => {
                       
                           if (err) console.error(err);
                           else console.log("Archivo descargado y eliminado exitosamente.");
                           

                   });
               } catch (err) {
                   console.error(err);
                   res.status(500).send("Error al acceder al archivo Excel generado.");
               }
           }
       });

       const correosResult = await sql.getdata_correos_reporte('2');
       correosResult.forEach(renglonactual => {
           enviarMailreportedist(renglonactual.correos);
       });
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
});
enviarMailreportedist=async(correos)=>{
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
      from:'sistemas@zayro.com',
      //to:'ca.she@logisteed-america.com, ja.diaz@logisteed-america.com, lfzamudio@zayro.com, lzamudio@zayro.com, distribution1@zayro.com, alule@zayro.com, avazquez@zayro.com',
      to:'programacion@zayro.com',
      //to: 'whmanager@zayro.com, teamleader@zayro.com, supervisor@zayro.com, revisadoresa@zayro.com, revisadoresb@zayro.com, revisadoresc@zayro.com, revisadoresd@zayro.com, revisadores@zayro.com, REVISADORESE@ZAYRO.COM, avazquez@zayro.com, cgonzalez@zayro.com, auxiliarimportacion102@zayro.com, distribution1@zayro.com,cchavana@zayro.com',
      //cc:'sistemas@zayro.com',
      subject:'Distribution Report',
      attachments:[
         {filename:'Distribution Report.xlsx',
         path:'./src/excel/Distribution Report.xlsx'}],
      text:'Please find attached the report',
   }
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos); 
} 
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdataclientesIMMEXfraccionesautorizadas', async function(req, res, next) {
   try {
       const result = await sql.sp_totaldeclientesrevisarfraccionesIMMEX();
         let resulta;
         let contador=1;
      for (let i = 0; i < result.length; i++) {
      const numeroCliente = result[i].NumeroCliente;
      console.log(numeroCliente);
      contador=contador+1;
      resulta=await sql.sp_obtenerfraccionesIMMEXnoautorizadas(numeroCliente)
      if (resulta.length>=1)
      {
         const correos = await  sql.sp_obtenerejecutivogerentesubcliente(numeroCliente);
         await enviarMailfraccionesIMMEX(correos[0].Correos,resulta)
      }
    }
    if (contador=result.length){
      res.json('Reportes Enviados')
    }

   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");
   }
});
enviarMailfraccionesIMMEX=async(correos,resultados)=>{
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
   const rows = resultados.map(r => `
    <tr>
      <td>${r.Referencia}</td>
      <td>${r.Cliente}</td>
      <td>${r.Factura}</td>
      <td>${r['Numero de Parte']}</td>
      <td>${r.Descripcion}</td>
      <td>${r.Fraccion_Actual}</td>
      <td>${r.FechaTrafico}</td>
      <td>${r.Partida}</td>
    </tr>
  `).join('');

  const htmlBody = `
    <p>Adjunto encontrará el reporte de fracciones IMMEX no autorizadas para el cliente.</p>
    <table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse">
      <thead>
        <tr>
          <th>Referencia</th>
          <th>Cliente</th>
          <th>Factura</th>
          <th>Número de Parte</th>
          <th>Descripción</th>
          <th>Fracción Actual</th>
          <th>Fecha Tráfico</th>
          <th>Partida</th>
        </tr>
      </thead>
      <tbody>
        ${rows}
      </tbody>
    </table>
  `;

  const mensaje = {
    from: 'sistemas@zayro.com',
    to: correos,
    cc:'programacion@zayro.com;gerenciati@zayro.com',
    subject: 'Reporte de fracciones IMMEX no autorizadas ' + resultados[0].Cliente,
    text: 'Por favor, revise el reporte de fracciones IMMEX no autorizadas adjunto o en el cuerpo del mensaje.',
    html: htmlBody,
  };
   const transport = nodemailer.createTransport(config);
   transport.verify().then(()=>console.log("Correo enviado...")).catch((error)=>console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if (error) {
        console.error('Error al enviar el correo:', error);
      } else {
        console.log('Correo enviado:', info.response);
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close()

   }); 
   //console.log(correos); 
} 
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_enviaranexo24semanalrochester', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
      const wb = new xl.Workbook();
      const nombreArchivo = "Reporte Anexo 24";
      const ws = wb.addWorksheet("Importación");
      const wsExpo = wb.addWorksheet("Exportación");

      const estiloTitulo = wb.createStyle({
          font: {
              name: 'Arial',
              color: '#FFFFFF',
              size: 10,
              bold: true,
          },
          fill: {
              type: 'pattern',
              patternType: 'solid',
              fgColor: '#630b57',
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
      });
      const estilocontenido = wb.createStyle({
          font: {
              name: 'Arial',
              color: '#000000',
              size: 10,
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
      });
      const estiloMoneda = wb.createStyle({
         numberFormat: '"$"#,##0.00', // formato con símbolo de dólar y dos decimales
         alignment: {
             horizontal: 'right'
         }
     });
      const columnas = [
         "Pedimento","Aduana","Clave","Fecha de Pago","Proveedor","Factura",	
         "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Origen","Tratado"
         ,"Cantidad UMComercial","UMComercial",	"Valor Aduana",	"Valor Comercial",	
         "TIGI","FP IGI","FP IVA","FP IEPS",	"Tipo de cambio"
     ];
      columnas.forEach((columna, index) => {
         ws.cell(1, index + 1).string(columna).style(estiloTitulo);
      });

     
      const result = await sql.sp_ObtenerPedimentos_Semanal(430);
      //console.log(result); 
      //PARA IMPORTASCION
      const facturas = await mysql.sp_ObtenerDatosFacturaSemanal(result[0].Pedimentos,1)

      //console.log(facturas); 
      var numfila=2;
      for (const factura of facturas) {

         //console.log('Pedimento:', factura.Pedimento);
         const informacionreporte=await sql.sp_ObtenerInformacionPedimento(430,1,factura.Pedimento,0) 
         //console.log(informacionreporte) 
         if (informacionreporte.length>0){
            
            for (const reporte of informacionreporte)
            {
               ws.cell(numfila, 1).string(reporte.Pedimento.toString()).style(estilocontenido);
               ws.cell(numfila, 2).string(reporte.Aduana.toString()).style(estilocontenido);
               ws.cell(numfila, 3).string(reporte.Clave.toString()).style(estilocontenido);

               ws.cell(numfila,4).string(factura.FechadePago).style(estilocontenido);
               ws.cell(numfila, 5).string(factura.Proveedor.toString()).style(estilocontenido);
               ws.cell(numfila, 6).string(reporte.Factura.toString()).style(estilocontenido);
               ws.cell(numfila,7).string(factura.FechaFactura).style(estilocontenido);

               ws.cell(numfila, 8).string(reporte.Producto.toString()).style(estilocontenido);
               ws.cell(numfila, 9).string(reporte.Fraccion.toString().substring(0, 8)).style(estilocontenido)
               ws.cell(numfila,10).string(reporte.OrigenDestino.toString()).style(estilocontenido);
               ws.cell(numfila,11).string(reporte.Tratado.toString()).style(estilocontenido);

               ws.cell(numfila,12).number(Number(reporte.CantidadUMComercial)).style(estilocontenido);
               ws.cell(numfila,13).string(factura.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
               ws.cell(numfila, 14).number(Number(factura.ValorAduana)).style(estiloMoneda);
               ws.cell(numfila, 15).number(Number(reporte.ValorComercial)).style(estiloMoneda);

               ws.cell(numfila,16).number(Number(reporte.TIGI)).style(estilocontenido);
               ws.cell(numfila,17).number(Number(reporte.FPIGI)).style(estilocontenido);
               ws.cell(numfila,18).number(Number(reporte.FPIVA)).style(estilocontenido);
               ws.cell(numfila,19).number(Number(reporte.FPIEPS)).style(estilocontenido);

               ws.cell(numfila,20).number(Number(factura.Tipodecambio)).style(estilocontenido);

               numfila++;
            }
            

         }

     }

     const columnasExpo = [
      "Pedimento","Aduana","Clave","Fecha de Pago","Cliente","Factura",	
      "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Destino"
      ,"Cantidad UMComercial","UMComercial",	"Valor Comercial",	
      "Valor USD",	"Tipo de cambio"
  ];
   columnasExpo.forEach((columna, index) => {
      wsExpo.cell(1, index + 1).string(columna).style(estiloTitulo);
   });
     const facturasExpo = await mysql.sp_ObtenerDatosFacturaexpoSemanal(result[0].Pedimentos,2)

     //console.log(facturas); 
     var numfilaExpo=2;
     for (const facturaexpo  of facturasExpo ) {

        console.log('Pedimento:', facturaexpo.Pedimento);
        const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(430,0,facturaexpo.Pedimento,facturaexpo.Partida) 
        //console.log(informacionreporteExpo ) 
        if (informacionreporteExpo.length>0){
         for (const reporteexpo of informacionreporteExpo)
            {
               wsExpo.cell(numfilaExpo, 1).string(reporteexpo.Pedimento.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 2).string(reporteexpo.Aduana.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 3).string(reporteexpo.Clave.toString()).style(estilocontenido);

               wsExpo.cell(numfilaExpo,4).string(facturaexpo.FechadePago).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 5).string(facturaexpo.Proveedor.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 6).string(reporteexpo.Factura.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo,7).string(facturaexpo.FechaFactura).style(estilocontenido);

               wsExpo.cell(numfilaExpo, 8).string(reporteexpo.Producto.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 9).string(reporteexpo.Fraccion.toString().substring(0, 8)).style(estilocontenido);

               wsExpo.cell(numfilaExpo,10).string(reporteexpo.OrigenDestino.toString()).style(estilocontenido);
            
               wsExpo.cell(numfilaExpo,11).number(Number(reporteexpo.CantidadUMComercial)).style(estilocontenido);
               wsExpo.cell(numfilaExpo,12).string(facturaexpo.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
               let valorComercial = Number(reporteexpo.ValorComercial);
               let tipoCambio = Number(facturaexpo.Tipodecambio);
               let resultado = valorComercial * tipoCambio;

               wsExpo.cell(numfilaExpo, 13).number(resultado).style(estiloMoneda);
               //wsExpo.cell(numfilaExpo,13).number(Number(reporteexpo.ValorComercial)).style(estiloMoneda);
               wsExpo.cell(numfilaExpo,14).number(Number(facturaexpo.ValorDolares)).style(estiloMoneda);

               wsExpo.cell(numfilaExpo,15).number(Number(facturaexpo.Tipodecambio)).style(estilocontenido);

               numfilaExpo++;
            }
           
        }
    }
     const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
      //Guardar
      await wb.write(pathExcel, (err, stats) => {
         if (err) {
            console.error('Error al guardar el archivo de Excel:', err);
            
         } else {
            console.log('Archivo de Excel guardado exitosamente en:', pathExcel);
            // Descargar el archivo de Excel
            res.download(pathExcel, nombreArchivo+'.xlsx', (err) => {
               if (err) {
                     console.error('Error al descargar el archivo:', err);
                     // Manejar el error
               } else {
                     console.log('Archivo descargado exitoso');
                   
               }
            });
         }
   });
   /*
     //PARA EXPORTACION
     const facturasExpo = await mysql.sp_ObtenerDatosFactura(result[0].Pedimentos,2)

      //console.log(facturas); 
      for (const facturaExpo  of facturasExpo ) {

         console.log('Pedimento:', facturaExpo.Pedimento);
         const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(2316,0,facturaExpo.Pedimento) 
         console.log(informacionreporteExpo ) 
         if (informacionreporteExpo.length>0){
            
         }
     }
     */
      await enviarMailAnexo24semanalrochester(nombreArchivo,transport);
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");  
   }
   

});
enviarMailAnexo24semanalrochester= async(nombreArchivo,transport) => {
      const meses = [
         "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
      ];
      
      const fechaActual = new Date();
      fechaActual.setMonth(fechaActual.getMonth() - 1); // Retrocede un mes
      
      const mesAnterior = meses[fechaActual.getMonth()];
      const año = fechaActual.getFullYear();
   const mensaje = {
      from:'sistemas@zayro.com',
      /*to: 'Programacion@zayro.com',*/
     to: 'jcadena@rochestersensors.com; mgarcia@rochestersensors.com; jackie.quinonez@rochestersensors.com',
      cc: 'avazquez@zayro.com; sistemas@zayro.com ',
      subject : `REPORTES SEMANALES PARA ANEXO 24 || ZAYRO || ${mesAnterior} ${año}`,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }],
      text: 'HOLA BUEN DÍA SE ANEXA ARCHIVO DE LOS REPORTES',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_enviaranexo24semanalthyssenkrup', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let cliente=1742;
      let transport = nodemailer.createTransport(config); 
      const wb = new xl.Workbook();
      const nombreArchivo = "Reporte Anexo 24";
      const ws = wb.addWorksheet("Importación");
      const wsExpo = wb.addWorksheet("Exportación");

      const estiloTitulo = wb.createStyle({
          font: {
              name: 'Arial',
              color: '#FFFFFF',
              size: 10,
              bold: true,
          },
          fill: {
              type: 'pattern',
              patternType: 'solid',
              fgColor: '#630b57',
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
      });
      const estilocontenido = wb.createStyle({
          font: {
              name: 'Arial',
              color: '#000000',
              size: 10,
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
          },
      });
      const estiloMoneda = wb.createStyle({
         numberFormat: '"$"#,##0.00', // formato con símbolo de dólar y dos decimales
         alignment: {
             horizontal: 'right'
         }
     });
      const columnas = [
         "Pedimento","Aduana","Clave","Fecha de Pago","Proveedor","Factura",	
         "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Origen","Tratado"
         ,"Cantidad UMComercial","UMComercial",	"Valor Aduana",	"Valor Comercial",	
         "TIGI","FP IGI","FP IVA","FP IEPS",	"Tipo de cambio",	"Partida Pedimento"
     ];
      columnas.forEach((columna, index) => {
         ws.cell(1, index + 1).string(columna).style(estiloTitulo);
      });

     
      const result = await sql.sp_ObtenerPedimentos_Semanal(cliente);
      //console.log(result); 
      //PARA IMPORTASCION
      const facturas = await mysql.sp_ObtenerDatosFacturaSemanal_Thyssen(result[0].Pedimentos,1)

      //console.log(facturas); 
      var numfila=2;
      for (const factura of facturas) {

         //console.log('Pedimento:', factura.Pedimento);
         const informacionreporte=await sql.sp_ObtenerInformacionPedimento(cliente,1,factura.Pedimento,0,factura.Partida) 
         //console.log(informacionreporte) 
         if (informacionreporte.length>0){
            
            for (const reporte of informacionreporte)
            {
               ws.cell(numfila, 1).string(reporte.Pedimento.toString()).style(estilocontenido);
               ws.cell(numfila, 2).string(reporte.Aduana.toString()).style(estilocontenido);
               ws.cell(numfila, 3).string(reporte.Clave.toString()).style(estilocontenido);

               ws.cell(numfila,4).string(factura.FechadePago).style(estilocontenido);
               ws.cell(numfila, 5).string(factura.Proveedor.toString()).style(estilocontenido);
               ws.cell(numfila, 6).string(reporte.Factura.toString()).style(estilocontenido);
               ws.cell(numfila,7).string(factura.FechaFactura).style(estilocontenido);

               ws.cell(numfila, 8).string(reporte.Producto.toString()).style(estilocontenido);
               ws.cell(numfila, 9).string(reporte.Fraccion.toString().substring(0, 8)).style(estilocontenido)
               ws.cell(numfila,10).string(reporte.OrigenDestino.toString()).style(estilocontenido);
               ws.cell(numfila,11).string(reporte.Tratado.toString()).style(estilocontenido);

               ws.cell(numfila,12).number(Number(reporte.CantidadUMComercial)).style(estilocontenido);
               ws.cell(numfila,13).string(factura.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
               ws.cell(numfila, 14).number(Number(factura.ValorAduana)).style(estiloMoneda);
               ws.cell(numfila, 15).number(Number(reporte.ValorComercial)).style(estiloMoneda);

               ws.cell(numfila,16).number(Number(reporte.TIGI)).style(estilocontenido);
               ws.cell(numfila,17).number(Number(reporte.FPIGI)).style(estilocontenido);
               ws.cell(numfila,18).number(Number(reporte.FPIVA)).style(estilocontenido);
               ws.cell(numfila,19).number(Number(reporte.FPIEPS)).style(estilocontenido);

               ws.cell(numfila,20).number(Number(factura.Tipodecambio)).style(estilocontenido);
               ws.cell(numfila,21).number(Number(reporte.Renglon)).style(estilocontenido);

               numfila++;
            }
            

         }

     }

     const columnasExpo = [
      "Pedimento","Aduana","Clave","Fecha de Pago","Cliente","Factura",	
      "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Destino"
      ,"Cantidad UMComercial","UMComercial",	"Valor Comercial",	
      "Valor USD",	"Tipo de cambio",	"Partida Pedimento"
  ];
   columnasExpo.forEach((columna, index) => {
      wsExpo.cell(1, index + 1).string(columna).style(estiloTitulo);
   });
     const facturasExpo = await mysql.sp_ObtenerDatosFacturaexpoSemanal_Thyssen(result[0].Pedimentos,2)

     //console.log(facturas); 
     var numfilaExpo=2;
     for (const facturaexpo  of facturasExpo ) {

        console.log('Pedimento:', facturaexpo.Pedimento);
        const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(cliente,0,facturaexpo.Pedimento,facturaexpo.Partida) 
        //console.log(informacionreporteExpo ) 
        if (informacionreporteExpo.length>0){
         for (const reporteexpo of informacionreporteExpo)
            {
               wsExpo.cell(numfilaExpo, 1).string(reporteexpo.Pedimento.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 2).string(reporteexpo.Aduana.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 3).string(reporteexpo.Clave.toString()).style(estilocontenido);

               wsExpo.cell(numfilaExpo,4).string(facturaexpo.FechadePago).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 5).string(facturaexpo.Proveedor.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 6).string(reporteexpo.Factura.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo,7).string(facturaexpo.FechaFactura).style(estilocontenido);

               wsExpo.cell(numfilaExpo, 8).string(reporteexpo.Producto.toString()).style(estilocontenido);
               wsExpo.cell(numfilaExpo, 9).string(reporteexpo.Fraccion.toString().substring(0, 8)).style(estilocontenido);

               wsExpo.cell(numfilaExpo,10).string(reporteexpo.OrigenDestino.toString()).style(estilocontenido);
            
               wsExpo.cell(numfilaExpo,11).number(Number(reporteexpo.CantidadUMComercial)).style(estilocontenido);
               wsExpo.cell(numfilaExpo,12).string(facturaexpo.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
               let valorComercial = Number(reporteexpo.ValorComercial);
               let tipoCambio = Number(facturaexpo.Tipodecambio);
               let resultado = valorComercial * tipoCambio;

               wsExpo.cell(numfilaExpo, 13).number(resultado).style(estiloMoneda);
               //wsExpo.cell(numfilaExpo,13).number(Number(reporteexpo.ValorComercial)).style(estiloMoneda);
               wsExpo.cell(numfilaExpo,14).number(Number(facturaexpo.ValorDolares)).style(estiloMoneda);

               wsExpo.cell(numfilaExpo,15).number(Number(facturaexpo.Tipodecambio)).style(estilocontenido);
               wsExpo.cell(numfilaExpo,16).number(Number(reporteexpo.Renglon)).style(estilocontenido);

               numfilaExpo++;
            }
           
        }
    }
     const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
      //Guardar
      await wb.write(pathExcel, (err, stats) => {
         if (err) {
            console.error('Error al guardar el archivo de Excel:', err);
            
         } else {
            console.log('Archivo de Excel guardado exitosamente en:', pathExcel);
            // Descargar el archivo de Excel
            res.download(pathExcel, nombreArchivo+'.xlsx', (err) => {
               if (err) {
                     console.error('Error al descargar el archivo:', err);
                     // Manejar el error
               } else {
                     console.log('Archivo descargado exitoso');
                   
               }
            });
         }
   });
   /*
     //PARA EXPORTACION
     const facturasExpo = await mysql.sp_ObtenerDatosFactura(result[0].Pedimentos,2)

      //console.log(facturas); 
      for (const facturaExpo  of facturasExpo ) {

         console.log('Pedimento:', facturaExpo.Pedimento);
         const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(2316,0,facturaExpo.Pedimento) 
         console.log(informacionreporteExpo ) 
         if (informacionreporteExpo.length>0){
            
         }
     }
     */
      await enviarMailAnexo24semanalthyssenkrup(nombreArchivo,transport);
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");  
   }
   

});
enviarMailAnexo24semanalthyssenkrup= async(nombreArchivo,transport) => {
      const meses = [
         "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
      ];
      
      const fechaActual = new Date();
      fechaActual.setMonth(fechaActual.getMonth()); // Retrocede un mes
      
      const mesAnterior = meses[fechaActual.getMonth()];
      const año = fechaActual.getFullYear();
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'mariana.gomez@thyssenkrupp-automotive.com;salvador.nieves-facundo@thyssenkrupp.com ',
      cc: 'avazquez@zayro.com; exportacion202@zayro.com;cchavana@zayro.com;sistemas@zayro.com ',
      subject : `REPORTES SEMANALES PARA ANEXO 24 || ZAYRO || ${mesAnterior} ${año}`,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }],
      text: 'HOLA BUEN DÍA SE ANEXA ARCHIVO DE LOS REPORTES',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/getdata_enviaranexo24semanalgeneral', async function(req, res, next) {
   try {
      let config = {
         host:process.env.hostemail,
         port:process.env.portemail,
         secure: true,
         auth: {
            user:process.env.useremail,
            pass:process.env.passemail
         },
         tls: {
            rejectUnauthorized: false
         }
      }
      let transport = nodemailer.createTransport(config); 
      const clientes=await sql.sp_obtenerclientesreporteanexo24();
      console.log(clientes.length)
      for (let i = 0; i < clientes.length; i++) {
         const c = clientes[i];
         console.log(`Cliente #${i + 1}:`);
         console.log(`  Número:      ${c.Numero}`);
         console.log(`  Cliente_id:  ${c.Cliente_id}`);
         console.log(`  Nombre:      ${c.Nom}`);
         console.log(`  RFC:         ${c.RFC}`);
         const wb = new xl.Workbook();
         const nombreArchivo = "Reporte Anexo 24 "+c.Numero;
         const ws = wb.addWorksheet("Importación");
         const wsExpo = wb.addWorksheet("Exportación");

         const estiloTitulo = wb.createStyle({
            font: {
               name: 'Arial',
               color: '#FFFFFF',
               size: 10,
               bold: true,
            },
            fill: {
               type: 'pattern',
               patternType: 'solid',
               fgColor: '#630b57',
            },
            alignment: {
               horizontal: 'center',
               vertical: 'center',
            },
         });
         const estilocontenido = wb.createStyle({
            font: {
               name: 'Arial',
               color: '#000000',
               size: 10,
            },
            alignment: {
               horizontal: 'center',
               vertical: 'center',
            },
         });
         const estiloMoneda = wb.createStyle({
            numberFormat: '"$"#,##0.00', // formato con símbolo de dólar y dos decimales
            alignment: {
               horizontal: 'right'
            }
      });
         const columnas = [
            "Pedimento","Aduana","Clave","Fecha de Pago","Proveedor","Factura",	
            "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Origen","Tratado"
            ,"Cantidad UMComercial","UMComercial",	"Valor Aduana",	"Valor Comercial",	
            "TIGI","FP IGI","FP IVA","FP IEPS",	"Tipo de cambio"
      ];
         columnas.forEach((columna, index) => {
            ws.cell(1, index + 1).string(columna).style(estiloTitulo);
         });
         let bandera=0;
      
         const result = await sql.sp_ObtenerPedimentos(c.Cliente_id);
         //console.log(result); 
         //PARA IMPORTACION
         //console.log(result.length); 
         if (result[0].Pedimentos !== null) {

            
      

            const facturas = await mysql.sp_ObtenerDatosFactura(result[0].Pedimentos,1)


            if (facturas.length>0){
               bandera=1;
            var numfila=2;
            for (const factura of facturas) {

               //console.log('Pedimento:', factura.Pedimento);
               const informacionreporte=await sql.sp_ObtenerInformacionPedimento(c.Cliente_id,1,factura.Pedimento,0) 
               //console.log(informacionreporte) 
               if (informacionreporte.length>0){
                  
                  for (const reporte of informacionreporte)
                  {
                     ws.cell(numfila, 1).string(reporte.Pedimento.toString()).style(estilocontenido);
                     ws.cell(numfila, 2).string(reporte.Aduana.toString()).style(estilocontenido);
                     ws.cell(numfila, 3).string(reporte.Clave.toString()).style(estilocontenido);

                     ws.cell(numfila,4).string(factura.FechadePago).style(estilocontenido);
                     ws.cell(numfila, 5).string(factura.Proveedor.toString()).style(estilocontenido);
                     ws.cell(numfila, 6).string(reporte.Factura.toString()).style(estilocontenido);
                     ws.cell(numfila,7).string(factura.FechaFactura).style(estilocontenido);

                     ws.cell(numfila, 8).string(reporte.Producto.toString()).style(estilocontenido);
                     ws.cell(numfila, 9).string(reporte.Fraccion.toString().substring(0, 8)).style(estilocontenido)
                     ws.cell(numfila,10).string(reporte.OrigenDestino.toString()).style(estilocontenido);
                     ws.cell(numfila,11).string(reporte.Tratado.toString()).style(estilocontenido);

                     ws.cell(numfila,12).number(Number(reporte.CantidadUMComercial)).style(estilocontenido);
                     ws.cell(numfila,13).string(factura.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
                     ws.cell(numfila, 14).number(Number(factura.ValorAduana)).style(estiloMoneda);
                     ws.cell(numfila, 15).number(Number(reporte.ValorComercial)).style(estiloMoneda);

                     ws.cell(numfila,16).number(Number(reporte.TIGI)).style(estilocontenido);
                     ws.cell(numfila,17).number(Number(reporte.FPIGI)).style(estilocontenido);
                     ws.cell(numfila,18).number(Number(reporte.FPIVA)).style(estilocontenido);
                     ws.cell(numfila,19).number(Number(reporte.FPIEPS)).style(estilocontenido);

                     ws.cell(numfila,20).number(Number(factura.Tipodecambio)).style(estilocontenido);
                     ws.cell(numfila,21).number(Number(reporte.Renglon)).style(estilocontenido);

                     numfila++;
                  }
                  

               }

            }
            }
               const columnasExpo = [
                  "Pedimento","Aduana","Clave","Fecha de Pago","Cliente","Factura",	
                  "Fecha Factura	",	"Clave de Insumo (NP)",	"Fraccion",	"Destino"
                  ,"Cantidad UMComercial","UMComercial",	"Valor Comercial",	
                  "Valor USD",	"Tipo de cambio"
            ];
            columnasExpo.forEach((columna, index) => {
               wsExpo.cell(1, index + 1).string(columna).style(estiloTitulo);
            });
            const facturasExpo = await mysql.sp_ObtenerDatosFacturaexpo(result[0].Pedimentos,2)
            if (facturasExpo.length>0){
            bandera=1;
            //console.log(facturas); 
            var numfilaExpo=2;
            for (const facturaexpo  of facturasExpo ) {

               //console.log('Pedimento:', facturaexpo.Pedimento);
               const informacionreporteExpo =await sql.sp_ObtenerInformacionPedimento(c.Cliente_id,0,facturaexpo.Pedimento,facturaexpo.Partida) 
               //console.log(informacionreporteExpo ) 
               if (informacionreporteExpo.length>0){
                  for (const reporteexpo of informacionreporteExpo)
                     {
                        wsExpo.cell(numfilaExpo, 1).string(reporteexpo.Pedimento.toString()).style(estilocontenido);
                        wsExpo.cell(numfilaExpo, 2).string(reporteexpo.Aduana.toString()).style(estilocontenido);
                        wsExpo.cell(numfilaExpo, 3).string(reporteexpo.Clave.toString()).style(estilocontenido);

                        wsExpo.cell(numfilaExpo,4).string(facturaexpo.FechadePago).style(estilocontenido);
                        wsExpo.cell(numfilaExpo, 5).string(facturaexpo.Proveedor.toString()).style(estilocontenido);
                        wsExpo.cell(numfilaExpo, 6).string(reporteexpo.Factura.toString()).style(estilocontenido);
                        wsExpo.cell(numfilaExpo,7).string(facturaexpo.FechaFactura).style(estilocontenido);

                        wsExpo.cell(numfilaExpo, 8).string(reporteexpo.Producto.toString()).style(estilocontenido);
                        wsExpo.cell(numfilaExpo, 9).string(reporteexpo.Fraccion.toString().substring(0, 8)).style(estilocontenido);

                        wsExpo.cell(numfilaExpo,10).string(reporteexpo.OrigenDestino.toString()).style(estilocontenido);
                     
                        wsExpo.cell(numfilaExpo,11).number(Number(reporteexpo.CantidadUMComercial)).style(estilocontenido);
                        wsExpo.cell(numfilaExpo,12).string(facturaexpo.UnidadMedidaComercial.toString().padStart(2, '0')).style(estilocontenido);
                        let valorComercial = Number(reporteexpo.ValorComercial);
                        let tipoCambio = Number(facturaexpo.Tipodecambio);
                        let resultado = valorComercial * tipoCambio;

                        wsExpo.cell(numfilaExpo, 13).number(resultado).style(estiloMoneda);
                        //wsExpo.cell(numfilaExpo,13).number(Number(reporteexpo.ValorComercial)).style(estiloMoneda);
                        wsExpo.cell(numfilaExpo,14).number(Number(facturaexpo.ValorDolares)).style(estiloMoneda);

                        wsExpo.cell(numfilaExpo,15).number(Number(facturaexpo.Tipodecambio)).style(estilocontenido);
                        wsExpo.cell(numfilaExpo,16).number(Number(reporteexpo.Renglon)).style(estilocontenido);

                        numfilaExpo++;
                     }
                  
               }
            }
            }
            if (bandera==1)
            {
               bandera=0;
         
               const pathExcel=path.join(__dirname,'excel',nombreArchivo+'.xlsx');
                  //Guardar
                  await wb.write(pathExcel, (err, stats) => {
                     if (err) {
                        console.error('Error al guardar el archivo de Excel:', err);
                        
                     } else {
                        console.log('Archivo de Excel guardado exitosamente en:', pathExcel);
                     
                        /*res.download(pathExcel, nombreArchivo+'.xlsx', (err) => {
                           if (err) {
                                 console.error('Error al descargar el archivo:', err);
                                 // Manejar el error
                           } else {*/
                                 console.log('Archivo descargado exitoso');
                              
                        /*   }
                        });*/
                     }
               });
            }

         
         }

      }
      if (i == (clientes.length-1)){
         res.json('reportes anexo 24 enviados')
      }

   
     await enviarMailAnexo24semanalthyssenkrup(nombreArchivo,transport);
   } catch (err) {
       console.error('EL ERROR ES ' + err);
       res.status(500).send("Error al obtener los datos de la base de datos.");  
   }

});
enviarMailAnexo24semanalgeneral= async(nombreArchivo,transport) => {
      const meses = [
         "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
         "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
      ];
      
      const fechaActual = new Date();
      fechaActual.setMonth(fechaActual.getMonth() - 1); // Retrocede un mes
      
      const mesAnterior = meses[fechaActual.getMonth()];
      const año = fechaActual.getFullYear();
   const mensaje = {
      from:'sistemas@zayro.com',
      to: 'programacion@zayro.com',
      //cc: 'avazquez@zayro.com;sistemas@zayro.com ',
      subject : `REPORTES MENSUALES PARA ANEXO 24 || ZAYRO || ${mesAnterior} ${año}`,
      attachments: [
         {
            filename: nombreArchivo +'.xlsx',
            path: './src/excel/' + nombreArchivo + '.xlsx',
         }],
      text: 'HOLA BUEN DÍA SE ANEXA ARCHIVO DE LOS REPORTES',
   }
   console.log(mensaje)
   transport.verify().then(() => console.log("Correo Enviado...")).catch((error) => console.log(error));
   transport.sendMail(mensaje,(error, info) => {
      if(error) {
         console.error('Error al enviar el correo:', error)
      } else {
         console.log('Correo enviado:', info.response);
      }

      transport.close()
   });
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/

app.get('/api/procesar-correo', async (req, res) => {
   const comandoLeerMail = 'python leermailData.py';
 
   exec(comandoLeerMail, (error, stdout, stderr) => {
     if (error) {
       console.error(`Error al ejecutar el script de leermail: ${error.message}`);
       return res.status(500).json({ error: 'Error al ejecutar el script de leermail' });
     }
     if (stderr) {
       console.error(`Error en la salida estándar de leermail: ${stderr}`);
       return res.status(500).json({ error: 'Error en la salida estándar de leermail' });
     }
    //console.log(`Salida del script de leermail: ${stdout}`);
 
     const carpeta = './archivos_adjuntos';
     fs.readdir(carpeta, (error, archivos) => {
       if (error) {
         console.error(`Error al leer la carpeta: ${error}`);
         return res.status(500).json({ error: 'Error al leer la carpeta' });
       }
 
       const nombresArchivos = { archivos: archivos };
 
       fs.writeFile('nombres_archivos.json', JSON.stringify(nombresArchivos), (error) => {
         if (error) {
           console.error(`Error al escribir en el archivo JSON: ${error}`);
           return res.status(500).json({ error: 'Error al escribir en el archivo JSON' });
         }
         const nombresArchivosJSON = JSON.stringify(nombresArchivos);
         if (nombresArchivosJSON.length > 2) {
           //console.log('Nombres de archivos almacenados en nombres_archivos.json' + JSON.stringify(nombresArchivos));
           archivos.forEach(async (nombreArchivo) => {
             const nombreSinExtension = nombreArchivo.slice(0, -5); // Elimina los últimos 5 caracteres (.xlsx)
             //console.log("Valor individual:", nombreSinExtension);
             const result=await sql.altafacturasencontradas(nombreSinExtension)
             console.log(result)
           });
 
         }
         else {
           console.log('Sin info');

         }
 
         const comandoEliminarArchivos = `python eliminar_archivos.py ${carpeta}`;
         exec(comandoEliminarArchivos, (error, stdout, stderr) => {
           if (error) {
             console.error(`Error al ejecutar el script de eliminar_archivos: ${error}`);
             return res.status(500).json({ error: 'Error al ejecutar el script de eliminar_archivos' });
           }
           if (stderr) {
             console.error(`Error en la salida estándar de eliminar_archivos: ${stderr}`);
             return res.status(500).json({ error: 'Error en la salida estándar de eliminar_archivos' });
           }
           console.log(`Salida del script de eliminar_archivos: ${stdout}`);
           return res.status(200).json({ message: 'Proceso de correo completado exitosamente' });
         });
       });
     });
   });
});
const columnasTablas = {
  Inci: [
    'Patente','Pedimento','SeccionAduanera','ConsecutivoRemesa','NumeroSeleccion',
    'FechaInicioReconocimiento','HoraInicioReconocimiento','FechaFinReconocimiento','HoraFinReconocimiento',
    'Fraccion','SecuenciaFraccion','ClaveDocumento','TipoOperacion','GradoIncidencia','FechaSeleccion'
  ],
  Sel: [
    'Patente','Pedimento','SeccionAduanera','ConsecutivoRemesa','NumeroSeleccion','FechaSeleccion','HoraSeleccion','SemaforoFiscal','ClaveDocumento','TipoOperacion'
  ],
  tabla501: [
    'Patente','Pedimento','SeccionAduanera','TipoOperacion','ClaveDocumento',
    'SeccionAduaneraEntrada','CurpContribuyente','Rfc','CurpAgenteA','TipoCambio',
    'TotalFletes','TotalSeguros','TotalEmbalajes','TotalIncrementables','TotalDeducibles',
    'PesoBrutoMercancia','MedioTransporteSalida','MedioTransporteArribo','MedioTransporteEntrada_Salida',
    'DestinoMercancia','NombreContribuyente','CalleContribuyente','NumInteriorContribuyente',
    'NumExteriorContribuyente','CPContribuyente','MunicipioContribuyente','EntidadFedContribuyente',
    'PaisContribuyente','TipoPedimento','FechaRecepcionPedimento','FechaPagoReal'
  ],
  tabla502: [
    'Patente','Pedimento','SeccionAduanera','RfcTransportista','CurpTransportista','NombreTransportista','PaisTransporte','IdentificadorTransporte','FechaPagoReal'
  ],
  tabla503: [
    'Patente','Pedimento','SeccionAduanera','NumeroGuia','TipoGuia','FechaPagoReal'
  ],
  tabla504: [
    'Patente','Pedimento','SeccionAduanera','NumContenedor','TipoContenedor','FechaPagoReal'
  ],
 
tabla505: [
  'Patente','Pedimento','SeccionAduanera','FechaFacturacion','NumeroFactura',
  'TerminoFacturacion','MonedaFacturacion','ValorDolares','ValorMonedaExtranjera',
  'PaisFacturacion','EntidadFedFacturacion','IndentFiscalProveedor','ProveedorMercancial',
  'CalleProveedor','NumInteriorProveedor','NumExteriorProveedor','CpProveedor',
  'MunicipioProveedor','FechaPagoReal'
],

  tabla506: [
    'Patente','Pedimento','SeccionAduanera','TipoFecha','FechaOperacion','FechaValidacionPagoR'
  ],
  tabla507: [
    'Patente','Pedimento','SeccionAduanera','ClaveCaso','IdentificadorCaso','TipoPedimento','ComplementoCaso','FechaPagoReal'
  ],
  tabla508: [
    'Patente','Pedimento','SeccionAduanera','InstitucionEmisora','NumeroCuenta','FolioConstancia','FechaConstancia','TipoCuenta','ClaveGarantia','ValorUnitarioTitulo','TotalGarantia','CantidadUnidades','TitulosAsignados','FechaPagoReal'
  ],
  tabla509: [
    'Patente','Pedimento','SeccionAduanera','ClaveContribucion','TasaContribucion','TipoTasa','TipoPedimento','FechaPagoReal'
  ],
  tabla510: [
    'Patente','Pedimento','SeccionAduanera','ClaveContribucion','FormaPago','ImportePago','TipoPedimento','FechaPagoReal'
  ],
  tabla511: [
    'Patente','Pedimento','SeccionAduanera','SecuenciaObservacion','Observaciones','TipoPedimento','FechaPagoReal'
  ],
  tabla512: [
    'Patente','Pedimento','SeccionAduanera','PatenteAduanalOrig','PedimentoOriginal','SeccionAduaneraDespOrig','DocumentoOriginal','FechaOperacionOrig','FraccionOriginal','UnidadMedida','MercanciaDescargada','TipoPedimento','FechaPagoReal'
  ],
  tabla513: [
    'Patente','Pedimento','SeccionAduanera','PatenteAutorizadaOrig','NumPedimentoOrig','SeccionAduanaOrig','FechaPagoOriginal','TipoContribucionCompensada','ImporteCompensado','TipoPedimento','FechaPagoReal'
  ],
  // === 520: usa el nombre correcto ===
tabla520: [
  'Patente','Pedimento','SeccionAduanera','IdentFiscalDestinatario',
  'NombreDestinatarioMercancia','CalleDestinatario','NumInteriorDestinatario',
  'NumExteriorDestinatario','CpDestinatario','MunicipioDestinatario',
  'PaisDestinatario','FechaPagoReal'
],
   tabla551: [
   'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','SubdivisionFraccion',
   'DescripcionMercancia','PrecioUnitario','ValorAduana','ValorComercial','ValorDolares',
   'CantidadUMComercial','UnidadMedidaComercial','CantidadUMTarifa','UnidadMedidaTarifa',
   'ValorAgregaddo','ClaveVinculacion','MetodoValorizacion','CodigoMercanciaProducto',
   'MarcaMercanciaProducto','ModeloMercanciaProducto','PaisOrigenDestino','PaisCompradorVendedor',
   'EntidadFedOrigen','EntidadFedDestino','EntidadFedComprador','EntidadFedVendedor',
   'TipoOperacion','ClaveDocumento','FechaPagoReal'
    
  ],
  tabla552: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','VinNumeroSerie','KilometrajeVehiculo','FechaPagoReal'
  ],
  tabla553: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','ClavePermiso','FirmaDescargo','NumeroPermiso','ValorComercialDolares','CantidadMUMTarifa','FechaPagoReal'
  ],
  tabla554: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','ClaveCaso','IdentificadorCaso','ComplementoCaso','FechaPagoReal'
  ],
  tabla555: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','InstitucionEmisora','NumeroCuenta','FolioConstancia','FechaConstancia','ClaveGarantia','ValorUnitarioTitulo','TotalGarantia','CantidadUnidadesMedida','TitulosAsignados','FechaPagoReal'
  ],
  tabla556: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','ClaveContribucion','TasaContribucion','TipoTasa','FechaPagoReal'
  ],
  tabla557: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','ClaveContribucion','FormaPago','ImportePago','FechaPagoReal'
  ],
  tabla558: [
    'Patente','Pedimento','SeccionAduanera','Fraccion','SecuenciaFraccion','SecuenciaObservacion','Observaciones','FechaPagoReal'
  ],
  tabla701: [
    'Patente','Pedimento','SeccionAduanera','ClaveDocumento','FechaPago','PedimentoAnterior','PatenteAnterior','SeccionAduaneraAnterior','DocumentoAnterior','FechaOperacionAnterior','PedimentoOriginal','PatenteAduanalOriginal','SeccionAduaneraDespOrig','FechaPagoReal'
  ],
  tabla702: [
    'Patente','Pedimento','SeccionAduanera','ClaveContribucion','FormaPago','ImportePago','TipoPedimento','FechaPagoReal'
  ]
};

// ---- 2. Alias automáticos (agrega aquí cualquier variante real de tus archivos) ----
const aliasColumnas = {

  FechaValidacionPagoR: 'FechaPagoReal',
  FechaOperacion: 'Fechaoperacion',
  FraccionOriginal: 'Fraccionoriginal',
  ValorAgregado: 'ValorAgregaddo',
  PatenteAduanalOrig: 'PatenteAduanalOriginal',

 // === Nuevos (corrigen los .asc que recibes) ===
  // 505
  ProveedorMercancia: 'ProveedorMercancial',
  CalleProveerdor: 'CalleProveedor',                 // por si aparece esta variante
  EntidadFederativa: 'EntidadFedFacturacion',        // por si te mandan esta abreviación

    // 520
  IndentFiscalDestinatario: 'IdentFiscalDestinatario',
  MunicpioDestinatario: 'MunicipioDestinatario',
  IdentidadFiscalProveedor: 'IndentFiscalProveedor',

  // 702 (variantes posibles)
'Tipo_Pedimento': 'TipoPedimento',
  'TIPO_PEDIMENTO': 'TipoPedimento',
  'Tipo pedimento': 'TipoPedimento',
  'Tipopedimento' : 'TipoPedimento',
  'TP'            : 'TipoPedimento',
  'Tipo de Pedimento': 'TipoPedimento',   // extra
  'Tipo de pedimento': 'TipoPedimento',   // extra

  // por si en 505 te llega la otra variante de ident fiscal:
  IdentidadFiscalProveedor: 'IndentFiscalProveedor',
};

// ---- 3. Define columnas numéricas por tabla (ajusta a tu realidad) ----
const columnasNumericas = {
  tabla501: ['TipoCambio', 'TotalFletes', 'TotalSeguros', 'TotalEmbalajes', 'TotalIncrementables', 'TotalDeducibles', 'PesoBrutoMercancia'],
  tabla505: ['ValorDolares', 'ValorMonedaExtranjera'],
  tabla551: ['PrecioUnitario', 'ValorAduana', 'ValorComercial', 'ValorDolares', 'CantidadUMComercial', 'CantidadUMTarifa', 'ValorAgregaddo'],
  tabla702: ['ClaveContribucion','FormaPago','ImportePago','TipoPedimento'],
};

const ASC_DIR = path.join(__dirname, 'archivos_adjuntos');


function stripBOM(s) {
  if (s == null) return s;
  if (typeof s !== 'string') return s;
  if (s.charCodeAt(0) === 0xFEFF) return s.slice(1);
  return s.replace(/^\uFEFF/, '');
}

// --- normaliza un header a minúsculas y alias canónico
function normalizaCol(col) {
  const limpio = stripBOM(String(col || '').trim());
  const canon  = aliasColumnas[limpio] || limpio;
  return canon.toLowerCase();
}

// --- NORMALIZA HEADERS (versión única; elimina la duplicada más abajo)
function normalizaHeaders(headers) {
  return headers.map(h => {
    const limpio = stripBOM(String(h || '').trim());
    return aliasColumnas[limpio] || limpio;
  });
}

function detectaTablaPorHeader(headers) {
  const normHeaders = headers.map(h => normalizaCol(stripBOM(h)));

  // Firma mínima OBLIGATORIA para 702
  const firma702 = [
    'patente','pedimento','seccionaduanera',
    'clavecontribucion','formapago','importepago',
    'tipopedimento','fechapagoreal'
  ];
  if (firma702.every(c => normHeaders.includes(c))) return 'tabla702';

  let candidato = null, best = -1, faltantes = [], extras = [];
  for (const [tabla, columnas] of Object.entries(columnasTablas)) {
    const colsNorm = columnas.map(normalizaCol);
    const comunes = colsNorm.filter(c => normHeaders.includes(c));
    if (comunes.length === colsNorm.length) return tabla;

    if (comunes.length > best) {
      candidato = tabla;
      best = comunes.length;
      faltantes = colsNorm.filter(c => !normHeaders.includes(c));
      extras    = normHeaders.filter(c => !colsNorm.includes(c));
    }
  }
  if (candidato) {
    console.warn(`[ASC] Header parece de "${candidato}" pero faltan: ${faltantes.join(', ')}; extras: ${extras.join(', ')}`);
    console.warn(`[ASC] Headers normalizados: ${JSON.stringify(normHeaders)}`);
  } else {
    console.warn(`[ASC] Headers desconocidos: ${JSON.stringify(normHeaders)}`);
  }
  return null;
}


// --- Limpia numéricos (null si vacío/"null", cambia "," por ".") ---
function limpiaValoresNumericos(tabla, obj) {
  (columnasNumericas[tabla] || []).forEach(col => {
    let v = obj[col];
    if (v === '' || v === '-' || v === null || v?.toLowerCase?.() === 'null') obj[col] = null;
    else if (typeof v === 'string') obj[col] = v.replace(/,/g, '.');
  });
  return obj;
}

function normalizaHeaders(headers) {
  return headers.map(h => {
    const limpio = stripBOM(String(h || '').trim());
    return aliasColumnas[limpio] || limpio;
  });
}
const archivosNoValidados={}
/* ------------ HELPERS PARA REPORTE ------------ */
function calcVaciosPorTabla(tablasJSON = {}) {
  const resultado = {};
  for (const [tabla, rows] of Object.entries(tablasJSON)) {
    if (!Array.isArray(rows) || rows.length === 0) continue;
    const cols = Object.keys(rows[0] ?? {});
    const cont = Object.fromEntries(cols.map(c => [c, 0]));
    let totalCeldas = 0, totalVacios = 0;

    for (const r of rows) {
      for (const c of cols) {
        totalCeldas++;
        const v = r[c];
        const vacio = v === '' || v === null || v === undefined ||
                      (typeof v === 'string' && (v.trim() === '' || v.trim().toLowerCase() === 'null' || v.trim() === '-'));
        if (vacio) {
          cont[c]++;
          totalVacios++;
        }
      }
    }
    const porCol = cols
      .map(c => ({
        col: c,
        vacios: cont[c],
        porcentaje: rows.length ? Math.round((cont[c] / rows.length) * 100) : 0
      }))
      .filter(x => x.vacios > 0)
      .sort((a, b) => b.vacios - a.vacios)
      .slice(0, 5); // Top 5 columnas con más vacíos

    resultado[tabla] = { totalVacios, totalCeldas, porCol };
  }
  return resultado;
}

function fmt(num) {
  return new Intl.NumberFormat('es-MX').format(num ?? 0);
}

/* ------------ MAIL HELPER (NUEVO) ------------ */
async function enviarResumenCarga({
  resumenOK = [],            // [{ tabla, registros }]
  resumenVacio = [],         // ['tablaXXX', ...]
  archivosNoValidados = [],  // ['123.asc', ...]
  exito = true,
  errorMsg = '',
  tablasJSON = {}
}) {
  const kpiTablasConDatos    = resumenOK.length;
  const kpiFilasInsertadas   = resumenOK.reduce((acc, r) => acc + (r.registros || 0), 0);
  const kpiTablasVacias      = resumenVacio.length;
  const kpiArchivosInvalidos = archivosNoValidados.length;

  const vacios     = calcVaciosPorTabla(tablasJSON);
  const totalNulls = Object.values(vacios).reduce((a, info) => a + (info?.totalVacios || 0), 0);
  const hayVacios  = totalNulls > 0;

  const fechaStr = new Date().toLocaleString('es-MX', { hour12: false });
  const brand    = '#0F62FE';
  const gray700  = '#374151';
  const gray600  = '#4b5563';
  const gray500  = '#6b7280';
  const gray200  = '#e5e7eb';
  const gray100  = '#f3f4f6';

  // ---------- HTML formal / minimal ----------
  const html = `
  <div style="margin:0;padding:20px;background:#f7f7f8;font-family:Segoe UI,Arial,sans-serif;color:${gray700};">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:760px;margin:0 auto;background:#ffffff;border:1px solid ${gray200};border-radius:12px;overflow:hidden;">
      <tr><td style="background:${brand};height:6px;"></td></tr>
      <tr>
        <td style="padding:16px 20px;border-bottom:1px solid ${gray200};">
          <div style="font-size:18px;font-weight:700;letter-spacing:.2px;">Reporte de Carga ASC ${exito ? '<span style="color:#1b7f42">· OK</span>' : '<span style="color:#b42318">· ERROR</span>'}</div>
          <div style="font-size:12px;color:${gray500};margin-top:2px;">${fechaStr}</div>
        </td>
      </tr>

      <tr>
        <td style="padding:16px 20px;">
          <div style="font-size:13px;color:${gray600};line-height:1.5;">
            Este informe resume la carga de archivos <strong>.asc</strong>. Los valores vacíos
            (<code style="background:${gray100};padding:0 4px;border-radius:4px;">''</code>,
             <code style="background:${gray100};padding:0 4px;border-radius:4px;">-</code>,
             <code style="background:${gray100};padding:0 4px;border-radius:4px;">"null"</code>)
            se insertaron como <strong>NULL</strong>.
          </div>
        </td>
      </tr>

      <tr>
        <td style="padding:8px 12px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:8px;">
            <tr>
              <td style="width:25%;border:1px solid ${gray200};border-radius:10px;padding:10px;">
                <div style="font-size:12px;color:${gray500};">Tablas con datos</div>
                <div style="font-size:22px;font-weight:700;margin-top:4px;">${fmt(kpiTablasConDatos)}</div>
              </td>
              <td style="width:25%;border:1px solid ${gray200};border-radius:10px;padding:10px;">
                <div style="font-size:12px;color:${gray500};">Filas insertadas</div>
                <div style="font-size:22px;font-weight:700;margin-top:4px;">${fmt(kpiFilasInsertadas)}</div>
              </td>
              <td style="width:25%;border:1px solid ${gray200};border-radius:10px;padding:10px;">
                <div style="font-size:12px;color:${gray500};">Campos vacíos (NULL)</div>
                <div style="font-size:22px;font-weight:700;margin-top:4px;">${fmt(totalNulls)}</div>
              </td>
              <td style="width:25%;border:1px solid ${gray200};border-radius:10px;padding:10px;">
                <div style="font-size:12px;color:${gray500};">Archivos no válidos</div>
                <div style="font-size:22px;font-weight:700;margin-top:4px;">${fmt(kpiArchivosInvalidos)}</div>
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <tr>
        <td style="padding:6px 20px 14px;">
          <div style="font-weight:700;margin:8px 0 6px;font-size:14px;">Tablas con registros</div>
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${gray200};border-radius:8px;overflow:hidden;font-size:13px;">
            <thead>
              <tr style="background:${gray100};">
                <th align="left"  style="padding:10px;border-bottom:1px solid ${gray200};">Tabla</th>
                <th align="right" style="padding:10px;border-bottom:1px solid ${gray200};"># filas</th>
              </tr>
            </thead>
            <tbody>
              ${
                resumenOK.length
                ? resumenOK.map(r => `
                  <tr>
                    <td style="padding:10px;border-bottom:1px solid ${gray200};">${r.tabla}</td>
                    <td style="padding:10px;border-bottom:1px solid ${gray200};text-align:right;">${fmt(r.registros)}</td>
                  </tr>
                `).join('')
                : `<tr><td colspan="2" style="padding:12px;color:${gray500};">— Ninguna —</td></tr>`
              }
            </tbody>
          </table>
        </td>
      </tr>

      <tr>
        <td style="padding:0 20px 14px;">
          <div style="font-weight:700;margin:10px 0 6px;font-size:14px;">Campos vacíos convertidos a NULL</div>
          ${
            Object.keys(vacios).length && hayVacios
            ? Object.entries(vacios).map(([tabla, info]) => {
                if (!info || info.totalVacios === 0) return '';
                const filas = (tablasJSON?.[tabla]?.length) || 0;
                return `
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin:8px 0;border:1px solid ${gray200};border-radius:8px;overflow:hidden;font-size:12px;">
                  <tr style="background:${gray100};">
                    <td colspan="3" style="padding:8px 10px;border-bottom:1px solid ${gray200};">
                      <strong>${tabla}</strong>
                      <span style="color:${gray500}"> · Vacíos: ${fmt(info.totalVacios)} · Registros: ${fmt(filas)}</span>
                    </td>
                  </tr>
                  <tr>
                    <th align="left"  style="padding:8px 10px;border-bottom:1px solid ${gray200};">Columna</th>
                    <th align="right" style="padding:8px 10px;border-bottom:1px solid ${gray200};"># vacíos</th>
                    <th align="right" style="padding:8px 10px;border-bottom:1px solid ${gray200};">% sobre filas</th>
                  </tr>
                  ${
                    info.porCol.map(c => `
                      <tr>
                        <td style="padding:8px 10px;border-bottom:1px solid ${gray200};">${c.col}</td>
                        <td style="padding:8px 10px;border-bottom:1px solid ${gray200};text-align:right;">${fmt(c.vacios)}</td>
                        <td style="padding:8px 10px;border-bottom:1px solid ${gray200};text-align:right;">${fmt(c.porcentaje)}%</td>
                      </tr>
                    `).join('')
                  }
                </table>`;
              }).join('')
            : `<div style="padding:10px;border:1px solid ${gray200};border-radius:8px;color:${gray500};">— No se detectaron campos vacíos —</div>`
          }
          <div style="margin-top:6px;font-size:12px;color:${gray500}">
            Nota: los valores vacíos indicados arriba se insertaron como <strong>NULL</strong>.
          </div>
        </td>
      </tr>

      <tr>
        <td style="padding:0 20px 16px;">
          <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:separate;border-spacing:0;">
            <tr>
              <td style="vertical-align:top;padding-right:8px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${gray200};border-radius:8px;">
                  <tr><td style="font-weight:700;padding:10px;border-bottom:1px solid ${gray200};">Tablas sin registros</td></tr>
                  <tr><td style="padding:10px;font-size:13px;color:${gray700};">${kpiTablasVacias ? resumenVacio.join(', ') : `<span style="color:${gray500}">— Ninguna —</span>`}</td></tr>
                </table>
              </td>
              <td style="vertical-align:top;padding-left:8px;">
                <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid ${gray200};border-radius:8px;">
                  <tr><td style="font-weight:700;padding:10px;border-bottom:1px solid ${gray200};">Archivos no válidos</td></tr>
                  <tr><td style="padding:10px;font-size:13px;color:${gray700};">${kpiArchivosInvalidos ? archivosNoValidados.join(', ') : `<span style="color:${gray500}">— Ninguno —</span>`}</td></tr>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>

      <tr>
        <td style="padding:12px 20px;border-top:1px solid ${gray200};">
          <div style="font-weight:700;color:${exito ? '#1b7f42' : '#b42318'};font-size:14px;">
            ${exito ? '✔ Carga completada sin errores.' : '✖ La carga presentó errores.'}
          </div>
          ${!exito ? `<div style="margin-top:6px;font-size:12px;color:${gray500}">${errorMsg}</div>` : ''}
          <div style="margin-top:10px;font-size:12px;color:${gray500}">Reporte automático · ${fechaStr}</div>
        </td>
      </tr>
    </table>
  </div>`;

  // ---------- Texto plano ----------
  let text = `Reporte de Carga ASC ${exito ? '(OK)' : '(ERROR)'}\n${fechaStr}\n\n`;
  text += `Tablas con datos: ${kpiTablasConDatos}\n`;
  text += `Filas insertadas: ${kpiFilasInsertadas}\n`;
  text += `Campos vacíos (NULL): ${totalNulls}\n`;
  text += `Tablas vacías: ${kpiTablasVacias}\n`;
  text += `Archivos no válidos: ${kpiArchivosInvalidos}\n\n`;
  text += `Tablas con registros:\n`;
  if (resumenOK.length) resumenOK.forEach(r => { text += `  - ${r.tabla}: ${r.registros}\n`; });
  else text += `  — Ninguna —\n`;
  text += `\nCampos vacíos convertidos a NULL:\n`;
  if (hayVacios) {
    for (const [tabla, info] of Object.entries(vacios)) {
      if (!info || info.totalVacios === 0) continue;
      const filas = (tablasJSON?.[tabla]?.length) || 0;
      text += `  - ${tabla}: vacíos=${info.totalVacios} / registros=${filas}\n`;
      info.porCol.forEach(c => { text += `      · ${c.col}: ${c.vacios} (${c.porcentaje}%)\n`; });
    }
  } else {
    text += `  — No se detectaron campos vacíos —\n`;
  }
  text += `\nTablas sin registros: ${kpiTablasVacias ? resumenVacio.join(', ') : '— Ninguna —'}\n`;
  text += `Archivos no válidos: ${kpiArchivosInvalidos ? archivosNoValidados.join(', ') : '— Ninguno —'}\n`;
  text += `\n${exito ? '✔ Carga completada sin errores.' : '✖ La carga presentó errores: ' + errorMsg}\n`;


  /* ---------- envío ---------- */
  const transport = nodemailer.createTransport({
    host  : process.env.hostemail,
    port  : process.env.portemail,
    secure: true,
    auth  : { user: process.env.useremail, pass: process.env.passemail },
    tls   : { rejectUnauthorized:false }
  });

  await transport.sendMail({
    from   : `"Carga ASC" <${process.env.useremail}>`,
    to     : process.env.destinatarios || 'programacion2@zayro.com',
    subject: exito ? 'Carga de Archivos DataStage – OK' : 'Carga de Archivos DataStage – ERROR',
    text,
    html
  });
}


// ----------- ENDPOINT ROBUSTO /procesarasc ------------------
app.get('/api/procesarasc', async (req, res) => {
  try {

   await new Promise((resolve, reject) => {
      exec('python leermaildata.py', (err, stdout, stderr) => {
        if (err) {
          console.error('Error ejecutando script Python:', stderr || err);
          return reject(new Error('Error ejecutando script Python.'));
        }
        resolve();
      });
    });
/* 1)  lee carpeta */
    if(!fs.existsSync(ASC_DIR)) fs.mkdirSync(ASC_DIR);
    const files = fs.readdirSync(ASC_DIR).filter(f => f.toLowerCase().endsWith('.asc'));

    const tablasJSON={}, archivosNoValidos=[];
    for(const file of files){
      const lines = fs.readFileSync(path.join(ASC_DIR,file),'utf8')
                      .split(/\r?\n/).filter(l=>l.trim());
      if(lines.length<2) continue;

      const headers = lines[0].split('|').map(h=>h.trim());
      const tabla   = detectaTablaPorHeader(headers);
      if(!tabla){ archivosNoValidos.push(file); continue; }

      const headersNorm = normalizaHeaders(headers);
      for(const ln of lines.slice(1)){
        const vals = ln.split('|').map(v=>v.trim());
        const obj={};
        columnasTablas[tabla].forEach(col=>{
          const idx = headersNorm.findIndex(h=>h.toLowerCase()===col.toLowerCase());
          obj[col]= idx>-1 ? vals[idx] : '';
        });
        limpiaValoresNumericos(tabla,obj);
        (tablasJSON[tabla]=tablasJSON[tabla]||[]).push(obj);
      }
    }

    /* 2) resumen */
    const resumenOK     = Object.entries(tablasJSON).map(([t,a])=>({tabla:t,registros:a.length}));
    const tablasVacias  = Object.keys(columnasTablas)
                                .filter(t => !(t in tablasJSON));

    /* 3) ejecuta SP */
    let ok=true, errMsg='';
    try{ await sqlData.ejecutaSPconJSON(tablasJSON); }
    catch(e){ ok=false; errMsg=e.message||'Fallo SP'; }

    /* 4) correo */
    await enviarResumenCarga({
      resumenOK,
      resumenVacio: tablasVacias,
      archivosNoValidados: archivosNoValidos,
      exito: ok,
      errorMsg: errMsg
    });

    /* 5) respuesta HTTP */
    if(ok) res.json({ok:true,resumenOK,tablasVacias,archivosNoValidados});
    else   res.status(500).json({ok:false,error:errMsg,resumenOK,tablasVacias,archivosNoValidados});

  }catch(e){
    console.error(e);
    res.status(500).json({ok:false,error:e.message||e});
  }
});

/*************************************************************************************** */
app.get('/api/revisarexisteboxid', async (req, res) => {
  const comandoLeerMail = 'python leermailboxid.py';

  exec(comandoLeerMail, (error, stdout, stderr) => {
    if (error) {
        console.error(`Error al ejecutar el script de leermail: ${error.message}`);
      return res.status(500).json({ error: 'Error al ejecutar el script de leermail' });
    }
    if (stderr) {
      console.error(`Error en la salida estándar de leermail: ${stderr}`);
      return res.status(500).json({ error: 'Error en la salida estándar de leermail' });
    }

    const carpeta = './archivos_boxid';
    fs.readdir(carpeta, async (error, archivos) => {
      if (error) {
        console.error(`Error al leer la carpeta: ${error}`);
        return res.status(500).json({ error: 'Error al leer la carpeta' });
      }

      if (archivos.length === 0) {
        return res.status(400).json({ error: 'No se encontraron archivos en la carpeta' });
      }

      // Procesar cada archivo Excel
      for (const archivo of archivos) {
        const rutaArchivo = path.join(carpeta, archivo);

        // Solo procesar archivos Excel
        if (rutaArchivo.endsWith('.xlsx')) {
          console.log(`Procesando archivo: ${archivo}`);
          if (!archivo.startsWith('LCN')) {
            console.log(`El archivo ${archivo} no comienza con 'LCN'. Ignorando...`);
            // Si se va por el else, eliminamos el archivo
            fs.unlink(rutaArchivo, (err) => {
              if (err) {
                console.error(`Error al eliminar el archivo ${archivo}:`, err);
              } else {
                console.log(`Archivo ${archivo} eliminado correctamente.`);
              }
            });
            
          }
          else {
            
          

          // Leer el archivo Excel
          const workbook = xlsx.readFile(rutaArchivo);
          const sheetName = workbook.SheetNames[0]; // Obtener la primera hoja
          const worksheet = workbook.Sheets[sheetName];
          const resultados = [];

          // Comenzar la lectura desde la fila 3
          let rowNumber = 3;
          while (true) {
            const cellAddress = `A${rowNumber}`; // Definir la dirección de la celda
            const cell = worksheet[cellAddress];
            const cellAddressR = `R${rowNumber}`; // Dirección de la celda en la columna 'R'
            const cellR = worksheet[cellAddressR]; // Leer la celda de la columna 'R'
            
            if (!cell || !cell.v) break; // Si la celda está vacía, detener la lectura
            
            // Obtener el valor de la celda en la columna 'R'
            const columnEValue =cellR ? cellR.v : null;

            resultados.push({
              columnE: columnEValue,
              cellAddress: cellAddress.replace('A', 'R') // Guardar la dirección de la celda
            });

            rowNumber++;
          }
         
          try { 
            
            const resul = await sql.sp_existeboxid(resultados);
           
            if (!resul || !Array.isArray(resul)) {
                console.error('La respuesta del procedimiento almacenado no es un array.');
                return res.status(500).json({ error: 'Error en la respuesta del procedimiento almacenado' });
            }
            const resulMap = resul.reduce((map, r) => {
              map[r.Box_ID.trim().toUpperCase()] = { HB201: r.HB201, NO_PALLET: r.NO_PALLET,CUARENTENA: r.CUARENTENA }; // Normalizar Box_ID y almacenar HB201 y NO_PALLET
              return map;
          }, {});
          /****************************************** */
          const resulhb103 = await sql.sp_existeboxid103(resultados);
          
            
            const resulMap103 = resulhb103.reduce((map, r) => {
              // Solo almacenar el Box_ID en el mapa, usando null como valor o podrías poner true si prefieres
              map[r.Box_ID.trim().toUpperCase()] = null;
              return map;
          }, {});
          
          /***************************************** */


          if (!worksheet['!ref']) {
            worksheet['!ref'] = 'A1:Z100';  // Ajusta según el tamaño de la hoja
          }
          // Obtener el rango actual de la hoja
          const range = xlsx.utils.decode_range(worksheet['!ref']);
          
          // Ajustar el rango si es necesario para que incluya la columna Y y la fila 2
          range.e.c = Math.max(range.e.c, 26); 
          range.e.r = Math.max(range.e.r, 1);  // Fila 2
          
          // Actualizar el rango de la hoja
          worksheet['!ref'] = xlsx.utils.encode_range(range);
          
          // Asignar valores a las celdas Y2 y Z2
          worksheet['Y2'] = { v: 'HB201', t: 's' };  // Nombre de la columna Y
          worksheet['Z2'] = { v: 'NO_PALLET', t: 's' };  // Nombre de la columna Z
          worksheet['AA2'] = { v: 'CUARENTENA', t: 's' };  // Nombre de la columna Z
          
          
          xlsx.writeFile(workbook, rutaArchivo);
          resultados.forEach(r => {
            const cleanedColumnE = String(r.columnE).trim().toUpperCase();  // Normalizar datos
            const cell = worksheet[r.cellAddress]; // Obtener la celda correspondiente
        
            if (cell) {
                const resulData = resulMap103[cleanedColumnE]; // Verificar si el Box_ID existe en el mapa
                const rowNumber = r.cellAddress.match(/\d+/)[0]; // Obtener el número de fila
        
                if (resulData !== undefined) {
                    // Si el Box_ID existe, pintar de verde
                    cell.s = {
                        fill: {
                            fgColor: { rgb: '00FF00' } // Verde
                        }
                    };
        
                   
        
                } else {
                    // Si el Box_ID no existe, pintar de rojo
                    cell.s = {
                        fill: {
                            fgColor: { rgb: 'FF0000' } // Rojo
                        }
                    };
                }
        
            } else {
                //console.warn(`Celda ${r.cellAddress} no encontrada. Valor: ${r.columnE}`);
            }
        });        
          resultados.forEach(r => {
            const cleanedColumnE = String(r.columnE).trim().toUpperCase();  // Normalizar datos
            const cell = worksheet[r.cellAddress]; // Obtener la celda correspondiente
        
            if (cell) {
                const resulData = resulMap[cleanedColumnE]; // Obtener el valor de HB201 para el Box_ID correspondiente
                const rowNumber = r.cellAddress.match(/\d+/)[0];
                const targetCell = worksheet[`Y${rowNumber}`];
                if (resulData !== undefined) {
                    const { HB201, NO_PALLET, CUARENTENA } = resulData; // Obtener HB201 y NO_PALLET
                    if(CUARENTENA){
                      worksheet[`AA${rowNumber}`] = { v: 1 };
                    }else{
                      worksheet[`AA${rowNumber}`] = { v: 0 };
                    }
                    
                    
                    // Si HB201 es 1, pintar verde; si es 0, pintar rojo
                    if (HB201) {
                      worksheet[`Y${rowNumber}`] = { v: 1 || 0 }; 
                      worksheet[`Z${rowNumber}`] = { v: NO_PALLET || '' };
                      
                        
                        
                        //console.log(worksheet)
                        //console.log(`Valores asignados a fila ${rowNumber}: HB201 = ${HB201}, NO_PALLET = ${NO_PALLET}`);
                       
                    } else {
                      worksheet[`Y${rowNumber}`] = { v: 0 }; 
                      worksheet[`Z${rowNumber}`] = { v: '' };
                     
                        
                    }
                    
        
                } else {
                  worksheet[`Y${rowNumber}`] = { v: '' }; 
                  worksheet[`Z${rowNumber}`] = { v: '' };
                  
                };
                    //console.warn(`No se encontró HB201 para el valor: ${cleanedColumnE}`);
                }
            
          });
          resultados.forEach(r => {
            const cleanedColumnE = String(r.columnE).trim().toUpperCase();  // Normalizar datos
            const cell = worksheet[r.cellAddress]; // Obtener la celda correspondiente
        
            if (cell) {
                const resulData = resulMap[cleanedColumnE]; // Obtener el valor de HB201 para el Box_ID correspondiente
                const rowNumber = r.cellAddress.match(/\d+/)[0];
                const targetCell = worksheet[`Y${rowNumber}`];
                if (resulData !== undefined) {
                    const { HB201, NO_PALLET, CUARENTENA } = resulData; // Obtener HB201 y NO_PALLET
                    if(CUARENTENA){
                      worksheet[`AA${rowNumber}`] = { v: 1 };
                    }else{
                      worksheet[`AA${rowNumber}`] = { v: 0 };
                    }
                    
                    
                    // Si HB201 es 1, pintar verde; si es 0, pintar rojo
                    if (HB201) {
                     
                      targetCell.s = {
                            fill: {
                                fgColor: { rgb: '00FF00' } // Verde
                            }
                        };  
                        
                        
                        
                        //console.log(worksheet)
                        //console.log(`Valores asignados a fila ${rowNumber}: HB201 = ${HB201}, NO_PALLET = ${NO_PALLET}`);
                       
                    } else {
                      if(CUARENTENA){
                        targetCell.s = {
                          fill: {
                              fgColor: { rgb: 'FFFF00' } // Rojo
                          }
                      };

                      }else{
                        targetCell.s = {
                          fill: {
                              fgColor: { rgb: 'FF0000' } // Rojo
                          }
                      };

                      }
                      
                        
                    }
                    
        
                } else {
                 
                  targetCell.s = {
                    fill: {
                        fgColor: { rgb: 'FFFF00' } //Amarillo
                    }
                };
                    //console.warn(`No se encontró HB201 para el valor: ${cleanedColumnE}`);
                }
            } else {
                console.warn(`Celda ${r.cellAddress} no encontrada. Valor: ${r.columnE}`);
            }
          }); 
          const newWorkbook = new ExcelJS.Workbook(); // Crear un nuevo workbook
            const newSheet = newWorkbook.addWorksheet('Box ID que no existen'); // Agregar nueva hoja

            // Agregar cabeceras a la nueva hoja
            newSheet.addRow(['Box_ID', 'PO_LCN', 'POD_LCN', 'HB201', 'NO_PALLET', 'CUARENTENA']);

            // Obtener los resultados de los Box_ID que no existen
            const noExistResults = await sql.sp_noexisteboxid(resultados);
              if (noExistResults && noExistResults.length > 0) {
              noExistResults.forEach(row => {
                newSheet.addRow([row.Box_ID, row.PO_LCN, row.POD_LCN, row.HB201, row.NO_PALLET, row.CUARENTENA]);
              });
            } else {
              console.log("noExistResults está vacío, no se agregarán filas.");
            }

            // Guardar la nueva hoja como archivo Excel
            const tempFilePath = path.join('BOX_ID_No_Existen'+archivo);
            await newWorkbook.xlsx.writeFile(tempFilePath);
            console.log(`Nueva hoja guardada en: ${tempFilePath}`);

            // Guardar el archivo modificado
            await xlsx.writeFile(workbook, rutaArchivo);
            console.log(`Archivo modificado guardado: ${rutaArchivo}`);
            await enviarMail(['sistemas@zayro.com'], archivo, rutaArchivo,carpeta,tempFilePath ); // Pasar los parámetros necesarios
            return res.status(200).json({ message: 'Proceso de correo completado exitosamente' });
          } catch (e) {
            console.error(`Error al llamar al procedimiento almacenado: ${e.message}`);
          }
        }
      }
    }
    });
    

  });
});
const enviarMail = async (correos, nombreArchivo, pathExcel, carpeta, tempFilePath ) => {
  const config = {
      host: process.env.hostemail,
      port: process.env.portemail,
      secure: true,
      auth: {
          user: process.env.useremail,
          pass: process.env.passemail
      },
      tls: {
          rejectUnauthorized: false,
      },
  };

  const zipFilePath = pathExcel.replace(/\.xlsx$/, '.zip'); // Cambia la extensión a .zip

  // Comprimir ambos archivos antes de enviarlos
  await compressFile([pathExcel, tempFilePath], zipFilePath);

  const mensaje = {
      from: 'sistemas@zayro.com',
      to: correos.join(','), // Convertir la lista de correos en una cadena
      subject: 'REPORT COMPARE KMC-ZAYRO',
      attachments: [
          {
              filename: nombreArchivo.replace(/\.xlsx$/, '.zip'), // Usar el nombre del archivo proporcionado
              path: zipFilePath, // Usar la ruta del archivo Excel comprimido
          },
      ],
      text: 'REPORT COMPARE KMC-ZAYRO',
  };

  const transport = nodemailer.createTransport(config);
  transport.verify().then(() => console.log("Configuración de correo verificada.")).catch((error) => console.log(error));
  
  transport.sendMail(mensaje, (error, info) => {
      if (error) {
          console.error('Error al enviar el correo:', error);
      } else {
          console.log('Correo enviado:', info.response);
          const comandoEliminarArchivos = `python eliminar_archivos.py ${carpeta}`;
          exec(comandoEliminarArchivos, (error, stdout, stderr) => {
            if (error) {
              console.error(`Error al ejecutar el script de eliminar_archivos: ${error}`);
            }
            if (stderr) {
              console.error(`Error en la salida estándar de eliminar_archivos: ${stderr}`);
            }
            console.log(`Salida del script de eliminar_archivos: ${stdout}`);
          });
      }
      
      // Cierra el transporte después de enviar el correo
      transport.close();
  });
};

function compressFile(inputFilePaths, outputZipPath) {
  return new Promise((resolve, reject) => {
    const output = fs.createWriteStream(outputZipPath);
    const archive = archiver('zip', {
      zlib: { level: 9 } // Nivel de compresión
    });

    output.on('close', () => {
      console.log(`Archivo comprimido: ${outputZipPath} (${archive.pointer()} bytes)`);
      resolve();
    });

    archive.on('error', (err) => {
      reject(err);
    });

    archive.pipe(output);

    // Agregar los archivos al zip
    inputFilePaths.forEach(filePath => {
      archive.file(filePath, { name: path.basename(filePath) });
    });

    archive.finalize();
  });
}

///---------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------
//---------------------------------------------------------------------------------------------------
app.get('/api/descargarfacturaskmx', async (req, res) => {
  const comandoLeerMail = 'python leermailfacturaskmx.py';
  let bandera=false;
  exec(comandoLeerMail, (error, stdout, stderr) => {
    if (error) {
        console.error(`Error al ejecutar el script de leermail: ${error.message}`);
      return res.status(500).json({ error: 'Error al ejecutar el script de leermail' });
    }
    if (stderr) {
      console.error(`Error en la salida estándar de leermail: ${stderr}`);
      return res.status(500).json({ error: 'Error en la salida estándar de leermail' });
    }

    //const carpeta = './FacturasKMX';
    const carpeta = path.resolve(__dirname, '..', 'FacturasKMX'); // raíz del proyecto
    console.log(`Carpeta: ${carpeta}`);
   if (!fs.existsSync(carpeta)) fs.mkdirSync(carpeta, { recursive: true });
    fs.readdir(carpeta, async (error, archivos) => {
      if (error) {
        console.error(`Error al leer la carpeta: ${error}`);
        return res.status(500).json({ error: 'Error al leer la carpeta' });
      }

      if (archivos.length === 0) {
        return res.status(400).json({ error: 'No se encontraron archivos en la carpeta' });
      }

      // Procesar cada archivo Excel
      for (const archivo of archivos) {
        const rutaArchivo = path.join(carpeta, archivo);

        // Solo procesar archivos Excel
        if (rutaArchivo.endsWith('.xlsx')) {
          console.log(`Procesando archivo: ${archivo}`);
          
          const a=await sqlram.altafactura(archivo)
          const resultadofactura= await sqlram.consultarfactura(archivo)
          
          //console.log(resultadofactura)
          if (resultadofactura[0].enviado==1)
          {
            // Leer el archivo Excel
            try {
            
              const workbook = new ExcelJS.Workbook();
              workbook.calcProperties.fullCalcOnLoad = true;

              //await workbook.xlsx.load(rutaArchivo);
              await workbook.xlsx.readFile(rutaArchivo);
              const resultados = [];
              let valorN4Asignado = false;
          
              
              // Procesar las hojas EXP_PAC_D
              workbook.eachSheet(sheet => {
                if (sheet.name.startsWith('EXP_PAC_D')) {
                  if (!valorN4Asignado) {
                    valorN4Asignado = true;
                    const valorN4 = sheet.getCell('N4').value;
                    resultados.push({ hoja: sheet.name, valorN4 });
                  }
          
                  // Iniciar la lectura desde la fila 8
                  for (let rowNumber = 8; rowNumber <= sheet.rowCount; rowNumber++) {
                    const row = sheet.getRow(rowNumber);
                    const columnAValue = row.getCell('A').value;
                    if (columnAValue === null) break;
          
                    resultados.push({
                      hoja: sheet.name,
                      columnA: columnAValue,
                      columnGW: row.getCell('AT').value,
                    });
                  }
                }
              });
              
              
              // Eliminar todas las hojas que no empiezan con 'IMP_ATT'
              workbook.eachSheet(sheet => {
                if (!sheet.name.startsWith('IMP_ATT')) {
                  workbook.removeWorksheet(sheet.id);
                }
              });
          
              //console.log(resultados)
          
              const primerResultado = resultados[0];
              const valorN4PrimerResultado = primerResultado.valorN4;
              //res.json(resultados)
              //console.log(resultados)
              //console.log(valorN4PrimerResultado +'usuario '+user+'resultados '+resultados)
              const respuesta = await sql.facturakmx(valorN4PrimerResultado,'sitemas@zayro.com',resultados);
              //console.log(respuesta)
              //const ws  = workbook.addWorksheet(valorN4PrimerResultado);
              
              //const wb = new ExcelJS.Workbook();
              //const ws = workbook.addWorksheet(valorN4PrimerResultado);
              // Agregar una nueva hoja en blanco al inicio del libro
              const ws  = workbook.addWorksheet(valorN4PrimerResultado);
              // Agregar valor en la celda B4
              ws.getCell('B4').value = 'INVOICE NO. '+valorN4PrimerResultado;
          
              // Establecer los títulos de las columnas
              const columnTitles = [
                'LOG BOX ID', '     PART     ','MASTER KMS', 'QTY', 'UNIT OF', 'G/W', '        DATE IN        ', '      DATE OUT      ',
                'DAYS', 'WEEKS', 'STORAGE', 'STORAGE', 'STORAGE',
                'IN/OUT', 'IN/OUT', 'IN/OUT'
              ];
              ws.getRow(6).values = columnTitles;
              ws.getRow(6).font = {
                name: 'Arial',
                color: '#000000',
                size: 10,
                bold: true,
              };
              const columnTitles2 = [
                '', '', '', '', '', '', '',
                '', '','', 'PALLETIZED', 'LOOSE', 'DOUBLE SKIDS',
                'PALLETIZED', 'LOOSE', 'DOUBLE SKIDS'
              ];
              ws.getRow(7).values = columnTitles2;
              ws.getRow(7).font = {
                name: 'Arial',
                color: '#000000',
                size: 10,
                bold: true,
              };
          
              // Centrar el contenido de las columnas y autoajustar el ancho de las columnas
              for (let i = 1; i <= columnTitles.length; i++) {
                ws.getColumn(i).alignment = { horizontal: 'center' };
              }
          
              ws.columns.forEach(column => {
                let maxLength = 0;
                column.eachCell({ includeEmpty: true }, cell => {
                  const columnLength = cell.value ? cell.value.toString().length : 10;
                  if (columnLength > maxLength) {
                    maxLength = columnLength;
                  }
                });
                column.width = maxLength < 10 ? 10 : maxLength + 2; // Establecer un ancho mínimo de columna
              });
              
              
              // Llenar los datos de respuesta en el archivo Excel
              let numfila = 8;
              const dollarFormat = '$#,##0.00'; // Formato de moneda
              respuesta.forEach(renglon => {
                      // Si LOG_BOX_ID no es 'ZZZZZZZ', agregamos los datos a las celdas correspondientes
                if (renglon.LOG_BOX_ID === 'ZZZZZZZ') {
                  ws.getCell(numfila, 2).value = 'Total de Master KMS';
                  ws.getCell(numfila, 3).value = renglon.MasterKMS ;



                  ws.getCell(numfila, 10).value = renglon.PART || '';
                  
                  ws.getCell(numfila, 11).value = { formula: 'SUM(' + 'K8:' + 'K' + (numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 11).numFmt = dollarFormat;
                  ws.getCell(numfila, 12).value = { formula: 'SUM(' + 'L' + '8:' + 'L' + (numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 12).numFmt = dollarFormat;
                  ws.getCell(numfila, 13).value = { formula: 'SUM(' + 'M' + '8:' + 'M' + (numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 13).numFmt = dollarFormat;
                  ws.getCell(numfila, 14).value = { formula: 'SUM(' + 'N' + '8:' + 'N' + (numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 14).numFmt = dollarFormat;
                  ws.getCell(numfila, 15).value = { formula: 'SUM(' + 'O' + '8:' + 'O' +(numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 15).numFmt = dollarFormat;
                  ws.getCell(numfila, 16).value = { formula: 'SUM(' + 'P' + '8:' + 'P' + (numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 16).numFmt = dollarFormat;
                  ws.getCell(numfila, 17).value = { formula: 'SUM(' + 'Q' + '8:' + 'Q' + (numfila - 1) + ')', result: 0 };
                  ws.getCell(numfila, 17).numFmt = dollarFormat;
                  workbook.eachSheet(sheet => {
                    if (sheet.name.endsWith("(PAGE1)")) {
                      const valorA3 = renglon.ADITIONAL !== 0 ? renglon.ADITIONAL / 4 : 0;
                      sheet.getCell('A3').value = valorA3;
                      sheet.getCell('A4').value = renglon.ADITIONAL !== 0 ? renglon.ADITIONAL : '-';

                    }
                  });
                
                }
                else{
                  ws.getCell(numfila, 1).value = renglon.LOG_BOX_ID || '';
                  ws.getCell(numfila, 2).value = renglon.PART || '';
                  ws.getCell(numfila, 3).value = renglon.MasterKMS || '';
                  ws.getCell(numfila, 4).value = renglon.QTY || 0;
                  ws.getCell(numfila, 5).value = renglon.UNIT_OF || ''
                  resultados.forEach(res => {
                    if (renglon.LOG_BOX_ID === res.columnA) {
                        const columnGWValue = res.columnGW || 0; // Obtener el valor de la columna GW
                        const formattedValue = parseFloat(columnGWValue).toFixed(2); // Formatear el valor con 2 decimales
                        const valueWithUnit = `${formattedValue} KGS`; // Concatenar "KGS" al valor formateado
                        ws.getCell(numfila, 6).value = valueWithUnit; // Asignar el valor formateado con la unidad a la celda
                    }
                  });
                  ws.getCell(numfila, 7).value = renglon.DATE_IN;
                  ws.getCell(numfila, 8).value = renglon.DATE_OUT;
                  ws.getCell(numfila, 9).value = renglon.DAYS || 0;
                  ws.getCell(numfila, 10).value = renglon.WEEKS || 0;
          
                  ws.getCell(numfila, 11).value = renglon.SP !== 0 ? renglon.SP : '-';
                  ws.getCell(numfila, 11).numFmt = dollarFormat;
          
                  ws.getCell(numfila, 12).value = renglon.SL !== 0 ? renglon.SL : '-';
                  ws.getCell(numfila, 12).numFmt = dollarFormat;
          
                  ws.getCell(numfila, 13).value = renglon.SD !== 0 ? renglon.SD : '-';
                  ws.getCell(numfila, 13).numFmt = dollarFormat;
          
                  ws.getCell(numfila, 14).value = renglon.IOP !== 0 ? renglon.IOP : '-';
                  ws.getCell(numfila, 14).numFmt = dollarFormat;
          
                  ws.getCell(numfila, 15).value = renglon.IOL !== 0 ? renglon.IOL : '-';
                  ws.getCell(numfila, 15).numFmt = dollarFormat;
          
                  ws.getCell(numfila, 16).value = renglon.IOD !== 0 ? renglon.IOD : '-';
                  ws.getCell(numfila, 16).numFmt = dollarFormat;
          
                 /* ws.getCell(numfila, 16).value = renglon.ADITIONAL !== 0 ? renglon.ADITIONAL : '-';
                  ws.getCell(numfila, 16).numFmt = dollarFormat;*/
          
                }
                
                  numfila++;
                  
              });  
              /********************************************************************* */
              const impAttData = {};
              let isFirstSheet = true; // Variable para controlar si es la primera hoja o no
            
              // Leer todas las hojas que empiecen con "IMP_ATT"
              workbook.eachSheet(sheet => {
                if (sheet.name.startsWith('IMP_ATT')) {
                  const sheetData = [];
                  let headers = [];
          
                  // Leer cada fila desde la fila 7
                  for (let rowNumber = 7; rowNumber <= sheet.rowCount; rowNumber++) {
                    // Saltar la fila 8 si es la primera hoja
                    if (isFirstSheet && rowNumber === 8) {
                      continue;
                    }
          
                    const row = sheet.getRow(rowNumber);
                    const rowData = {};
          
                    // Leer cada celda en la fila
                    row.eachCell((cell, colNumber) => {
                      const columnName = headers[colNumber - 1] ? headers[colNumber - 1].replace(/\s/g, '') : ''; // Verificar si el encabezado existe antes de reemplazar los espacios
                      const cellValue = cell.value;
          
                      // Guardar los encabezados si estamos en la fila 7
                      if (rowNumber === 7) {
                        headers.push(cellValue);
                      } else {
                        // Guardar el valor de la celda bajo el nombre de la columna correspondiente
                        rowData[columnName] = cellValue;
                      }
                    });
          
                    // Solo agregar datos si la fila no está vacía y tiene al menos una celda con valor
                    if (Object.keys(rowData).length > 0) {
                      sheetData.push(rowData);
                    }
                  }
          
                  // Agregar los datos de la hoja actual al objeto impAttData
                  impAttData[sheet.name] = sheetData;
          
                  // Después de leer la primera hoja, cambia isFirstSheet a falso para que no se salte la fila 8 en las siguientes hojas
                  isFirstSheet = false;
                }
              });
              //res.json(impAttData);
              await sqlSISTEMAS.guardafacturakmx(valorN4PrimerResultado,impAttData)
              
          
              /********************************************************************* */
              //const excelFilePath = path.join(__dirname, `${valorN4PrimerResultado}_Invoice.xlsx`);
              const excelFilePath = path.join(carpeta, `${valorN4PrimerResultado}_Invoice.xlsx`);
               await workbook.xlsx.writeFile(excelFilePath);
               await fs.promises.access(excelFilePath);         // confirma que existe
               await sqlram.actualizarfactura(archivo);
               bandera = true;
               console.log('Archivo guardado correctamente en:', excelFilePath);
          
          
          
            } catch (error) {
              console.error(error);
              //res.redirect('./cargafacturaskmx?error=true');
            }
          }
          else {
            // Si se va por el else, eliminamos el archivo
            fs.unlink(rutaArchivo, (err) => {
              if (err) {
                console.error(`Error al eliminar el archivo ${archivo}:`, err);
              } else {
                console.log(`Archivo ${archivo} eliminado correctamente.`);
              }
            });
          }


         
        }  
      }
      //console.log(bandera)
      if (bandera==true){
        await enviarMailFacturasKMX('Sistemas@zayro,com',carpeta)

        return res.status(200).json({ message: 'Proceso de correo completado exitosamente' });
      }
      else{
        return res.status(200).json({ message: 'No hay archivos pendientes de enviar' });
      }
      
    });
    

  });
});
const enviarMailFacturasKMX = async (correos, carpeta) => {
  const port = parseInt(process.env.portemail, 10) || 465;
  const transport = nodemailer.createTransport({
    host: process.env.hostemail,
    port,
    secure: port === 465, // 465 = SSL; otros puertos = STARTTLS
    auth: {
      user: process.env.useremail,
      pass: process.env.passemail
    },
    tls: { rejectUnauthorized: false },
  });

  try {
    // Asegura que la carpeta exista
    const baseDir = path.resolve(carpeta || path.resolve(__dirname, '..', 'FacturasKMX'));
    await fs.promises.mkdir(baseDir, { recursive: true });

    // Verifica que haya archivos .xlsx para enviar
    const files = await fs.promises.readdir(baseDir);
    const xlsxFiles = files.filter(f => f.toLowerCase().endsWith('.xlsx'));
    if (xlsxFiles.length === 0) {
      console.warn('[KMX] No hay archivos .xlsx en:', baseDir);
      return { ok: false, message: 'No hay facturas para enviar' };
    }

    // Fecha y ruta del ZIP (fuera de la carpeta comprimida)
    const fecha = new Date().toISOString().split('T')[0]; // YYYY-MM-DD
    const zipFilePath = path.join(path.dirname(baseDir), `facturas_${fecha}.zip`);

    // Comprimir la carpeta baseDir -> zipFilePath
    await compressFolder(baseDir, zipFilePath);


    

    const mensaje = {
      from: 'sistemas@zayro.com',
      to: 'sistemas@zayro.com',
      subject: `FACTURAS KMX_${fecha}`,
      text: 'FACTURAS KMX GENERADAS AUTOMÁTICAMENTE',
      attachments: [
        { filename: path.basename(zipFilePath), path: zipFilePath }
      ],
    };

    await transport.verify();
    const info = await transport.sendMail(mensaje);
    console.log('Correo enviado:', info.messageId);

    // Limpieza: borra contenido de la carpeta y luego el ZIP
    await limpiarCarpeta(baseDir);
    await fs.promises.unlink(zipFilePath);

    return { ok: true, message: 'Correo enviado y limpieza realizada' };
  } catch (error) {
    console.error('Error al enviar facturas KMX:', error);
    throw error;
  } finally {
    transport.close();
  }
};


function compressFolder(inputFolderPath, outputZipPath) {
    return new Promise((resolve, reject) => {
        const output = fs.createWriteStream(outputZipPath);
        const archive = archiver('zip', {
            zlib: { level: 9 } // Nivel de compresión
        });

        output.on('close', () => {
            console.log(`Archivo comprimido: ${outputZipPath} (${archive.pointer()} bytes)`);
            resolve();
        });

        archive.on('error', (err) => {
            reject(err);
        });

        archive.pipe(output);
        // Agregar todos los archivos de la carpeta
        archive.directory(inputFolderPath, false); // false para no incluir la carpeta en el zip
        archive.finalize();
    });
}
async function limpiarCarpeta(carpeta) {
    const archivos = await fs.promises.readdir(carpeta);
    const eliminarArchivos = archivos.map(archivo => fs.promises.unlink(path.join(carpeta, archivo)));
    await Promise.all(eliminarArchivos);
    console.log(`Carpeta ${carpeta} limpiada.`);
}
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
app.get('/api/generar-token', async (req, res) => {
  const { usuario, password } = req.query;

  if (!usuario || !password) {
    return res.status(400).send('Faltan credenciales');
  }
  try {
    const resultado=await sqlram.sp_loginaceesostoken(usuario,password);

    if(resultado[0].Resultado==1){
        const resultadotoken=await sqlram.sp_altaToken(usuario);
        const token =resultadotoken[0].token;
        res.json({token});
    }else{  
        res.status(400).send('Usuario o contraseña incorrectos');

    }
  } catch (err) {
    console.error('Error al generar token:', err);
    res.status(500).send('Error interno');
  }
  
});
/**************************************************************************************/
/**************************************************************************************/
/**************************************************************************************/
// Clave secreta (guárdala en .env)

const JWT_SECRET = process.env.JWT_SECRET || 'Zayroserver2025##';
const JWT_ISSUER = process.env.JWT_ISSUER || 'zayrocom';
const JWT_AUDIENCE = process.env.JWT_AUDIENCE || 'apizayrocom';
// Rate limit específico para login (anti fuerza bruta)
const loginLimiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 30,
  standardHeaders: true,
  legacyHeaders: false
});
// Schema de validación del body
const LoginSchema = z.object({
  usuario: z.string().min(1, 'usuario requerido'),
  password: z.string().min(1, 'password requerida')
});
//Se quita el api de login en la ruta para que no se pueda acceder desde fuera
app.post('/auth/login', loginLimiter, async (req, res) => {
  try {
    // ✅ POST: usa body, no query
    const parsed = LoginSchema.safeParse(req.body);
    if (!parsed.success) {
      return res.status(400).json({ error: 'validation', details: parsed.error.issues });
    }
    const { usuario, password } = parsed.data;
    // Llama a tu SP (ajusta al formato real de retorno)
    const resultado = await sqlram.sp_loginaceesostoken(usuario, password);

    // Normaliza el resultado (ejemplos de posibles formatos)
    const row = Array.isArray(resultado) ? resultado[0] :
                (resultado?.recordset?.[0] || resultado?.[0]);
    if (!row) return res.status(401).json({ error: 'Credenciales inválidas' });

    // Asume que el SP regresa { Resultado: 1, Role: 'admin', Scope: 'read:sica write:sica', UserId: '123' }
    if (Number(row.Resultado) !== 1) {
      return res.status(401).json({ error: 'Credenciales inválidas' });
    }

    const role = row.Role || 'user';
    const scope = row.Scope || '';        // espacio separado: "read:sica write:sica"
    const sub   = row.UserId || usuario;  // subject del token

    // Crea JWT corto
    const token = jwt.sign(
      { sub, role, scope },
      JWT_SECRET,
      {
        expiresIn: '15m',
        issuer: JWT_ISSUER,
        audience: JWT_AUDIENCE,
        algorithm: 'HS256'
      }
    );

    return res.json({
      access_token: token,
      token_type: 'Bearer',
      expires_in: 900
    });
  } catch (err) {
    console.error('Error en /auth/login:', err);
    return res.status(500).json({ error: 'internal_error' });
  }
});
// ====================== Error handler (AL FINAL) =======================
app.use(errorHandler);

// ====================== Arranque del servidor ==========================
server.listen(PORT, () => {
  console.log(`Servidor escuchando en http://localhost:${PORT}`);
});