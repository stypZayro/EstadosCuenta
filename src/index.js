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
const PORT = process.env.PORT || 3015;
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
      //to: 'cobranza@zayro.com;sistemas@zayro.com;'+correos,
      to: 'programacion@zayro.com;',
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
      //to: 'cobranza@zayro.com;sistemas@zayro.com;'+correos,
      to: 'programacion@zayro.com;',
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