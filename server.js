const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const multer = require('multer');
const path = require('path');
const stream = require('stream');
const { google } = require('googleapis');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

// -------------------------------------------------------------
// CONFIGURACIÓN DE GOOGLE DRIVE Y ARCHIVOS LOCALES
// -------------------------------------------------------------
const authConfig = process.env.GOOGLE_CREDENTIALS
  ? { credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS), scopes: ['https://www.googleapis.com/auth/drive'] }
  : { keyFile: path.join(__dirname, 'credentials.json'), scopes: ['https://www.googleapis.com/auth/drive'] };

const auth = new google.auth.GoogleAuth(authConfig);
const drive = google.drive({ version: 'v3', auth });

const FILE_ID_BASE = '1Y2p2S6NXHIfAutb_PsqJAFvvyqHSGz62';
const PATH_USERS = path.join(__dirname, 'docs', 'usuarios.xlsx');
const PATH_OBS = path.join(__dirname, 'docs', 'observaciones.xlsx');

function leerExcelLocal(ruta) {
  const fs = require('fs');
  if (!fs.existsSync(ruta)) return [];
  try {
    const libro = xlsx.readFile(ruta, { cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss' });
    const hoja = libro.Sheets[libro.SheetNames[0]];
    return xlsx.utils.sheet_to_json(hoja, { raw: false, defval: '' });
  } catch (err) {
    console.error('Error al leer Excel local:', err);
    return [];
  }
}

async function leerExcelDrive(fileId) {
  try {
    const res = await drive.files.get(
      { fileId: fileId, alt: 'media' },
      { responseType: 'arraybuffer' }
    );
    const libro = xlsx.read(res.data, { cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss' });
    const hoja = libro.Sheets[libro.SheetNames[0]];
    return xlsx.utils.sheet_to_json(hoja, { raw: false, defval: '' });
  } catch (err) {
    console.error('Error al leer Excel desde Drive:', err.message);
    return [];
  }
}

async function guardarExcelDrive(fileId, datos) {
  try {
    const hoja = xlsx.utils.json_to_sheet(datos);
    const libro = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(libro, hoja, 'Hoja1');
    const buffer = xlsx.write(libro, { type: 'buffer', bookType: 'xlsx' });

    const bufferStream = new stream.PassThrough();
    bufferStream.end(buffer);

    await drive.files.update({
      fileId: fileId,
      media: {
        mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        body: bufferStream,
      },
    });
    console.log('✅ Archivo sincronizado y guardado correctamente en Google Drive');
  } catch (err) {
    console.error('Error al guardar Excel en Drive:', err.message);
    throw err;
  }
}

// -------------------------------------------------------------
// ENDPOINTS DE LA API
// -------------------------------------------------------------

app.post('/api/login', (req, res) => {
  const { usuario, clave } = req.body;
  const usuarios = leerExcelLocal(PATH_USERS);
  const encontrado = usuarios.find(u => 
    String(u.USERNAME).trim() === String(usuario).trim() && 
    String(u.CONTRASEÑA).trim() === String(clave).trim()
  );
  if (encontrado) {
    res.json({ ok: true, perfil: encontrado.ROL, nombre: encontrado.USERNAME });
  } else {
    res.status(401).json({ ok: false, mensaje: 'Credenciales incorrectas' });
  }
});

app.get('/api/listas', (req, res) => {
  let observaciones = leerExcelLocal(PATH_OBS).map(o => o.OBSERVACION).filter(Boolean);
  observaciones = observaciones.filter(obs => obs.toLowerCase() !== 'pendiente');
  res.json({ observaciones });
});

app.get('/api/datos', async (req, res) => {
  try {
    let datos = await leerExcelDrive(FILE_ID_BASE);
    if (!datos || datos.length === 0) return res.json([]);
    
    datos = datos.map(d => {
      let fechaProg = d['FECHA DE PROGRAMACIÓN'] || d['FECHA DE PROGRAMACION'] || d['FECHA PROG.'] || '';
      return { ...d, 'FECHA DE PROGRAMACIÓN': fechaProg };
    });

    res.json(datos);
  } catch (error) {
    console.error('Error en /api/datos:', error);
    res.status(500).json([]);
  }
});

// Endpoint de Edición: Unifica y actualiza TODAS las filas con el mismo ID de TAREA
app.put('/api/editar', async (req, res) => {
  try {
    const items = Array.isArray(req.body) ? req.body : [req.body];
    let datos = await leerExcelDrive(FILE_ID_BASE);
    let editados = 0;

    items.forEach(item => {
      const { tarea, operador, observacion, fechaRegistro } = item;
      if (!observacion || observacion.toLowerCase() === 'pendiente') return;

      const tareaBuscada = String(tarea).trim();

      datos.forEach(d => {
        if (String(d.TAREA).trim() === tareaBuscada) {
          d.Operador = operador;
          d.Observacion = observacion;
          d['Fecha de Registro'] = fechaRegistro;
          d.NuevoRegistro = false;
          editados++;
        }
      });
    });

    await guardarExcelDrive(FILE_ID_BASE, datos);
    res.json({ ok: true, mensaje: `Se sincronizaron ${editados} registros asociados al ID de Tarea en Google Drive.` });
  } catch (error) {
    res.status(500).json({ ok: false, mensaje: 'Error al actualizar: ' + error.message });
  }
});

app.get('/api/descargar', async (req, res) => {
  try {
    let datos = await leerExcelDrive(FILE_ID_BASE);
    if (!datos || datos.length === 0) {
      return res.status(404).json({ mensaje: 'Aún no existe una base de datos para descargar.' });
    }

    const hoja = xlsx.utils.json_to_sheet(datos);
    const libro = xlsx.utils.book_new();
    xlsx.utils.book_append_sheet(libro, hoja, 'Base_Gestion');

    const buffer = xlsx.write(libro, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=Base_Gestion_Completa.xlsx');
    res.send(buffer);
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al generar la descarga: ' + error.message });
  }
});

const carga = multer({ storage: multer.memoryStorage() });
app.post('/api/cargar-excel', carga.single('archivo'), async (req, res) => {
  const { tipo } = req.body;
  if (tipo !== 'base') {
    return res.status(400).json({ mensaje: 'Tipo de archivo no válido' });
  }

  try {
    const libro = xlsx.read(req.file.buffer, { cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss' });
    const hoja = libro.Sheets[libro.SheetNames[0]];
    const datosNuevos = xlsx.utils.sheet_to_json(hoja, { raw: false, defval: '' });

    let baseActual = await leerExcelDrive(FILE_ID_BASE);
    
    // Mapa rápido para preservar gestiones previas si el ID de Tarea ya estaba gestionado
    const mapaGestionados = new Map();
    baseActual.forEach(b => {
      const t = String(b.TAREA || '').trim();
      const obs = String(b.Observacion || '').trim();
      if (t && obs && obs.toLowerCase() !== 'pendiente' && obs.toLowerCase() !== '-- seleccionar --') {
        mapaGestionados.set(t, {
          Operador: b.Operador,
          Observacion: b.Observacion,
          'Fecha de Registro': b['Fecha de Registro'] || ''
        });
      }
    });

    let nuevaBaseProcesada = [];
    let nuevosPendientes = 0;
    let recuperadosGestionados = 0;

    datosNuevos.forEach(nuevo => {
      const norm = {};
      Object.keys(nuevo).forEach(k => norm[k.trim().toUpperCase()] = nuevo[k]);

      const tarea = norm['TAREA'] || nuevo['TAREA'] || nuevo['Tarea'];
      if (!tarea) return;

      const tareaStr = String(tarea).trim();
      const gestionPrev = mapaGestionados.get(tareaStr);

      if (gestionPrev) {
        nuevaBaseProcesada.push({
          ...nuevo,
          TAREA: tareaStr,
          Operador: gestionPrev.Operador,
          Observacion: gestionPrev.Observacion,
          'Fecha de Registro': gestionPrev['Fecha de Registro'],
          NuevoRegistro: false
        });
        recuperadosGestionados++;
      } else {
        const obsSubida = (norm['OBSERVACION'] || norm['OBSERVACIÓN'] || '').toString().trim();
        const opSubido = (norm['OPERADOR'] || '').toString().trim();
        const fechaRegSubida = (norm['FECHA DE REGISTRO'] || '').toString().trim();
        const estaGestionado = obsSubida !== '' && obsSubida.toLowerCase() !== 'pendiente';

        nuevaBaseProcesada.push({
          ...nuevo,
          TAREA: tareaStr,
          Operador: opSubido,
          Observacion: estaGestionado ? obsSubida : 'Pendiente',
          'Fecha de Registro': fechaRegSubida,
          NuevoRegistro: !estaGestionado
        });

        if (estaGestionado) recuperadosGestionados++;
        else nuevosPendientes++;
      }
    });

    await guardarExcelDrive(FILE_ID_BASE, nuevaBaseProcesada);
    res.json({ 
      ok: true, 
      mensaje: `✅ Archivo procesado con éxito. Nuevos pendientes: ${nuevosPendientes} | Gestionados preservados/cargados: ${recuperadosGestionados}.` 
    });
  } catch (error) {
    res.status(500).json({ ok: false, mensaje: 'Error al procesar el archivo: ' + error.message });
  }
});

app.listen(PORT, () => console.log(`🚀 Servidor Isertel corriendo en puerto ${PORT}`));