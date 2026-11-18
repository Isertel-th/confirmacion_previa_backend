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
    console.log('✅ Archivo sincronizado correctamente en Google Drive');
  } catch (err) {
    console.error('Error al guardar Excel en Drive:', err.message);
    throw err;
  }
}

// Login
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

// Listas de observaciones
app.get('/api/listas', (req, res) => {
  let observaciones = leerExcelLocal(PATH_OBS).map(o => o.OBSERVACION).filter(Boolean);
  observaciones = observaciones.filter(obs => obs.toLowerCase() !== 'pendiente');
  res.json({ observaciones });
});

// Obtener datos de Google Drive con normalización robusta de claves
app.get('/api/datos', async (req, res) => {
  try {
    let datos = await leerExcelDrive(FILE_ID_BASE);
    if (!datos || datos.length === 0) return res.json([]);

    const datosNormalizados = datos.map(d => {
      const norm = {};
      Object.keys(d).forEach(k => {
        norm[k.trim()] = d[k];
        norm[k.trim().toUpperCase()] = d[k];
      });

      return {
        TAREA: norm['TAREA'] || norm['Task'] || norm['TASK'] || '',
        ORDEN: norm['ORDEN'] || norm['Order'] || norm['ORDER'] || '',
        CIUDAD: norm['CIUDAD'] || norm['Ciudad'] || '',
        TECNICO: norm['TECNICO'] || norm['TÉCNICO'] || norm['Tecnico'] || '',
        CONTRATO: norm['CONTRATO'] || norm['Contrato'] || '',
        CLIENTE: norm['CLIENTE'] || norm['Cliente'] || '',
        'FECHA DE PROGRAMACIÓN': norm['FECHA DE PROGRAMACIÓN'] || norm['FECHA DE PROGRAMACION'] || norm['FECHA PROG.'] || norm['FECHA'] || '',
        Operador: norm['OPERADOR'] || norm['Operador'] || '',
        'Fecha de Registro': norm['FECHA DE REGISTRO'] || norm['Fecha de Registro'] || '',
        Observacion: norm['OBSERVACION'] || norm['OBSERVACIÓN'] || norm['Observacion'] || 'Pendiente',
        NuevoRegistro: norm['NUEVOREGISTRO'] || norm['NuevoRegistro'] || false
      };
    });

    res.json(datosNormalizados);
  } catch (error) {
    console.error('Error en /api/datos:', error);
    res.status(500).json([]);
  }
});

// Editar registros
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
        const tActual = String(d.TAREA || d.Task || d.TASK || '').trim();
        if (tActual === tareaBuscada) {
          d.Operador = operador;
          d.Observacion = observacion;
          d['Fecha de Registro'] = fechaRegistro;
          d.NuevoRegistro = false;
          editados++;
        }
      });
    });

    await guardarExcelDrive(FILE_ID_BASE, datos);
    res.json({ ok: true, mensaje: `Se sincronizaron ${editados} registros asociados al ID de Tarea.` });
  } catch (error) {
    res.status(500).json({ ok: false, mensaje: 'Error al actualizar: ' + error.message });
  }
});

// Descargar Excel Completo
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

// Cargar / Integrar nuevo Excel
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
    
    const mapaGestionados = new Map();
    baseActual.forEach(b => {
      const normB = {};
      Object.keys(b).forEach(k => normB[k.trim().toUpperCase()] = b[k]);
      const t = String(normB['TAREA'] || normB['TASK'] || '').trim();
      const obs = String(normB['OBSERVACION'] || normB['OBSERVACIÓN'] || '').trim();
      if (t && obs && obs.toLowerCase() !== 'pendiente' && obs.toLowerCase() !== '-- seleccionar --') {
        mapaGestionados.set(t, {
          Operador: normB['OPERADOR'] || '',
          Observacion: obs,
          'Fecha de Registro': normB['FECHA DE REGISTRO'] || ''
        });
      }
    });

    let nuevaBaseProcesada = [];
    let nuevosPendientes = 0;
    let recuperadosGestionados = 0;

    datosNuevos.forEach(nuevo => {
      const norm = {};
      Object.keys(nuevo).forEach(k => norm[k.trim().toUpperCase()] = nuevo[k]);

      const tarea = norm['TAREA'] || norm['TASK'] || '';
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
      mensaje: `✅ Archivo procesado. Nuevos pendientes: ${nuevosPendientes} | Gestionados preservados: ${recuperadosGestionados}.` 
    });
  } catch (error) {
    res.status(500).json({ ok: false, mensaje: 'Error al procesar el archivo: ' + error.message });
  }
});

app.listen(PORT, () => console.log(`🚀 Servidor Isertel en puerto ${PORT}`));