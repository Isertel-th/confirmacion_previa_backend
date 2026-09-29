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
// CONFIGURACIÓN DE GOOGLE DRIVE (RENDER / LOCAL)
// -------------------------------------------------------------
const authConfig = process.env.GOOGLE_CREDENTIALS
  ? { credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS), scopes: ['https://www.googleapis.com/auth/drive'] }
  : { keyFile: path.join(__dirname, 'credentials.json'), scopes: ['https://www.googleapis.com/auth/drive'] };

const auth = new google.auth.GoogleAuth(authConfig);
const drive = google.drive({ version: 'v3', auth });

// ID Real de tu archivo Excel en Google Drive
const FILE_ID_BASE = '1Y2p2S6NXHIfAutb_PsqJAFvvyqHSGz62';

// Archivos locales para usuarios y observaciones
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

// Función para LEER Excel desde Google Drive
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

// Función para GUARDAR / ACTUALIZAR Excel en Google Drive
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
    console.log('✅ Archivo actualizado correctamente en Google Drive');
  } catch (err) {
    console.error('Error al guardar Excel en Drive:', err.message);
  }
}

// -------------------------------------------------------------
// ENDPOINTS Y RUTAS DE LA API
// -------------------------------------------------------------

// Login de usuarios
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

// Obtener datos desde Google Drive
app.get('/api/datos', async (req, res) => {
  try {
    const { perfil, horaInicio } = req.query;
    let datos = await leerExcelDrive(FILE_ID_BASE);

    if (!datos || datos.length === 0) return res.json([]);
    
    datos = datos.map(d => {
      let fechaProg = d['FECHA DE PROGRAMACIÓN'] || d['FECHA DE PROGRAMACION'] || d['FECHA PROG.'] || '';
      if (fechaProg instanceof Date) {
        fechaProg = fechaProg.toLocaleString('es-EC', { timeZone: 'America/Guayaquil' });
      }
      return { ...d, 'FECHA DE PROGRAMACIÓN': fechaProg };
    });

    if (perfil !== 'admin' && horaInicio) {
      const inicio = new Date(horaInicio);
      // Margen ajustado a 30 días (720 horas)
      const corte = new Date(inicio.getTime() - 720 * 60 * 60 * 1000);
      datos = datos.filter(d => {
        const fecha = d['FECHA DE PROGRAMACIÓN'];
        if (!fecha) return true;
        const fechaObj = new Date(fecha);
        return isNaN(fechaObj.getTime()) || fechaObj >= corte;
      });
    }

    res.json(datos);
  } catch (error) {
    console.error('Error en /api/datos:', error);
    res.status(500).json([]);
  }
});

// Editar múltiples registros y guardar en Google Drive
app.put('/api/editar', async (req, res) => {
  try {
    const items = Array.isArray(req.body) ? req.body : [req.body];
    let datos = await leerExcelDrive(FILE_ID_BASE);
    let editados = 0;

    items.forEach(item => {
      const { tarea, operador, observacion, fechaRegistro } = item;
      if (!observacion || observacion.toLowerCase() === 'pendiente') return;

      const indice = datos.findIndex(d => String(d.TAREA).trim() === String(tarea).trim());
      if (indice !== -1) {
        datos[indice].Operador = operador;
        datos[indice].Observacion = observacion;
        datos[indice]['Fecha de Registro'] = fechaRegistro;
        datos[indice].NuevoRegistro = false;
        editados++;
      }
    });

    await guardarExcelDrive(FILE_ID_BASE, datos);
    res.json({ ok: true, mensaje: `Se guardaron ${editados} registros correctamente en Google Drive.` });
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al actualizar: ' + error.message });
  }
});

// Descargar Excel Completo desde Google Drive
app.get('/api/descargar', async (req, res) => {
  try {
    let datos = await leerExcelDrive(FILE_ID_BASE);
    if (!datos || datos.length === 0) {
      return res.status(404).json({ mensaje: 'Aún no existe una base de datos para descargar.' });
    }

    const datosExportar = datos.map(d => ({
      TAREA: d.TAREA || '',
      ORDEN: d.ORDEN || '',
      CIUDAD: d.CIUDAD || '',
      TECNICO: d.TECNICO || d['TÉCNICO'] || '',
      CONTRATO: d.CONTRATO || '',
      CLIENTE: d.CLIENTE || '',
      'FECHA DE PROGRAMACIÓN': d['FECHA DE PROGRAMACIÓN'] || d['FECHA DE PROGRAMACION'] || '',
      Operador: d.Operador || '',
      'Fecha de Registro': d['Fecha de Registro'] || '',
      Observacion: d.Observacion || 'Pendiente'
    }));

    const hoja = xlsx.utils.json_to_sheet(datosExportar);
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

// Vaciar base a cero en Google Drive
app.post('/api/reset-base', async (req, res) => {
  try {
    await guardarExcelDrive(FILE_ID_BASE, []);
    res.json({ ok: true, mensaje: 'Base de datos en Google Drive vaciada por completo.' });
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al vaciar base: ' + error.message });
  }
});

// Cargar / Integrar nuevos Excel en la base de Google Drive
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
    baseActual = baseActual.map(d => ({ ...d, NuevoRegistro: false }));

    let nuevosPendientes = 0;
    let recuperadosGestionados = 0;

    datosNuevos.forEach(nuevo => {
      const norm = {};
      Object.keys(nuevo).forEach(k => norm[k.trim().toUpperCase()] = nuevo[k]);

      const tarea = norm['TAREA'] || nuevo['TAREA'] || nuevo['Tarea'];
      if (!tarea) return;

      const obsSubida = (nuevo['Observacion'] || nuevo['OBSERVACION'] || norm['OBSERVACION'] || '').toString().trim();
      const opSubido = (nuevo['Operador'] || nuevo['OPERADOR'] || norm['OPERADOR'] || '').toString().trim();
      const fechaRegSubida = (nuevo['Fecha de Registro'] || nuevo['FECHA DE REGISTRO'] || norm['FECHA DE REGISTRO'] || '').toString().trim();

      const existe = baseActual.find(b => String(b.TAREA).trim() === String(tarea).trim());

      if (!existe) {
        const estaGestionado = obsSubida !== '' && obsSubida.toLowerCase() !== 'pendiente';

        baseActual.push({
          TAREA: tarea,
          ORDEN: norm['ORDEN'] || '',
          CIUDAD: norm['CIUDAD'] || '',
          TECNICO: norm['TECNICO'] || norm['TÉCNICO'] || '',
          CONTRATO: norm['CONTRATO'] || '',
          CLIENTE: norm['CLIENTE'] || '',
          'FECHA DE PROGRAMACIÓN': norm['FECHA DE PROGRAMACIÓN'] || norm['FECHA DE PROGRAMACION'] || norm['FECHA PROG.'] || '',
          Operador: opSubido,
          Observacion: obsSubida || 'Pendiente',
          'Fecha de Registro': fechaRegSubida,
          NuevoRegistro: !estaGestionado
        });

        if (estaGestionado) recuperadosGestionados++;
        else nuevosPendientes++;
      }
    });

    await guardarExcelDrive(FILE_ID_BASE, baseActual);
    res.json({ 
      ok: true, 
      mensaje: `Archivo procesado en Drive. Nuevos pendientes: ${nuevosPendientes}. Gestionados cargados/recuperados: ${recuperadosGestionados}.` 
    });
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al procesar el archivo: ' + error.message });
  }
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));