const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const multer = require('multer');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

const PATH_BASE = path.join(__dirname, 'docs', 'base_datos.xlsx');
const PATH_USERS = path.join(__dirname, 'docs', 'usuarios.xlsx');
const PATH_OBS = path.join(__dirname, 'docs', 'observaciones.xlsx');

const carpetaDocs = path.join(__dirname, 'docs');
if (!fs.existsSync(carpetaDocs)) fs.mkdirSync(carpetaDocs, { recursive: true });

function leerExcel(ruta) {
  if (!fs.existsSync(ruta)) return [];
  try {
    const libro = xlsx.readFile(ruta, { cellDates: true, dateNF: 'yyyy-mm-dd hh:mm:ss' });
    const hoja = libro.Sheets[libro.SheetNames[0]];
    return xlsx.utils.sheet_to_json(hoja, { raw: false, defval: '' });
  } catch (err) {
    console.error('Error al leer Excel:', err);
    return [];
  }
}

function guardarExcel(ruta, datos) {
  const hoja = xlsx.utils.json_to_sheet(datos);
  const libro = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(libro, hoja, 'Hoja1');
  xlsx.writeFile(libro, ruta);
}

// Login
app.post('/api/login', (req, res) => {
  const { usuario, clave } = req.body;
  const usuarios = leerExcel(PATH_USERS);
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

// Listas de observaciones (Se remueve 'Pendiente' del desplegable)
app.get('/api/listas', (req, res) => {
  let observaciones = leerExcel(PATH_OBS).map(o => o.OBSERVACION).filter(Boolean);
  observaciones = observaciones.filter(obs => obs.toLowerCase() !== 'pendiente');
  res.json({ observaciones });
});

// Obtener datos
app.get('/api/datos', (req, res) => {
  const { perfil, horaInicio } = req.query;
  let datos = leerExcel(PATH_BASE);

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
    // Margen ajustado a 48 horas (2 días) para que la información persista más tiempo cargada
    const corte = new Date(inicio.getTime() - 48 * 60 * 60 * 1000);
    datos = datos.filter(d => {
      const fecha = d['FECHA DE PROGRAMACIÓN'];
      if (!fecha) return true;
      const fechaObj = new Date(fecha);
      return isNaN(fechaObj.getTime()) || fechaObj >= corte;
    });
  }

  res.json(datos);
});

// Editar múltiples registros o uno solo
app.put('/api/editar', (req, res) => {
  const items = Array.isArray(req.body) ? req.body : [req.body];
  let datos = leerExcel(PATH_BASE);
  let editados = 0;

  items.forEach(item => {
    const { tarea, operador, observacion, fechaRegistro } = item;
    if (!observacion || observacion.toLowerCase() === 'pendiente') return;

    const indice = datos.findIndex(d => String(d.TAREA).trim() === String(tarea).trim());
    if (indice !== -1) {
      // Si la tarea ya estaba gestionada, no permite sobreescritura si se envia algo invalido
      datos[indice].Operador = operador;
      datos[indice].Observacion = observacion;
      datos[indice]['Fecha de Registro'] = fechaRegistro;
      datos[indice].NuevoRegistro = false;
      editados++;
    }
  });

  guardarExcel(PATH_BASE, datos);
  res.json({ ok: true, mensaje: `Se guardaron ${editados} registros correctamente.` });
});

// Descargar Excel Completo
app.get('/api/descargar', (req, res) => {
  if (!fs.existsSync(PATH_BASE)) {
    return res.status(404).json({ mensaje: 'Aún no existe una base de datos para descargar.' });
  }
  let datos = leerExcel(PATH_BASE);

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
});

// Vaciar base a cero
app.post('/api/reset-base', (req, res) => {
  try {
    if (fs.existsSync(PATH_BASE)) fs.unlinkSync(PATH_BASE);
    res.json({ ok: true, mensaje: 'Base de datos vaciada por completo.' });
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al vaciar base: ' + error.message });
  }
});

// Cargar/Integrar Excel Inteligente
const carga = multer({ storage: multer.memoryStorage() });
app.post('/api/cargar-excel', carga.single('archivo'), (req, res) => {
  const { tipo } = req.body;
  if (tipo !== 'base' && tipo !== 'observaciones') {
    return res.status(400).json({ mensaje: 'Tipo de archivo no válido' });
  }

  try {
    const libro = xlsx.read(req.file.buffer);
    const hoja = libro.Sheets[libro.SheetNames[0]];
    const datosNuevos = xlsx.utils.sheet_to_json(hoja, { defval: '' });

    if (tipo === 'base') {
      let baseActual = fs.existsSync(PATH_BASE) ? leerExcel(PATH_BASE) : [];
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

      guardarExcel(PATH_BASE, baseActual);
      res.json({ 
        ok: true, 
        mensaje: `Archivo procesado. Nuevos pendientes: ${nuevosPendientes}. Gestionados cargados/recuperados: ${recuperadosGestionados}.` 
      });
    } else {
      guardarExcel(PATH_OBS, datosNuevos);
      res.json({ ok: true, mensaje: 'Lista de observaciones actualizada.' });
    }
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al procesar el archivo: ' + error.message });
  }
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));