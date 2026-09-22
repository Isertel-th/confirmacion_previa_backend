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

// Rutas de archivos en el servidor (almacenamiento persistente/local)
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

// Listas de observaciones
app.get('/api/listas', (req, res) => {
  let observaciones = leerExcel(PATH_OBS).map(o => o.OBSERVACION).filter(Boolean);
  if (!observaciones.includes('Pendiente')) observaciones.unshift('Pendiente');
  res.json({ observaciones });
});

// Obtener datos de la base cargada por Admin
app.get('/api/datos', (req, res) => {
  const { perfil, horaInicio } = req.query;
  let datos = leerExcel(PATH_BASE);

  if (!datos || datos.length === 0) {
    return res.json([]);
  }
  
  datos = datos.map(d => {
    let fechaProg = d['FECHA DE PROGRAMACIÓN'] || d['FECHA DE PROGRAMACION'] || d['FECHA PROG.'] || '';
    if (fechaProg instanceof Date) {
      fechaProg = fechaProg.toLocaleString('es-EC', { timeZone: 'America/Guayaquil' });
    }
    return {
      ...d,
      'FECHA DE PROGRAMACIÓN': fechaProg
    };
  });

  if (perfil !== 'admin' && horaInicio) {
    const inicio = new Date(horaInicio);
    const corte = new Date(inicio.getTime() - 2 * 60 * 60 * 1000);
    datos = datos.filter(d => {
      const fecha = d['FECHA DE PROGRAMACIÓN'];
      if (!fecha) return true;
      const fechaObj = new Date(fecha);
      return isNaN(fechaObj.getTime()) || fechaObj >= corte;
    });
  }

  res.json(datos);
});

// Editar registro (Guardar gestión)
app.put('/api/editar', (req, res) => {
  const { tarea, operador, observacion, fechaRegistro } = req.body;
  let datos = leerExcel(PATH_BASE);
  
  const indice = datos.findIndex(d => String(d.TAREA).trim() === String(tarea).trim());
  if (indice === -1) {
    return res.status(404).json({ mensaje: 'Tarea no encontrada' });
  }

  datos[indice].Operador = operador;
  datos[indice].Observacion = observacion;
  datos[indice]['Fecha de Registro'] = fechaRegistro;
  datos[indice].NuevoRegistro = false; // Quita la etiqueta roja al gestionarse

  guardarExcel(PATH_BASE, datos);
  res.json({ ok: true, mensaje: 'Guardado correctamente' });
});

// Descargar Excel
app.get('/api/descargar', (req, res) => {
  if (!fs.existsSync(PATH_BASE)) return res.status(404).json({ mensaje: 'Aún no se ha cargado ninguna base de datos.' });
  res.download(PATH_BASE, 'base_datos_actualizada.xlsx');
});

// Subir Excel Inteligente (Administrador)
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
      
      // Desmarcar etiqueta "nuevo registro" en datos viejos
      baseActual = baseActual.map(d => ({ ...d, NuevoRegistro: false }));
      
      let agregados = 0;

      if (baseActual.length === 0) {
        // Primera carga completa de base por el Admin
        baseActual = datosNuevos.map(nuevo => {
          const norm = {};
          Object.keys(nuevo).forEach(k => norm[k.trim().toUpperCase()] = nuevo[k]);
          agregados++;
          return {
            TAREA: norm['TAREA'] || '',
            ORDEN: norm['ORDEN'] || '',
            CIUDAD: norm['CIUDAD'] || '',
            TECNICO: norm['TECNICO'] || norm['TÉCNICO'] || '',
            CONTRATO: norm['CONTRATO'] || '',
            CLIENTE: norm['CLIENTE'] || '',
            'FECHA DE PROGRAMACIÓN': norm['FECHA DE PROGRAMACIÓN'] || norm['FECHA DE PROGRAMACION'] || norm['FECHA PROG.'] || '',
            Operador: '',
            Observacion: 'Pendiente',
            'Fecha de Registro': '',
            NuevoRegistro: true
          };
        });
      } else {
        // Carga subsecuente: Fusionar evitando duplicados por TAREA
        datosNuevos.forEach(nuevo => {
          const norm = {};
          Object.keys(nuevo).forEach(k => norm[k.trim().toUpperCase()] = nuevo[k]);
          
          const tarea = norm['TAREA'];
          if (!tarea) return;

          const existe = baseActual.find(b => String(b.TAREA).trim() === String(tarea).trim());
          
          if (!existe) {
            baseActual.push({
              TAREA: tarea,
              ORDEN: norm['ORDEN'] || '',
              CIUDAD: norm['CIUDAD'] || '',
              TECNICO: norm['TECNICO'] || norm['TÉCNICO'] || '',
              CONTRATO: norm['CONTRATO'] || '',
              CLIENTE: norm['CLIENTE'] || '',
              'FECHA DE PROGRAMACIÓN': norm['FECHA DE PROGRAMACIÓN'] || norm['FECHA DE PROGRAMACION'] || norm['FECHA PROG.'] || '',
              Operador: '',
              Observacion: 'Pendiente',
              'Fecha de Registro': '',
              NuevoRegistro: true
            });
            agregados++;
          }
        });
      }

      guardarExcel(PATH_BASE, baseActual);
      res.json({ ok: true, mensaje: `Base cargada. Se añadieron ${agregados} tareas nuevas como pendientes.` });
    } else {
      guardarExcel(PATH_OBS, datosNuevos);
      res.json({ ok: true, mensaje: `Lista de observaciones actualizada correctamente.` });
    }
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al procesar el archivo: ' + error.message });
  }
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));