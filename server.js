const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const multer = require('multer');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());
app.use('/docs', express.static(path.join(__dirname, 'docs')));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, './docs/'),
  filename: (req, file, cb) => cb(null, file.originalname)
});
const upload = multer({ storage });

// Rutas de archivos
const PATH_BASE = './docs/base_datos.xlsx';
const PATH_OP = './docs/operadores.xlsx';
const PATH_OBS = './docs/observaciones.xlsx';
const PATH_USERS = './docs/usuarios.xlsx';

// Colores por lote
const COLORES = ['#e3f2fd', '#c8e6c9', '#fff9c4', '#f8bbd9', '#e1bee8', '#d7ccc8'];

// Leer Excel
const leerExcel = (ruta) => {
  if (!fs.existsSync(ruta)) return [];
  const wb = XLSX.readFile(ruta);
  const hoja = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(hoja);
};

// Guardar Excel
const guardarExcel = (ruta, datos) => {
  const hoja = XLSX.utils.json_to_sheet(datos);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, hoja, 'Datos');
  XLSX.writeFile(wb, ruta);
};

// Asignar colores por lote según hora de ingreso
const asignarColores = (registros, horaInicioSistema) => {
  return registros.map((reg, idx) => {
    const horaReg = new Date(reg.fecha_ingreso || horaInicioSistema);
    const horasTranscurridas = (horaReg - horaInicioSistema) / (1000 * 60 * 60);
    const lote = Math.floor(horasTranscurridas / 2);
    return { ...reg, color: COLORES[lote % COLORES.length], lote };
  });
};

// Login
app.post('/api/login', (req, res) => {
  const { usuario, clave } = req.body;
  const usuarios = leerExcel(PATH_USERS);
  const encontrado = usuarios.find(u => 
    u.usuario === usuario && u.clave === clave
  );
  if (encontrado) {
    res.json({ ok: true, perfil: encontrado.perfil, nombre: encontrado.usuario });
  } else {
    res.status(401).json({ ok: false, mensaje: 'Credenciales incorrectas' });
  }
});

// Obtener datos según perfil y tiempo
app.get('/api/datos', (req, res) => {
  const { perfil, horaInicio } = req.query;
  let datos = leerExcel(PATH_BASE);
  const horaInicioSistema = new Date(horaInicio || Date.now());

  // Asignar fecha_ingreso si no existe
  datos = datos.map(d => {
    if (!d.fecha_ingreso) {
      d.fecha_ingreso = horaInicioSistema.toISOString();
    }
    return d;
  });

  // Filtrar: user solo ve últimos 2h; admin ve todo
  if (perfil !== 'admin') {
    const ahora = new Date();
    const limite = new Date(ahora.getTime() - 2 * 60 * 60 * 1000);
    datos = datos.filter(d => new Date(d.fecha_ingreso) >= limite);
  }

  // Sin tareas repetidas
  const vistos = new Set();
  const unicos = datos.filter(d => {
    if (vistos.has(d.tarea)) return false;
    vistos.add(d.tarea);
    return true;
  });

  res.json(asignarColores(unicos, horaInicioSistema));
});

// Listas desplegables
app.get('/api/listas', (req, res) => {
  res.json({
    operadores: leerExcel(PATH_OP).map(o => o.operador),
    observaciones: leerExcel(PATH_OBS).map(o => o.observacion)
  });
});

// Editar registro
app.put('/api/editar', (req, res) => {
  const { tarea, operador, observacion } = req.body;
  const datos = leerExcel(PATH_BASE);
  const idx = datos.findIndex(d => d.tarea === tarea);
  if (idx !== -1) {
    datos[idx].operador = operador;
    datos[idx].observacion = observacion;
    guardarExcel(PATH_BASE, datos);
    res.json({ ok: true });
  } else {
    res.status(404).json({ ok: false });
  }
});

// Subir Excel (solo admin)
app.post('/api/cargar-excel', upload.single('archivo'), (req, res) => {
  res.json({ ok: true, mensaje: 'Archivo cargado correctamente' });
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));