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

// Rutas de archivos — en Render se guardan en el disco persistente
const PATH_BASE = path.join(__dirname, 'docs', 'base_datos.xlsx');
const PATH_USERS = path.join(__dirname, 'docs', 'usuarios.xlsx');
const PATH_OP = path.join(__dirname, 'docs', 'operadores.xlsx');
const PATH_OBS = path.join(__dirname, 'docs', 'observaciones.xlsx');

// Asegurar que la carpeta docs exista
const carpetaDocs = path.join(__dirname, 'docs');
if (!fs.existsSync(carpetaDocs)) fs.mkdirSync(carpetaDocs, { recursive: true });

function leerExcel(ruta) {
  if (!fs.existsSync(ruta)) return [];
  const libro = xlsx.readFile(ruta);
  const hoja = libro.Sheets[libro.SheetNames[0]];
  return xlsx.utils.sheet_to_json(hoja);
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
    u.USERNAME === usuario && u.CONTRASEÑA === clave
  );
  if (encontrado) {
    res.json({ ok: true, perfil: encontrado.ROL, nombre: encontrado.USERNAME });
  } else {
    res.status(401).json({ ok: false, mensaje: 'Credenciales incorrectas' });
  }
});

// Listas
app.get('/api/listas', (req, res) => {
  const operadores = leerExcel(PATH_OP).map(o => o.OPERADORES).filter(Boolean);
  const observaciones = leerExcel(PATH_OBS).map(o => o.OBSERVACION).filter(Boolean);
  if (!observaciones.includes('Pendiente')) observaciones.unshift('Pendiente');
  res.json({ operadores, observaciones });
});

function asignarColor(indice) {
  const colores = [
    '#e3f2fd', '#c8e6c9', '#fff9c4',
    '#f8bbd9', '#e1bee8', '#d7ccc8'
  ];
  return colores[indice % colores.length];
}

// Obtener datos
app.get('/api/datos', (req, res) => {
  const { perfil, horaInicio } = req.query;
  let datos = leerExcel(PATH_BASE);
  
  datos = datos.map((d, i) => {
    // Lectura flexible para CLIENTE
    const cliente = d.CLIENTE || d.Cliente || d.cliente || d['CLIENTE'] || '';

    // Lectura flexible para FECHA DE PROGRAMACIÓN
    const fechaProg = d['FECHA DE PROGRAMACIÓN'] 
      || d['FECHA DE PROGRAMACION'] 
      || d['Fecha de Programación'] 
      || d['FECHA PROG.'] 
      || d['Fecha Prog.'] 
      || '';

    return {
      TAREA: d.TAREA || d.Tarea || '',
      ORDEN: d.ORDEN || d.Orden || '',
      CIUDAD: d.CIUDAD || d.Ciudad || '',
      TECNICO: d.TECNICO || d.TÉCNICO || d.Técnico || '',
      CONTRATO: d.CONTRATO || d.Contrato || '',
      CLIENTE: cliente,
      'FECHA DE PROGRAMACIÓN': fechaProg,
      Operador: d.OPERADOR || d.Operador || '',
      Observacion: d.OBSERVACION || d.Observacion || 'Pendiente',
      lote: Math.floor(i / 50),
      color: asignarColor(Math.floor(i / 50))
    };
  });

  if (perfil !== 'admin' && horaInicio) {
    const inicio = new Date(horaInicio);
    const corte = new Date(inicio.getTime() - 2 * 60 * 60 * 1000);
    datos = datos.filter(d => {
      const fecha = d['FECHA DE PROGRAMACIÓN'];
      if (!fecha) return true;
      return new Date(fecha) >= corte;
    });
  }

  res.json(datos);
});

// Editar registro
app.put('/api/editar', (req, res) => {
  const { tarea, operador, observacion } = req.body;
  let datos = leerExcel(PATH_BASE);
  
  const indice = datos.findIndex(d => d.TAREA === tarea);
  if (indice === -1) {
    return res.status(404).json({ mensaje: 'Tarea no encontrada' });
  }

  datos[indice].Operador = operador;
  datos[indice].Observacion = observacion;

  guardarExcel(PATH_BASE, datos);
  res.json({ ok: true, mensaje: 'Guardado correctamente' });
});

// Descargar Excel
app.get('/api/descargar', (req, res) => {
  if (!fs.existsSync(PATH_BASE)) {
    return res.status(404).json({ mensaje: 'Archivo no encontrado' });
  }
  res.download(PATH_BASE, 'base_datos_actualizada.xlsx');
});

// Subir y reemplazar Excel — SIN NECESIDAD DE GIT
const carga = multer({ storage: multer.memoryStorage() });
app.post('/api/cargar-excel', carga.single('archivo'), (req, res) => {
  const { tipo } = req.body;
  const rutas = {
    base: PATH_BASE,
    operadores: PATH_OP,
    observaciones: PATH_OBS
  };
  const ruta = rutas[tipo];
  
  if (!ruta) {
    return res.status(400).json({ mensaje: 'Tipo de archivo no válido' });
  }

  try {
    const libro = xlsx.read(req.file.buffer);
    const hoja = libro.Sheets[libro.SheetNames[0]];
    const datos = xlsx.utils.sheet_to_json(hoja);
    
    guardarExcel(ruta, datos);
    
    res.json({ 
      ok: true, 
      mensaje: `Archivo ${tipo} actualizado — ${datos.length} registros cargados` 
    });
  } catch (error) {
    res.status(500).json({ mensaje: 'Error al procesar el archivo: ' + error.message });
  }
});

app.listen(PORT, () => console.log(`Servidor en puerto ${PORT}`));