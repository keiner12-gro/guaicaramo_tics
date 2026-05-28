require('dotenv').config();
const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

// Configurar multer para archivos Excel
const upload = multer({
    storage: multer.memoryStorage(),
    limits: { fileSize: 50 * 1024 * 1024 }, // 50MB max
    fileFilter: (req, file, cb) => {
        const allowedMimes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];
        
        if (allowedMimes.includes(file.mimetype) || file.originalname.endsWith('.xlsx') || file.originalname.endsWith('.xls')) {
            cb(null, true);
        } else {
            cb(new Error('Solo se permiten archivos Excel (.xls, .xlsx)'));
        }
    }
});

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname)); // To serve index.html from root

// Database Connection
const pool = mysql.createPool({
    host: process.env.MYSQLHOST,
    user: process.env.MYSQLUSER,
    password: process.env.MYSQLPASSWORD,
    database: process.env.MYSQLDATABASE,
    port: process.env.MYSQLPORT,
    ssl: {
        rejectUnauthorized: false
    },
    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

const normalizarFecha = (valor) => {
    if (valor === null || valor === undefined || valor === '') return null;

    const texto = String(valor).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) return texto;

    const serial = Number(texto);
    if (!Number.isFinite(serial) || serial < 1 || serial > 60000) return texto || null;

    const fecha = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
    return fecha.toISOString().slice(0, 10);
};

const ensureEquipoColumns = async () => {
    const columnas = [
        ['nombre_equipo', 'VARCHAR(255) NULL'],
        ['anydesk', 'VARCHAR(100) NULL']
    ];

    for (const [columna, definicion] of columnas) {
        const [existente] = await pool.query(
            `SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS
             WHERE TABLE_SCHEMA = ? AND TABLE_NAME = 'equipos' AND COLUMN_NAME = ?`,
            [process.env.DB_NAME, columna]
        );

        if (existente.length === 0) {
            await pool.query(`ALTER TABLE equipos ADD COLUMN ${columna} ${definicion}`);
        }
    }
};

// Routes for Elementos (existing)
app.get('/api/elementos', async (req, res) => {
    try {
        const { search } = req.query;
        let query = "SELECT id, cantidad, modelo, marca, serial, placa, descripcion, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') AS fechaIngreso, DATE_FORMAT(fecha_baja, '%Y-%m-%d') AS fechaBaja FROM elementos";
        let params = [];

        if (search) {
            query += ` WHERE descripcion LIKE ? OR marca LIKE ? OR serial LIKE ? OR placa LIKE ? OR modelo LIKE ?`;
            const searchVal = `%${search}%`;
            params = [searchVal, searchVal, searchVal, searchVal, searchVal];
        }

        const [rows] = await pool.query(query, params);
        res.json(rows);
    } catch (error) {
        console.error('Error fetching elements:', error);
        res.status(500).json({ error: 'Error al obtener elementos' });
    }
});

app.post('/api/elementos', async (req, res) => {
    try {
        const { cantidad, modelo, descripcion, marca, serial, placa, fechaIngreso, fechaBaja } = req.body;
        const query = 'INSERT INTO elementos (cantidad, modelo, descripcion, marca, serial, placa, fecha_ingreso, fecha_baja) VALUES (?, ?, ?, ?, ?, ?, ?, ?)';
        const [result] = await pool.query(query, [cantidad, modelo, descripcion, marca, serial, placa, fechaIngreso || null, fechaBaja || null]);
        res.status(201).json({ id: result.insertId, message: 'Elemento registrado con éxito' });
    } catch (error) {
        console.error('Error inserting element:', error);
        res.status(500).json({ error: 'Error al registrar elemento' });
    }
});

app.delete('/api/elementos/:id', async (req, res) => {
    try {
        const { id } = req.params;
        await pool.query('DELETE FROM elementos WHERE id = ?', [id]);
        res.json({ message: 'Elemento eliminado con éxito' });
    } catch (error) {
        console.error('Error deleting element:', error);
        res.status(500).json({ error: 'Error al eliminar elemento' });
    }
});

// Routes for Equipos (new)
app.get('/api/equipos', async (req, res) => {
    try {
        const { search } = req.query;
        let query = "SELECT id, marca, modelo, estado, nombre_equipo, DATE_FORMAT(fecha_compra, '%Y-%m-%d') AS fechaCompra, placa, usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk, DATE_FORMAT(fecha_ultimo_mantenimiento, '%Y-%m-%d') AS fechaUltimoMantenimiento, DATE_FORMAT(fecha_proximo_mantenimiento, '%Y-%m-%d') AS fechaProximoMantenimiento FROM equipos";
        let params = [];

        if (search) {
            query += ` WHERE marca LIKE ? OR modelo LIKE ? OR estado LIKE ? OR nombre_equipo LIKE ? OR numero_serie LIKE ? OR usuario LIKE ? OR correo LIKE ? OR placa LIKE ? OR sistema_operativo LIKE ? OR ubicacion LIKE ? OR anydesk LIKE ?`;
            const searchVal = `%${search}%`;
            params = [searchVal, searchVal, searchVal, searchVal, searchVal, searchVal, searchVal, searchVal, searchVal, searchVal, searchVal];
        }

        const [rows] = await pool.query(query, params);
        res.json(rows);
    } catch (error) {
        console.error('Error fetching equipos:', error);
        res.status(500).json({ error: 'Error al obtener equipos' });
    }
});

app.post('/api/equipos', async (req, res) => {
    try {
        const { 
            marca, modelo, estado, nombre_equipo, fecha_compra, placa,
            usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk,
            fecha_ultimo_mantenimiento, fecha_proximo_mantenimiento 
        } = req.body;

        const query = `INSERT INTO equipos 
            (marca, modelo, estado, nombre_equipo, fecha_compra, placa, usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk, fecha_ultimo_mantenimiento, fecha_proximo_mantenimiento)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

        const [result] = await pool.query(query, [
            marca, modelo, estado, nombre_equipo, normalizarFecha(fecha_compra), placa,
            usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk,
            normalizarFecha(fecha_ultimo_mantenimiento), normalizarFecha(fecha_proximo_mantenimiento)
        ]);

        res.status(201).json({ id: result.insertId, message: 'Equipo registrado con éxito' });
    } catch (error) {
        console.error('Error inserting equipo:', error);
        res.status(500).json({ error: 'Error al registrar equipo' });
    }
});

// Ruta para carga de Excel con multer - MEJORADA CON VALIDACIONES
app.post('/api/equipos/upload', upload.single('file'), async (req, res) => {
    try {
        console.log('=== INICIANDO CARGA DE ARCHIVO ===');
        
        if (!req.file) {
            console.error('No file uploaded');
            return res.status(400).json({ error: 'No se subió archivo' });
        }

        console.log(`Archivo recibido: ${req.file.originalname}, tamaño: ${req.file.size} bytes`);

        // Validar que el buffer existe y tiene contenido
        if (!req.file.buffer || req.file.buffer.length === 0) {
            console.error('Buffer is empty');
            return res.status(400).json({ error: 'El archivo está vacío' });
        }

        // Leer el archivo Excel desde memoria - CORREGIDO: type: 'array'
        let workbook;
        try {
            workbook = XLSX.read(req.file.buffer, { type: 'array' });
            console.log(`Workbook leído. Hojas: ${workbook.SheetNames.join(', ')}`);
        } catch (readError) {
            console.error('Error reading Excel:', readError.message);
            return res.status(400).json({ error: 'Error al leer el archivo Excel. Verifica que sea un archivo válido.' });
        }

        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
            console.error('No sheets found in workbook');
            return res.status(400).json({ error: 'El archivo Excel no contiene hojas' });
        }

        const sheetName = workbook.SheetNames[0];
        console.log(`Procesando hoja: ${sheetName}`);
        
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            console.error('Worksheet is null or undefined');
            return res.status(400).json({ error: 'No se pudo acceder a la hoja del archivo' });
        }

        let jsonData;
        try {
            jsonData = XLSX.utils.sheet_to_json(worksheet);
            console.log(`Datos extraídos: ${jsonData.length} filas`);
        } catch (parseError) {
            console.error('Error parsing JSON from sheet:', parseError.message);
            return res.status(400).json({ error: 'Error al convertir los datos del Excel' });
        }

        if (jsonData.length === 0) {
            console.error('No data rows found');
            return res.status(400).json({ error: 'El archivo está vacío o no contiene datos válidos' });
        }

        console.log(`Primera fila del Excel:`, jsonData[0]);

        // Mapear función (igual a la del frontend)
        const mapearEquipo = (fila) => {
            const texto = (valor) => String(valor ?? "").trim();
            const normalizar = (valor) => texto(valor)
                .normalize("NFD")
                .replace(/[\u0300-\u036f]/g, "")
                .toLowerCase()
                .replace(/[^a-z0-9]/g, "");

            const valorFila = (nombres) => {
                const buscados = nombres.map(normalizar);
                const clave = Object.keys(fila).find((key) => buscados.includes(normalizar(key)));
                return clave ? texto(fila[clave]) : "";
            };

            const fechaExcel = (valor) => {
                const limpio = texto(valor);
                if (!limpio) return "";
                if (/^\d{4}-\d{2}-\d{2}$/.test(limpio)) return limpio;

                const serial = Number(limpio);
                if (!Number.isFinite(serial) || serial < 1 || serial > 60000) return limpio;

                const fecha = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
                return fecha.toISOString().slice(0, 10);
            };

            return {
                marca: valorFila(["Marca"]),
                modelo: valorFila(["Modelo", "Modelo CPU", "Tipo", "Tipo de PC", "Tipo de equipo"]),
                estado: valorFila(["Estado"]),
                nombre_equipo: valorFila(["Nombre del equipo", "Nombre equipo", "Nombre nuevo de equipo", "Equipo", "Nombre PC", "Host", "Hostname"]),
                fecha_compra: fechaExcel(valorFila(["Fecha de compra", "Fecha compra", "Compra"])),
                placa: valorFila(["Placa", "Placa CPU", "Placa TICS", "Placa TICS/Activo"]),
                usuario: valorFila(["Responsable", "Nombre responsable", "Funcionario", "Nombre", "", "__EMPTY", "__EMPTY_1"]) || valorFila(["Usuario"]),
                correo: valorFila(["Correo", "Email", "Correo electronico", "Correo electrónico"]),
                sistema_operativo: valorFila(["Sistema operativo", "SO", "S.O."]),
                numero_serie: valorFila(["Serial", "Serial CPU", "Serie", "Numero de serie", "Numero serie", "Número de serie", "S/N"]),
                ubicacion: valorFila(["Ubicacion", "Ubicación", "Sede"]),
                anydesk: valorFila(["AnyDesk", "Anydesk", "Anidex", "Anidesk"]),
                fecha_ultimo_mantenimiento: valorFila(["fecha_ultimo_mantenimiento", "Ultimo mantenimiento", "Último mantenimiento", "Manto anterior", "Fecha mantenimiento"]),
                fecha_proximo_mantenimiento: valorFila(["fecha_proximo_mantenimiento", "Proximo mantenimiento", "Próximo mantenimiento", "Proxima revision", "Próxima revisión"])
            };
        };

        const equiposMapeados = jsonData.map(mapearEquipo);
        console.log(`Equipos mapeados: ${equiposMapeados.length}`);

        // Insertar en la base de datos
        const query = `INSERT INTO equipos 
            (marca, modelo, estado, nombre_equipo, fecha_compra, placa, usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk, fecha_ultimo_mantenimiento, fecha_proximo_mantenimiento)
            VALUES ?`;

        const values = equiposMapeados.map(e => [
            e.marca, e.modelo, e.estado, e.nombre_equipo, normalizarFecha(e.fecha_compra),
            e.placa, e.usuario, e.correo, e.sistema_operativo, e.numero_serie, e.ubicacion, e.anydesk,
            normalizarFecha(e.fecha_ultimo_mantenimiento), normalizarFecha(e.fecha_proximo_mantenimiento)
        ]);

        try {
            await pool.query(query, [values]);
            console.log(`✅ ${equiposMapeados.length} equipos insertados exitosamente`);
        } catch (dbError) {
            console.error('Database insertion error:', dbError.message);
            return res.status(400).json({ error: 'Error al guardar en la base de datos: ' + dbError.message });
        }

        res.status(201).json({ 
            message: `${equiposMapeados.length} equipos importados con éxito`,
            equipos: equiposMapeados 
        });
        
    } catch (error) {
        console.error('=== ERROR PROCESANDO EXCEL ===');
        console.error('Stack:', error.stack);
        console.error('Mensaje:', error.message);
        res.status(400).json({ error: 'No se pudo procesar el archivo Excel. Verifica el formato: ' + error.message });
    }
});

// Bulk insert for Excel import (JSON)
app.post('/api/equipos/bulk', async (req, res) => {
    try {
        const { equipos } = req.body;
        if (!Array.isArray(equipos) || equipos.length === 0) {
            return res.status(400).json({ error: 'No hay datos para importar' });
        }

        const query = `INSERT INTO equipos 
            (marca, modelo, estado, nombre_equipo, fecha_compra, placa, usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk, fecha_ultimo_mantenimiento, fecha_proximo_mantenimiento)
            VALUES ?`;

        const values = equipos.map(e => [
            e.marca, e.modelo, e.estado, e.nombre_equipo, normalizarFecha(e.fecha_compra),
            e.placa, e.usuario, e.correo, e.sistema_operativo, e.numero_serie, e.ubicacion, e.anydesk,
            normalizarFecha(e.fecha_ultimo_mantenimiento), normalizarFecha(e.fecha_proximo_mantenimiento)
        ]);

        await pool.query(query, [values]);
        res.status(201).json({ message: `${equipos.length} equipos importados con éxito` });
    } catch (error) {
        console.error('Error bulk inserting equipos:', error);
        res.status(500).json({ error: 'Error al importar equipos' });
    }
});

app.delete('/api/equipos/all', async (req, res) => {
    try {
        await pool.query('DELETE FROM equipos');
        res.json({ message: 'Todos los equipos han sido eliminados con éxito' });
    } catch (error) {
        console.error('Error deleting all equipos:', error);
        res.status(500).json({ error: 'Error al eliminar todos los equipos' });
    }
});

app.delete('/api/equipos/:id', async (req, res) => {
    try {
        const { id } = req.params;
        await pool.query('DELETE FROM equipos WHERE id = ?', [id]);
        res.json({ message: 'Equipo eliminado con éxito' });
    } catch (error) {
        console.error('Error deleting equipo:', error);
        res.status(500).json({ error: 'Error al eliminar equipo' });
    }
});


// Start server
app.listen(PORT, async () => {
    try {
        await ensureEquipoColumns();
    } catch (error) {
        console.error('No se pudieron preparar las columnas de equipos:', error.message);
    }
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
