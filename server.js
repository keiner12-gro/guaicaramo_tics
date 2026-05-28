require('dotenv').config();

const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

/* ======================================================
   MULTER CONFIG
====================================================== */

const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 50 * 1024 * 1024
    },
    fileFilter: (req, file, cb) => {
        const allowedMimes = [
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-excel'
        ];

        if (
            allowedMimes.includes(file.mimetype) ||
            file.originalname.endsWith('.xlsx') ||
            file.originalname.endsWith('.xls')
        ) {
            cb(null, true);
        } else {
            cb(new Error('Solo se permiten archivos Excel'));
        }
    }
});

/* ======================================================
   MIDDLEWARE
====================================================== */

app.use(cors());
app.use(express.json());

app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

/* ======================================================
   MYSQL CONNECTION
====================================================== */

console.log('================================');
console.log('DB_HOST:', process.env.DB_HOST);
console.log('DB_PORT:', process.env.DB_PORT);
console.log('DB_NAME:', process.env.DB_NAME);
console.log('================================');

const pool = mysql.createPool({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,

    port: Number(process.env.DB_PORT),

    ssl: {
        rejectUnauthorized: false
    },

    connectTimeout: 60000,

    waitForConnections: true,
    connectionLimit: 10,
    queueLimit: 0
});

/* ======================================================
   TEST MYSQL CONNECTION
====================================================== */

(async () => {
    try {
        const conn = await pool.getConnection();

        console.log('================================');
        console.log('✅ MYSQL CONECTADO');
        console.log('================================');

        conn.release();
    } catch (err) {
        console.log('================================');
        console.log('❌ ERROR MYSQL');
        console.log(err);
        console.log('================================');
    }
})();

/* ======================================================
   FECHA
====================================================== */

const normalizarFecha = (valor) => {
    if (valor === null || valor === undefined || valor === '') {
        return null;
    }

    const texto = String(valor).trim();

    // 1. Si ya es YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {
        return texto;
    }

    // 2. Si es un serial de Excel (número)
    const serial = Number(texto);
    if (Number.isFinite(serial) && serial >= 1 && serial <= 100000) {
        try {
            const fecha = new Date(
                Date.UTC(1899, 11, 30) + serial * 86400000
            );
            return fecha.toISOString().slice(0, 10);
        } catch (e) {
            return null;
        }
    }

    // 3. Intentar parsear como fecha genérica
    const d = new Date(texto);
    if (!isNaN(d.getTime())) {
        try {
            return d.toISOString().slice(0, 10);
        } catch (e) {
            return null;
        }
    }

    // 4. Si no es nada de lo anterior (como 'P'), devolver null para no romper SQL
    return null;
};

/* ======================================================
   CREAR COLUMNAS SI NO EXISTEN
====================================================== */

const ensureEquipoColumns = async () => {
    const columnas = [
        ['nombre_equipo', 'VARCHAR(255) NULL'],
        ['anydesk', 'VARCHAR(255) NULL'],
        ['fecha_ultimo_mantenimiento', 'DATE NULL'],
        ['fecha_proximo_mantenimiento', 'DATE NULL']
    ];

    for (const [columna, definicion] of columnas) {

        const [existente] = await pool.query(
            `
            SELECT COLUMN_NAME
            FROM INFORMATION_SCHEMA.COLUMNS
            WHERE TABLE_SCHEMA = ?
            AND TABLE_NAME = 'equipos'
            AND COLUMN_NAME = ?
            `,
            [process.env.DB_NAME, columna]
        );

        if (existente.length === 0) {

            console.log(`➕ Creando columna: ${columna}`);

            await pool.query(
                `
                ALTER TABLE equipos
                ADD COLUMN ${columna} ${definicion}
                `
            );
        }
    }
};

/* ======================================================
   RUTAS ELEMENTOS
====================================================== */

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
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

app.post('/api/elementos', async (req, res) => {
    try {
        const { cantidad, modelo, descripcion, marca, serial, placa, fechaIngreso, fechaBaja } = req.body;
        const query = 'INSERT INTO elementos (cantidad, modelo, descripcion, marca, serial, placa, fecha_ingreso, fecha_baja) VALUES (?, ?, ?, ?, ?, ?, ?, ?)';
        const [result] = await pool.query(query, [cantidad, modelo, descripcion, marca, serial, placa, fechaIngreso || null, fechaBaja || null]);
        res.status(201).json({ id: result.insertId, message: 'Elemento registrado con éxito' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/elementos/:id', async (req, res) => {
    try {
        const { id } = req.params;
        await pool.query('DELETE FROM elementos WHERE id = ?', [id]);
        res.json({ message: 'Elemento eliminado con éxito' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

/* ======================================================
   OBTENER EQUIPOS
====================================================== */

app.get('/api/equipos', async (req, res) => {

    try {
        const { search } = req.query;
        let query = `
            SELECT
                id,
                marca,
                modelo,
                estado,
                nombre_equipo,
                DATE_FORMAT(fecha_compra, '%Y-%m-%d') AS fechaCompra,
                placa,
                usuario,
                correo,
                sistema_operativo,
                numero_serie,
                ubicacion,
                anydesk,
                DATE_FORMAT(fecha_ultimo_mantenimiento, '%Y-%m-%d') AS fechaUltimoMantenimiento,
                DATE_FORMAT(fecha_proximo_mantenimiento, '%Y-%m-%d') AS fechaProximoMantenimiento
            FROM equipos
        `;
        let params = [];

        if (search) {
            query += ` WHERE marca LIKE ? OR modelo LIKE ? OR estado LIKE ? OR nombre_equipo LIKE ? OR numero_serie LIKE ? OR usuario LIKE ? OR correo LIKE ? OR placa LIKE ? OR sistema_operativo LIKE ? OR ubicacion LIKE ? OR anydesk LIKE ?`;
            const searchVal = `%${search}%`;
            params = Array(11).fill(searchVal);
        }

        const [rows] = await pool.query(query, params);
        res.json(rows);

    } catch (error) {

        console.error(error);

        res.status(500).json({
            error: error.message
        });
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
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

app.delete('/api/equipos/all', async (req, res) => {
    try {
        await pool.query('DELETE FROM equipos');
        res.json({ message: 'Todos los equipos han sido eliminados con éxito' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
    }
});

/* ======================================================
   SUBIR EXCEL: Lógica ultra-robusta y filtrado estricto
====================================================== */
app.post('/api/equipos/upload', upload.single('file'), async (req, res) => {
    try {
        if (!req.file) return res.status(400).json({ error: 'No se subió archivo' });

        console.log('\n--- INICIO DE IMPORTACIÓN ---');
        const workbook = XLSX.read(req.file.buffer, { type: 'array' });
        
        // 1. Encontrar la mejor hoja (la que tenga más coincidencias de encabezados)
        let mejorHoja = workbook.SheetNames[0];
        let maxCoincidencias = 0;

        workbook.SheetNames.forEach(name => {
            const sheet = workbook.Sheets[name];
            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, range: 0 }); // Leer primeras filas
            const headers = (rows[0] || []).map(h => String(h).toLowerCase().trim());
            
            const coincidencias = headers.filter(h => 
                ['marca', 'modelo', 'serie', 'usuario', 'placa'].includes(h)
            ).length;

            if (coincidencias > maxCoincidencias) {
                maxCoincidencias = coincidencias;
                mejorHoja = name;
            }
        });

        console.log(`Hoja seleccionada: "${mejorHoja}" (Coincidencias: ${maxCoincidencias})`);

        const worksheet = workbook.Sheets[mejorHoja];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
        
        console.log(`Total filas en bruto: ${rawData.length}`);
        if (rawData.length > 0) {
            console.log('Ejemplo de columnas detectadas:', Object.keys(rawData[0]));
        }

        // 2. Filtrado estricto de filas
        const jsonData = rawData.filter((fila, index) => {
            const marcaKey = Object.keys(fila).find(k => k.toLowerCase().includes('marca'));
            const modeloKey = Object.keys(fila).find(k => k.toLowerCase().includes('modelo'));
            
            const marca = String(fila[marcaKey] || '').trim().toUpperCase();
            const modelo = String(fila[modeloKey] || '').trim().toUpperCase();

            // REGLAS DE EXCLUSIÓN MÁS FUERTES:
            const esBasura = 
                !marca || 
                marca === 'MARCA' || 
                marca === 'N/A' || 
                marca === 'EDAD' || // Específicamente excluimos "edad" si aparece en la columna marca
                marca.length < 2;

            const tieneInfoMinima = marca.length > 1 && modelo.length > 0;

            if (index < 5) {
                console.log(`Fila ${index} debug: Marca="${marca}", Modelo="${modelo}" -> Pasa: ${!esBasura && tieneInfoMinima}`);
            }
            
            return !esBasura && tieneInfoMinima;
        });

        console.log(`Filas después del filtro: ${jsonData.length}`);

        const mapearEquipo = (fila) => {
            const texto = (v) => String(v ?? '').trim();
            const normalizar = (v) => texto(v).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');

            const valorFila = (nombres) => {
                const buscados = nombres.map(normalizar);
                const clave = Object.keys(fila).find(k => buscados.includes(normalizar(k)));
                // IMPORTANTE: Si la clave es __EMPTY, la ignoramos para evitar "edad" u otros datos basura
                if (clave && clave.startsWith('__EMPTY')) return '';
                return clave ? texto(fila[clave]) : '';
            };

            return {
                marca: valorFila(['Marca', 'Fabricante']),
                modelo: valorFila(['Modelo', 'Modelo CPU', 'Tipo']),
                estado: valorFila(['Estado', 'Condicion']),
                nombre_equipo: valorFila(['Nombre equipo', 'Hostname', 'Nombre PC']),
                fecha_compra: valorFila(['Fecha compra', 'Compra']),
                placa: valorFila(['Placa', 'Activo']),
                usuario: valorFila(['Responsable', 'Usuario', 'Funcionario', 'Nombre']),
                correo: valorFila(['Correo', 'Email']),
                sistema_operativo: valorFila(['Sistema operativo', 'SO']),
                numero_serie: valorFila(['Serial', 'Serie', 'S/N']),
                ubicacion: valorFila(['Ubicacion', 'Sede', 'Oficina']),
                anydesk: valorFila(['AnyDesk']),
                fecha_ultimo_mantenimiento: valorFila(['Fecha mantenimiento', 'Ultimo mantenimiento']),
                fecha_proximo_mantenimiento: valorFila(['Proxima revision', 'Proximo mantenimiento'])
            };
        };

        const equiposMapeados = jsonData.map(mapearEquipo);

        const values = equiposMapeados.map(e => [
            e.marca, e.modelo, e.estado, e.nombre_equipo,
            normalizarFecha(e.fecha_compra),
            e.placa, e.usuario, e.correo, e.sistema_operativo, e.numero_serie, e.ubicacion, e.anydesk,
            normalizarFecha(e.fecha_ultimo_mantenimiento),
            normalizarFecha(e.fecha_proximo_mantenimiento)
        ]);

        const query = `INSERT INTO equipos 
            (marca, modelo, estado, nombre_equipo, fecha_compra, placa, usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk, fecha_ultimo_mantenimiento, fecha_proximo_mantenimiento)
            VALUES ?`;

        const [result] = await pool.query(query, [values]);

        console.log(`✅ ÉXITO: ${result.affectedRows} filas insertadas.`);
        console.log('--- FIN DE IMPORTACIÓN ---\n');

        res.status(201).json({ 
            message: `Importación exitosa: ${result.affectedRows} equipos registrados.`,
            detalles: { total_leidas: rawData.length, procesadas: result.affectedRows }
        });
    } catch (error) {
        console.error('❌ ERROR EN UPLOAD:', error);
        res.status(500).json({ error: error.message });
    }
});

/* ======================================================
   ELIMINAR EQUIPO
====================================================== */

app.delete('/api/equipos/:id', async (req, res) => {

    try {

        const { id } = req.params;

        await pool.query(
            'DELETE FROM equipos WHERE id = ?',
            [id]
        );

        res.json({
            message: 'Equipo eliminado'
        });

    } catch (error) {

        console.error(error);

        res.status(500).json({
            error: error.message
        });
    }
});

/* ======================================================
   START SERVER
====================================================== */

app.listen(PORT, async () => {

    try {

        await ensureEquipoColumns();

    } catch (error) {

        console.error(
            'No se pudieron preparar las columnas:',
            error.message
        );
    }

    console.log(`Servidor corriendo en puerto ${PORT}`);
});