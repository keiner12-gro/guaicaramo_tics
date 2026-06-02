require('dotenv').config();

const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const path = require('path');
const multer = require('multer');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');

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
   VARIABLES DE ENTORNO (CONFIGURACIÓN)
====================================================== */
console.log('\n======================================================');
console.log('       ⚙️  CONFIGURACIÓN DE ENTORNO DETECTADA');
console.log('======================================================');
console.log(`📡 PUERTO:    ${process.env.PORT || 3000}`);
console.log(`🏠 DB HOST:   ${process.env.DB_HOST}`);
console.log(`🔌 DB PORT:   ${process.env.DB_PORT}`);
console.log(`👤 DB USER:   ${process.env.DB_USER}`);
console.log(`🗄️  DB NAME:   ${process.env.DB_NAME}`);
console.log('======================================================\n');

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
        console.log('✅ BASE DE DATOS: CONEXIÓN EXITOSA');
        conn.release();
    } catch (err) {
        console.error('❌ BASE DE DATOS: ERROR DE CONEXIÓN');
        console.error(err.message);
    }
})();

/* ======================================================
   EXPORTAR A EXCEL (ESTILIZADO Y FILTRADO)
====================================================== */

const aplicarEstilosExcel = (worksheet, title) => {
    // Estilo para el encabezado
    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FF4F81BD' }
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    // Bordes para todas las celdas con datos
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell) => {
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };
        });
    });

    // Autoajustar ancho de columnas
    worksheet.columns.forEach(column => {
        let maxColumnLength = 0;
        column.eachCell({ includeEmpty: true }, cell => {
            const columnLength = cell.value ? cell.value.toString().length : 10;
            if (columnLength > maxColumnLength) {
                maxColumnLength = columnLength;
            }
        });
        column.width = maxColumnLength < 12 ? 12 : maxColumnLength + 2;
    });
};

app.get('/api/export/elementos', async (req, res) => {
    try {
        const { search } = req.query;
        let query = "SELECT cantidad, modelo, marca, serial, placa, descripcion, DATE_FORMAT(fecha_ingreso, '%Y-%m-%d') AS fecha_ingreso, DATE_FORMAT(fecha_baja, '%Y-%m-%d') AS fecha_baja FROM elementos";
        let params = [];

        if (search) {
            query += ` WHERE descripcion LIKE ? OR marca LIKE ? OR serial LIKE ? OR placa LIKE ? OR modelo LIKE ?`;
            const searchVal = `%${search}%`;
            params = [searchVal, searchVal, searchVal, searchVal, searchVal];
        }

        const [rows] = await pool.query(query, params);
        
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Elementos');

        if (rows.length > 0) {
            worksheet.columns = Object.keys(rows[0]).map(key => ({
                header: key.toUpperCase().replace('_', ' '),
                key: key
            }));
            worksheet.addRows(rows);
            aplicarEstilosExcel(worksheet, 'Reporte de Elementos');
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="reporte_elementos_${new Date().getTime()}.xlsx"`);

        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Error al exportar elementos' });
    }
});

app.get('/api/export/equipos', async (req, res) => {
    try {
        const { search } = req.query;
        let query = `
            SELECT
                marca,
                modelo,
                estado,
                nombre_equipo,
                DATE_FORMAT(fecha_compra, '%Y-%m-%d') AS fecha_compra,
                placa,
                usuario,
                correo,
                sistema_operativo,
                numero_serie,
                ubicacion,
                anydesk,
                DATE_FORMAT(fecha_ultimo_mantenimiento, '%Y-%m-%d') AS fecha_ultimo_manto,
                DATE_FORMAT(fecha_proximo_mantenimiento, '%Y-%m-%d') AS fecha_proximo_manto
            FROM equipos
        `;
        let params = [];

        if (search) {
            query += ` WHERE marca LIKE ? OR modelo LIKE ? OR estado LIKE ? OR nombre_equipo LIKE ? OR numero_serie LIKE ? OR usuario LIKE ? OR correo LIKE ? OR placa LIKE ? OR sistema_operativo LIKE ? OR ubicacion LIKE ? OR anydesk LIKE ?`;
            const searchVal = `%${search}%`;
            params = Array(11).fill(searchVal);
        }

        const [rows] = await pool.query(query, params);
        
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Equipos');

        if (rows.length > 0) {
            worksheet.columns = Object.keys(rows[0]).map(key => ({
                header: key.toUpperCase().replace(/_/g, ' '),
                key: key
            }));
            worksheet.addRows(rows);
            aplicarEstilosExcel(worksheet, 'Reporte de Equipos');
        }

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', `attachment; filename="reporte_equipos_${new Date().getTime()}.xlsx"`);

        await workbook.xlsx.write(res);
        res.end();
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: 'Error al exportar equipos' });
    }
});

/* ======================================================
   FECHA
====================================================== */

const normalizarFecha = (valor) => {
    if (valor === null || valor === undefined || valor === '') {
        return null;
    }

    const texto = String(valor).trim().toLowerCase();
    if (texto === 'n/a' || texto === 'na' || texto === '') return null;

    // 1. Si ya es YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(String(valor).trim())) {
        return String(valor).trim();
    }

    // 2. Si es un serial de Excel (número)
    const serial = Number(texto);
    // Se ajusta el umbral a 25569 (01/01/1970) para evitar que años simples (como 2024) se procesen como días
    if (Number.isFinite(serial) && serial >= 25569 && serial <= 100000) {
        try {
            const fecha = new Date(
                Date.UTC(1899, 11, 30) + serial * 86400000
            );
            return fecha.toISOString().slice(0, 10);
        } catch (e) {
            return null;
        }
    }

    // 3. Evitar interpretar años sueltos como fechas (ej: 2016)
    if (/^\d{4}$/.test(texto)) {
        return null;
    }

    // 4. Intentar parsear como fecha genérica
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

app.put('/api/elementos/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { cantidad, modelo, descripcion, marca, serial, placa, fechaIngreso, fechaBaja } = req.body;
        const query = `UPDATE elementos SET cantidad = ?, modelo = ?, descripcion = ?, marca = ?, serial = ?, placa = ?, fecha_ingreso = ?, fecha_baja = ? WHERE id = ?`;
        await pool.query(query, [cantidad, modelo, descripcion, marca, serial, placa, fechaIngreso || null, fechaBaja || null, id]);
        res.json({ message: 'Elemento actualizado con éxito' });
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
   SUBIR EXCEL: Lógica multi-hoja y mapeo flexible
====================================================== */
app.post('/api/equipos/upload', upload.single('file'), async (req, res) => {
    try {

        if (!req.file) {
            return res.status(400).json({
                error: 'No se subió archivo'
            });
        }

        console.log('\n========== IMPORTANDO EXCEL ==========');

        const workbook = XLSX.read(
            req.file.buffer,
            { type: 'buffer' }
        );

        console.log('Hojas encontradas:');
        console.log(workbook.SheetNames);

        let todosLosRegistros = [];

        // LEER TODAS LAS HOJAS
        for (const nombreHoja of workbook.SheetNames) {

            const worksheet = workbook.Sheets[nombreHoja];

            const datosHoja = XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: ''
                }
            );

            console.log(
                `Hoja "${nombreHoja}" -> ${datosHoja.length} filas`
            );

            todosLosRegistros.push(...datosHoja);
        }

        console.log(
            `Total registros encontrados: ${todosLosRegistros.length}`
        );

        if (todosLosRegistros.length === 0) {
            return res.status(400).json({
                error: 'No se encontraron datos'
            });
        }

        console.log('COLUMNAS DETECTADAS:');
        console.log(
            Object.keys(todosLosRegistros[0])
        );

        const normalizarTexto = (texto) =>
            String(texto || '')
                .normalize('NFD')
                .replace(/[\u0300-\u036f]/g, '')
                .toLowerCase()
                .replace(/[^a-z0-9]/g, '');

        const obtenerValor = (fila, nombres) => {

            const buscados = nombres.map(
                n => normalizarTexto(n)
            );

            const clave = Object.keys(fila).find(
                k => buscados.includes(
                    normalizarTexto(k)
                )
            );

            return clave
                ? String(fila[clave] || '').trim()
                : '';
        };

        const equipos = todosLosRegistros.map(fila => ({

            marca: obtenerValor(fila, [
                'marca',
                'fabricante'
            ]),

            modelo: obtenerValor(fila, [
                'modelo cpu',
                'modelo',
                'tipo'
            ]),

            estado: obtenerValor(fila, [
                'estado'
            ]),

            nombre_equipo: obtenerValor(fila, [
                'nombre nuevo de equipo',
                'nombre equipo',
                'nombre del equipo',
                'equipo',
                'hostname',
                'nombre pc'
            ]),

            fecha_compra: obtenerValor(fila, [
                'fecha compra',
                'compra'
            ]),

            placa: obtenerValor(fila, [
                'placa cpu',
                'placa',
                'placa tics',
                'activo'
            ]),

            usuario: obtenerValor(fila, [
                'usuario',
                'responsable',
                'funcionario',
                'nombre'
            ]),

            correo: obtenerValor(fila, [
                'correo',
                'email'
            ]),

            sistema_operativo: obtenerValor(fila, [
                'sistema operativo',
                'so'
            ]),

            numero_serie: obtenerValor(fila, [
                'serial cpu',
                'serial',
                'serie',
                'numero serie',
                's/n'
            ]),

            ubicacion: obtenerValor(fila, [
                'ubicacion',
                'ubicación',
                'sede',
                'oficina'
            ]),

            anydesk: obtenerValor(fila, [
                'anydesk'
            ]),

            fecha_ultimo_mantenimiento: obtenerValor(fila, [
                'fecha mantenimiento',
                'ultimo mantenimiento'
            ]),

            fecha_proximo_mantenimiento: obtenerValor(fila, [
                'proxima revision',
                'proximo mantenimiento'
            ])
        }));

        const equiposValidos = equipos.filter(e =>

            e.marca ||
            e.modelo ||
            e.nombre_equipo ||
            e.numero_serie ||
            e.usuario
        );

        console.log(
            `Equipos válidos: ${equiposValidos.length}`
        );

        if (equiposValidos.length === 0) {
            return res.status(400).json({
                error: 'No se detectaron equipos válidos'
            });
        }

        const values = equiposValidos.map(e => [

            e.marca,
            e.modelo,
            e.estado,
            e.nombre_equipo,

            normalizarFecha(
                e.fecha_compra
            ),

            e.placa,
            e.usuario,
            e.correo,
            e.sistema_operativo,
            e.numero_serie,
            e.ubicacion,
            e.anydesk,

            normalizarFecha(
                e.fecha_ultimo_mantenimiento
            ),

            normalizarFecha(
                e.fecha_proximo_mantenimiento
            )
        ]);

        const query = `
            INSERT INTO equipos (
                marca,
                modelo,
                estado,
                nombre_equipo,
                fecha_compra,
                placa,
                usuario,
                correo,
                sistema_operativo,
                numero_serie,
                ubicacion,
                anydesk,
                fecha_ultimo_mantenimiento,
                fecha_proximo_mantenimiento
            )
            VALUES ?
        `;

        const [result] = await pool.query(
            query,
            [values]
        );

        console.log(
            `Insertados: ${result.affectedRows}`
        );

        res.status(201).json({
            message: `${result.affectedRows} equipos importados correctamente`
        });

    } catch (error) {

        console.error(error);

        res.status(500).json({
            error: error.message
        });
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

app.put('/api/equipos/:id', async (req, res) => {
    try {
        const { id } = req.params;
        const { 
            marca, modelo, estado, nombre_equipo, fecha_compra, placa,
            usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk,
            fecha_ultimo_mantenimiento, fecha_proximo_mantenimiento 
        } = req.body;

        const query = `UPDATE equipos SET 
            marca = ?, modelo = ?, estado = ?, nombre_equipo = ?, fecha_compra = ?, placa = ?, 
            usuario = ?, correo = ?, sistema_operativo = ?, numero_serie = ?, ubicacion = ?, 
            anydesk = ?, fecha_ultimo_mantenimiento = ?, fecha_proximo_mantenimiento = ? 
            WHERE id = ?`;

        await pool.query(query, [
            marca, modelo, estado, nombre_equipo, normalizarFecha(fecha_compra), placa,
            usuario, correo, sistema_operativo, numero_serie, ubicacion, anydesk,
            normalizarFecha(fecha_ultimo_mantenimiento), normalizarFecha(fecha_proximo_mantenimiento), id
        ]);

        res.json({ message: 'Equipo actualizado con éxito' });
    } catch (error) {
        console.error(error);
        res.status(500).json({ error: error.message });
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