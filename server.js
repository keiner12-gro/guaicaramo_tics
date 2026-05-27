require('dotenv').config();
const express = require('express');
const mysql = require('mysql2/promise');
const cors = require('cors');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname)); // To serve index.html from root

// Database Connection
const pool = mysql.createPool({
    host: process.env.DB_HOST,
    user: process.env.DB_USER,
    password: process.env.DB_PASSWORD,
    database: process.env.DB_NAME,
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

// Bulk insert for Excel import
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
