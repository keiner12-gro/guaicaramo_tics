-- phpMyAdmin SQL Dump
-- version 5.2.1
-- https://www.phpmyadmin.net/
--
-- Servidor: 127.0.0.1
-- Tiempo de generación: 27-05-2026 a las 16:22:33
-- Versión del servidor: 10.4.32-MariaDB
-- Versión de PHP: 8.0.30

SET SQL_MODE = "NO_AUTO_VALUE_ON_ZERO";
START TRANSACTION;
SET time_zone = "+00:00";


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8mb4 */;

--
-- Base de datos: `bodega`
--

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `elementos`
--

CREATE TABLE `elementos` (
  `id` int(11) NOT NULL,
  `cantidad` int(11) NOT NULL,
  `modelo` varchar(255) NOT NULL,
  `marca` varchar(100) NOT NULL,
  `serial` varchar(100) NOT NULL,
  `placa` varchar(100) NOT NULL,
  `descripcion` varchar(100) DEFAULT NULL,
  `fecha_ingreso` date NOT NULL,
  `fecha_baja` date DEFAULT NULL
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Volcado de datos para la tabla `elementos`
--

INSERT INTO `elementos` (`id`, `cantidad`, `modelo`, `marca`, `serial`, `placa`, `descripcion`, `fecha_ingreso`, `fecha_baja`) VALUES
(8, 1, 'monitor', 'hp', 'N/A', 'N/A', 'monitor traido de bogota', '2026-05-15', NULL),
(9, 1, 'mause', 'hp', 'N/A', 'N/A', 'mouse traido de bogota', '2026-05-20', NULL);

-- --------------------------------------------------------

--
-- Estructura de tabla para la tabla `equipos`
--

CREATE TABLE `equipos` (
  `id` int(11) NOT NULL,
  `marca` varchar(100) DEFAULT 'N/A',
  `modelo` varchar(100) DEFAULT 'N/A',
  `estado` varchar(50) DEFAULT 'N/A',
  `nombre_equipo` varchar(150) DEFAULT 'N/A',
  `fecha_compra` date DEFAULT NULL,
  `placa` varchar(100) DEFAULT 'N/A',
  `usuario` varchar(100) DEFAULT 'N/A',
  `correo` varchar(150) DEFAULT 'N/A',
  `sistema_operativo` varchar(100) DEFAULT 'N/A',
  `numero_serie` varchar(100) DEFAULT 'N/A',
  `ubicacion` varchar(100) DEFAULT 'N/A',
  `anydesk` varchar(100) DEFAULT 'N/A',
  `fecha_ultimo_mantenimiento` date DEFAULT NULL,
  `fecha_proximo_mantenimiento` date DEFAULT NULL,
  `fecha_registro` timestamp NOT NULL DEFAULT current_timestamp()
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_general_ci;

--
-- Índices para tablas volcadas
--

--
-- Indices de la tabla `elementos`
--
ALTER TABLE `elementos`
  ADD PRIMARY KEY (`id`);

--
-- Indices de la tabla `equipos`
--
ALTER TABLE `equipos`
  ADD PRIMARY KEY (`id`);

--
-- AUTO_INCREMENT de las tablas volcadas
--

--
-- AUTO_INCREMENT de la tabla `elementos`
--
ALTER TABLE `elementos`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=10;

--
-- AUTO_INCREMENT de la tabla `equipos`
--
ALTER TABLE `equipos`
  MODIFY `id` int(11) NOT NULL AUTO_INCREMENT, AUTO_INCREMENT=68;
COMMIT;

/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@OLD_CHARACTER_SET_CONNECTION */;