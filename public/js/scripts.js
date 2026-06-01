const API_URL = "/api/elementos";
const API_EQUIPOS_URL = "/api/equipos";

const obtenerElementos = async (busqueda = "") => {
    try {
        const url = busqueda ? `${API_URL}?search=${encodeURIComponent(busqueda)}` : API_URL;
        const respuesta = await fetch(url);
        if (!respuesta.ok) throw new Error("Error al obtener datos");
        return await respuesta.json();
    } catch (error) {
        console.error("Error al obtener elementos del servidor:", error);
        return [];
    }
};

const obtenerEquipos = async (busqueda = "") => {
    try {
        const url = busqueda ? `${API_EQUIPOS_URL}?search=${encodeURIComponent(busqueda)}` : API_EQUIPOS_URL;
        const respuesta = await fetch(url);
        if (!respuesta.ok) throw new Error("Error al obtener equipos");
        return await respuesta.json();
    } catch (error) {
        console.error("Error al obtener equipos del servidor:", error);
        return [];
    }
};

const texto = (valor) => String(valor ?? "").trim();

const fechaExcel = (valor) => {
    const limpio = texto(valor);
    if (!limpio) return "";

    // 1. Si ya es YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(limpio)) return limpio;

    // 2. Si es un serial de Excel
    const serial = Number(limpio);
    if (Number.isFinite(serial) && serial >= 1 && serial <= 100000) {
        try {
            const fecha = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
            return fecha.toISOString().slice(0, 10);
        } catch (e) {
            return "";
        }
    }

    // 3. Intentar parsear como fecha genérica
    const d = new Date(limpio);
    if (!isNaN(d.getTime())) {
        try {
            return d.toISOString().slice(0, 10);
        } catch (e) {
            return "";
        }
    }

    // 4. Si no es nada de lo anterior, devolver vacío para evitar errores
    return "";
};

const valorFila = (fila, nombres) => {
    const normalizar = (valor) => texto(valor)
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase()
        .replace(/[^a-z0-9]/g, "");
    const buscados = nombres.map(normalizar);
    const clave = Object.keys(fila).find((key) => buscados.includes(normalizar(key)));
    return clave ? texto(fila[clave]) : "";
};

const mapearEquipo = (fila) => ({
    marca: valorFila(fila, ["Marca"]),
    modelo: valorFila(fila, ["Modelo", "Modelo CPU", "Tipo", "Tipo de PC", "Tipo de equipo"]),
    estado: valorFila(fila, ["Estado"]),
    nombre_equipo: valorFila(fila, ["Nombre del equipo", "Nombre equipo", "Nombre nuevo de equipo", "Equipo", "Nombre PC", "Host", "Hostname"]),
    fecha_compra: fechaExcel(valorFila(fila, ["Fecha de compra", "Fecha compra", "Compra"])),
    placa: valorFila(fila, ["Placa", "Placa CPU", "Placa TICS", "Placa TICS/Activo"]),
    usuario: valorFila(fila, ["Responsable", "Nombre responsable", "Funcionario", "Nombre", "", "__EMPTY", "__EMPTY_1"]) || valorFila(fila, ["Usuario"]),
    correo: valorFila(fila, ["Correo", "Email", "Correo electronico", "Correo electrónico"]),
    sistema_operativo: valorFila(fila, ["Sistema operativo", "SO", "S.O."]),
    numero_serie: valorFila(fila, ["Serial", "Serial CPU", "Serie", "Numero de serie", "Numero serie", "Número de serie", "S/N"]),
    ubicacion: valorFila(fila, ["Ubicacion", "Ubicación", "Sede"]),
    anydesk: valorFila(fila, ["AnyDesk", "Anydesk", "Anidex", "Anidesk"]),
    fecha_ultimo_mantenimiento: fechaExcel(valorFila(fila, ["fecha_ultimo_mantenimiento", "Ultimo mantenimiento", "Último mantenimiento", "Manto anterior", "Fecha mantenimiento"])),
    fecha_proximo_mantenimiento: fechaExcel(valorFila(fila, ["fecha_proximo_mantenimiento", "Proximo mantenimiento", "Próximo mantenimiento", "Proxima revision", "Próxima revisión"]))
});

const alerta = (icono, titulo, mensaje) => {
    if (window.Swal) {
        return Swal.fire({
            icon: icono,
            title: titulo,
            text: mensaje,
            confirmButtonColor: "#2563eb"
        });
    }
    alert(`${titulo}\n${mensaje}`);
    return Promise.resolve();
};

const confirmarAlerta = async (titulo, mensaje) => {
    if (window.Swal) {
        const resultado = await Swal.fire({
            icon: "warning",
            title: titulo,
            text: mensaje,
            showCancelButton: true,
            confirmButtonColor: "#dc2626",
            cancelButtonColor: "#2563eb",
            confirmButtonText: "Si, eliminar",
            cancelButtonText: "Cancelar"
        });
        return resultado.isConfirmed;
    }
    return window.confirm(`${titulo}\n${mensaje}`);
};

const columnasElementos = [
    { titulo: "Cantidad", campo: "cantidad" },
    { titulo: "Modelo", campo: "modelo" },
    { titulo: "Marca", campo: "marca" },
    { titulo: "Serial", campo: "serial" },
    { titulo: "Placa", campo: "placa" },
    { titulo: "Descripcion", campo: "descripcion" },
    { titulo: "Fecha de ingreso", campo: "fechaIngreso" },
    { titulo: "Fecha de baja", campo: "fechaBaja" }
];

const columnasEquipos = [
    { titulo: "Marca", campo: "marca" },
    { titulo: "Modelo", campo: "modelo" },
    { titulo: "Estado", campo: "estado" },
    { titulo: "Nombre equipo", campo: "nombre_equipo" },
    { titulo: "Compra", campo: "fechaCompra" },
    { titulo: "Placa", campo: "placa" },
    { titulo: "SO", campo: "sistema_operativo" },
    { titulo: "Ubicación", campo: "ubicacion" },
    { titulo: "Responsable", campo: "usuario" },
    { titulo: "Serial", campo: "numero_serie" },
    { titulo: "AnyDesk", campo: "anydesk" },
    { titulo: "Correo", campo: "correo" },
    { titulo: "Manto. Ant.", campo: "fechaUltimoMantenimiento" },
    { titulo: "Próx. Manto.", campo: "fechaProximoMantenimiento" }
];

const renderizarElementos = (elementos, permitirEliminar = false) => {
    // 🛡️ BLINDAJE: Si no es un arreglo válido, lo volvemos un arreglo vacío
    if (!elementos || !Array.isArray(elementos)) {
        console.warn("renderizarElementos recibió datos inválidos:", elementos);
        elementos = []; 
    }

    const cuerpoTabla = document.getElementById("elementos-body");
    if (!cuerpoTabla) return;

    cuerpoTabla.textContent = "";

    if (elementos.length === 0) {
        const fila = document.createElement("tr");
        const celda = document.createElement("td");
        celda.colSpan = permitirEliminar ? 9 : 8;
        celda.className = "px-6 py-4 text-center text-body";
        celda.textContent = "No hay elementos registrados.";
        fila.appendChild(celda);
        cuerpoTabla.appendChild(fila);
        return;
    }

    elementos.forEach((elemento) => {
        const fila = document.createElement("tr");
        fila.className = "border-b border-default hover:bg-neutral-secondary-soft";

        const campos = ["cantidad", "modelo", "marca", "serial", "placa", "descripcion", "fechaIngreso", "fechaBaja"];

        campos.forEach((campo) => {
            const celda = document.createElement("td");
            celda.className = "px-6 py-4";
            celda.textContent = elemento[campo] || "-";
            fila.appendChild(celda);
        });

        if (permitirEliminar) {
            const celdaAccion = document.createElement("td");
            celdaAccion.className = "px-6 py-4";
            const boton = document.createElement("button");
            boton.type = "button";
            boton.className = "boton-eliminar";
            boton.textContent = "Eliminar";
            boton.onclick = () => eliminarElemento(elemento.id, elemento.modelo || elemento.descripcion);
            celdaAccion.appendChild(boton);
            fila.appendChild(celdaAccion);
        }

        cuerpoTabla.appendChild(fila);
    });
};

const renderizarEquipos = (equipos, permitirEliminar = false) => {
    // 🛡️ BLINDAJE: Si no es un arreglo válido, lo volvemos un arreglo vacío
    if (!equipos || !Array.isArray(equipos)) {
        console.warn("renderizarEquipos recibió datos inválidos:", equipos);
        equipos = []; 
    }

    const cuerpoTabla = document.getElementById("equipos-body") || document.getElementById("equipos-eliminar-body");
    const contador = document.getElementById("contador-equipos");
    if (!cuerpoTabla) return;

    cuerpoTabla.textContent = "";
    if (contador) contador.textContent = `Total: ${equipos.length} equipos encontrados`;

    if (equipos.length === 0) {
        const fila = document.createElement("tr");
        const celda = document.createElement("td");
        celda.colSpan = permitirEliminar ? 7 : 14;
        celda.className = "px-6 py-10 text-center text-gray-500 font-medium bg-gray-50";
        celda.textContent = "No hay equipos registrados.";
        fila.appendChild(celda);
        cuerpoTabla.appendChild(fila);
        return;
    }

    equipos.forEach((equipo) => {
        const fila = document.createElement("tr");
        fila.className = "hover:bg-blue-50/50 transition-colors";

        if (permitirEliminar) {
            const campos = [
                { campo: "marca", extra: "font-bold text-gray-900" },
                { campo: "modelo", extra: "text-gray-700" },
                { campo: "nombre_equipo", extra: "font-semibold text-blue-600" },
                { campo: "numero_serie", extra: "font-mono text-gray-600" },
                { campo: "usuario", extra: "font-medium text-gray-800" },
                { campo: "placa", extra: "font-bold text-red-500" }
            ];
            
            campos.forEach(({ campo, extra }) => {
                const celda = document.createElement("td");
                celda.className = `px-8 py-5 border-b border-gray-100 ${extra}`;
                celda.textContent = equipo[campo] || "-";
                fila.appendChild(celda);
            });
            
            const celdaAccion = document.createElement("td");
            celdaAccion.className = "px-8 py-5 border-b border-gray-100 text-center";
            const boton = document.createElement("button");
            boton.type = "button";
            boton.className = "boton-eliminar scale-110 hover:scale-125 transition-transform duration-200";
            boton.textContent = "Eliminar";
            boton.onclick = () => eliminarEquipo(equipo.id, `${equipo.marca} ${equipo.modelo}`);
            celdaAccion.appendChild(boton);
            fila.appendChild(celdaAccion);
        } else {
            const crearCelda = (contenido, extraClass = "") => {
                const celda = document.createElement("td");
                celda.className = `px-6 py-4 border-b border-gray-100 ${extraClass}`;
                if (contenido instanceof Node) celda.appendChild(contenido);
                else celda.textContent = contenido || "-";
                return celda;
            };

            fila.appendChild(crearCelda(equipo.marca, "font-semibold text-gray-900"));
            fila.appendChild(crearCelda(equipo.modelo));

            const spanEstado = document.createElement("span");
            const est = String(equipo.estado).toLowerCase();
            spanEstado.className = "badge-estado " + (
                est.includes("nuevo") ? "badge-nuevo" :
                est.includes("usado") ? "badge-usado" :
                est.includes("manto") || est.includes("repar") ? "badge-manto" : "badge-default"
            );
            spanEstado.textContent = equipo.estado || "N/A";
            fila.appendChild(crearCelda(spanEstado));

            fila.appendChild(crearCelda(equipo.nombre_equipo));
            fila.appendChild(crearCelda(equipo.fechaCompra));
            fila.appendChild(crearCelda(equipo.placa, "font-mono font-bold text-blue-700"));
            fila.appendChild(crearCelda(equipo.sistema_operativo));
            fila.appendChild(crearCelda(equipo.ubicacion));
            fila.appendChild(crearCelda(equipo.usuario, "font-semibold text-gray-800"));
            fila.appendChild(crearCelda(equipo.numero_serie));
            fila.appendChild(crearCelda(equipo.anydesk));
            fila.appendChild(crearCelda(equipo.correo, "text-blue-600 italic"));
            fila.appendChild(crearCelda(equipo.fechaUltimoMantenimiento));
            fila.appendChild(crearCelda(equipo.fechaProximoMantenimiento, "text-red-600 font-bold"));
        }

        cuerpoTabla.appendChild(fila);
    });
};

const registrarElemento = () => {
    const formulario = document.getElementById("form-registro");
    if (!formulario) return;

    formulario.addEventListener("submit", async (evento) => {
        evento.preventDefault();

        const nuevoElemento = {
            cantidad: texto(document.getElementById("texto")?.value),
            modelo: texto(document.getElementById("modelo")?.value),
            marca: texto(document.getElementById("texto3")?.value),
            serial: texto(document.getElementById("texto4")?.value),
            placa: texto(document.getElementById("texto5")?.value),
            descripcion: texto(document.getElementById("texto2")?.value),
            fechaIngreso: texto(document.getElementById("fecha_ingreso")?.value),
            fechaBaja: texto(document.getElementById("fecha_de_baja")?.value)
        };

        try {
            const respuesta = await fetch(API_URL, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(nuevoElemento)
            });

            if (!respuesta.ok) throw new Error("Error en el registro");

            formulario.reset();
            await alerta("success", "Registro exitoso", "Elemento registrado en el servidor.");
        } catch (error) {
            await alerta("error", "Error", "No se pudo registrar el elemento.");
        }
    });
};

const registrarEquipo = () => {
    const formulario = document.getElementById("form-registro-equipo");
    if (!formulario) return;

    formulario.addEventListener("submit", async (evento) => {
        evento.preventDefault();

        const equipo = mapearEquipo({
            marca: document.getElementById("marca")?.value,
            modelo: document.getElementById("modelo")?.value,
            estado: document.getElementById("estado")?.value,
            nombre_equipo: document.getElementById("nombre_equipo")?.value,
            fecha_compra: document.getElementById("fecha_compra")?.value,
            placa: document.getElementById("placa")?.value,
            usuario: document.getElementById("usuario")?.value,
            correo: document.getElementById("correo")?.value,
            sistema_operativo: document.getElementById("sistema_operativo")?.value,
            numero_serie: document.getElementById("numero_serie")?.value,
            ubicacion: document.getElementById("ubicacion")?.value,
            anydesk: document.getElementById("anydesk")?.value,
            fecha_ultimo_mantenimiento: document.getElementById("fecha_ultimo_mantenimiento")?.value,
            fecha_proximo_mantenimiento: document.getElementById("fecha_proximo_mantenimiento")?.value
        });

        try {
            const respuesta = await fetch(API_EQUIPOS_URL, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(equipo)
            });

            if (!respuesta.ok) throw new Error("Error en el registro");

            // Solo mostramos el log si metimos datos en un arreglo válido
            mostrarLogImportacion([equipo]);
            formulario.reset();
            await alerta("success", "Registro exitoso", "Equipo registrado correctamente.");
        } catch (error) {
            await alerta("error", "Error", "No se pudo registrar el equipo.");
        }
    });
};

const mostrarLogImportacion = (equipos) => {
    // 🛡️ BLINDAJE ESTRELLA: Si no es un arreglo válido, paramos aquí suavemente
    if (!equipos || !Array.isArray(equipos)) {
        console.warn("El servidor no envió el detalle de equipos, pero la importación se realizó.");
        return; 
    }

    const logSection = document.getElementById("import-log");
    const logBody = document.getElementById("log-body");
    if (!logSection || !logBody) return;

    logSection.classList.remove("hidden");
    logBody.innerHTML = "";
    
    equipos.forEach(e => {
        const row = `
            <tr class="border-b hover:bg-gray-50">
                <td class="px-4 py-2 font-medium text-gray-900">${e.marca || "-"}</td>
                <td class="px-4 py-2">${e.modelo || "-"}</td>
                <td class="px-4 py-2">${e.numero_serie || "-"}</td>
                <td class="px-4 py-2"><span class="px-2 py-1 text-xs font-semibold rounded bg-green-100 text-green-800">Guardado</span></td>
            </tr>
        `;
        logBody.insertAdjacentHTML('beforeend', row);
    });
};

const inicializarExcelImport = () => {
    const inputExcel = document.getElementById("excel-import");
    if (!inputExcel) return;

    inputExcel.addEventListener("change", async (e) => {
        const file = e.target.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        try {
            const respuesta = await fetch('/api/equipos/upload', {
                method: "POST",
                body: formData
            });

            const resultado = await respuesta.json();

            if (!respuesta.ok) {
                throw new Error(resultado.error || 'Error al subir archivo');
            }

            // 🛡️ BLINDAJE 2: Revisamos qué nos mandó el backend
            if (resultado.equipos && Array.isArray(resultado.equipos)) {
                mostrarLogImportacion(resultado.equipos);
                await alerta("success", "Importación Exitosa", `${resultado.equipos.length} equipos guardados automáticamente.`);
            } else {
                await alerta("success", "Importación Exitosa", resultado.message || "Equipos guardados correctamente.");
            }
            
            inputExcel.value = "";

            // Actualizamos la tabla principal si está visible
            const esVistaConsulta = Boolean(document.getElementById("equipos-body"));
            const esVistaEliminar = Boolean(document.getElementById("tabla-equipos-eliminar"));
            
            if (esVistaConsulta || esVistaEliminar) {
                const equiposActualizados = await obtenerEquipos();
                renderizarEquipos(equiposActualizados, esVistaEliminar);
            }

        } catch (error) {
            console.error("Error Excel:", error);
            await alerta("error", "Error", error.message || "No se pudo procesar el archivo Excel. Verifica el formato.");
        }
    });
};

const inicializarConsulta = async () => {
    const formulario = document.getElementById("form");
    const entrada = document.getElementById("dato1");
    const esVistaEliminar = Boolean(document.getElementById("tabla-elementos-eliminar"));
    const esVistaConsulta = Boolean(document.getElementById("tabla-elementos"));

    if (!esVistaConsulta && !esVistaEliminar) return;

    const elementos = await obtenerElementos();
    renderizarElementos(elementos, esVistaEliminar);

    if (formulario && entrada) {
        formulario.addEventListener("submit", async (evento) => {
            evento.preventDefault();
            const filtrados = await obtenerElementos(entrada.value);
            renderizarElementos(filtrados, esVistaEliminar);
        });
    }
};

const inicializarConsultaEquipos = async () => {
    const formulario = document.getElementById("form-busqueda-equipo");
    const entrada = document.getElementById("busqueda-equipo");
    const esVistaEliminar = Boolean(document.getElementById("tabla-equipos-eliminar"));
    const esVistaConsulta = Boolean(document.getElementById("equipos-body"));

    if (!esVistaConsulta && !esVistaEliminar) return;

    const equipos = await obtenerEquipos();
    renderizarEquipos(equipos, esVistaEliminar);

    if (formulario && entrada) {
        formulario.addEventListener("submit", async (evento) => {
            evento.preventDefault();
            const filtrados = await obtenerEquipos(entrada.value);
            renderizarEquipos(filtrados, esVistaEliminar);
        });
    }
};

const eliminarElemento = async (id, nombre) => {
    const confirmar = await confirmarAlerta("Confirmar eliminación", `¿Deseas eliminar "${nombre}"?`);
    if (!confirmar) return;

    try {
        const respuesta = await fetch(`${API_URL}/${id}`, { method: "DELETE" });
        if (!respuesta.ok) throw new Error("Error al eliminar");
        await alerta("success", "Eliminado", "Elemento eliminado correctamente.");
        const elementos = await obtenerElementos();
        renderizarElementos(elementos, true);
    } catch (error) {
        await alerta("error", "Error", "No se pudo eliminar el elemento.");
    }
};

const eliminarEquipo = async (id, nombre) => {
    const confirmar = await confirmarAlerta("Confirmar eliminación", `¿Deseas eliminar el equipo "${nombre}"?`);
    if (!confirmar) return;

    try {
        const respuesta = await fetch(`${API_EQUIPOS_URL}/${id}`, { method: "DELETE" });
        if (!respuesta.ok) throw new Error("Error al eliminar");
        await alerta("success", "Eliminado", "Equipo eliminado correctamente.");
        const equipos = await obtenerEquipos();
        renderizarEquipos(equipos, true);
    } catch (error) {
        await alerta("error", "Error", "No se pudo eliminar el equipo.");
    }
};

const eliminarTodosEquipos = async () => {
    const confirmar = await confirmarAlerta("¡ADVERTENCIA!", "¿Estás seguro de que deseas eliminar TODOS los equipos? Esta acción no se puede deshacer.");
    if (!confirmar) return;

    try {
        const respuesta = await fetch(`${API_EQUIPOS_URL}/all`, { method: "DELETE" });
        if (!respuesta.ok) throw new Error("Error al eliminar todos");
        await alerta("success", "Eliminados", "Todos los equipos han sido eliminados.");
        renderizarEquipos([], true);
    } catch (error) {
        await alerta("error", "Error", "No se pudieron eliminar los equipos.");
    }
};

const iniciarExportaciones = () => {
    const btnExcelEle = document.getElementById("exportar-excel");
    const btnExcelEqu = document.getElementById("exportar-equipos-excel");

    if (btnExcelEle) {
        btnExcelEle.onclick = () => {
            window.location.href = '/api/export/elementos';
        };
    }

    if (btnExcelEqu) {
        btnExcelEqu.onclick = () => {
            window.location.href = '/api/export/equipos';
        };
    }
};

document.addEventListener("DOMContentLoaded", () => {
    registrarElemento();
    registrarEquipo();
    inicializarConsulta();
    inicializarConsultaEquipos();
    inicializarExcelImport();
    iniciarExportaciones();

    const btnEliminarTodo = document.getElementById("eliminar-todos-equipos");
    if (btnEliminarTodo) {
        btnEliminarTodo.onclick = eliminarTodosEquipos;
    }
});