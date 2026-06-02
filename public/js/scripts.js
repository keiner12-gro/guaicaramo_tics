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
    if (!limpio || limpio.toLowerCase() === 'n/a') return "";

    // 1. Si ya es YYYY-MM-DD
    if (/^\d{4}-\d{2}-\d{2}$/.test(limpio)) return limpio;

    // 2. Si es un serial de Excel
    const serial = Number(limpio);
    // Evitamos interpretar años (ej. 2024) como seriales de días
    if (Number.isFinite(serial) && serial >= 25569 && serial <= 100000) {
        try {
            const fecha = new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
            return fecha.toISOString().slice(0, 10);
            const year = fecha.getUTCFullYear();
            if (year < 1980) { // Asumiendo que fechas anteriores a 1980 son consideradas "mal"
                return "";
            }
        } catch (e) {
            return "";
        }
    }

    // 3. Evitar interpretar años sueltos como fechas (ej: 2016)
    if (/^\d{4}$/.test(limpio)) return "";

    // 4. Intentar parsear como fecha genérica
    const d = new Date(limpio);
    if (!isNaN(d.getTime())) {
        try {
            const year = d.getFullYear();
            if (year < 1980) { // Asumiendo que fechas anteriores a 1980 son consideradas "mal"
                return "";
            }
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
            showClass: {
                popup: 'animate__animated animate__fadeInUp animate__faster'
            },
            hideClass: {
                popup: 'animate__animated animate__fadeOutDown animate__faster'
            },
            confirmButtonColor: "var(--primary)",
            customClass: {
                popup: 'swal2-popup-custom',
                title: 'swal2-title-custom',
                confirmButton: 'swal2-confirm-custom'
            }
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
            confirmButtonColor: "var(--danger)",
            cancelButtonColor: "var(--primary)",
            confirmButtonText: "Si, eliminar",
            cancelButtonText: "Cancelar",
            customClass: {
                popup: 'swal2-popup-custom',
                title: 'swal2-title-custom',
                confirmButton: 'swal2-confirm-custom',
                cancelButton: 'swal2-confirm-custom'
            }
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

// Global state for single equipment view
let allEquipos = [];
let currentEquipoIndex = 0;

const renderizarElementos = (elementos) => {
    if (!elementos || !Array.isArray(elementos)) elementos = []; 

    const contenedor = document.getElementById("fichas-elementos");
    if (!contenedor) return;

    contenedor.textContent = "";

    if (elementos.length === 0) {
        contenedor.innerHTML = `<div class="col-span-full py-20 text-center text-gray-500 font-bold bg-white/50 rounded-2xl">No hay elementos registrados.</div>`;
        return;
    }

    elementos.forEach((elemento) => {
        const ficha = document.createElement("div");
        ficha.className = "ficha-card animate__animated animate__fadeInUp";
        
        ficha.innerHTML = `
            <div class="ficha-header">
                <div>
                    <h3 class="ficha-title">${elemento.marca || "Sin Marca"}</h3>
                    <p class="ficha-subtitle">${elemento.modelo || "Sin Modelo"}</p>
                </div>
                <span class="px-3 py-1 bg-primary/10 text-primary rounded-lg font-bold text-xs">CANT: ${elemento.cantidad || 0}</span>
            </div>
            <div class="ficha-grid">
                <div class="ficha-item">
                    <span class="ficha-label">Serial</span>
                    <span class="ficha-valor">${elemento.serial || "-"}</span>
                </div>
                <div class="ficha-item">
                    <span class="ficha-label">Placa</span>
                    <span class="ficha-valor">${elemento.placa || "-"}</span>
                </div>
                <div class="ficha-item">
                    <span class="ficha-label">Ingreso</span>
                    <span class="ficha-valor">${elemento.fechaIngreso || "-"}</span>
                </div>
                <div class="ficha-item">
                    <span class="ficha-label">Baja</span>
                    <span class="ficha-valor">${elemento.fechaBaja || "-"}</span>
                </div>
                <div class="ficha-item col-span-2">
                    <span class="ficha-label">Descripción</span>
                    <span class="ficha-valor italic text-gray-600">${elemento.descripcion || "-"}</span>
                </div>
            </div>
            <div class="ficha-acciones">
                <button onclick='abrirModal(${JSON.stringify(elemento).replace(/'/g, "&apos;")})' class="btn-ficha-edit">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
                    Editar
                </button>
                <button onclick="eliminarElemento(${elemento.id}, '${elemento.modelo || elemento.descripcion}')" class="btn-ficha-delete">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                </button>
            </div>
        `;
        contenedor.appendChild(ficha);
    });
};

const mostrarEquipoActual = () => {
    const contenedor = document.getElementById("equipo-detalle-container");
    const contador = document.getElementById("contador-equipos");
    const btnPrev = document.getElementById("prev-equipo");
    const btnNext = document.getElementById("next-equipo");
    const indexDisplay = document.getElementById("equipo-index-display");

    if (!contenedor || !contador || !btnPrev || !btnNext || !indexDisplay) return;

    contenedor.textContent = "";

    if (allEquipos.length === 0) {
        contenedor.innerHTML = `<div class="col-span-full py-20 text-center text-gray-500 font-bold bg-white/50 rounded-2xl">No hay equipos registrados.</div>`;
        contador.textContent = `Total: 0 equipos encontrados`;
        indexDisplay.textContent = "";
        btnPrev.disabled = true;
        btnNext.disabled = true;
        return;
    }

    // Ensure currentEquipoIndex is within bounds
    if (currentEquipoIndex < 0) currentEquipoIndex = 0;
    if (currentEquipoIndex >= allEquipos.length) currentEquipoIndex = allEquipos.length - 1;

    const equipo = allEquipos[currentEquipoIndex];
    const est = String(equipo.estado || "N/A").toLowerCase();
    const badgeClass = est.includes("nuevo") ? "badge-nuevo" : est.includes("usado") ? "badge-usado" : "badge-manto";

    contenedor.innerHTML = `
        <div class="equipo-detalle-wrapper animate__animated animate__fadeIn">
            <div class="card-equipo-moderna">
                <!-- HEADER: Marca, Nombre, Modelo, Estado, Fecha -->
                <div class="seccion-header">
                    <div class="flex justify-between items-start mb-4">
                        <div>
                            <span class="etiqueta-pro">Marca / Modelo</span>
                            <h2 class="text-2xl font-black text-bg-dark">${equipo.marca || "N/A"}</h2>
                            <p class="text-primary font-bold text-sm">${equipo.modelo || "Sin Modelo"}</p>
                        </div>
                        <span class="badge-estado ${badgeClass}">${equipo.estado || "N/A"}</span>
                    </div>
                    <div class="flex justify-between items-end">
                        <div class="campo-info">
                            <span class="etiqueta-pro">Nombre del Equipo</span>
                            <span class="valor-pro text-lg">${equipo.nombre_equipo || "-"}</span>
                        </div>
                        <div class="campo-info text-right">
                            <span class="etiqueta-pro">Fecha Compra</span>
                            <span class="valor-pro text-sm">${equipo.fechaCompra || "N/A"}</span>
                        </div>
                    </div>
                </div>

                <!-- ESPECIFICACIONES: Placa, Serial, SO, Ubicación -->
                <div class="seccion-cuerpo">
                    <div class="grid-campos">
                        <div class="campo-info">
                            <span class="etiqueta-pro">Placa TICS</span>
                            <span class="valor-pro valor-destacado text-red-600">${equipo.placa || "-"}</span>
                        </div>
                        <div class="campo-info">
                            <span class="etiqueta-pro">Número de Serial</span>
                            <span class="valor-pro valor-destacado">${equipo.numero_serie || "-"}</span>
                        </div>
                        <div class="campo-info">
                            <span class="etiqueta-pro">Sistema Operativo</span>
                            <span class="valor-pro">${equipo.sistema_operativo || "-"}</span>
                        </div>
                        <div class="campo-info">
                            <span class="etiqueta-pro">Ubicación / Sede</span>
                            <span class="valor-pro">${equipo.ubicacion || "-"}</span>
                        </div>
                    </div>
                </div>

                <!-- USUARIO: Nombre, Correo, Cédula -->
                <div class="seccion-cuerpo seccion-gris">
                    <div class="campo-info mb-4">
                        <span class="etiqueta-pro">Responsable Directo</span>
                        <span class="valor-pro text-lg">👤 ${equipo.usuario || "Sin Asignar"}</span>
                    </div>
                    <div class="grid-campos">
                        <div class="campo-info">
                            <span class="etiqueta-pro">Correo Corporativo</span>
                            <span class="valor-pro text-sm text-blue-600 italic">${equipo.correo || "-"}</span>
                        </div>
                        <div class="campo-info">
                            <span class="etiqueta-pro">Identificación (CC)</span>
                            <span class="valor-pro text-sm">${equipo.cedula || "N/A"}</span>
                        </div>
                    </div>
                </div>

                <!-- MANTENIMIENTO: Último y Próximo -->
                <div class="seccion-cuerpo">
                    <div class="grid-campos">
                        <div class="campo-info">
                            <span class="etiqueta-pro">Último Mantenimiento</span>
                            <span class="valor-pro">${equipo.fechaUltimoMantenimiento || "-"}</span>
                        </div>
                        <div class="campo-info">
                            <span class="etiqueta-pro">Próximo Mantenimiento</span>
                            <span class="valor-pro text-red-600 font-bold">${equipo.fechaProximoMantenimiento || "-"}</span>
                        </div>
                    </div>
                </div>

                <!-- ACCIONES -->
                <div class="p-6 bg-gray-50 flex justify-between gap-4">
                    <button onclick="eliminarEquipo(${equipo.id}, '${equipo.marca} ${equipo.modelo}')" class="btn-ficha-delete w-12 h-12 flex items-center justify-center">
                        <svg class="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16"></path></svg>
                    </button>
                <button onclick='abrirModalEquipo(${JSON.stringify(equipo).replace(/'/g, "&apos;")})' class="btn-ficha-edit">
                    <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>
                    Editar Equipo
                </button>
                </div>
            </div>
        </div>
    `;

    contador.textContent = `Total: ${allEquipos.length} equipos encontrados`;
    indexDisplay.innerHTML = `<span class="index-pills">${currentEquipoIndex + 1} / ${allEquipos.length}</span>`;

    btnPrev.disabled = currentEquipoIndex === 0;
    btnNext.disabled = currentEquipoIndex === allEquipos.length - 1;
};

const navegarEquipo = (direccion) => {
    currentEquipoIndex += direccion;
    mostrarEquipoActual();
};

// --- GESTIÓN DE MODALES Y ACTUALIZACIÓN ---

const abrirModal = (elemento) => {
    const modal = document.getElementById("modal-edicion");
    if (!modal) return;
    
    document.getElementById("edit-id").value = elemento.id;
    document.getElementById("edit-cantidad").value = elemento.cantidad;
    document.getElementById("edit-modelo").value = elemento.modelo;
    document.getElementById("edit-marca").value = elemento.marca;
    document.getElementById("edit-serial").value = elemento.serial;
    document.getElementById("edit-placa").value = elemento.placa;
    document.getElementById("edit-fechaIngreso").value = elemento.fechaIngreso;
    document.getElementById("edit-fechaBaja").value = elemento.fechaBaja;
    document.getElementById("edit-descripcion").value = elemento.descripcion;
    
    modal.classList.add("active");
};

const cerrarModal = () => {
    document.getElementById("modal-edicion")?.classList.remove("active");
};

const abrirModalEquipo = (equipo) => {
    const modal = document.getElementById("modal-edicion-equipo");
    if (!modal) return;
    
    document.getElementById("edit-equipo-id").value = equipo.id;
    document.getElementById("edit-equipo-marca").value = equipo.marca;
    document.getElementById("edit-equipo-modelo").value = equipo.modelo;
    document.getElementById("edit-equipo-estado").value = equipo.estado;
    document.getElementById("edit-equipo-nombre").value = equipo.nombre_equipo;
    document.getElementById("edit-equipo-placa").value = equipo.placa;
    document.getElementById("edit-equipo-so").value = equipo.sistema_operativo;
    document.getElementById("edit-equipo-ubicacion").value = equipo.ubicacion;
    document.getElementById("edit-equipo-usuario").value = equipo.usuario;
    document.getElementById("edit-equipo-serial").value = equipo.numero_serie;
    document.getElementById("edit-equipo-anydesk").value = equipo.anydesk;
    document.getElementById("edit-equipo-correo").value = equipo.correo;
    document.getElementById("edit-equipo-compra").value = equipo.fechaCompra;
    document.getElementById("edit-equipo-manto-ant").value = equipo.fechaUltimoMantenimiento;
    document.getElementById("edit-equipo-manto-prox").value = equipo.fechaProximoMantenimiento;
    
    modal.classList.add("active");
};

const cerrarModalEquipo = () => {
    document.getElementById("modal-edicion-equipo")?.classList.remove("active");
};

// --- LOGICA DE ACTUALIZACIÓN ---

const inicializarEdicion = () => {
    const formEle = document.getElementById("form-edicion");
    if (formEle) {
        formEle.onsubmit = async (e) => {
            e.preventDefault();
            const id = document.getElementById("edit-id").value;
            const data = {
                cantidad: document.getElementById("edit-cantidad").value,
                modelo: document.getElementById("edit-modelo").value,
                marca: document.getElementById("edit-marca").value,
                serial: document.getElementById("edit-serial").value,
                placa: document.getElementById("edit-placa").value,
                fechaIngreso: document.getElementById("edit-fechaIngreso").value,
                fechaBaja: document.getElementById("edit-fechaBaja").value,
                descripcion: document.getElementById("edit-descripcion").value
            };
            
            try {
                const res = await fetch(`${API_URL}/${id}`, {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(data)
                });
                if (!res.ok) throw new Error("Error al actualizar");
                cerrarModal();
                await alerta("success", "Actualizado", "Elemento actualizado correctamente.");
                inicializarConsulta();
            } catch (err) {
                alerta("error", "Error", err.message);
            }
        };
    }

    const formEqu = document.getElementById("form-edicion-equipo");
    if (formEqu) {
        formEqu.onsubmit = async (e) => {
            e.preventDefault();
            const id = document.getElementById("edit-equipo-id").value;
            const data = {
                marca: document.getElementById("edit-equipo-marca").value,
                modelo: document.getElementById("edit-equipo-modelo").value,
                estado: document.getElementById("edit-equipo-estado").value,
                nombre_equipo: document.getElementById("edit-equipo-nombre").value,
                placa: document.getElementById("edit-equipo-placa").value,
                sistema_operativo: document.getElementById("edit-equipo-so").value,
                ubicacion: document.getElementById("edit-equipo-ubicacion").value,
                usuario: document.getElementById("edit-equipo-usuario").value,
                numero_serie: document.getElementById("edit-equipo-serial").value,
                anydesk: document.getElementById("edit-equipo-anydesk").value,
                correo: document.getElementById("edit-equipo-correo").value,
                fecha_compra: document.getElementById("edit-equipo-compra").value,
                fecha_ultimo_mantenimiento: document.getElementById("edit-equipo-manto-ant").value,
                fecha_proximo_mantenimiento: document.getElementById("edit-equipo-manto-prox").value
            };
            
            try {
                const res = await fetch(`${API_EQUIPOS_URL}/${id}`, {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify(data)
                });
                if (!res.ok) throw new Error("Error al actualizar");
                cerrarModalEquipo();
                await alerta("success", "Actualizado", "Equipo actualizado correctamente.");
                inicializarConsultaEquipos();
            } catch (err) {
                alerta("error", "Error", err.message);
            }
        };
    }
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
    const contenedor = document.getElementById("fichas-elementos");

    if (!contenedor) return;

    const elementos = await obtenerElementos();
    renderizarElementos(elementos);

    if (formulario && entrada) {
        formulario.addEventListener("submit", async (evento) => {
            evento.preventDefault();
            const filtrados = await obtenerElementos(entrada.value);
            renderizarElementos(filtrados);
        });
    }
};

const inicializarConsultaEquipos = async (busqueda = "") => {
    const formulario = document.getElementById("form-busqueda-equipo");
    const entrada = document.getElementById("busqueda-equipo");
    const contenedor = document.getElementById("equipo-detalle-container");

    if (!contenedor) return;

    allEquipos = await obtenerEquipos(busqueda);
    currentEquipoIndex = 0; // Reset index on new search
    mostrarEquipoActual();

    if (formulario && entrada) {
        formulario.addEventListener("submit", async (evento) => {
            evento.preventDefault();
            inicializarConsultaEquipos(entrada.value); // Re-initialize with search
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
        inicializarConsulta(); // Re-render elements
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
        await alerta("success", "Eliminado", "Equipo eliminado correctamente."); // Re-render current equipment view
        inicializarConsultaEquipos(document.getElementById("busqueda-equipo")?.value || "");
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
        await alerta("success", "Eliminados", "Todos los equipos han sido eliminados."); // Re-render current equipment view
        inicializarConsultaEquipos(""); // Clear search and re-render
    } catch (error) {
        await alerta("error", "Error", "No se pudieron eliminar los equipos.");
    }
};

const iniciarExportaciones = () => {
    const btnExcelEle = document.getElementById("exportar-excel");
    const btnExcelEqu = document.getElementById("exportar-equipos-excel");

    if (btnExcelEle) {
        btnExcelEle.onclick = () => {
            window.location.href = `/api/export/elementos`;
        };
    }

    if (btnExcelEqu) {
        btnExcelEqu.onclick = () => {
            window.location.href = `/api/export/equipos`;
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
    inicializarEdicion();

    const btnEliminarTodo = document.getElementById("eliminar-todos-equipos");
    if (btnEliminarTodo) {
        btnEliminarTodo.onclick = eliminarTodosEquipos;
    }

    const btnPrevEquipo = document.getElementById("prev-equipo");
    if (btnPrevEquipo) {
        btnPrevEquipo.onclick = () => navegarEquipo(-1);
    }

    const btnNextEquipo = document.getElementById("next-equipo");
    if (btnNextEquipo) {
        btnNextEquipo.onclick = () => navegarEquipo(1);
    }
});