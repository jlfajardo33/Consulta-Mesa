let invitados = [];

// Cargar el Excel
async function cargarInvitados() {

    try {

        const respuesta = await fetch("Invitados - Mesa.xlsx");

        const archivo = await respuesta.arrayBuffer();

        const libro = XLSX.read(archivo);

        const hoja = libro.Sheets[libro.SheetNames[0]];

        invitados = XLSX.utils.sheet_to_json(hoja);

        console.log("Invitados cargados:", invitados);

    } catch (error) {

        console.error("Error al cargar el Excel:", error);

    }

}

cargarInvitados();

const buscador = document.getElementById("buscar");
const resultados = document.getElementById("resultados");
const mesa = document.getElementById("mesa");

// Buscar mientras escribe
buscador.addEventListener("input", buscar);

function limpiarTexto(texto) {

    return texto
        .toString()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .toLowerCase();

}

function buscar() {

    const texto = limpiarTexto(buscador.value);

    resultados.innerHTML = "";
    mesa.innerHTML = "";

    if (texto.length < 3) return;

    const encontrados = invitados.filter(invitado => {

        return limpiarTexto(invitado.Invitados).includes(texto);

    });

    if (encontrados.length === 0) {

        resultados.innerHTML = `
            <div class="resultado">
                No encontramos ese nombre.
            </div>
        `;

        return;

    }

    encontrados.forEach(invitado => {

        const tarjeta = document.createElement("div");

        tarjeta.className = "resultado";

        tarjeta.innerHTML = `
            👤 ${invitado.Invitados}
        `;

        tarjeta.onclick = () => mostrarMesa(invitado);

        resultados.appendChild(tarjeta);

    });

}

function mostrarMesa(invitado){

    const modal = document.getElementById("modalMesa");
    const contenido = document.getElementById("contenidoModal");

    contenido.innerHTML = `
        <h3>José & Daniela</h3>

        <p class="nombre">
            ${invitado.Invitados}
        </p>

        <p class="tituloMesa">
            Tu mesa es
        </p>

        <div class="numeroMesa">
            ${invitado.Mesa}
        </div>

        <p class="gracias">
            Gracias por acompañarnos
            en este día tan especial.
        </p>

        <button onclick="cerrarModal()">
            Cerrar
        </button>
    `;

    modal.style.display = "flex";

    setTimeout(() => {
        cerrarModal();
        }, 8000);
}

function nuevaBusqueda() {

    buscador.value = "";

    resultados.innerHTML = "";

    mesa.innerHTML = "";

    buscador.focus();

}

function cerrarModal(){

    document.getElementById("modalMesa").style.display = "none";

    buscador.value = "";
    resultados.innerHTML = "";
    buscador.focus();
}

function abrirDistribucion(){

    const modal = document.getElementById("modalMesa");
    const contenido = document.getElementById("contenidoModal");

    contenido.innerHTML = `
        <h3>Distribución de Mesas</h3>

        <img
            src="img/distribucion.png"
            class="imagen-modal"
            alt="Distribución de mesas">

        <button onclick="cerrarModal()">
            Regresar
        </button>
    `;

    modal.style.display = "flex";
}

