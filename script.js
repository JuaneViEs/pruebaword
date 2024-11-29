document.addEventListener('DOMContentLoaded', () => {
    // Eventos para el login
    document.getElementById('entrarBtn').addEventListener('click', login);
    document.getElementById('password').addEventListener('keypress', (event) => {
        if (event.key === 'Enter') {
            login();
        }
    });

    // Evento para mostrar el contenedor de cédula al seleccionar un archivo
    document.getElementById('fileInput').addEventListener('change', toggleCedulaContainer);

    // Eventos para buscar persona y regresar
    document.getElementById('buscarBtn').addEventListener('click', buscarPersona);
    document.getElementById('regresarBtn').addEventListener('click', regresar);

    // Evento para generar constancia
    document.getElementById('generateBtn').addEventListener('click', generarConstanciaWord);

    // Evento para regresar al formulario desde la constancia
    document.getElementById('regresarConstanciaBtn').addEventListener('click', regresarConstancia);
});

let personas = []; // Almacena los datos del CSV

// Función de Login
function login() {
    const passwordInput = document.getElementById('password').value;
    if (passwordInput === 'Incess2024') {
        toggleVisibility('formContainer', 'loginContainer', 'passwordError', false);
    } else {
        toggleVisibility('passwordError', false);
    }
}

// Función para mostrar y ocultar elementos
function toggleVisibility(...elements) {
    elements.forEach(element => {
        if (typeof element === 'string') {
            const elem = document.getElementById(element);
            if (elem) {
                elem.classList.toggle('hidden');
            }
        }
    });
}

// Función para mostrar el contenedor de cédula al seleccionar un archivo
function toggleCedulaContainer() {
    const file = document.getElementById('fileInput').files[0];
    const cedulaContainer = document.getElementById('cedulaContainer');
    if (file) {
        cedulaContainer.classList.remove('hidden');
    } else {
        cedulaContainer.classList.add('hidden');
    }
}

// Función de regreso al formulario
function regresar() {
    toggleVisibility('cedulaContainer', 'resultado', 'generateBtn');
    document.getElementById('fileInput').value = ''; // Reiniciar el input de archivo
}

// Función para buscar persona
function buscarPersona() {
    const fileInput = document.getElementById('fileInput');
    const cedulaInput = document.getElementById('cedulaInput').value.trim();

    if (!fileInput.files.length) {
        alert('Por favor, selecciona un archivo CSV.');
        return;
    }
    if (!cedulaInput) {
        alert('Por favor, ingrese una Cédula o ID.');
        return;
    }

    const reader = new FileReader();
    reader.onload = function(event) {
        const data = event.target.result;
        
        // Usamos PapaParse para leer el archivo CSV
        const parsed = Papa.parse(data, { header: true, skipEmptyLines: true });
        personas = parsed.data;

        // Encontramos a la persona por cédula
        const persona = personas.find(p => p['Cedula'] === cedulaInput);
        if (persona) {
            const nombre = persona['Nombre'] || 'Nombre no disponible';
            const apellido = persona['Apellido'] || 'Apellido no disponible';

            // Mostrar resultado y habilitar generar constancia
            document.getElementById('resultado').textContent = `Persona encontrada: ${nombre} ${apellido}`;
            toggleVisibility('resultado', 'generateBtn');
        } else {
            alert('No se encontró una persona con esa Cédula o ID.');
            toggleVisibility('resultado', 'generateBtn', true);
        }
    };
    reader.readAsText(fileInput.files[0]);
}

// Función para generar la constancia en Word
function generarConstanciaWord() {
    const cedulaInput = document.getElementById('cedulaInput').value.trim();
    const persona = personas.find(p => p['Cedula'] === cedulaInput);

    if (!persona) {
        alert("No se ha encontrado la persona. Por favor, verifica la cédula.");
        return;
    }

    // Cargar la plantilla .docx
    fetch('path/to/plantilla.docx') // Cambia esta ruta a la ubicación real de tu plantilla
        .then(response => response.arrayBuffer())
        .then(arrayBuffer => {
            const zip = new JSZip();
            zip.loadAsync(arrayBuffer).then(function (doc) {
                const templater = new Docxtemplater();
                templater.loadZip(doc);

                // Datos a insertar en el documento
                const today = new Date();
                const day = today.getDate();
                const monthNames = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
                const month = monthNames[today.getMonth()];
                const year = today.getFullYear();

                const fechaTexto = `los ${day} días del mes de ${month} de ${year}`;

                const unidadCurricular = persona['Denominacion de la Formacion'] || 'No disponible';

                const textData = {
                    nombre: persona['Nombre'].toLowerCase(),
                    apellido: persona['Apellido'].toLowerCase(),
                    cedula: persona['Cedula'],
                    unidadCurricular: unidadCurricular,
                    horas: persona['Horas'],
                    fechaInicio: persona['Fecha de Inicio'],
                    fechaCierre: persona['Fecha de Cierre'],
                    fechaTexto: fechaTexto
                };

                // Establecer los datos en el template
                templater.setData(textData);

                try {
                    templater.render();
                } catch (error) {
                    console.error(error);
                    alert('Error al generar el documento.');
                    return;
                }

                const out = templater.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
                saveAs(out, `constancia_${persona['Nombre']}_${persona['Apellido']}.docx`);
            });
        })
        .catch(err => console.error('Error al cargar la plantilla .docx:', err));
}

// Función para regresar a la pantalla de formulario
function regresarConstancia() {
    toggleVisibility('constanciaContainer');
    toggleVisibility('formContainer');
}

