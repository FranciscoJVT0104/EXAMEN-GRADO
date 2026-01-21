const fileInput = document.getElementById("fileInput");
const output = document.getElementById("output");
const exportBtn = document.getElementById("exportBtn");
const exportExcelBtn = document.getElementById("exportExcelBtn");

const clearBtn = document.getElementById("clearBtn");

clearBtn.addEventListener("click", () => {
    // Limpiar textarea
    output.value = "";

    // Limpiar input de archivo
    fileInput.value = "";

    // Limpiar nombre del archivo mostrado
    const fileNameSpan = document.getElementById("fileName");
    if (fileNameSpan) fileNameSpan.textContent = "";

    // Reiniciar variables globales
    textoOriginal = "";
    textoProcesadoTXT = "";
});


/* ===============================
   VARIABLES GLOBALES
================================ */
let textoOriginal = "";
let textoProcesadoTXT = "";

/* ===============================
   NORMALIZAR TEXTO WORD
================================ */
function normalizarLineas(texto) {
    return texto
        .split("\n")
        .map(l =>
            l
                .replace(/\t/g, " ")       // eliminar tabs
                .replace(/^•\s*/g, "")     // quitar viñetas
                .replace(/\s+/g, " ")      // normalizar espacios
                .trim()                     // quitar espacios al inicio y fin
        )
        .filter(l => l !== "");
}

/* ===============================
   UTILIDADES PREGUNTA / OPCIONES
================================ */
const esOpcion = linea => /^[A-E]\)/.test(linea);

const esPregunta = (linea, lineaAnterior) => {
    return (
        /^\d+\./.test(linea) ||                          // pregunta numerada
        (!esOpcion(linea) && (
            !lineaAnterior ||                             // primera línea
            (esOpcion(lineaAnterior) && lineaAnterior.startsWith("E)"))  // después de E) nueva pregunta
        ))
    );
};

/* ===============================
   LECTURA DEL WORD
================================ */
fileInput.addEventListener("click", function () {
    this.value = "";
});

fileInput.addEventListener("change", function () {
    const file = this.files[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = function (event) {
        mammoth.extractRawText({ arrayBuffer: event.target.result })
            .then(result => {
                textoOriginal = result.value;
                textoProcesadoTXT = procesarTextoTXT(textoOriginal);
                output.value = textoProcesadoTXT;
            })
            .catch(err => console.error(err));
    };

    reader.readAsArrayBuffer(file);
});

/* ===============================
   EXPORTAR TXT
================================ */
exportBtn.addEventListener("click", () => {
    if (!textoProcesadoTXT) return;

    const blob = new Blob([textoProcesadoTXT], { type: "text/plain" });
    const url = URL.createObjectURL(blob);

    const a = document.createElement("a");
    a.href = url;
    a.download = "examen.txt";
    a.click();

    URL.revokeObjectURL(url);
});

/* ===============================
   PROCESAR TEXTO PARA TXT
================================ */
function procesarTextoTXT(texto) {
    const lineas = normalizarLineas(texto);
    let resultado = [];

    let i = 0;
    let numeroPregunta = 1;

    while (i < lineas.length) {
        let linea = lineas[i];

        // ===== ALUMNOS =====
        if (/^ALUMNOS/i.test(linea)) {
            numeroPregunta = 1;

            // limpiar nombre en la misma línea
            let alumno = linea
                .replace(/ALUMNOS\s*:/i, "")
                .replace(/^[_\s]+/, "")
                .replace(/,/g, "")   // QUITAR COMAS
                .trim();
            if (alumno) resultado.push(alumno);

            i++;
            // limpiar nombres en líneas siguientes
            while (
                i < lineas.length &&
                !/^ALUMNOS/i.test(lineas[i]) &&
                !/^TEMA/i.test(lineas[i]) &&
                !esPregunta(lineas[i], lineas[i - 1])
            ) {
                let nombre = lineas[i]
                    .replace(/^[_\s]+/, "")
                    .replace(/,/g, "")   // QUITAR COMAS
                    .trim();
                if (nombre) resultado.push(nombre);
                i++;
            }
            continue;
        }

        // ===== TEMA =====
        if (/^TEMA/i.test(linea)) {
            resultado.push("TEMA: " + linea.replace(/TEMA\s*:/i, "").trim());
            i++;
            continue;
        }

        // ===== PREGUNTA =====
        if (esPregunta(linea, lineas[i - 1])) {
            let pregunta = linea.replace(/^\d+\.\s*/, "").replace(/^[A-E]\)\s*/, "");
            resultado.push(`${numeroPregunta}. ${pregunta}`);
            i++;

            while (
                i < lineas.length &&
                !esPregunta(lineas[i], lineas[i - 1]) &&
                !/^ALUMNOS/i.test(lineas[i])
            ) {
                if (esOpcion(lineas[i])) {
                    resultado.push(lineas[i]);
                }
                i++;
            }

            resultado.push("");
            numeroPregunta++;
            continue;
        }

        i++;
    }

    return resultado.join("\n");
}

/* ===============================
   EXPORTAR EXCEL
================================ */
exportExcelBtn.addEventListener("click", () => {
    if (!textoOriginal) return;

    const filas = procesarTextoExcel(textoOriginal);

    if (!filas.length) {
        alert("No se encontró contenido válido para Excel");
        return;
    }

    const ws = XLSX.utils.aoa_to_sheet(filas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "EXAMEN");

    XLSX.writeFile(wb, "examenes_grado.xlsx");
});

/* ===============================
   PROCESAR TEXTO PARA EXCEL
================================ */
function procesarTextoExcel(texto) {
    const lineas = normalizarLineas(texto);

    let filas = [];
    let alumnos = [];
    let preguntas = [];

    let numeroCategoria = 1;
    let numeroPregunta = 1;

    function cerrarBloque() {
        if (!alumnos.length || !preguntas.length) return;

        alumnos.forEach(alumno => {
            filas.push([
                `$CATEGORY: $course$/top/EXAMENES DE GRADO/${String(numeroCategoria++).padStart(2, "0")}. ${alumno.replace(/,/g, "")}`
            ]);

            filas.push([""]);
        });

        preguntas.forEach(p => {
            filas.push([`::e_${p.num}::${p.texto}{`]);

            p.opciones.forEach((op, i) => {
                // eliminar letras A)-E) y dejar solo texto
                let textoLimpio = op.replace(/^[A-E]\)\s*/, "");
                filas.push([(i === 0 ? "=" : "~") + textoLimpio]);
            });

            filas.push(["}"]);
            filas.push([""]);
        });

        alumnos = [];
        preguntas = [];
        numeroPregunta = 1;
    }

    let i = 0;
    while (i < lineas.length) {
        let linea = lineas[i];

        // ===== ALUMNOS =====
        if (/^ALUMNOS/i.test(linea)) {
            cerrarBloque();
            alumnos = [];
            preguntas = [];
            numeroPregunta = 1;

            // alumno en la misma línea
            let mismoRenglon = linea
                .replace(/ALUMNOS\s*:/i, "")
                .replace(/^[_\s]+/, "")   // eliminar guiones bajos y espacios al inicio
                .trim();
            if (mismoRenglon) alumnos.push(mismoRenglon);

            i++;
            // alumnos en líneas siguientes
            while (
                i < lineas.length &&
                !/^ALUMNOS/i.test(lineas[i]) &&
                !/^TEMA/i.test(lineas[i]) &&
                !esPregunta(lineas[i], lineas[i - 1])
            ) {
                let nombre = lineas[i]
                    .replace(/^[_\s]+/, "")
                    .replace(/,/g, "")   // QUITAR COMAS
                    .trim();
                if (nombre) alumnos.push(nombre);
                i++;
            }
            continue;
        }

        // ===== PREGUNTA =====
        if (esPregunta(linea, lineas[i - 1])) {
            let textoPregunta = linea.replace(/^\d+\.\s*/, "").replace(/^[A-E]\)\s*/, "");
            let opciones = [];

            i++;
            while (
                i < lineas.length &&
                !esPregunta(lineas[i], lineas[i - 1]) &&
                !/^ALUMNOS/i.test(lineas[i])
            ) {
                if (esOpcion(lineas[i])) {
                    opciones.push(lineas[i]);
                }
                i++;
            }

            preguntas.push({
                num: numeroPregunta++,
                texto: textoPregunta,
                opciones
            });

            continue;
        }

        i++;
    }

    cerrarBloque();
    return filas;
}
