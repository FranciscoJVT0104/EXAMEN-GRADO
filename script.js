const fileInput = document.getElementById("fileInput");
const output = document.getElementById("output");
const exportBtn = document.getElementById("exportBtn");
const exportExcelBtn = document.getElementById("exportExcelBtn");

/* ===============================
   VARIABLES GLOBALES
================================ */
let textoOriginal = "";        // TEXTO CRUDO DEL WORD
let textoProcesadoTXT = "";   // TEXTO LIMPIO PARA TXT

/* ===============================
   LECTURA DEL WORD
================================ */
// 🔹 Limpia el input al hacer click para permitir mismo nombre
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
    const lineas = texto.split("\n").map(l => l.trim()).filter(l => l !== "");
    let resultado = [];

    let i = 0;
    let numeroPregunta = 1;

    while (i < lineas.length) {
        let linea = lineas[i];

        // ===== ALUMNOS =====
        if (/^ALUMNOS/i.test(linea)) {
            numeroPregunta = 1;

            let alumno = linea
                .replace(/ALUMNOS\s*:/i, "")
                .replace(/_/g, "")
                .trim();

            if (alumno) resultado.push(alumno);

            i++;

            while (
                i < lineas.length &&
                !/^ALUMNOS/i.test(lineas[i]) &&
                !lineas[i].startsWith("¿")
            ) {
                let posible = lineas[i].replace(/_/g, "").trim();
                if (posible) resultado.push(posible);
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
        if (linea.startsWith("¿")) {
            resultado.push(`${numeroPregunta}. ${linea}`);
            i++;

            const letras = ["A", "B", "C", "D", "E", "F"];
            let idx = 0;

            while (
                i < lineas.length &&
                !lineas[i].startsWith("¿") &&
                !/^ALUMNOS/i.test(lineas[i])
            ) {
                resultado.push(`${letras[idx++]}) ${lineas[i]}`);
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
    const lineas = texto.split("\n").map(l => l.trim()).filter(l => l !== "");

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
                filas.push([(i === 0 ? "=" : "~") + op]);
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

        // ===== NUEVO BLOQUE =====
        if (/^ALUMNOS/i.test(linea)) {
            cerrarBloque();
            alumnos = [];
            preguntas = [];
            numeroPregunta = 1;

            // 🔹 alumno en la MISMA línea
            let mismoRenglon = linea
                .replace(/ALUMNOS\s*:/i, "")
                .replace(/_/g, "")
                .trim();

            if (mismoRenglon) alumnos.push(mismoRenglon);

            i++;

            // 🔹 alumnos en líneas siguientes
            while (
                i < lineas.length &&
                !/^ALUMNOS/i.test(lineas[i]) &&
                !lineas[i].startsWith("¿") &&
                !/^TEMA/i.test(lineas[i])
            ) {
                let nombre = lineas[i].replace(/_/g, "").trim();
                if (nombre) alumnos.push(nombre);
                i++;
            }

            continue;
        }

        // ===== PREGUNTA =====
        if (linea.startsWith("¿")) {
            let textoPregunta = linea;
            let opciones = [];

            i++;

            while (
                i < lineas.length &&
                !lineas[i].startsWith("¿") &&
                !/^ALUMNOS/i.test(lineas[i])
            ) {
                opciones.push(lineas[i]);
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
