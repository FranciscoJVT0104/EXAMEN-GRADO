const fileInput = document.getElementById("fileInput");
const output = document.getElementById("output");
const exportBtn = document.getElementById("exportBtn");
const exportExcelBtn = document.getElementById("exportExcelBtn");
const clearBtn = document.getElementById("clearBtn");
const exportAlumnosBtn = document.getElementById("exportAlumnosBtn");

/* ===============================
   VARIABLES
================================ */
let textoOriginal = "";
let textoProcesadoTXT = "";

/* ===============================
   NORMALIZAR TEXTO
================================ */
function normalizarLineas(texto) {
    return texto
        .split("\n")
        .map(l =>
            l
                .replace(/\t/g, " ")
                .replace(/^•\s*/g, "")
                .replace(/\s+/g, " ")
                .trim()
        )
        .filter(l => l !== "");
}

/* ===============================
   DETECTORES
================================ */
const esOpcion = linea => /^[a-eA-E]\s*\)/.test(linea);

const esPregunta = (linea, anterior) => {
    if (/^\d+\./.test(linea)) return true;
    if (anterior && /^[eE]\s*\)/.test(anterior)) return true;
    return false;
};

const esEncabezadoAlumnos = linea =>
    /^(ALUMNOS?|ALUMNA?|ALUMNO\(A\)|ALUMNOS\(AS\)|APELLIDOS Y NOMBRES)\s*:/i.test(linea);

/* ===============================
   LECTURA ARCHIVOS
================================ */
fileInput.addEventListener("click", function () {
    this.value = "";
});

fileInput.addEventListener("change", function () {
    const file = this.files[0];
    if (!file) return;

    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "docx") leerWord(file);
    else if (ext === "txt") leerTXT(file);
    else alert("Formato no soportado");
});

function leerWord(file) {
    const reader = new FileReader();
    reader.onload = e => {
        mammoth.extractRawText({ arrayBuffer: e.target.result })
            .then(r => {
                textoOriginal = r.value;
                textoProcesadoTXT = procesarTextoTXT(textoOriginal);
                output.value = textoProcesadoTXT;
            });
    };
    reader.readAsArrayBuffer(file);
}

function leerTXT(file) {
    const reader = new FileReader();
    reader.onload = e => {
        textoOriginal = e.target.result;
        textoProcesadoTXT = procesarTextoTXT(textoOriginal);
        output.value = textoProcesadoTXT;
    };
    reader.readAsText(file, "UTF-8");
}

/* ===============================
   EXPORTAR TXT
================================ */
exportBtn.addEventListener("click", () => {
    if (!textoProcesadoTXT) return;
    const blob = new Blob([textoProcesadoTXT], { type: "text/plain" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "examen.txt";
    a.click();
});

/* ===============================
   PROCESAR TXT
================================ */
function procesarTextoTXT(texto) {
    const lineas = normalizarLineas(texto);
    let r = [], i = 0, n = 1;

    while (i < lineas.length) {
        let l = lineas[i];

        if (esEncabezadoAlumnos(l)) {
            n = 1;

            let nombre = l
                .replace(/^(ALUMNOS?|ALUMNA?|ALUMNO\(A\)|ALUMNOS\(AS\)|APELLIDOS Y NOMBRES)\s*:/i, "")
                .replace(/^[_\s]+/, "")
                .replace(/,/g, "")
                .trim();
            if (nombre) r.push(nombre);

            i++;
            while (
                i < lineas.length &&
                !/^TEMA/i.test(lineas[i]) &&
                !esPregunta(lineas[i], lineas[i - 1])
            ) {
                let x = lineas[i]
                    .replace(/^[_\s]+/, "")
                    .replace(/,/g, "")
                    .trim();
                if (x) r.push(x);
                i++;
            }
            continue;
        }

        if (/^TEMA/i.test(l)) {
            r.push("TEMA: " + l.replace(/TEMA\s*:/i, "").trim());
            i++;
            continue;
        }

        if (esPregunta(l, lineas[i - 1])) {
            r.push(`${n}. ${l.replace(/^\d+\.\s*/, "")}`);
            i++;
            while (i < lineas.length && !esPregunta(lineas[i], lineas[i - 1])) {
                if (esOpcion(lineas[i])) r.push(lineas[i]);
                i++;
            }
            r.push("");
            n++;
            continue;
        }

        i++;
    }

    return r.join("\n");
}

/* ===============================
   EXPORTAR EXCEL
================================ */
exportExcelBtn.addEventListener("click", () => {
    if (!textoOriginal) return;
    const filas = procesarTextoExcel(textoOriginal);
    if (!filas.length) return alert("No hay datos válidos");
    const ws = XLSX.utils.aoa_to_sheet(filas);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "EXAMEN");
    XLSX.writeFile(wb, "examenes_grado.xlsx");
});

/* ===============================
   PROCESAR EXCEL
================================ */
function procesarTextoExcel(texto) {
    const l = normalizarLineas(texto);
    let filas = [], alumnos = [], preguntas = [];
    let cat = 1, num = 1;

    function cerrar() {
        if (!alumnos.length || !preguntas.length) return;

        alumnos.forEach(a => {
            filas.push([
                `$CATEGORY: $course$/top/EXAMENES DE GRADO/${String(cat++).padStart(2, "0")}. ${a}`
            ]);
            filas.push([""]);
        });

        preguntas.forEach(p => {
            filas.push([`::e_${p.num}::${p.texto}{`]);
            p.opciones.forEach((o, i) => {
                let limpio = o.replace(/^[a-eA-E]\s*\)\s*/, "");
                filas.push([(i === 0 ? "=" : "~") + limpio]);
            });
            filas.push(["}"]);
            filas.push([""]);
        });

        alumnos = [];
        preguntas = [];
        num = 1;
    }

    let i = 0;
    while (i < l.length) {
        if (esEncabezadoAlumnos(l[i])) {
            cerrar();

            let nombre = l[i]
                .replace(/^(ALUMNOS?|ALUMNA?|ALUMNO\(A\)|ALUMNOS\(AS\)|APELLIDOS Y NOMBRES)\s*:/i, "")
                .replace(/^[_\s]+/, "")
                .replace(/,/g, "")
                .trim();
            if (nombre) alumnos.push(nombre);

            i++;
            while (
                i < l.length &&
                !/^TEMA/i.test(l[i]) &&
                !esPregunta(l[i], l[i - 1])
            ) {
                let x = l[i]
                    .replace(/^[_\s]+/, "")
                    .replace(/,/g, "")
                    .trim();
                if (x) alumnos.push(x);
                i++;
            }
            continue;
        }

        if (esPregunta(l[i], l[i - 1])) {
            let textoPregunta = l[i].replace(/^\d+\.\s*/, "");
            let ops = [];
            i++;
            while (i < l.length && !esPregunta(l[i], l[i - 1])) {
                if (esOpcion(l[i])) ops.push(l[i]);
                i++;
            }
            preguntas.push({ num: num++, texto: textoPregunta, opciones: ops });
            continue;
        }

        i++;
    }

    cerrar();
    return filas;
}

/* ===============================
   LIMPIAR
================================ */
clearBtn.addEventListener("click", () => {
    output.value = "";
    fileInput.value = "";
    document.getElementById("fileName").textContent = "";
    textoOriginal = "";
    textoProcesadoTXT = "";
});

/* ===============================
   EXPORTAR LISTA DE ALUMNOS
================================ */
function limpiarNombreAlumno(texto) {

    let limpio = texto
        .replace(/^(ALUMNOS?|ALUMNAS?|ALUMNOS\(AS\)|ALUMNO\(A\)|ALUMNA\(O\)|APELLIDOS Y NOMBRES)\s*:/i, "")
        .replace(/^[_•\-\s]+/, "")
        .replace(/,/g, "")
        .trim();

    // evitar encabezados sin nombre
    if (
        /^(ALUMNOS?|ALUMNAS?|ALUMNOS\(AS\)|ALUMNO\(A\)|ALUMNA\(O\)|APELLIDOS Y NOMBRES)$/i.test(limpio)
    ) {
        return "";
    }

    return limpio;
}

exportAlumnosBtn.addEventListener("click", () => {

    if (!textoOriginal) {
        alert("Primero carga un archivo.");
        return;
    }

    const lineas = normalizarLineas(textoOriginal);

    let alumnos = [];

    for (let i = 0; i < lineas.length; i++) {

        if (esEncabezadoAlumnos(lineas[i])) {

            // alumno en la misma línea
            let nombre = limpiarNombreAlumno(lineas[i]);

            if (nombre) alumnos.push(nombre);

            i++;

            // alumnos en líneas siguientes
            while (
                i < lineas.length &&
                !esEncabezadoAlumnos(lineas[i]) &&
                !/^TEMA/i.test(lineas[i]) &&
                !esPregunta(lineas[i], lineas[i - 1])
            ) {

                let alumno = limpiarNombreAlumno(lineas[i]);

                if (alumno) alumnos.push(alumno);

                i++;
            }
        }
    }

    if (!alumnos.length) {
        alert("No se encontraron alumnos.");
        return;
    }

    // eliminar duplicados por seguridad
    alumnos = [...new Set(alumnos)];

    // enumerar
    const textoLista = alumnos
        .map((a, index) => `${index + 1}. ${a}`)
        .join("\n");

    // descargar
    const blob = new Blob([textoLista], { type: "text/plain;charset=utf-8;" });

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = "lista_alumnos.txt";
    link.click();

    URL.revokeObjectURL(link.href);

});
