const fileInput = document.getElementById("fileInput"),
    visibleInput = document.getElementById("visibleInput"),
    convertButton = document.getElementById("convertButton"),
    donwloadButton = document.getElementById("donwloadButton"),
    closeButton = document.getElementById("closeButton"),
    loader = document.getElementsByClassName("mdi-loading")[0];

let generatedWorkbook = null;
let lastBlobUrl = null;

function setalert(message, type) {
    toastr.options = { "closeButton": true, "progressBar": true, "positionClass": "toast-bottom-right", "preventDuplicates": false, "showDuration": "300", "hideDuration": "1000", "timeOut": "1000", "extendedTimeOut": "1000" };
    toastr[type](message);
}

function extractDataHeaderAndData(jsonData) {
    const header = [];
    const data = [];

    jsonData.forEach((item) => {
        const row = [];
        for (const key in item) {
            if (!header.includes(key)) { header.push(key); }
            let value = item[key];
            if (typeof value === "number" && value.toString().length > 15) { value = value.toLocaleString('fullwide', { useGrouping: false }); }
            if (typeof value === "string" && /^__EMPTY_\d*$/.test(value)) { value = ""; }
            if (value === null || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
                row.push(value);
            } else {
                row.push("");
            }
        }
        data.push(row);
    });

    return { header, data };
};

function getColumnWidths(header, data) {
    const colWidths = header.map((col, idx) => {
        let maxLen = String(col).length;
        for (let i = 0; i < data.length; i++) {
            const val = data[i][idx];
            if (val !== null && val !== undefined) {
                const len = String(val).length;
                if (len > maxLen) maxLen = len;
            }
        }
        return { width: Math.min(maxLen + 2, 50) };
    });
    return colWidths;
};

function normalizeJsonData(jsonData) {
    if (Array.isArray(jsonData)) {
        return jsonData;
    }

    if (typeof jsonData === "object" && jsonData !== null) {
        return [jsonData];
    }

    console.warn("Data JSON tidak valid. Data akan diabaikan.");
    return [];
};

function detectNumberColumns(data, header) {
    const numberCols = [];
    const textCols = [];
    const generalCols = [];
    for (let colIdx = 0; colIdx < header.length; colIdx++) {
        let isNumberCol = true;
        let mustBeText = false;
        let mustBeGeneral = false;
        for (let rowIdx = 0; rowIdx < data.length; rowIdx++) {
            const val = data[rowIdx][colIdx];
            if (val === null || val === undefined || val === "") continue;
            if (typeof val === "number") {
                if (val.toString().length > 15) mustBeGeneral = true;
                continue;
            }
            if (typeof val === "string" && /^[0-9]+$/.test(val)) {
                if (val.length > 1 && val[0] === "0") mustBeText = true;
                if (val.length > 15) mustBeGeneral = true;
                continue;
            }
            isNumberCol = false;
            break;
        }
        if (isNumberCol && mustBeText) {
            textCols.push(colIdx);
        } else if (isNumberCol && mustBeGeneral) {
            generalCols.push(colIdx);
        } else if (isNumberCol) {
            numberCols.push(colIdx);
        }
    }
    return { numberCols, textCols, generalCols };
}

async function exportToExcelFromJson(jsonData) {
    const workbook = new ExcelJS.Workbook();
    let sheetCounter = 1;

    function processWorksheet(worksheet, header, data) {
        worksheet.addRow(header);
        data.forEach((row) => worksheet.addRow(row));

        const { numberCols, textCols, generalCols } = detectNumberColumns(data, header);

        header.forEach((_, colIndex) => {
            const cell = worksheet.getRow(1).getCell(colIndex + 1);
            cell.font = { bold: true };
            cell.alignment = { horizontal: "left" };
        });

        for (let rowIdx = 2; rowIdx <= worksheet.rowCount; rowIdx++) {
            const row = worksheet.getRow(rowIdx);
            row.eachCell((cell, colIdx) => {
                cell.alignment = { horizontal: "left" };
                if (textCols.includes(colIdx - 1)) {
                    cell.numFmt = "@";
                    if (cell.value !== null && cell.value !== undefined) {
                        cell.value = cell.value.toString();
                    }
                }
                else if (generalCols.includes(colIdx - 1)) { cell.numFmt = "0"; }

                else if (numberCols.includes(colIdx - 1)) {
                    cell.numFmt = "0";
                    if (typeof cell.value === "string" && /^[0-9]+$/.test(cell.value) && cell.value.length <= 15) {
                        cell.value = Number(cell.value);
                    }
                }
            });
        }

        worksheet.columns = getColumnWidths(header, data);
    }

    for (const key in jsonData) {
        if (Array.isArray(jsonData[key])) {
            const normalizedData = normalizeJsonData(jsonData[key]);
            const { header, data } = extractDataHeaderAndData(normalizedData);
            const worksheet = workbook.addWorksheet(key.length > 31 ? `Sheet${sheetCounter++}` : key);
            processWorksheet(worksheet, header, data);
        }
    }

    if (Array.isArray(jsonData)) {
        const normalizedData = normalizeJsonData(jsonData);
        const { header, data } = extractDataHeaderAndData(normalizedData);
        const worksheet = workbook.addWorksheet(`Sheet${sheetCounter++}`);
        processWorksheet(worksheet, header, data);
    }

    if (typeof jsonData === "object" && !Array.isArray(jsonData)) {
        const normalizedData = normalizeJsonData(jsonData);
        const { header, data } = extractDataHeaderAndData(normalizedData);
        const worksheet = workbook.addWorksheet(`Sheet${sheetCounter++}`);
        processWorksheet(worksheet, header, data);
    }

    generatedWorkbook = workbook;
    setTimeout(() => {
        loader.classList.add("d-none"); donwloadButton.disabled = false; donwloadButton.classList.remove("d-none"); convertButton.disabled = true;
        setalert("File berhasil dikonversi menjadi Excel.", "success");
    }, 1000);
};

function fixBigIntJson(jsonString) {
    return jsonString.replace(
        /(:\s*|\[\s*|,\s*)(0\d{1,}|[1-9]\d{15,})(?=\s*[,\}\]])/g,
        function (match, p1, p2) { return p1 + '"' + p2 + '"'; }
    );
}

convertButton.addEventListener("click", async () => {
    loader.classList.remove("d-none");
    console.log("clicked, generatedWorkbook:", generatedWorkbook); // Debug
    const file = fileInput.files[0];
    if (!file) {
        setTimeout(() => { loader.classList.add("d-none"); setalert("Silakan masukkan file.", "error"); }, 500); return;
    }

    const reader = new FileReader();
    reader.onload = async (event) => {
        let jsonData;
        try {
            const fixedJson = fixBigIntJson(event.target.result);
            jsonData = JSON.parse(fixedJson);
        } catch (error) {
            setTimeout(() => { loader.classList.add("d-none"); setalert("File JSON tidak valid. Silakan periksa format file.", "error"); }, 500); return;
        }

        await exportToExcelFromJson(jsonData);
    };
    reader.onerror = () => {
        setTimeout(() => { loader.classList.add("d-none"); setalert("Terjadi kesalahan saat membaca file. Silakan coba lagi.", "error"); }, 500); return;
    };
    reader.readAsText(file);
});

donwloadButton.addEventListener("click", async () => {
    if (!generatedWorkbook) {
        setalert("Tidak ada file untuk diunduh. Silakan konversi file terlebih dahulu.", "error");
        return;
    }
    const buffer = await generatedWorkbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const link = document.createElement("a");
    lastBlobUrl = URL.createObjectURL(blob);
    link.href = lastBlobUrl;
    const now = new Date();
    const filename = `ConvertedFile_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}_${String(now.getHours()).padStart(2, '0')}${String(now.getMinutes()).padStart(2, '0')}${String(now.getSeconds()).padStart(2, '0')}.xlsx`;
    link.download = filename;
    link.click();

    setalert("File berhasil diunduh.", "success");
    console.log("clicked, generatedWorkbook:", generatedWorkbook); // Debug
});


fileInput.addEventListener("change", (event) => {
    generatedWorkbook = null; donwloadButton.disabled = true; donwloadButton.classList.add("d-none"); convertButton.disabled = false; visibleInput.value = event.target.files[0].name;
    console.log("clicked, generatedWorkbook:", generatedWorkbook); // Debug
});

closeButton.addEventListener("click", () => {
    fileInput.value = ""; generatedWorkbook = null; donwloadButton.disabled = true; donwloadButton.classList.add("d-none"); convertButton.disabled = true; visibleInput.value = "Masukkan File";
    console.log("clicked, generatedWorkbook:", generatedWorkbook); // Debug 
});

window.addEventListener("beforeunload", () => {
    fileInput.value = "";
    generatedWorkbook = null;
    donwloadButton.disabled = true; donwloadButton.classList.add("d-none"); convertButton.disabled = true; visibleInput.value = "Masukkan File";
});