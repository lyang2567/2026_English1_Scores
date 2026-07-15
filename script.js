let students = [];
let excelLoaded = false;

async function loadExcel() {
    const status = document.getElementById("loadingStatus");
    const result = document.getElementById("result");

    try {
        const response = await fetch(
            "./data/English1_2026Spring_Template.xlsx?v=7"
        );

        if (!response.ok) {
            throw new Error(`Excel file error: ${response.status}`);
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        students = XLSX.utils.sheet_to_json(sheet, {
            defval: ""
        });

        excelLoaded = true;

        if (status) {
            status.textContent =
                `Student information loaded: ${students.length}
