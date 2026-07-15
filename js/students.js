let students = [];
let excelLoaded = false;

const excelPath =
    "data/English1_2026Spring_Template.xlsx";

async function loadExcel() {
    const status = document.getElementById("status");

    try {
        const response = await fetch(excelPath, {
            cache: "no-store"
        });

        if (!response.ok) {
            throw new Error(
                `Excel file could not be loaded: ${response.status}`
            );
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];

        students = XLSX.utils.sheet_to_json(sheet, {
            defval: ""
        });

        excelLoaded = true;
        status.textContent = "";
    } catch (error) {
        console.error(error);

        status.textContent =
            "Student information could not be loaded. Please try again later.";
    }
}

function searchStudent() {
    const input = document.getElementById("studentId");
    const result = document.getElementById("result");

    const enteredId = input.value.trim();

    if (!excelLoaded) {
        result.innerHTML = `
            <div class="card">
                <h2>⏳ Please wait</h2>
                <p>Student information is still loading.</p>
            </div>
        `;
        return;
    }

    if (enteredId === "") {
        result.innerHTML = `
            <div class="card">
                <h2>Please enter your Student ID.</h2>
            </div>
        `;
        return;
    }

    const student = students.find(item => {
        const excelId = String(item["Student ID"])
            .trim()
            .replace(/\.0$/, "");

        return excelId === enteredId;
    });

    if (!student) {
        result.innerHTML = `
            <div class="card">
                <h2>❌ Student Not Found</h2>
                <p>Please check your Student ID and try again.</p>
            </div>
        `;
        return;
    }

    const stickerCount =
        Number(student["Stickers"]) || 0;

    const medals =
        stickerCount > 0
            ? "🏅".repeat(stickerCount)
            : "No stickers yet";

    result.innerHTML = `
        <div class="card">

            <h2>${escapeHtml(student["Name"])}</h2>

            <p>
                <strong>Student ID</strong><br>
                ${escapeHtml(String(student["Student ID"]))}
            </p>

            <p>
                <strong>Presentation</strong><br>
                <span class="presentation-grade">
                    ${escapeHtml(student["Presentation"])}
                </span>
            </p>

            <p>
                <strong>Stickers</strong><br>
                <span class="stickers">
                    ${medals}
                </span>
            </p>

            <p>
                <strong>Message from Yanyan</strong>
            </p>

            <p class="message">
                ${formatMessage(student["Message from Yanyan"])}
            </p>

        </div>
    `;
}

function escapeHtml(value) {
    return String(value)
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

function formatMessage(value) {
    return escapeHtml(value)
        .replaceAll("\n", "<br>");
}

document.addEventListener("DOMContentLoaded", () => {
    const button =
        document.getElementById("searchButton");

    const input =
        document.getElementById("studentId");

    button.addEventListener("click", searchStudent);

    input.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            searchStudent();
        }
    });

    loadExcel();
});
