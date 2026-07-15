let students = [];
let excelLoaded = false;

async function loadExcel() {
    const status = document.getElementById("loadingStatus");
    const result = document.getElementById("result");

    try {
        const response = await fetch(
            "./data/English1_2026Spring_Template.xlsx?v=3"
        );

        if (!response.ok) {
            throw new Error(`Excel file error: ${response.status}`);
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

        students = XLSX.utils.sheet_to_json(firstSheet, {
            defval: ""
        });

        excelLoaded = true;

        console.log("Students:", students);

        if (status) {
            status.textContent =
                `Student information loaded: ${students.length} students`;
        }

    } catch (error) {
        console.error(error);

        if (status) {
            status.textContent = "Failed to load student information.";
        }

        result.innerHTML = `
            <div class="card">
                <h2>⚠️ Excel Loading Error</h2>
                <p>${error.message}</p>
            </div>
        `;
    }
}

function searchStudent() {
    const result = document.getElementById("result");
    const inputElement = document.getElementById("studentId");

    if (!inputElement) {
        alert("studentId input not found");
        return;
    }

    const enteredId = inputElement.value.trim();

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
                <h2>✏️ Please enter your Student ID</h2>
            </div>
        `;
        return;
    }

   const student = students.find(row => {
    const id = Object.values(row)[0];
    return String(id).trim() === enteredId;
});

    if (!student) {
        result.innerHTML = `
            <div class="card">
                <h2>❌ Student Not Found</h2>
                <p>Entered ID: ${escapeHtml(enteredId)}</p>
            </div>
        `;
        return;
    }

    const name =
        student["Name"] ??
        student["Student Name"] ??
        student["姓名"] ??
        "";

    const presentation =
        student["Presentation"] ??
        student["Presentation Grade"] ??
        student["Grade"] ??
        "";

    const stickers =
        student["Stickers"] ??
        student["Sticker"] ??
        student["贴画个数"] ??
        "";

    const comment =
        student["Yanyan's Comment"] ??
        student["Yanyan Comment"] ??
        student["Comment"] ??
        student["阳阳的话"] ??
        "";

    result.innerHTML = `
        <div class="card result-card">
            <h2>🌟 ${escapeHtml(name)}</h2>

            <div class="score-item">
                <h3>🎤 Presentation</h3>
                <p class="grade">${escapeHtml(presentation)}</p>
            </div>

            <div class="score-item">
                <h3>🏅 Stickers</h3>
                <p>${escapeHtml(stickers)} 🏅</p>
            </div>

            <div class="comment-box">
                <h3>💛 A Message from Yanyan</h3>
                <p>${escapeHtml(comment)}</p>
            </div>
        </div>
    `;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

document.addEventListener("DOMContentLoaded", () => {
    const button = document.getElementById("searchButton");
    const input = document.getElementById("studentId");

    if (button) {
        button.addEventListener("click", searchStudent);
    }

    if (input) {
        input.addEventListener("keydown", event => {
            if (event.key === "Enter") {
                searchStudent();
            }
        });
    }

    loadExcel();
});
