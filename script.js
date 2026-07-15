let students = [];
let excelLoaded = false;

// 页面打开后读取 Excel
async function loadExcel() {
    const result = document.getElementById("result");

    try {
        const response = await fetch(
            "./data/English1_2026Spring_Template.xlsx"
        );

        if (!response.ok) {
            throw new Error(
                `Excel file could not be loaded: ${response.status}`
            );
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        students = XLSX.utils.sheet_to_json(sheet, {
            defval: ""
        });

        excelLoaded = true;

        console.log("Excel loaded successfully.");
        console.log("Student data:", students);

    } catch (error) {
        excelLoaded = false;

        console.error("Excel loading error:", error);

        result.innerHTML = `
            <div class="card">
                <h2>⚠️ Data Loading Error</h2>
                <p>The student data could not be loaded.</p>
                <p>Please check the Excel filename and folder.</p>
            </div>
        `;
    }
}

// 查询学生
function searchStudent() {
    const result = document.getElementById("result");

    if (!excelLoaded) {
        result.innerHTML = `
            <div class="card">
                <h2>⏳ Please wait</h2>
                <p>The student data is still loading.</p>
            </div>
        `;
        return;
    }

    const input = document
        .getElementById("studentId")
        .value
        .trim();

    if (input === "") {
        result.innerHTML = `
            <div class="card">
                <h2>✏️ Enter Your Student ID</h2>
            </div>
        `;
        return;
    }

    const student = students.find(item => {
        const studentId =
            item["Student ID"] ??
            item["StudentID"] ??
            item["student ID"] ??
            item["学号"];

        return String(studentId).trim() === input;
    });

    if (!student) {
        result.innerHTML = `
            <div class="card">
                <h2>❌ Student Not Found</h2>
                <p>Please check your Student ID.</p>
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

// 防止 Excel 内容被当作 HTML 执行
function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

// 按 Enter 也可以搜索
document.addEventListener("DOMContentLoaded", () => {
    const input = document.getElementById("studentId");

    input.addEventListener("keydown", event => {
        if (event.key === "Enter") {
            searchStudent();
        }
    });

    loadExcel();
});
