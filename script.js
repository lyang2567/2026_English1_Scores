let students = [];

/* =========================
   读取学生 Excel 文件
========================= */

async function loadExcel() {
    const loadingStatus = document.getElementById("loadingStatus");
    const searchButton = document.getElementById("searchButton");

    try {
        loadingStatus.textContent = "Loading student information...";
        searchButton.disabled = true;

        const response = await fetch(
            "data/English1_2026Spring_Template.xlsx"
        );

        if (!response.ok) {
            throw new Error("Excel file could not be loaded.");
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

        loadingStatus.textContent = "";
        searchButton.disabled = false;

        console.log("Student information loaded:", students);

    } catch (error) {
        console.error(error);

        loadingStatus.textContent =
            "Student information could not be loaded.";

        searchButton.disabled = true;
    }
}


/* =========================
   查询学生
========================= */

function searchStudent() {
    const input = document.getElementById("studentId");
    const result = document.getElementById("result");

    const studentId = input.value.trim();

    if (studentId === "") {
        result.innerHTML = `
            <div class="card error-card">
                <h2>🌼 Please enter your Student ID.</h2>
            </div>
        `;
        return;
    }

    const student = students.find(item => {
        return String(item["Student ID"]).trim() === studentId;
    });

    if (!student) {
        result.innerHTML = `
            <div class="card error-card">
                <h2>Student Not Found</h2>

                <p>
                    Please check your Student ID and try again.
                </p>
            </div>
        `;
        return;
    }

    displayStudentResult(student);
}


/* =========================
   显示学生成绩卡
========================= */

function displayStudentResult(student) {
    const result = document.getElementById("result");

    const studentName =
        student["Name"] ||
        student["Student Name"] ||
        student["名前"] ||
        "Student";

    const presentationGrade =
        student["Presentation"] ||
        student["Presentation Grade"] ||
        student["Grade"] ||
        "-";

    const stickerNumber =
        student["Stickers"] ||
        student["Sticker"] ||
        student["Sticker Count"] ||
        "0";

    const message =
        student["Message"] ||
        student["Yanyan's Message"] ||
        student["Comment"] ||
        "";

    const gradeClass = getGradeClass(presentationGrade);

    result.innerHTML = `
        <div class="card result-card">

            <div class="student-name-area">
                <p class="welcome-text">
                    Welcome back,
                </p>

                <h2 class="student-name">
                    ${escapeHTML(studentName)} ✨
                </h2>
            </div>

            <div class="result-section">

                <p class="result-label">
                    Presentation
                </p>

                <div class="grade-circle ${gradeClass}">
                    ${escapeHTML(presentationGrade)}
                </div>

            </div>

            <div class="result-section">

                <p class="result-label">
                    Stickers
                </p>

                <div class="sticker-result">
                    🏅 × ${escapeHTML(stickerNumber)}
                </div>

            </div>

            <div class="message-section">

                <h3>
                    💌 A Message from Yanyan
                </h3>

                <p class="teacher-message">
                    ${formatMessage(message)}
                </p>

            </div>

            <div class="yanyan-footer">

                <p class="thank-you-message">
                    Thank you for being part of
                    My English Journey. ✨
                </p>

                <div class="yanyan-signature-area">

                    <img
                        src="images/yanyan.png"
                        alt="Yanyan"
                        class="yanyan-avatar"
                    >

                    <div class="yanyan-signature">

                        <span class="from-text">
                            With love,
                        </span>

                        <span class="signature-name">
                            Yanyan
                        </span>

                    </div>

                </div>

            </div>

        </div>
    `;
}


/* =========================
   成绩对应的颜色
========================= */

function getGradeClass(grade) {
    const normalizedGrade = String(grade)
        .trim()
        .toUpperCase();

    if (normalizedGrade === "A+") {
        return "grade-a-plus";
    }

    if (normalizedGrade === "A") {
        return "grade-a";
    }

    if (normalizedGrade === "B+") {
        return "grade-b-plus";
    }

    if (normalizedGrade === "B") {
        return "grade-b";
    }

    return "grade-default";
}


/* =========================
   安全显示 Excel 文字
========================= */

function escapeHTML(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}


/* =========================
   保留评语换行
========================= */

function formatMessage(message) {
    const safeMessage = escapeHTML(message);

    if (safeMessage.trim() === "") {
        return "Thank you for your wonderful work this semester!";
    }

    return safeMessage.replace(/\r?\n/g, "<br>");
}


/* =========================
   页面加载完成后的操作
========================= */

document.addEventListener("DOMContentLoaded", () => {
    const searchButton =
        document.getElementById("searchButton");

    const studentIdInput =
        document.getElementById("studentId");

    searchButton.addEventListener(
        "click",
        searchStudent
    );

    studentIdInput.addEventListener(
        "keydown",
        event => {
            if (event.key === "Enter") {
                searchStudent();
            }
        }
    );

    loadExcel();
});
