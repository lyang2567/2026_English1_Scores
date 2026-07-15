/* ==========================================
   My English Journey 2026 Spring
   File: script.js
========================================== */

let students = [];

/**
 * Safely escape Excel text before displaying it in HTML.
 */
function escapeHtml(value) {
    return String(value ?? "")
        .replaceAll("&", "&amp;")
        .replaceAll("<", "&lt;")
        .replaceAll(">", "&gt;")
        .replaceAll('"', "&quot;")
        .replaceAll("'", "&#039;");
}

/**
 * Show content in the result area.
 */
function showResult(message) {
    const result = document.getElementById("result");

    if (!result) {
        console.error("Result area was not found.");
        return;
    }

    result.innerHTML = message;
    result.classList.add("show");
}

/**
 * Find a column value using several possible header names.
 */
function getColumnValue(student, possibleHeaders, fallbackValue = "") {
    for (const header of possibleHeaders) {
        if (
            Object.prototype.hasOwnProperty.call(student, header) &&
            String(student[header]).trim() !== ""
        ) {
            return student[header];
        }
    }

    return fallbackValue;
}

/**
 * Return the matching CSS class for each grade.
 */
function getGradeClass(grade) {
    const normalizedGrade = String(grade)
        .trim()
        .toUpperCase()
        .replace(/\s/g, "");

    if (normalizedGrade === "A+") {
        return "grade-aplus";
    }

    if (normalizedGrade === "A") {
        return "grade-a";
    }

    if (normalizedGrade === "B+") {
        return "grade-bplus";
    }

    if (normalizedGrade === "B") {
        return "grade-b";
    }

    return "grade-default";
}

/**
 * Convert the sticker value into a medal display.
 *
 * Examples:
 * 5       → 🏅🏅🏅🏅🏅
 * "5"     → 🏅🏅🏅🏅🏅
 * "🏅🏅🏅" → 🏅🏅🏅
 */
function formatStickers(value) {
    const text = String(value ?? "").trim();

    if (text === "") {
        return "—";
    }

    const number = Number(text);

    if (Number.isInteger(number) && number >= 0 && number <= 30) {
        if (number === 0) {
            return "0";
        }

        return "🏅".repeat(number);
    }

    if (text.includes("🏅")) {
        return text;
    }

    return text + " 🏅";
}

/**
 * Load student information from Excel.
 */
async function loadExcel() {
    const status = document.getElementById("loadingStatus");
    const searchButton = document.getElementById("searchButton");

    searchButton.disabled = true;

    try {
        const response = await fetch(
            "./data/English1_2026Spring_Template.xlsx?v=30"
        );

        if (!response.ok) {
            throw new Error(
                "Excel file could not be loaded. HTTP status: " +
                response.status
            );
        }

        const buffer = await response.arrayBuffer();

        const workbook = XLSX.read(buffer, {
            type: "array"
        });

        const firstSheetName = workbook.SheetNames[0];

        if (!firstSheetName) {
            throw new Error("No worksheet was found in the Excel file.");
        }

        const sheet = workbook.Sheets[firstSheetName];

        students = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            raw: false
        });

        if (students.length === 0) {
            throw new Error("No student information was found.");
        }

        status.textContent =
            "Student information loaded: " +
            students.length +
            " students";

        searchButton.disabled = false;

    } catch (error) {
        console.error(error);

        status.textContent =
            "Failed to load student information.";

        showResult(
            "<div class='error-card'>" +
                "<h2>⚠️ Excel Loading Error</h2>" +
                "<p>" +
                    escapeHtml(error.message) +
                "</p>" +
            "</div>"
        );
    }
}

/**
 * Search for a student by Student ID.
 */
function searchStudent() {
    const input = document.getElementById("studentId");
    const enteredId = String(input.value).trim();

    if (enteredId === "") {
        showResult(
            "<div class='error-card'>" +
                "<h2>Student ID Required</h2>" +
                "<p>Please enter your Student ID.</p>" +
            "</div>"
        );

        input.focus();
        return;
    }

    if (students.length === 0) {
        showResult(
            "<div class='error-card'>" +
                "<h2>Student Information Is Not Ready</h2>" +
                "<p>Please refresh the page and try again.</p>" +
            "</div>"
        );

        return;
    }

    const headers = Object.keys(students[0]);

    const idHeader =
        headers.find(function (header) {
            const normalizedHeader = String(header)
                .replace(/\s/g, "")
                .replace(/_/g, "")
                .toLowerCase();

            return normalizedHeader.includes("studentid");
        }) ||
        headers.find(function (header) {
            return String(header)
                .toLowerCase()
                .includes("id");
        }) ||
        headers[0];

    const student = students.find(function (row) {
        return String(row[idHeader]).trim() === enteredId;
    });

    if (!student) {
        showResult(
            "<div class='error-card'>" +
                "<h2>❌ Student Not Found</h2>" +
                "<p>Please check your Student ID and try again.</p>" +
            "</div>"
        );

        return;
    }

    const values = Object.values(student);

    const name = getColumnValue(
        student,
        [
            "Name",
            "Student Name",
            "StudentName",
            "名前",
            "氏名"
        ],
        values[1] || ""
    );

    const presentation = getColumnValue(
        student,
        [
            "Presentation",
            "Presentation Grade",
            "Grade",
            "Result",
            "発表成績"
        ],
        values[2] || ""
    );

    const stickers = getColumnValue(
        student,
        [
            "Stickers",
            "Sticker",
            "Sticker Count",
            "貼画",
            "貼紙"
        ],
        values[3] || ""
    );

    const comment = getColumnValue(
        student,
        [
            "Yanyan's Comment",
            "Yanyan Comment",
            "Comment",
            "Message",
            "Teacher's Message",
            "Teacher Message",
            "陽陽のコメント"
        ],
        values[4] || ""
    );

    const gradeClass = getGradeClass(presentation);
    const stickerDisplay = formatStickers(stickers);

    showResult(
        "<article class='result-card'>" +

            "<h2 class='student-name'>" +
                "🌟 " +
                escapeHtml(name) +
            "</h2>" +

            "<section>" +
                "<h3 class='section-title'>" +
                    "🎤 Presentation Grade" +
                "</h3>" +

                "<p class='grade " +
                    gradeClass +
                "'>" +
                    escapeHtml(presentation || "—") +
                "</p>" +
            "</section>" +

            "<section class='sticker-area'>" +
                "<h3 class='section-title'>" +
                    "🏅 Stickers" +
                "</h3>" +

                "<p class='sticker-display'>" +
                    escapeHtml(stickerDisplay) +
                "</p>" +
            "</section>" +

            "<section class='comment-box'>" +
                "<h3>💛 Message from Yanyan</h3>" +

                "<p>" +
                    escapeHtml(comment || "Keep doing your best!") +
                "</p>" +
            "</section>" +

        "</article>"
    );
}

/**
 * Start the website after the page is ready.
 */
document.addEventListener("DOMContentLoaded", function () {
    const button = document.getElementById("searchButton");
    const input = document.getElementById("studentId");

    button.addEventListener("click", searchStudent);

    input.addEventListener("keydown", function (event) {
        if (event.key === "Enter") {
            searchStudent();
        }
    });

    loadExcel();
});
