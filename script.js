let students = [];

function showResult(message) {
    let result = document.getElementById("result");

    if (!result) {
        result = document.createElement("div");
        result.id = "result";

        const status = document.getElementById("loadingStatus");
        status.insertAdjacentElement("afterend", result);
    }

    result.innerHTML = message;
}

async function loadExcel() {
    const status = document.getElementById("loadingStatus");

    try {
        const response = await fetch(
            "./data/English1_2026Spring_Template.xlsx?v=20"
        );

        const buffer = await response.arrayBuffer();
        const workbook = XLSX.read(buffer, { type: "array" });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        students = XLSX.utils.sheet_to_json(sheet, {
            defval: "",
            raw: false
        });

        status.textContent =
            "Student information loaded: " +
            students.length +
            " students";
    } catch (error) {
        status.textContent = "Failed to load student information.";
        showResult("<p>Excel loading error: " + error.message + "</p>");
    }
}

function searchStudent() {
    const input = document.getElementById("studentId");
    const enteredId = String(input.value).trim();

    if (enteredId === "") {
        showResult("<p>Please enter your Student ID.</p>");
        return;
    }

    const headers = Object.keys(students[0] || {});

    const idHeader =
        headers.find(function (header) {
            return String(header)
                .replace(/\s/g, "")
                .toLowerCase()
                .includes("studentid");
        }) || headers[0];

    const student = students.find(function (row) {
        return String(row[idHeader]).trim() === enteredId;
    });

    if (!student) {
        showResult(
            "<div class='card'>" +
            "<h2>❌ Student Not Found</h2>" +
            "<p>Please check your Student ID.</p>" +
            "</div>"
        );
        return;
    }

    const values = Object.values(student);

    const name =
        student["Name"] ||
        student["Student Name"] ||
        values[1] ||
        "";

    const presentation =
        student["Presentation"] ||
        student["Grade"] ||
        values[2] ||
        "";

    const stickers =
        student["Stickers"] ||
        student["Sticker"] ||
        values[3] ||
        "";

    const comment =
        student["Yanyan's Comment"] ||
        student["Comment"] ||
        student["Message"] ||
        values[4] ||
        "";

    showResult(
        "<div class='card result-card'>" +
        "<h2>🌟 " + name + "</h2>" +
        "<h3>🎤 Presentation</h3>" +
        "<p class='grade'>" + presentation + "</p>" +
        "<h3>🏅 Stickers</h3>" +
        "<p>" + stickers + " 🏅</p>" +
        "<div class='comment-box'>" +
        "<h3>💛 Message from Yanyan</h3>" +
        "<p>" + comment + "</p>" +
        "</div>" +
        "</div>"
    );
}

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
