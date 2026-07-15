let students = [];

// 读取 Excel
async function loadExcel() {
    const response = await fetch("data/English1_2026Spring_Template.xlsx");
    const arrayBuffer = await response.arrayBuffer();

    const workbook = XLSX.read(arrayBuffer, {
        type: "array"
    });

    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    students = XLSX.utils.sheet_to_json(sheet);
}

loadExcel();

// 查询学生
function searchStudent() {

    const id = document.getElementById("studentId").value.trim();

    const student = students.find(s =>
        String(s["Student ID"]) === id
    );

    const result = document.getElementById("result");

    if (!student) {
        result.innerHTML = `
        <div class="card">
            <h2>❌ Student Not Found</h2>
        </div>`;
        return;
    }

    result.innerHTML = `
    <div class="card">
        <h2>${student.Name}</h2>

        <p><strong>Student ID：</strong>${student["Student ID"]}</p>

        <p><strong>Presentation：</strong>
        ${student.Presentation}</p>

        <p style="font-size:40px">
        ${student.Stickers}
        </p>

        <p>${student["Message from Yanyan"]}</p>
    </div>
    `;
}
