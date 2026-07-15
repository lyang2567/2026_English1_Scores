let students = [];

// 显示查询结果
function showResult(html) {
    let result = document.getElementById("result");

    // 如果 HTML 中没有 result 区域，就自动创建
    if (!result) {
        result = document.createElement("div");
        result.id = "result";
        result.className = "result";

        const loadingStatus = document.getElementById("loadingStatus");

        if (loadingStatus) {
            loadingStatus.insertAdjacentElement("afterend", result);
        } else {
            document.body.appendChild(result);
        }
    }

    result.innerHTML = html;
}

// 读取 Excel
async function loadExcel() {
    const loadingStatus = document.getElementById("loadingStatus");

    try {
        const response = await fetch(
            "./data/English1_2026Spring_Template.xlsx?v=10"
        );

        if (!response.ok) {
            throw new Error(
                "Excel file could not be loaded. HTTP " + response.status
            );
        }

        const arrayBuffer = await response.arrayBuffer();

        const workbook = XLSX.read(arrayBuffer, {
            type: "array"
        });

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        students = XLSX.utils.sheet_to_json(worksheet, {
            defval: "",
            raw: false
        });

        console.log("Student data:", students);
        console.log("Excel headers:", Object.keys(students[0] || {}));

        if (loadingStatus) {
            loadingStatus.textContent =
                "Student information loaded: " +
                students.length +
                " students";
        }
    } catch (error) {
        console.error(error);

        if (loadingStatus) {
            loadingStatus.textContent =
                "Failed to load student information.";
        }

        showResult(`
            <div style="margin-top:20px; padding:20px; background:#fff0f0; border-radius:15px;">
                <strong>Excel loading error</strong>
                <p>${String(error.message)}</p>
            </div>
        `);
    }
}

// 查询学生
function searchStudent() {
    try {
        const input = document.getElementById("studentId");

        if (!input) {
            throw new Error(
                'The input field id="studentId" was not found.'
            );
        }

        const enteredId = String(input.value).trim();

        if (enteredId === "") {
            showResult(`
                <div style="margin-top:20px; padding:20px; background:#fff8dc; border-radius:15px;">
                    Please enter your Student ID.
                </div>
            `);
            return;
        }

        if (students.length === 0) {
            showResult(`
                <div style="margin-top:20px; padding:20px; background:#fff8dc; border-radius:15px;">
                    Student information has not finished loading.
                </div>
            `);
            return;
        }

        // 自动寻找包含 Student ID 的表头
        const headers = Object.keys(students[0]);

        const idHeader =
            headers.find(header =>
                String(header)
                    .replace(/\s/g, "")
                    .toLowerCase()
                    .includes("studentid")
            ) || headers[0];

        const student = students.find(row => {
            return String(row[idHeader]).trim() === enteredId;
        });

        if (!student) {
            showResult(`
                <div style="margin-top:20px; padding:22px; background:#fff0f0; border-radius:15px;">
                    <h2>❌ Student Not Found</h2>
                    <p>Student ID: ${enteredId}</p>
                    <p>Please check the number and try again.</p>
                </div>
            `);
            return;
        }

        // 自动匹配各列
        const findValue = keywords => {
            const key = Object.keys(student).find(header => {
                const normalized = String(header)
                    .replace(/\s/g, "")
                    .toLowerCase();

                return keywords.some(keyword =>
                    normalized.includes(keyword)
                );
            });

            return key ? student[key] : "";
        };

        const name = findValue([
            "name",
            "studentname",
            "姓名"
        ]);

        const presentation = findValue([
            "presentation",
            "grade",
            "发表成绩"
        ]);

        const stickers = findValue([
            "stickers",
            "sticker",
            "贴画",
            "貼画"
        ]);

        const comment = findValue([
            "comment",
            "message",
            "yanyan",
            "阳阳",
            "陽陽"
        ]);

        showResult(`
            <div style="
                margin-top:25px;
                padding:25px;
                background:white;
                border-radius:22px;
                box-shadow:0 8px 25px rgba(0,0,0,0.08);
            ">
                <h2>🌟 ${name || "Student"}</h2>

                <div style="margin-top:18px;">
                    <h3>🎤 Presentation</h3>
                    <p style="font-size:32px; font-weight:bold;">
                        ${presentation || "—"}
                    </p>
                </div>

                <div style="margin-top:18px;">
                    <h3>🏅 Stickers</h3>
                    <p style="font-size:26px;">
                        ${stickers || "0"} 🏅
                    </p>
                </div>

                <div style="
                    margin-top:20px;
                    padding:18px;
                    background:#fff8e8;
                    border-radius:16px;
                ">
                    <h3>💛 Message from Yanyan</h3>
                    <p>${comment || "Great job!"}</p>
                </div>
            </div>
        `);
    } catch (error) {
        console.error(error);

        showResult(`
            <div style="margin-top:20px; padding:20px; background:#fff0f0; border-radius:15px;">
                <strong>Search error</strong>
                <p>${String(error.message)}</p>
            </div>
        `);
    }
}

// 页面加载完成后运行
document.addEventListener("DOMContentLoaded", function () {
    const button = document.getElementById("searchButton");
    const input = document.getElementById("studentId");

    if (button) {
        button.onclick = searchStudent;
    }

    if (input) {
        input.addEventListener("keydown", function (event) {
            if (event.key === "Enter") {
                searchStudent();
            }
        });
    }

    loadExcel();
});
