// ==========================================
// My English Journey 2026 Spring
// File: js/app.js
// ==========================================

function searchStudent() {

    const studentId = document
        .getElementById("studentId")
        .value
        .trim();

    const result = document.getElementById("result");

    if (studentId === "") {

        result.className = "result show";

        result.innerHTML = `
            <div class="not-found">
                Please enter your Student ID.
            </div>
        `;

        return;
    }

    const student = students[studentId];

    if (!student) {

        result.className = "result show";

        result.innerHTML = `
            <div class="not-found">
                Student ID not found.
            </div>
        `;

        return;
    }

    let stars = "";

    if (student.presentation >= 90) {

        stars = "⭐⭐⭐⭐⭐";

    } else if (student.presentation >= 80) {

        stars = "⭐⭐⭐⭐";

    } else if (student.presentation >= 70) {

        stars = "⭐⭐⭐";

    } else if (student.presentation >= 60) {

        stars = "⭐⭐";

    } else {

        stars = "⭐";

    }

    result.className = "result show";

    result.innerHTML = `

        <h3>👤 ${student.name}</h3>

        <div class="item">

            <div class="item-title">

                🎤 Presentation

            </div>

            <div class="score">

                ${student.presentation} / 100

            </div>

            <div>

                ${stars}

            </div>

        </div>

        <div class="item">

            <div class="item-title">

                ⭐ Stickers

            </div>

            <div class="sticker">

                ${student.stickers} Stickers

            </div>

        </div>

        <div class="item">

            <div class="item-title">

                💌 Message from Yanyan

            </div>

            <div class="comment">

                ${student.comment.replace(/\n/g,"<br>")}

            </div>

        </div>

    `;

}

document
    .getElementById("studentId")
    .addEventListener("keydown", function(event){

        if(event.key === "Enter"){

            searchStudent();

        }

    });
