// ==============================
// BNCP Smart Welfare Portal
// app.js
// ==============================

// Cloudflare Worker URL
const API_URL = "https://broken-poetry-c6d2.chogak1449.workers.dev/";

// ==============================
// 요소 가져오기
// ==============================
const chatBox = document.getElementById("chatBox");
const userInput = document.getElementById("userInput");

// ==============================
// 채팅 메시지 출력
// ==============================
function addMessage(sender, message) {

    const div = document.createElement("div");

    div.className = sender === "user" ? "user" : "ai";

    div.innerHTML = `
        <strong>${sender === "user" ? "🙋 나" : "🤖 AI"}</strong><br>
        ${message}
    `;

    chatBox.appendChild(div);
    chatBox.scrollTop = chatBox.scrollHeight;
}

// ==============================
// AI 호출
// ==============================
async function sendMessage() {

    const message = userInput.value.trim();

    if (message === "") return;

    addMessage("user", message);

    userInput.value = "";

    try {

        const response = await fetch(API_URL, {

            method: "POST",

            headers: {
                "Content-Type": "application/json"
            },

            body: JSON.stringify({
                message: message
            })

        });

        if (!response.ok) {
            throw new Error("API Error");
        }

        const data = await response.json();

        addMessage("ai", data.reply);

    } catch (error) {

        console.error(error);

        addMessage(
            "ai",
            "죄송합니다.<br>현재 AI 서버와 연결할 수 없습니다."
        );

    }

}

// ==============================
// Enter 키 전송
// ==============================
if (userInput) {

    userInput.addEventListener("keydown", function (e) {

        if (e.key === "Enter") {

            sendMessage();

        }

    });

}

// ==============================
// 실시간 날씨
// ==============================
async function loadWeather() {

    const weather = document.getElementById("weather");

    if (!weather) return;

    try {

        // Baghdad
        const latitude = 33.3152;
        const longitude = 44.3661;

        const url =
            `https://api.open-meteo.com/v1/forecast?latitude=${latitude}&longitude=${longitude}&current=temperature_2m,weather_code`;

        const response = await fetch(url);

        const data = await response.json();

        const temp = Math.round(data.current.temperature_2m);

        const code = data.current.weather_code;

        let icon = "☀️";
        let text = "맑음";

        if (code === 0) {
            icon = "☀️";
            text = "맑음";
        }
        else if (code === 1 || code === 2) {
            icon = "⛅";
            text = "부분적으로 흐림";
        }
        else if (code === 3) {
            icon = "☁️";
            text = "흐림";
        }
        else if (code >= 45 && code <= 48) {
            icon = "🌫️";
            text = "안개";
        }
        else if (code >= 51 && code <= 67) {
            icon = "🌦️";
            text = "비";
        }
        else if (code >= 71 && code <= 77) {
            icon = "❄️";
            text = "눈";
        }
        else if (code >= 80 && code <= 82) {
            icon = "🌧️";
            text = "소나기";
        }
        else if (code >= 95) {
            icon = "⛈️";
            text = "천둥번개";
        }

        weather.innerHTML = `
            <div style="font-size:18px;">
                ${icon} <strong>Baghdad</strong>
            </div>
            <div style="font-size:26px;font-weight:bold;margin-top:5px;">
                ${temp}℃
            </div>
            <div style="font-size:14px;color:#666;">
                ${text}
            </div>
        `;

    } catch (error) {

        console.error(error);

        weather.innerHTML = `
            ❌<br>
            날씨 정보를 불러올 수 없습니다.
        `;

    }

}

// ==============================
// 페이지 시작
// ==============================
window.onload = function () {

    loadWeather();

    // 10분마다 자동 갱신
    setInterval(loadWeather, 600000);

    addMessage(
        "ai",
        "안녕하세요 😊<br><br>BNCP AI Assistant입니다.<br>궁금한 사항을 입력해 주세요."
    );

};
