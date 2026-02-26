// 현재 활성 데이터 (기본값: data.js의 목데이터)
let activeThreats = THREATS;
let activeDetections = DETECTIONS;

document.addEventListener("DOMContentLoaded", () => {
  renderThreats();
  renderDetections();
  setupChatbot();
  setupSidebar();
  setupExcelUpload();
});

// ─── 엑셀 업로드 ────────────────────────────────────────────────────
function setupExcelUpload() {
  const input = document.getElementById("excel-input");
  const status = document.getElementById("excel-status");

  input.addEventListener("change", async (e) => {
    const file = e.target.files[0];
    if (!file) return;

    status.textContent = "⏳ 파싱 중...";
    status.className = "excel-status loading";

    try {
      const { threats, detections } = await parseExcel(file);

      if (threats.length === 0 && detections.length === 0) {
        status.textContent = "⚠️ 데이터를 찾을 수 없습니다. 시트명을 확인해주세요.";
        status.className = "excel-status error";
        return;
      }

      activeThreats = threats.length > 0 ? threats : THREATS;
      activeDetections = detections.length > 0 ? detections : DETECTIONS;

      renderThreats();
      renderDetections();

      status.textContent = `✅ ${file.name} — 뉴스 ${activeThreats.length}건 / 탐지 ${activeDetections.length}건 로드 완료`;
      status.className = "excel-status success";

      appendChatMessage(
        "bot",
        `📂 <b>${file.name}</b> 데이터가 반영되었습니다.<br>• 외부 위협동향: ${activeThreats.length}건<br>• 탐지현황: ${activeDetections.length}건`
      );
    } catch (err) {
      console.error(err);
      status.textContent = `❌ 파싱 오류: ${err.message}`;
      status.className = "excel-status error";
    }
    input.value = "";
  });
}

// ─── 외부 위협동향 렌더링 ───────────────────────────────────────────
function renderThreats() {
  const list = document.getElementById("threat-list");
  if (activeThreats.length === 0) {
    list.innerHTML = `<div class="empty-msg">데이터가 없습니다.</div>`;
    return;
  }
  list.innerHTML = activeThreats.map(
    (t) => `
    <div class="list-item threat-item" data-id="${t.id}" title="${t.source} · ${t.date}">
      <span class="item-dot"></span>
      <span class="item-text">${t.title} <span class="item-source">(${t.source})</span></span>
    </div>`
  ).join("");

  list.querySelectorAll(".threat-item").forEach((el) => {
    el.addEventListener("click", () => {
      clearActive(".threat-item");
      el.classList.add("active");
      const t = activeThreats.find((x) => x.id === +el.dataset.id);
      appendChatMessage("bot", formatNewsMessage(t));
      highlightMatchedDetections(t.id);
    });
  });
}

// ─── 탐지현황 렌더링 ────────────────────────────────────────────────
function renderDetections() {
  const list = document.getElementById("detection-list");
  if (activeDetections.length === 0) {
    list.innerHTML = `<div class="empty-msg">데이터가 없습니다.</div>`;
    return;
  }
  list.innerHTML = activeDetections.map(
    (d) => `
    <div class="list-item detection-item" data-id="${d.id}">
      <span class="item-dot ${getTypeClass(d.type)}"></span>
      <span class="item-text">${d.label}</span>
    </div>`
  ).join("");

  list.querySelectorAll(".detection-item").forEach((el) => {
    el.addEventListener("click", () => {
      clearActive(".detection-item");
      el.classList.add("active");
      const d = activeDetections.find((x) => x.id === +el.dataset.id);
      appendChatMessage("bot", formatDetectionMessage(d));
    });
  });
}

function getTypeClass(type) {
  if (type === "메일") return "dot-mail";
  if (type === "웹방화벽") return "dot-waf";
  return "dot-ndr";
}

function highlightMatchedDetections(threatId) {
  document.querySelectorAll(".detection-item").forEach((el) => {
    const d = activeDetections.find((x) => x.id === +el.dataset.id);
    el.classList.toggle("matched", d && d.threatId === threatId);
  });
}

// ─── 메시지 포맷 ────────────────────────────────────────────────────
function formatNewsMessage(t) {
  const tags = (t.tags || []).map((g) => `<span class="tag">${g}</span>`).join(" ");
  return `<div class="chat-card">
    <div class="chat-card-title">📰 ${t.title}</div>
    <div class="chat-card-meta">${t.source} · ${t.date} ${tags}</div>
    <div class="chat-card-body">${String(t.body || "").replace(/\n/g, "<br>")}</div>
  </div>`;
}

function formatDetectionMessage(d) {
  const rows = Object.entries(d.detail)
    .map(([k, v]) => `<tr><th>${k}</th><td>${v}</td></tr>`)
    .join("");
  return `<div class="chat-card">
    <div class="chat-card-title">🔍 이벤트 요약 설명</div>
    <div class="chat-card-meta">${d.label}</div>
    <table class="detail-table">${rows}</table>
  </div>`;
}

// ─── 챗봇 ───────────────────────────────────────────────────────────
function setupChatbot() {
  const input = document.getElementById("chat-input");
  const sendBtn = document.getElementById("btn-send");
  const attachBtn = document.getElementById("btn-attach");
  const fileInput = document.getElementById("file-input");

  sendBtn.addEventListener("click", sendMessage);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      sendMessage();
    }
  });
  attachBtn.addEventListener("click", () => fileInput.click());
  fileInput.addEventListener("change", (e) => {
    if (e.target.files[0]) {
      appendChatMessage("user", `📎 파일 첨부: ${e.target.files[0].name}`);
      appendChatMessage("bot", `<b>${e.target.files[0].name}</b> 파일을 수신했습니다.`);
      fileInput.value = "";
    }
  });

  appendChatMessage(
    "bot",
    "안녕하세요! 위협탐지 챗봇입니다.<br>• <b>외부 위협동향</b> 항목 클릭 → 뉴스 본문 표시<br>• <b>탐지현황</b> 항목 클릭 → 이벤트 요약 표시<br>• 상단 <b>엑셀 데이터 가져오기</b> 버튼으로 실제 데이터를 반영할 수 있습니다."
  );
}

function sendMessage() {
  const input = document.getElementById("chat-input");
  const text = input.value.trim();
  if (!text) return;
  appendChatMessage("user", text);
  input.value = "";
  setTimeout(() => appendChatMessage("bot", autoReply(text)), 400);
}

function autoReply(text) {
  const lower = text.toLowerCase();

  if (/(피싱|phishing|메일|mail)/.test(lower)) {
    const list = activeDetections.filter((d) => d.type === "메일");
    if (list.length === 0) return "현재 탐지된 메일 관련 위협이 없습니다.";
    return `현재 탐지된 <b>메일 관련 위협</b> ${list.length}건:<br>` + list.map((d) => `• ${d.label}`).join("<br>");
  }
  if (/(웹방화벽|waf|방화벽)/.test(lower)) {
    const list = activeDetections.filter((d) => d.type.includes("웹방화벽"));
    if (list.length === 0) return "현재 탐지된 웹방화벽 관련 이벤트가 없습니다.";
    return `현재 <b>웹방화벽</b> 관련 탐지 ${list.length}건:<br>` + list.map((d) => `• ${d.label}`).join("<br>");
  }
  if (/(ndr|엔디알)/.test(lower)) {
    const list = activeDetections.filter((d) => d.type.includes("NDR"));
    if (list.length === 0) return "현재 탐지된 NDR 관련 이벤트가 없습니다.";
    return `현재 <b>NDR</b> 관련 탐지 ${list.length}건:<br>` + list.map((d) => `• ${d.label}`).join("<br>");
  }
  if (/(전체|요약|현황|통계|summary)/.test(lower)) {
    const total = activeDetections.reduce((s, d) => s + d.count, 0);
    const mailCnt = activeDetections.filter((d) => d.type === "메일").reduce((s, d) => s + d.count, 0);
    return `<b>위협 탐지 현황 요약</b><br>
      • 외부 위협동향: ${activeThreats.length}건<br>
      • 총 탐지 이벤트: <b>${total}건</b><br>
      • 메일 피싱 유입: ${mailCnt}건<br>
      • 웹방화벽/NDR 차단·탐지: ${total - mailCnt}건`;
  }
  if (/(도움|help|사용법)/.test(lower)) {
    return "사용 방법:<br>① 상단 <b>엑셀 데이터 가져오기</b>로 .xlsx 업로드<br>② <b>외부 위협동향</b> 클릭 → 기사 본문 표시<br>③ <b>탐지현황</b> 클릭 → 이벤트 상세 표시<br>④ 키워드 질문: 피싱, NDR, 웹방화벽, 전체 요약";
  }

  // 뉴스 제목 키워드 검색
  const matched = activeThreats.filter((t) =>
    t.title.toLowerCase().includes(lower) || (t.tags || []).some((tag) => tag.toLowerCase().includes(lower))
  );
  if (matched.length > 0) {
    return `"<b>${text}</b>" 관련 위협동향 ${matched.length}건:<br>` +
      matched.map((t) => `• ${t.title} (${t.source})`).join("<br>");
  }

  return `"<b>${text}</b>"에 대한 정보를 찾지 못했습니다.<br>키워드 예시: 피싱, NDR, 웹방화벽, 전체 요약`;
}

// ─── 유틸 ───────────────────────────────────────────────────────────
function appendChatMessage(role, html) {
  const box = document.getElementById("chat-messages");
  const wrap = document.createElement("div");
  wrap.className = `chat-message ${role}`;
  const bubble = document.createElement("div");
  bubble.className = "bubble";
  bubble.innerHTML = html;
  wrap.appendChild(bubble);
  box.appendChild(wrap);
  box.scrollTop = box.scrollHeight;
}

function clearActive(selector) {
  document.querySelectorAll(selector).forEach((el) => el.classList.remove("active"));
  document.querySelectorAll(".detection-item").forEach((el) => el.classList.remove("matched"));
}

// ─── 사이드바 ────────────────────────────────────────────────────────
function setupSidebar() {
  document.querySelectorAll(".menu-item").forEach((item) => {
    item.addEventListener("click", () => {
      document.querySelectorAll(".menu-item").forEach((i) => i.classList.remove("active"));
      item.classList.add("active");
    });
  });
}
