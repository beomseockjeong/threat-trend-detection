/* ── 정보자산 식별 앱 ── */

(function () {
  "use strict";

  let filteredAssets = [...ASSETS];

  // ── 초기화 ─────────────────────────────────────────────
  function init() {
    renderSummaryCards();
    renderTable(ASSETS);
    bindEvents();
    showStartupModal();
  }

  // ── 시작 모달: fetch 시도 → 실패 시 파일 선택 UI 표시 ──
  async function showStartupModal() {
    // 로컬 서버 환경이면 fetch로 자동 로드 시도
    try {
      const imported = await parseAssetsFromUrl("./assets.xlsx");
      if (imported.length > 0) {
        applyImported(imported, "assets.xlsx");
        closeStartupModal();
        return;
      }
    } catch (_) {
      // file:// 환경 등 fetch 불가 → 시작 모달 유지
    }
    // fetch 실패 or 데이터 없음 → 모달 그대로 표시
  }

  function applyImported(imported, filename) {
    ASSETS.length = 0;
    imported.forEach((a) => ASSETS.push(a));
    renderSummaryCards();
    applyFilter();
    const status = document.getElementById("excel-status");
    status.textContent = `✅ ${filename} — ${imported.length}건 로드 완료`;
    status.className = "excel-status success";
  }

  function closeStartupModal() {
    document.getElementById("startup-modal").classList.remove("open");
  }

  // ── 새로고침 (서버 환경용) ───────────────────────────────
  async function autoLoadExcel() {
    const status = document.getElementById("excel-status");
    status.textContent = "⏳ assets.xlsx 로딩 중...";
    status.className = "excel-status loading";
    try {
      const imported = await parseAssetsFromUrl("./assets.xlsx");
      if (imported.length === 0) {
        status.textContent = "⚠️ 데이터를 찾을 수 없습니다.";
        status.className = "excel-status error";
        return;
      }
      applyImported(imported, "assets.xlsx");
    } catch (err) {
      status.textContent = "fetch 불가 — 업로드 버튼을 이용해주세요.";
      status.className = "excel-status error";
    }
  }

  // ── 요약 카드 ───────────────────────────────────────────
  function renderSummaryCards() {
    const total = ASSETS.length;
    const itCount = ASSETS.filter((a) => a.category === "IT").length;
    const otCount = ASSETS.filter((a) => a.category === "OT").length;

    const swKeys = ["edr", "eps", "dlp", "drm", "nac", "pms"];
    const swLabels = { edr: "EDR", eps: "EPS", dlp: "DLP", drm: "DRM", nac: "NAC", pms: "PMS" };

    const swRates = swKeys.map((k) => {
      const installed = ASSETS.filter((a) => a[k]).length;
      return { label: swLabels[k], rate: Math.round((installed / total) * 100) };
    });

    document.getElementById("card-total").textContent = total + "건";
    document.getElementById("card-it").textContent = itCount + "건";
    document.getElementById("card-ot").textContent = otCount + "건";

    const swContainer = document.getElementById("sw-rates");
    swContainer.innerHTML = swKeys
      .map((k) => {
        const installed = ASSETS.filter((a) => a[k]).length;
        const rate = Math.round((installed / total) * 100);
        return `<div class="sw-rate-item">
          <span class="sw-rate-label">${swLabels[k]}</span>
          <div class="sw-rate-bar-wrap">
            <div class="sw-rate-bar" style="width:${rate}%"></div>
          </div>
          <span class="sw-rate-pct">${rate}%</span>
        </div>`;
      })
      .join("");
  }

  // ── 테이블 렌더링 ────────────────────────────────────────
  function renderTable(assets) {
    filteredAssets = assets;
    const tbody = document.getElementById("asset-tbody");

    if (assets.length === 0) {
      tbody.innerHTML = `<tr><td colspan="17" class="asset-empty">검색 결과가 없습니다.</td></tr>`;
      document.getElementById("result-count").textContent = "0건";
      return;
    }

    document.getElementById("result-count").textContent = assets.length + "건";

    tbody.innerHTML = assets
      .map(
        (a) => `
      <tr class="asset-row" data-id="${a.id}">
        <td><span class="category-badge category-${a.category.toLowerCase()}">${a.category}</span></td>
        <td>${a.assetName}</td>
        <td>${a.ip}</td>
        <td class="font-mono">${a.mac}</td>
        <td>${a.hostname}</td>
        <td class="col-os">${a.os}</td>
        <td>${a.model || "-"}</td>
        <td>${a.manageDept}</td>
        <td>${a.manager}</td>
        <td>${a.operator}</td>
        <td><span class="status-badge">${a.status}</span></td>
        <td>${a.location}</td>
        <td>${swBadge(a.edr)}</td>
        <td>${swBadge(a.eps)}</td>
        <td>${swBadge(a.dlp)}</td>
        <td>${swBadge(a.drm)}</td>
        <td>${swBadge(a.nac)}</td>
        <td>${swBadge(a.pms)}</td>
      </tr>`
      )
      .join("");

    // 행 클릭 → 상세 모달
    tbody.querySelectorAll(".asset-row").forEach((row) => {
      row.addEventListener("click", () => {
        const id = parseInt(row.dataset.id, 10);
        openModal(ASSETS.find((a) => a.id === id));
      });
    });
  }

  function swBadge(val) {
    return val
      ? `<span class="sw-badge sw-on">O</span>`
      : `<span class="sw-badge sw-off">X</span>`;
  }

  // ── 검색 / 필터 ─────────────────────────────────────────
  function applyFilter() {
    const category = document.getElementById("filter-category").value;
    const keyword = document.getElementById("filter-keyword").value.trim().toLowerCase();

    const result = ASSETS.filter((a) => {
      const catMatch = category === "전체" || a.category === category;
      const kwMatch =
        !keyword ||
        a.ip.toLowerCase().includes(keyword) ||
        a.hostname.toLowerCase().includes(keyword) ||
        a.assetName.toLowerCase().includes(keyword) ||
        a.mac.toLowerCase().includes(keyword) ||
        a.os.toLowerCase().includes(keyword) ||
        a.manageDept.includes(keyword) ||
        a.manager.includes(keyword) ||
        a.location.includes(keyword);
      return catMatch && kwMatch;
    });

    renderTable(result);
  }

  // ── 상세 모달 ────────────────────────────────────────────
  function openModal(asset) {
    if (!asset) return;

    const swLabels = { edr: "EDR", eps: "EPS", dlp: "DLP", drm: "DRM", nac: "NAC", pms: "PMS" };
    const swRows = Object.entries(swLabels)
      .map(([k, label]) => `<tr><th>${label}</th><td>${swBadge(asset[k])}</td></tr>`)
      .join("");

    document.getElementById("modal-title").textContent = `[${asset.category}] ${asset.hostname}`;
    document.getElementById("modal-body").innerHTML = `
      <table class="detail-table">
        <tr><th>구분</th><td><span class="category-badge category-${asset.category.toLowerCase()}">${asset.category}</span></td></tr>
        <tr><th>자산정보</th><td>${asset.assetName}</td></tr>
        <tr><th>IP</th><td>${asset.ip}</td></tr>
        <tr><th>MAC</th><td class="font-mono">${asset.mac}</td></tr>
        <tr><th>호스트명</th><td>${asset.hostname}</td></tr>
        <tr><th>OS</th><td>${asset.os}</td></tr>
        <tr><th>모델명</th><td>${asset.model || "-"}</td></tr>
        <tr><th>관리부서</th><td>${asset.manageDept}</td></tr>
        <tr><th>관리담당자</th><td>${asset.manager}</td></tr>
        <tr><th>운영자</th><td>${asset.operator}</td></tr>
        <tr><th>상태</th><td><span class="status-badge">${asset.status}</span></td></tr>
        <tr><th>위치</th><td>${asset.location}</td></tr>
        ${swRows}
      </table>`;

    document.getElementById("asset-modal").classList.add("open");
  }

  function closeModal() {
    document.getElementById("asset-modal").classList.remove("open");
  }

  // ── 엑셀 다운로드 ────────────────────────────────────────
  function downloadExcel() {
    const headers = [
      "구분", "자산정보", "IP", "MAC", "호스트명",
      "OS", "모델명", "관리부서", "관리담당자", "운영자", "상태", "위치",
      "EDR", "EPS", "DLP", "DRM", "NAC", "PMS",
    ];

    const rows = filteredAssets.map((a) => [
      a.category, a.assetName, a.ip, a.mac, a.hostname,
      a.os, a.model, a.manageDept, a.manager, a.operator, a.status, a.location,
      a.edr ? "O" : "X",
      a.eps ? "O" : "X",
      a.dlp ? "O" : "X",
      a.drm ? "O" : "X",
      a.nac ? "O" : "X",
      a.pms ? "O" : "X",
    ]);

    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);

    // 열 너비 설정
    ws["!cols"] = [
      { wch: 6 }, { wch: 8 }, { wch: 14 }, { wch: 20 }, { wch: 16 },
      { wch: 30 }, { wch: 20 }, { wch: 14 }, { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 20 },
      { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 }, { wch: 6 },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "정보자산목록");

    const date = new Date().toISOString().slice(0, 10).replace(/-/g, "");
    XLSX.writeFile(wb, `정보자산목록_${date}.xlsx`);
  }

  // ── 엑셀 업로드 (공통) ──────────────────────────────────
  async function handleUpload(file, isStartup = false) {
    if (!file) return;
    const statusEl = document.getElementById("upload-status");
    try {
      const imported = await parseAssetsFile(file);
      if (imported.length === 0) {
        alert("유효한 데이터를 찾을 수 없습니다.");
        return;
      }
      applyImported(imported, file.name);
      if (isStartup) closeStartupModal();
      statusEl.textContent = `✅ ${imported.length}건 로드 완료`;
      statusEl.className = "upload-status success";
    } catch (err) {
      alert("파일 파싱 오류: " + err.message);
    }
  }

  // ── 이벤트 바인딩 ────────────────────────────────────────
  function bindEvents() {
    document.getElementById("btn-search").addEventListener("click", applyFilter);

    document.getElementById("filter-keyword").addEventListener("keydown", (e) => {
      if (e.key === "Enter") applyFilter();
    });

    document.getElementById("btn-refresh").addEventListener("click", autoLoadExcel);

    document.getElementById("btn-download").addEventListener("click", downloadExcel);

    document.getElementById("btn-upload").addEventListener("click", () => {
      document.getElementById("asset-file-input").click();
    });

    document.getElementById("asset-file-input").addEventListener("change", (e) => {
      handleUpload(e.target.files[0], false);
      e.target.value = "";
    });

    // 시작 모달 이벤트
    document.getElementById("startup-file-input").addEventListener("change", (e) => {
      handleUpload(e.target.files[0], true);
      e.target.value = "";
    });

    document.getElementById("startup-skip").addEventListener("click", () => {
      closeStartupModal();
    });

    document.getElementById("modal-close").addEventListener("click", closeModal);
    document.getElementById("asset-modal").addEventListener("click", (e) => {
      if (e.target === e.currentTarget) closeModal();
    });

    document.getElementById("btn-dashboard").addEventListener("click", () => {
      window.location.href = "index.html";
    });
  }

  // ── 실행 ────────────────────────────────────────────────
  document.addEventListener("DOMContentLoaded", init);
})();
