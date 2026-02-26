/**
 * 자산관리 엑셀 파서 (SheetJS 사용)
 *
 * 지원 시트명 (순서대로 탐색, 첫 번째 매칭 사용):
 *   "대상 시스템", "대상시스템", "IT", "OT", "자산목록", "기준 시트", "기준시트", Sheet1
 *
 * 컬럼 헤더 매핑 (대소문자·공백 무시):
 *   구분       → category   (IT / OT)
 *   자산정보   → assetName
 *   IP         → ip
 *   MAC        → mac
 *   호스트명   → hostname
 *   OS         → os  (OS버전, OS 버전 등도 인식)
 *   모델명     → model
 *   관리부서   → manageDept (운영부서, 운영 부서명 등도 인식)
 *   관리담당자 → manager    (운영담당자 등도 인식)
 *   운영자     → operator
 *   상태       → status
 *   위치       → location   (위치정보, 위치 정보 등도 인식)
 *   EDR        → edr        (O/Y/1 = true)
 *   EPS        → eps
 *   DLP        → dlp
 *   DRM        → drm
 *   NAC        → nac
 *   PMS        → pms
 */

// URL(fetch)로 엑셀 파일을 읽어 파싱
async function parseAssetsFromUrl(url) {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`${url} 파일을 찾을 수 없습니다`);
  const arrayBuffer = await response.arrayBuffer();
  return parseAssetsData(new Uint8Array(arrayBuffer));
}

// File 객체로 엑셀 파일을 읽어 파싱
function parseAssetsFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try { resolve(parseAssetsData(new Uint8Array(e.target.result))); }
      catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function parseAssetsData(data) {
  const wb = XLSX.read(data, { type: "array" });

  // ── 대상 시트 탐색 ────────────────────────────────────────────────
  const PREFERRED = [
    "대상 시스템", "대상시스템",
    "IT", "OT",
    "자산목록", "자산 목록",
    "기준 시트", "기준시트",
  ];

  let sheetName = wb.SheetNames.find((n) =>
    PREFERRED.includes(n.trim())
  ) || wb.SheetNames[0];

  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

  if (rows.length < 2) return [];

  // ── 헤더 행 탐색 (최대 5행까지) ──────────────────────────────────
  // 헤더가 병합 셀 등으로 인해 1행이 아닐 수 있음
  let headerRowIdx = 0;
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const cells = rows[i].map((c) => norm(c));
    if (
      cells.some((c) => c.includes("ip") || c.includes("구분") || c.includes("hostname") || c.includes("호스트"))
    ) {
      headerRowIdx = i;
      break;
    }
  }

  const header = rows[headerRowIdx].map((c) => norm(c));

  // ── 컬럼 인덱스 매핑 ─────────────────────────────────────────────
  const col = {
    category:   findIdx(header, ["구분"]),
    assetName:  findIdx(header, ["자산정보", "자산 정보", "자산명"]),
    ip:         findIdx(header, ["ip"]),
    mac:        findIdx(header, ["mac"]),
    hostname:   findIdx(header, ["호스트명", "hostname", "호스트"]),
    os:         findIdx(header, ["os", "os버전", "os 버전"]),
    model:      findIdx(header, ["모델명", "모델", "model"]),
    manageDept: findIdx(header, ["관리부서", "운영부서", "운영 부서명", "운영부서명", "관리 부서"]),
    manager:    findIdx(header, ["관리담당자", "관리 담당자", "운영담당자", "운영 담당자", "담당자"]),
    operator:   findIdx(header, ["운영자"]),
    status:     findIdx(header, ["상태"]),
    location:   findIdx(header, ["위치", "위치정보", "위치 정보"]),
    edr:        findIdx(header, ["edr"]),
    eps:        findIdx(header, ["eps"]),
    dlp:        findIdx(header, ["dlp"]),
    drm:        findIdx(header, ["drm"]),
    nac:        findIdx(header, ["nac"]),
    pms:        findIdx(header, ["pms"]),
  };

  // ── 데이터 행 파싱 ────────────────────────────────────────────────
  const assets = [];
  let id = 1;

  for (let i = headerRowIdx + 1; i < rows.length; i++) {
    const r = rows[i];

    // 완전히 빈 행 건너뜀
    if (r.every((c) => c === "" || c === null || c === undefined)) continue;

    const ip = get(r, col.ip);
    const hostname = get(r, col.hostname);
    const category = get(r, col.category).toUpperCase();

    // IP 또는 호스트명이 없으면 건너뜀
    if (!ip && !hostname) continue;

    assets.push({
      id: id++,
      category: (category === "IT" || category === "OT") ? category : "IT",
      assetName: get(r, col.assetName) || "OA",
      ip,
      mac:        get(r, col.mac),
      hostname,
      os:         get(r, col.os),
      model:      get(r, col.model),
      manageDept: get(r, col.manageDept),
      manager:    get(r, col.manager),
      operator:   get(r, col.operator),
      status:     get(r, col.status) || "운영중",
      location:   get(r, col.location),
      edr: isChecked(r, col.edr),
      eps: isChecked(r, col.eps),
      dlp: isChecked(r, col.dlp),
      drm: isChecked(r, col.drm),
      nac: isChecked(r, col.nac),
      pms: isChecked(r, col.pms),
    });
  }

  return assets;
}

// ── 유틸 ───────────────────────────────────────────────────────────
function norm(val) {
  return String(val || "").trim().toLowerCase().replace(/\s+/g, "");
}

function findIdx(header, names) {
  for (const name of names) {
    const n = norm(name);
    const idx = header.findIndex((h) => h === n);
    if (idx !== -1) return idx;
  }
  // 부분 일치 fallback
  for (const name of names) {
    const n = norm(name);
    const idx = header.findIndex((h) => h.includes(n) || n.includes(h));
    if (idx !== -1) return idx;
  }
  return -1;
}

function get(row, idx) {
  if (idx === -1) return "";
  return String(row[idx] ?? "").trim();
}

function isChecked(row, idx) {
  if (idx === -1) return false;
  const v = String(row[idx] ?? "").trim().toUpperCase();
  return v === "O" || v === "TRUE" || v === "1" || v === "Y" || v === "✓";
}
