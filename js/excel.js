/**
 * 엑셀 파일 파싱 모듈 (SheetJS 사용)
 *
 * 예상 시트 구조:
 *
 * [뉴스기사1, 뉴스기사2, ...]
 *   A:제목 | B:출처 | C:날짜 | D:본문 | E:태그(쉼표 구분)
 *
 * [스팸스나이퍼]
 *   A:기사명 | B:시간 | C:발신자 | D:메일제목 | E:수신자 | F:수신건수 | G:조치
 *
 * [NDR로그]
 *   A:기사명 | B:NDR_RuleName | C:소스IP | D:대상IP | E:탐지유형
 *   F:클라이언트IP | G:서버IP | H:탐지근거 | I:조치 | J:이벤트건수
 *
 * [웹방화벽로그]
 *   A:기사명 | B:클라이언트IP | C:서버IP | D:탐지근거 | E:조치 | F:이벤트건수
 */

function parseExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array", cellDates: true });

        const threats = [];
        const detections = [];
        let threatId = 1;
        let detectionId = 1;

        // ── 뉴스기사 시트 파싱 ──────────────────────────────────
        wb.SheetNames.forEach((name) => {
          if (!name.startsWith("뉴스기사")) return;
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: "" });
          // 1행 = 헤더, 2행부터 데이터
          for (let i = 1; i < rows.length; i++) {
            const [title, source, date, body, tagsRaw] = rows[i];
            if (!title) continue;
            const dateStr = formatDate(date);
            const tags = tagsRaw
              ? String(tagsRaw).split(",").map((t) => t.trim()).filter(Boolean)
              : [];
            threats.push({ id: threatId++, title: String(title), source: String(source || ""), date: dateStr, body: String(body || ""), tags });
          }
        });

        // ── 스팸스나이퍼 시트 파싱 ─────────────────────────────
        if (wb.Sheets["스팸스나이퍼"]) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets["스팸스나이퍼"], { header: 1, defval: "" });
          // 기사명별로 그룹핑
          const groups = {};
          for (let i = 1; i < rows.length; i++) {
            const [articleTitle, time, sender, mailTitle, recipient, count, action] = rows[i];
            if (!articleTitle) continue;
            const key = String(articleTitle);
            if (!groups[key]) groups[key] = { rows: [], count: 0 };
            groups[key].rows.push({ time, sender, mailTitle, recipient, count: +count || 0, action });
            groups[key].count += +count || 0;
          }

          Object.entries(groups).forEach(([articleTitle, g]) => {
            const matchedThreat = findMatchingThreat(threats, articleTitle);
            const first = g.rows[0];
            detections.push({
              id: detectionId++,
              threatId: matchedThreat ? matchedThreat.id : null,
              type: "메일",
              label: `[메일] ${articleTitle} 관련 메일 ${g.count}건 유입`,
              count: g.count,
              action: first.action || "유입",
              source: "스팸스나이퍼",
              detail: {
                시간: g.rows.map((r) => formatDate(r.time)).filter(Boolean).join(" ~ ") || String(first.time),
                발신자: g.rows.map((r) => r.sender).filter(Boolean).join(", "),
                제목: g.rows.map((r) => r.mailTitle).filter(Boolean).join(" / "),
                수신자: g.rows.map((r) => r.recipient).filter(Boolean).join(", "),
                수신건수: `${g.count}건`,
                조치: first.action || "스팸 격리",
              },
            });
          });
        }

        // ── NDR로그 시트 파싱 ──────────────────────────────────
        if (wb.Sheets["NDR로그"]) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets["NDR로그"], { header: 1, defval: "" });
          const groups = {};
          for (let i = 1; i < rows.length; i++) {
            const [articleTitle, ruleNameRaw, srcIP, dstIP, detType, clientIP, serverIP, basis, actionRaw, countRaw] = rows[i];
            if (!articleTitle) continue;
            const key = String(articleTitle);
            if (!groups[key]) groups[key] = { rows: [], count: 0 };
            groups[key].rows.push({ ruleNameRaw, srcIP, dstIP, detType, clientIP, serverIP, basis, actionRaw, count: +countRaw || 0 });
            groups[key].count += +countRaw || 0;
          }

          Object.entries(groups).forEach(([articleTitle, g]) => {
            const matchedThreat = findMatchingThreat(threats, articleTitle);
            const first = g.rows[0];
            const totalCount = g.count || g.rows.length;
            detections.push({
              id: detectionId++,
              threatId: matchedThreat ? matchedThreat.id : null,
              type: "NDR",
              label: `[NDR] ${articleTitle} 관련 이벤트 ${totalCount}건 탐지`,
              count: totalCount,
              action: "탐지",
              source: "NDR",
              detail: {
                로그출처: "NDR",
                기사명: articleTitle,
                NDR_RuleName: [...new Set(g.rows.map((r) => r.ruleNameRaw).filter(Boolean))].join(", "),
                "소스 IP": [...new Set(g.rows.map((r) => r.srcIP).filter(Boolean))].join(", "),
                "대상 IP": [...new Set(g.rows.map((r) => r.dstIP).filter(Boolean))].join(", "),
                탐지유형: [...new Set(g.rows.map((r) => r.detType).filter(Boolean))].join(", "),
                탐지근거: first.basis || "",
                매칭이벤트건수: `${totalCount}건`,
                조치량: `${totalCount}건 탐지`,
              },
            });
          });
        }

        // ── 웹방화벽로그 시트 파싱 ────────────────────────────
        if (wb.Sheets["웹방화벽로그"]) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets["웹방화벽로그"], { header: 1, defval: "" });
          const groups = {};
          for (let i = 1; i < rows.length; i++) {
            const [articleTitle, clientIP, serverIP, basis, actionRaw, countRaw] = rows[i];
            if (!articleTitle) continue;
            const key = String(articleTitle);
            if (!groups[key]) groups[key] = { rows: [], count: 0 };
            groups[key].rows.push({ clientIP, serverIP, basis, actionRaw, count: +countRaw || 0 });
            groups[key].count += +countRaw || 0;
          }

          Object.entries(groups).forEach(([articleTitle, g]) => {
            const matchedThreat = findMatchingThreat(threats, articleTitle);
            const first = g.rows[0];
            const totalCount = g.count || g.rows.length;

            // 같은 기사명으로 NDR 탐지가 이미 있는지 확인 → 있으면 합치기
            const existingNDR = detections.find(
              (d) => d.threatId === (matchedThreat ? matchedThreat.id : null) && d.type === "NDR"
            );
            if (existingNDR) {
              existingNDR.type = "NDR,웹방화벽";
              existingNDR.label = `[NDR, 웹방화벽] ${articleTitle} 관련 이벤트 ${existingNDR.count + totalCount}건 탐지/차단`;
              existingNDR.count += totalCount;
              existingNDR.action = "탐지/차단";
              existingNDR.detail["웹방화벽_클라이언트IP"] = [...new Set(g.rows.map((r) => r.clientIP).filter(Boolean))].join(", ");
              existingNDR.detail["웹방화벽_서버IP"] = [...new Set(g.rows.map((r) => r.serverIP).filter(Boolean))].join(", ");
              existingNDR.detail["웹방화벽_탐지근거"] = first.basis || "";
              existingNDR.detail["조치량"] = `NDR 탐지 ${existingNDR.count - totalCount}건 / 웹방화벽 차단 ${totalCount}건`;
            } else {
              detections.push({
                id: detectionId++,
                threatId: matchedThreat ? matchedThreat.id : null,
                type: "웹방화벽",
                label: `[웹방화벽] ${articleTitle} 관련 이벤트 ${totalCount}건 차단`,
                count: totalCount,
                action: "차단",
                source: "웹방화벽",
                detail: {
                  로그출처: "웹방화벽",
                  기사명: articleTitle,
                  매칭이벤트건수: `${totalCount}건`,
                  "클라이언트 IP": [...new Set(g.rows.map((r) => r.clientIP).filter(Boolean))].join(", "),
                  "서버 IP": [...new Set(g.rows.map((r) => r.serverIP).filter(Boolean))].join(", "),
                  탐지근거: first.basis || "",
                  조치량: `${totalCount}건 차단`,
                },
              });
            }
          });
        }

        resolve({ threats, detections });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ── 유틸 ──────────────────────────────────────────────────────────
function findMatchingThreat(threats, articleTitle) {
  if (!articleTitle) return null;
  const query = String(articleTitle).replace(/\s/g, "").toLowerCase();
  return (
    threats.find((t) => t.title.replace(/\s/g, "").toLowerCase().includes(query)) ||
    threats.find((t) => query.includes(t.title.replace(/\s/g, "").toLowerCase().slice(0, 8)))
  );
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return val.toLocaleDateString("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit" }).replace(/\. /g, "-").replace(".", "");
  }
  return String(val);
}
