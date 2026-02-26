/**
 * 엑셀 파일 파싱 모듈 (SheetJS 사용)
 *
 * 예상 시트 구조:
 *
 * [뉴스기사1, 뉴스기사2, ...]  ← 시트 1개 = 기사 1개
 *   A1: 제목+본문+출처+날짜+태그가 합쳐진 전체 텍스트
 *       (첫 번째 줄 → 제목, 나머지 → 본문으로 처리)
 *
 * [스팸스나이퍼]
 *   A:날짜 | B:메일종류 | C:모드 | D:전송결과 | E:첨부 | F:제목 | G:발신자 | H:발신자IP
 *   I:수신자 | J:메일크기 | K:필터링정보 | L:복구날짜 | M:서버IP | N:Vade Engine Spamcause
 *   → 뉴스기사 제목 키워드를 F(제목) 또는 G(발신자)에서 검색하여 매칭
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
        // 시트 1개 = 기사 1개, A1 셀에 제목+본문 전체 텍스트
        wb.SheetNames.forEach((name) => {
          if (!name.startsWith("뉴스기사")) return;
          const sheet = wb.Sheets[name];
          const a1 = sheet["A1"] ? String(sheet["A1"].v) : "";
          if (!a1) return;
          const lines = a1.split("\n").map((l) => l.trim()).filter(Boolean);
          const title = lines[0] || "";
          const body = lines.slice(1).join("\n");
          threats.push({ id: threatId++, title, source: "", date: "", body, tags: [] });
        });

        // ── 스팸스나이퍼 시트 파싱 ─────────────────────────────
        // 컬럼: A:날짜 B:메일종류 C:모드 D:전송결과 E:첨부 F:제목 G:발신자 H:발신자IP
        //        I:수신자 J:메일크기 K:필터링정보 L:복구날짜 M:서버IP N:Vade Engine Spamcause
        if (wb.Sheets["스팸스나이퍼"]) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets["스팸스나이퍼"], { header: 1, defval: "" });
          const spamRows = [];
          for (let i = 1; i < rows.length; i++) {
            const [date, mailType, mode, result, attachment, subject, sender, senderIP, recipient, mailSize, filterInfo, recoveryDate, serverIP, vadeCause] = rows[i];
            if (!date && !subject && !sender) continue;
            spamRows.push({
              date, mailType, mode, result, attachment,
              subject: String(subject || ""),
              sender: String(sender || ""),
              senderIP: String(senderIP || ""),
              recipient: String(recipient || ""),
              mailSize, filterInfo: String(filterInfo || ""),
              recoveryDate, serverIP, vadeCause: String(vadeCause || ""),
            });
          }

          // 뉴스기사 제목 키워드로 스팸 로그 매칭 (메일 제목 or 발신자)
          threats.forEach((threat) => {
            const keywords = extractKeywords(threat.title);
            const matched = spamRows.filter((row) =>
              keywords.some((kw) => row.subject.includes(kw) || row.sender.includes(kw))
            );
            if (matched.length === 0) return;
            detections.push({
              id: detectionId++,
              threatId: threat.id,
              type: "메일",
              label: `[메일] ${threat.title} 관련 메일 ${matched.length}건 유입`,
              count: matched.length,
              action: "유입",
              source: "스팸스나이퍼",
              detail: {
                날짜: [...new Set(matched.map((r) => formatDate(r.date)).filter(Boolean))].join(", "),
                발신자: [...new Set(matched.map((r) => r.sender).filter(Boolean))].join(", "),
                "메일 제목": matched.map((r) => r.subject).filter(Boolean).join(" / "),
                수신자: [...new Set(matched.map((r) => r.recipient).filter(Boolean))].join(", "),
                수신건수: `${matched.length}건`,
                필터링정보: [...new Set(matched.map((r) => r.filterInfo).filter(Boolean))].join(", "),
              },
            });
          });
        }

        // ── NDR로그 시트 파싱 ──────────────────────────────────
        // 컬럼: A:NDR_RuleName(사용자정의) B:RiskScore(사용자정의) C:로그소스
        //        D:시작시간 E:소스IP F:소스포트 G:대상IP H:대상포트
        // → 뉴스기사 제목 키워드를 A(NDR_RuleName)/C(로그소스)에서 검색하여 매칭
        if (wb.Sheets["NDR로그"]) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets["NDR로그"], { header: 1, defval: "" });
          const ndrRows = [];
          for (let i = 1; i < rows.length; i++) {
            const [ruleName, riskScore, logSource, startTime, srcIP, srcPort, dstIP, dstPort] = rows[i];
            if (!ruleName && !logSource && !srcIP) continue;
            ndrRows.push({
              ruleName: String(ruleName || ""),
              riskScore: +riskScore || 0,
              logSource: String(logSource || ""),
              startTime,
              srcIP: String(srcIP || ""),
              srcPort,
              dstIP: String(dstIP || ""),
              dstPort,
            });
          }

          // 뉴스기사 제목 키워드로 매칭 (NDR_RuleName / 로그소스)
          threats.forEach((threat) => {
            const keywords = extractKeywords(threat.title);
            const matched = ndrRows.filter((row) =>
              keywords.some((kw) => row.ruleName.includes(kw) || row.logSource.includes(kw))
            );
            if (matched.length === 0) return;
            const totalCount = matched.length;
            detections.push({
              id: detectionId++,
              threatId: threat.id,
              type: "NDR",
              label: `[NDR] ${threat.title} 관련 이벤트 ${totalCount}건 탐지`,
              count: totalCount,
              action: "탐지",
              source: "NDR",
              detail: {
                로그출처: "NDR",
                NDR_RuleName: [...new Set(matched.map((r) => r.ruleName).filter(Boolean))].join(", "),
                "로그 소스": [...new Set(matched.map((r) => r.logSource).filter(Boolean))].join(", "),
                "소스 IP": [...new Set(matched.map((r) => r.srcIP).filter(Boolean))].join(", "),
                "대상 IP": [...new Set(matched.map((r) => r.dstIP).filter(Boolean))].join(", "),
                매칭이벤트건수: `${totalCount}건`,
                조치량: `${totalCount}건 탐지`,
              },
            });
          });
        }

        // ── 웹방화벽로그 시트 파싱 ────────────────────────────
        // 컬럼: A:시간 B:클라이언트IP(국가) C:클라이언트포트 D:OriginIP(국가) E:서버IP
        //        F:서버포트 G:HTTP버전 H:URL도메인 I:요청 J:요청데이터길이 K:응답
        //        L:응답데이터길이 M:룰이름 N:패턴이름 O:탐지유형 P:탐지근거
        //        Q:OWASP취약점 R:국정원8대취약점 S:KISA홈페이지취약점 T:탐지개수
        //        U:메일 V:위험도 W:조치 X:TransactionID
        // → 뉴스기사 제목 키워드를 H(URL도메인)/M(룰이름)/N(패턴이름)/P(탐지근거)에서 검색하여 매칭
        if (wb.Sheets["웹방화벽로그"]) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets["웹방화벽로그"], { header: 1, defval: "" });
          const wafRows = [];
          for (let i = 1; i < rows.length; i++) {
            const [time, clientIP, clientPort, originIP, serverIP, serverPort, httpVer, urlDomain, request, reqLen, response, resLen, ruleName, patternName, detType, basis, owasp, gov8, kisa, countRaw, mail, riskLevel, action, transactionId] = rows[i];
            if (!time && !clientIP && !ruleName) continue;
            wafRows.push({
              time, clientIP: String(clientIP || ""), clientPort, originIP, serverIP: String(serverIP || ""),
              serverPort, httpVer, urlDomain: String(urlDomain || ""), request, reqLen, response, resLen,
              ruleName: String(ruleName || ""), patternName: String(patternName || ""),
              detType: String(detType || ""), basis: String(basis || ""),
              owasp, gov8, kisa, count: +countRaw || 1,
              mail, riskLevel, action: String(action || "차단"), transactionId,
            });
          }

          // 뉴스기사 제목 키워드로 매칭 (URL도메인/룰이름/패턴이름/탐지근거)
          threats.forEach((threat) => {
            const keywords = extractKeywords(threat.title);
            const matched = wafRows.filter((row) =>
              keywords.some((kw) =>
                row.urlDomain.includes(kw) || row.ruleName.includes(kw) ||
                row.patternName.includes(kw) || row.basis.includes(kw)
              )
            );
            if (matched.length === 0) return;

            const totalCount = matched.reduce((sum, r) => sum + r.count, 0);

            // 같은 위협에 NDR 탐지가 이미 있으면 합치기
            const existingNDR = detections.find((d) => d.threatId === threat.id && d.type === "NDR");
            if (existingNDR) {
              existingNDR.type = "NDR,웹방화벽";
              existingNDR.label = `[NDR, 웹방화벽] ${threat.title} 관련 이벤트 ${existingNDR.count + totalCount}건 탐지/차단`;
              existingNDR.count += totalCount;
              existingNDR.action = "탐지/차단";
              existingNDR.detail["웹방화벽_클라이언트IP"] = [...new Set(matched.map((r) => r.clientIP).filter(Boolean))].join(", ");
              existingNDR.detail["웹방화벽_서버IP"] = [...new Set(matched.map((r) => r.serverIP).filter(Boolean))].join(", ");
              existingNDR.detail["웹방화벽_URL도메인"] = [...new Set(matched.map((r) => r.urlDomain).filter(Boolean))].join(", ");
              existingNDR.detail["웹방화벽_룰이름"] = [...new Set(matched.map((r) => r.ruleName).filter(Boolean))].join(", ");
              existingNDR.detail["웹방화벽_탐지근거"] = [...new Set(matched.map((r) => r.basis).filter(Boolean))].join(", ");
              existingNDR.detail["조치량"] = `NDR 탐지 ${existingNDR.count - totalCount}건 / 웹방화벽 차단 ${totalCount}건`;
            } else {
              detections.push({
                id: detectionId++,
                threatId: threat.id,
                type: "웹방화벽",
                label: `[웹방화벽] ${threat.title} 관련 이벤트 ${totalCount}건 차단`,
                count: totalCount,
                action: "차단",
                source: "웹방화벽",
                detail: {
                  로그출처: "웹방화벽",
                  매칭이벤트건수: `${totalCount}건`,
                  "클라이언트 IP": [...new Set(matched.map((r) => r.clientIP).filter(Boolean))].join(", "),
                  "서버 IP": [...new Set(matched.map((r) => r.serverIP).filter(Boolean))].join(", "),
                  "URL 도메인": [...new Set(matched.map((r) => r.urlDomain).filter(Boolean))].join(", "),
                  룰이름: [...new Set(matched.map((r) => r.ruleName).filter(Boolean))].join(", "),
                  패턴이름: [...new Set(matched.map((r) => r.patternName).filter(Boolean))].join(", "),
                  탐지근거: [...new Set(matched.map((r) => r.basis).filter(Boolean))].join(", "),
                  조치: [...new Set(matched.map((r) => r.action).filter(Boolean))].join(", "),
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

function extractKeywords(title) {
  const stopWords = new Set(["및", "의", "을", "를", "이", "가", "은", "는", "에", "서", "로", "와", "과", "도", "만", "관련", "대상", "통해", "위한", "사용", "통한", "위해", "있는", "있어", "하는", "하여", "으로", "에서", "부터", "까지", "대한", "통해서"]);
  return title
    .split(/[\s,·\-\/\[\]\(\)]+/)
    .map((w) => w.trim())
    .filter((w) => w.length >= 2 && !stopWords.has(w));
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return val.toLocaleDateString("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit" }).replace(/\. /g, "-").replace(".", "");
  }
  return String(val);
}
