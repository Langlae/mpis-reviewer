---
marp: true
paginate: true
---

<style>
/* ── 팔레트 ─────────────────────────────────────────── */
:root {
  --navy:    #0D1B2A;
  --navy2:   #162537;
  --navy3:   #1E3047;
  --text:    #E8EDF2;
  --muted:   #7A8FA6;
  --cyan:    #4CC9F0;
  --red:     #E63946;
  --green:   #22C55E;
  --gold:    #F7B731;
  --border:  rgba(76,201,240,0.18);
}

/* ── 공통 ───────────────────────────────────────────── */
section {
  background: var(--navy);
  color: var(--text);
  font-family: 'Apple SD Gothic Neo','Malgun Gothic','Noto Sans KR',sans-serif;
  padding: 52px 64px 64px;
  display: flex;
  flex-direction: column;
  width: 1280px;
  height: 720px;
  box-sizing: border-box;
  position: relative;
  overflow: hidden;
}

/* dot grid 배경 */
section::before {
  content: '';
  position: absolute; inset: 0;
  background-image: radial-gradient(circle, rgba(76,201,240,0.06) 1px, transparent 1px);
  background-size: 36px 36px;
  pointer-events: none;
  z-index: 0;
}

/* 모든 자식은 grid 위에 */
section > * { position: relative; z-index: 1; }

/* 페이지 번호 */
section::after {
  font-size: 11px;
  color: var(--muted);
  letter-spacing: .12em;
  font-family: 'Consolas','Courier New',monospace;
}

/* ── 공통 헤더 ──────────────────────────────────────── */
.eyebrow {
  font-size: 10px;
  letter-spacing: .22em;
  color: var(--cyan);
  text-transform: uppercase;
  font-family: 'Consolas','Courier New',monospace;
  display: flex;
  align-items: center;
  gap: 10px;
  margin-bottom: 6px;
}
.eyebrow::before {
  content: '';
  display: inline-block;
  width: 24px; height: 2px;
  background: var(--cyan);
  flex-shrink: 0;
}

h2 {
  font-size: 36px;
  font-weight: 800;
  letter-spacing: -.02em;
  line-height: 1.2;
  margin: 0 0 18px;
  color: var(--text);
}

hr {
  border: none;
  border-top: 1px solid rgba(76,201,240,0.18);
  margin: 0 0 20px;
}

/* ── 1. 표지 ─────────────────────────────────────────
   레이아웃: 좌우 분할, 오른쪽 패널 */
section.cover {
  padding: 0;
  flex-direction: row;
}

.cover-left {
  flex: 1;
  display: flex;
  flex-direction: column;
  justify-content: center;
  padding: 60px 56px;
  position: relative;
  z-index: 1;
}

.cover-right {
  width: 420px;
  background: var(--navy3);
  border-left: 3px solid var(--cyan);
  display: flex;
  flex-direction: column;
  align-items: center;
  justify-content: center;
  padding: 40px 32px;
  position: relative;
  z-index: 1;
  gap: 12px;
}

.cover-tag {
  display: inline-block;
  border: 1px solid var(--border);
  border-radius: 4px;
  padding: 4px 12px;
  font-size: 10px;
  letter-spacing: .2em;
  color: var(--cyan);
  font-family: 'Consolas','Courier New',monospace;
  margin-bottom: 24px;
}

.cover h1 {
  font-size: 46px;
  font-weight: 900;
  letter-spacing: -.025em;
  line-height: 1.18;
  margin: 0 0 12px;
  color: var(--text);
}

.cover h1 em {
  font-style: normal;
  color: var(--cyan);
  display: block;
}

.cover-sub {
  font-size: 14px;
  color: var(--muted);
  line-height: 1.8;
  margin: 0 0 36px;
}

.cover-pills {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}

.pill {
  background: var(--navy2);
  border: 1px solid var(--border);
  border-radius: 20px;
  padding: 5px 14px;
  font-size: 12px;
  color: var(--muted);
}

.pill strong { color: var(--text); }

/* 오른쪽 패널 — 문서 시각화 */
.viz-box {
  width: 100%;
  background: var(--navy2);
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
}

.viz-head {
  background: var(--cyan);
  padding: 8px 14px;
  font-size: 10px;
  letter-spacing: .18em;
  color: var(--navy);
  font-family: 'Consolas','Courier New',monospace;
  font-weight: 700;
}

.viz-body {
  padding: 16px 14px;
  display: flex;
  flex-direction: column;
  gap: 7px;
}

.vl {
  height: 8px;
  border-radius: 2px;
  background: var(--navy3);
}
.vl.sm   { width: 55%; }
.vl.md   { width: 78%; }
.vl.full { width: 100%; }
.vl.red  { background: rgba(230,57,70,0.2); width: 66%; }
.vl.grn  { background: rgba(34,197,94,0.2); width: 80%; }

.viz-beam {
  height: 2px;
  background: linear-gradient(90deg, transparent, var(--cyan), transparent);
  margin: 2px 0;
  box-shadow: 0 0 12px var(--cyan);
}

.viz-status {
  background: rgba(34,197,94,0.1);
  border: 1px solid rgba(34,197,94,0.35);
  border-radius: 4px;
  padding: 6px 10px;
  font-size: 10px;
  color: var(--green);
  font-family: 'Consolas','Courier New',monospace;
  letter-spacing: .12em;
  margin-top: 4px;
}

.cover-bar {
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 4px;
  background: linear-gradient(90deg, var(--red), var(--cyan), var(--green));
  z-index: 2;
}

/* ── 2. 필요성 ──────────────────────────────────────
   레이아웃: 번호가 크게 왼쪽에, 내용이 오른쪽에 (세로 리스트) */
section.problem { padding: 0; }

.problem-header {
  background: var(--navy2);
  padding: 28px 64px 24px;
  border-bottom: 1px solid var(--border);
}

.problem-rows {
  flex: 1;
  display: flex;
  flex-direction: column;
}

.prob-row {
  flex: 1;
  display: flex;
  align-items: stretch;
  border-bottom: 1px solid var(--border);
}
.prob-row:last-child { border-bottom: none; }
.prob-row:nth-child(even) { background: rgba(22,37,55,0.5); }

.prob-num-block {
  width: 96px;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--red);
  font-size: 36px;
  font-weight: 900;
  color: #fff;
  flex-shrink: 0;
  letter-spacing: -.02em;
}
.prob-row:nth-child(even) .prob-num-block {
  background: #B01C28;
}

.prob-content {
  flex: 1;
  padding: 14px 28px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  gap: 4px;
}

.prob-content h3 {
  font-size: 15px;
  font-weight: 700;
  color: var(--text);
  margin: 0;
}

.prob-content p {
  font-size: 12.5px;
  color: var(--muted);
  line-height: 1.65;
  margin: 0;
}

/* ── 3. 목적 ────────────────────────────────────────
   레이아웃: 전폭 수평 카드 (아이콘 | 제목 | 설명) */
section.purpose { padding: 0; }

.purpose-header {
  background: var(--navy2);
  padding: 28px 64px 24px;
  border-bottom: 1px solid var(--border);
}

.purpose-cards {
  flex: 1;
  display: flex;
  flex-direction: column;
}

.pur-card {
  flex: 1;
  display: flex;
  align-items: center;
  gap: 0;
  border-bottom: 1px solid var(--border);
  position: relative;
  overflow: hidden;
}
.pur-card:last-child { border-bottom: none; }
.pur-card:nth-child(even) { background: rgba(22,37,55,0.4); }

.pur-accent-bar {
  width: 5px;
  align-self: stretch;
  flex-shrink: 0;
}

.pur-icon-box {
  width: 72px;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 26px;
  flex-shrink: 0;
  margin: 0 8px;
}

.pur-divider {
  width: 1px;
  height: 40px;
  background: var(--border);
  flex-shrink: 0;
}

.pur-title {
  width: 230px;
  flex-shrink: 0;
  padding: 16px 20px;
  font-size: 14px;
  font-weight: 700;
  line-height: 1.4;
}

.pur-desc {
  flex: 1;
  padding: 16px 28px;
  font-size: 12.5px;
  color: var(--muted);
  line-height: 1.7;
  border-left: 1px solid var(--border);
}

.pur-index {
  width: 72px;
  text-align: right;
  padding-right: 24px;
  font-size: 32px;
  font-weight: 900;
  opacity: .08;
  font-family: 'Consolas','Courier New',monospace;
  flex-shrink: 0;
}

/* ── 4. 프로세스 ────────────────────────────────────
   레이아웃: 지그재그 (홀수=위, 짝수=아래) */
section.process {
  padding: 48px 52px 56px;
}

.process-area {
  flex: 1;
  position: relative;
  display: flex;
  align-items: center;
}

.process-steps {
  display: flex;
  width: 100%;
  position: relative;
}

/* 중앙 수평선 */
.process-steps::before {
  content: '';
  position: absolute;
  top: 50%; left: 0; right: 0;
  height: 1px;
  background: var(--border);
  transform: translateY(-50%);
  z-index: 0;
}

.p-step {
  flex: 1;
  display: flex;
  flex-direction: column;
  align-items: center;
  position: relative;
  z-index: 1;
}

/* 홀수 step: 원 위쪽에 텍스트 */
.p-step:nth-child(odd) {
  flex-direction: column-reverse;
  padding-bottom: 20px;
}
.p-step:nth-child(even) {
  flex-direction: column;
  padding-top: 20px;
}

.p-circle {
  width: 62px; height: 62px;
  border-radius: 50%;
  background: var(--navy2);
  border: 2px solid var(--cyan);
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 24px;
  position: relative;
  flex-shrink: 0;
}

.p-badge {
  position: absolute;
  top: -4px; right: -4px;
  width: 20px; height: 20px;
  border-radius: 50%;
  background: var(--cyan);
  color: var(--navy);
  font-size: 8px;
  font-weight: 700;
  display: flex;
  align-items: center;
  justify-content: center;
  font-family: 'Consolas','Courier New',monospace;
}

/* 수직 연결선 */
.p-step:nth-child(odd) .p-vline {
  width: 1px; height: 30px;
  background: var(--border);
  margin-top: 0; margin-bottom: 0;
}
.p-step:nth-child(even) .p-vline {
  width: 1px; height: 30px;
  background: var(--border);
}

.p-text { text-align: center; padding: 8px 4px; }
.p-text .num { font-size: 9px; color: var(--cyan); letter-spacing: .15em; font-family: 'Consolas','Courier New',monospace; }
.p-text .title { font-size: 12.5px; font-weight: 700; line-height: 1.4; margin: 2px 0; }
.p-text .desc { font-size: 10.5px; color: var(--muted); line-height: 1.55; }

/* ── 5. 결과물 ──────────────────────────────────────
   레이아웃: 상단 기본정보 + 2열 결과 패널 */
section.result { padding: 36px 52px 48px; }

.info-strip {
  display: flex;
  background: var(--navy3);
  border: 1px solid var(--border);
  border-radius: 6px;
  padding: 10px 0;
  margin-bottom: 14px;
  flex-shrink: 0;
}

.info-field {
  flex: 1;
  padding: 4px 18px;
  display: flex;
  flex-direction: column;
  gap: 2px;
  border-right: 1px solid var(--border);
}
.info-field:last-child { border-right: none; }
.info-label { font-size: 8px; color: var(--muted); letter-spacing: .14em; font-family: 'Consolas','Courier New',monospace; }
.info-val   { font-size: 11px; font-weight: 700; }

.result-panels {
  flex: 1;
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 14px;
}

.res-panel {
  background: var(--navy2);
  border: 1px solid var(--border);
  border-radius: 8px;
  overflow: hidden;
  display: flex;
  flex-direction: column;
}

.res-head {
  padding: 10px 16px;
  font-size: 12.5px;
  font-weight: 700;
  border-bottom: 1px solid var(--border);
  display: flex;
  align-items: center;
  gap: 8px;
}

.res-head.ok  { background: rgba(34,197,94,0.08);  color: var(--green); }
.res-head.bad { background: rgba(230,57,70,0.08); color: var(--red); }

.res-item {
  padding: 12px 16px;
  border-bottom: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  gap: 5px;
}
.res-item:last-child { border-bottom: none; }
.ri-title { font-size: 11.5px; font-weight: 700; }
.ri-meta  { font-size: 10px; color: var(--muted); }
.ri-op {
  padding: 7px 10px;
  border-radius: 4px;
  font-size: 10.5px;
  line-height: 1.6;
}
.ri-op.ok  { background: rgba(34,197,94,0.08);  color: #86EFAC; border-left: 2px solid var(--green); }
.ri-op.bad { background: rgba(230,57,70,0.08); color: #FCA5A5; border-left: 2px solid var(--red); }

/* ── 6. 기대효과 ────────────────────────────────────
   레이아웃: 좌측 타이틀 패널 + 우측 2×2 대형 수치 */
section.effect { padding: 0; }

.effect-left {
  position: absolute;
  left: 0; top: 0; bottom: 0;
  width: 220px;
  background: var(--navy2);
  border-right: 1px solid var(--border);
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  text-align: center;
  padding: 32px 24px;
  z-index: 1;
}

.effect-left::before {
  content: '';
  position: absolute;
  left: 0; top: 0; bottom: 0;
  width: 5px;
  background: var(--cyan);
}

.effect-left h2 {
  font-size: 22px;
  margin: 0 0 10px;
}
.effect-left p {
  font-size: 13px;
  color: var(--muted);
  line-height: 1.6;
  margin: 0;
}

.effect-grid {
  position: absolute;
  left: 220px; right: 0; top: 0; bottom: 0;
  display: grid;
  grid-template-columns: 1fr 1fr;
  grid-template-rows: 1fr 1fr;
  z-index: 1;
}

.eff-card {
  display: flex;
  flex-direction: column;
  justify-content: flex-end;
  padding: 28px 32px;
  border-right: 1px solid var(--border);
  border-bottom: 1px solid var(--border);
  position: relative;
  overflow: hidden;
}
.eff-card:nth-child(even) { border-right: none; }
.eff-card:nth-child(3),
.eff-card:nth-child(4) { border-bottom: none; }
.eff-card:nth-child(2) { background: rgba(22,37,55,0.5); }
.eff-card:nth-child(3) { background: rgba(22,37,55,0.5); }

.eff-bg-num {
  position: absolute;
  top: -10px; right: 16px;
  font-size: 110px;
  font-weight: 900;
  opacity: .04;
  line-height: 1;
  font-family: 'Consolas','Courier New',monospace;
  pointer-events: none;
}

.eff-num {
  font-size: 54px;
  font-weight: 900;
  letter-spacing: -.03em;
  line-height: 1;
  margin-bottom: 8px;
}
.eff-card:nth-child(1) .eff-num { color: var(--cyan); }
.eff-card:nth-child(2) .eff-num { color: var(--red); }
.eff-card:nth-child(3) .eff-num { color: var(--gold); }
.eff-card:nth-child(4) .eff-num { color: var(--green); }

.eff-title { font-size: 14px; font-weight: 700; margin-bottom: 6px; }
.eff-desc  { font-size: 11.5px; color: var(--muted); line-height: 1.6; }

.eff-bar {
  position: absolute;
  bottom: 0; left: 0; right: 0;
  height: 3px;
}
.eff-card:nth-child(1) .eff-bar { background: var(--cyan); }
.eff-card:nth-child(2) .eff-bar { background: var(--red); }
.eff-card:nth-child(3) .eff-bar { background: var(--gold); }
.eff-card:nth-child(4) .eff-bar { background: var(--green); }
</style>

---

<!-- _class: cover -->

<div class="cover-left">
  <div class="cover-tag">AI · 기계설비 성능점검 · 2026</div>
  <h1>기계설비<br>성능점검 보고서<br><em>AI 검토자문 시스템</em></h1>
  <p class="cover-sub">방대한 점검 항목을 100% 누락 없이 추출하고,<br>계측 사진 수치를 AI가 직접 판독하여<br>보고서의 적정성을 자동으로 검증합니다.</p>
  <div class="cover-pills">
    <span class="pill"><strong>🤖</strong> Google Gemini AI</span>
    <span class="pill"><strong>📊</strong> 자동 판정</span>
    <span class="pill"><strong>📥</strong> 엑셀 출력</span>
  </div>
</div>

<div class="cover-right">
  <div class="viz-box">
    <div class="viz-head">▶  AI SCANNING</div>
    <div class="viz-body">
      <div class="vl full"></div>
      <div class="vl sm"></div>
      <div class="viz-beam"></div>
      <div class="vl md"></div>
      <div class="vl red"></div>
      <div class="vl full"></div>
      <div class="viz-beam"></div>
      <div class="vl grn"></div>
      <div class="vl sm"></div>
      <div class="vl md"></div>
      <div class="vl red"></div>
      <div class="viz-beam"></div>
      <div class="vl full"></div>
      <div class="vl grn"></div>
      <div class="vl sm"></div>
      <div class="viz-status">● 분석 중 · 전체 항목 추출</div>
    </div>
  </div>
</div>

<div class="cover-bar"></div>

---

<!-- _class: problem -->

<div class="problem-header">
  <div class="eyebrow">배경 및 필요성</div>
  <h2>왜 지금 이 시스템이 필요한가?</h2>
</div>

<div class="problem-rows">
  <div class="prob-row">
    <div class="prob-num-block">01</div>
    <div class="prob-content">
      <h3>수작업 검토의 한계</h3>
      <p>성능점검 보고서 한 건에 수십~수백 개의 점검 항목이 포함됩니다. 전문가가 수작업으로 검토할 경우 시간·인력 부족으로 누락이 발생하고, 형식적 검토에 그치는 경우가 많습니다.</p>
    </div>
  </div>
  <div class="prob-row">
    <div class="prob-num-block">02</div>
    <div class="prob-content">
      <h3>사진 속 수치 미검증</h3>
      <p>보고서에 첨부된 계측기 사진의 수치와 텍스트 기재값이 일치하는지 육안으로 확인하기 어렵습니다. 불일치가 있어도 검토자가 놓치는 경우가 빈번합니다.</p>
    </div>
  </div>
  <div class="prob-row">
    <div class="prob-num-block">03</div>
    <div class="prob-content">
      <h3>검토자 전문성 편차</h3>
      <p>검토자의 경험과 전문 지식에 따라 동일한 보고서에 대해 판정 결과가 달라질 수 있습니다. 일관된 기준 적용이 어렵습니다.</p>
    </div>
  </div>
  <div class="prob-row">
    <div class="prob-num-block">04</div>
    <div class="prob-content">
      <h3>사후 보완 소통 비효율</h3>
      <p>보완이 필요한 항목을 공문으로 통보하고 재검토하는 과정이 반복됩니다. 의견 수정과 재확인에 많은 시간이 소요됩니다.</p>
    </div>
  </div>
</div>

---

<!-- _class: purpose -->

<div class="purpose-header">
  <div class="eyebrow">시스템 목적</div>
  <h2>이 시스템이 해결하는 것</h2>
</div>

<div class="purpose-cards">
  <div class="pur-card">
    <div class="pur-accent-bar" style="background:#4CC9F0"></div>
    <div class="pur-icon-box">🔍</div>
    <div class="pur-divider"></div>
    <div class="pur-title" style="color:#4CC9F0">점검 항목<br>100% 자동 추출</div>
    <div class="pur-desc">업로드된 PDF 또는 이미지 전체를 AI가 처음부터 끝까지 분석하여 단 하나의 항목도 누락 없이 추출합니다.</div>
    <div class="pur-index">01</div>
  </div>
  <div class="pur-card">
    <div class="pur-accent-bar" style="background:#E63946"></div>
    <div class="pur-icon-box">📷</div>
    <div class="pur-divider"></div>
    <div class="pur-title" style="color:#E63946">계측기 사진 수치<br>자동 판독</div>
    <div class="pur-desc">첨부된 계측기 사진에서 측정값을 직접 읽어 보고서의 텍스트 기재값과 자동으로 교차 검증합니다.</div>
    <div class="pur-index">02</div>
  </div>
  <div class="pur-card">
    <div class="pur-accent-bar" style="background:#F7B731"></div>
    <div class="pur-icon-box">⚖️</div>
    <div class="pur-divider"></div>
    <div class="pur-title" style="color:#F7B731">적정 / 보완필요<br>자동 판정</div>
    <div class="pur-desc">점검 기준과 실측 수치를 근거로 각 항목에 대해 [적정] 또는 [보완필요] 판정을 일관되게 내립니다.</div>
    <div class="pur-index">03</div>
  </div>
  <div class="pur-card">
    <div class="pur-accent-bar" style="background:#22C55E"></div>
    <div class="pur-icon-box">📝</div>
    <div class="pur-divider"></div>
    <div class="pur-title" style="color:#22C55E">전문가 수준<br>검토자문의견 생성</div>
    <div class="pur-desc">판정 근거와 보완 방향을 포함한 상세 의견을 자동 생성하고, 수정 지시를 반영하여 즉시 재작성합니다.</div>
    <div class="pur-index">04</div>
  </div>
</div>

---

<!-- _class: process -->

<div class="eyebrow">작동 프로세스</div>
<h2>업로드에서 결과 출력까지 — 6단계</h2>
<hr>

<div class="process-area">
  <div class="process-steps">

    <div class="p-step">
      <div class="p-text">
        <div class="num">STEP 01</div>
        <div class="title">보고서<br>업로드</div>
        <div class="desc">PDF 또는<br>이미지 다중 업로드</div>
      </div>
      <div class="p-vline"></div>
      <div class="p-circle">📄<div class="p-badge">01</div></div>
    </div>

    <div class="p-step">
      <div class="p-circle">🤖<div class="p-badge">02</div></div>
      <div class="p-vline"></div>
      <div class="p-text">
        <div class="num">STEP 02</div>
        <div class="title">AI 전체<br>문서 분석</div>
        <div class="desc">Gemini AI<br>모델 선택 가능</div>
      </div>
    </div>

    <div class="p-step">
      <div class="p-text">
        <div class="num">STEP 03</div>
        <div class="title">수치<br>판독 비교</div>
        <div class="desc">사진값 vs 기재값<br>교차검증</div>
      </div>
      <div class="p-vline"></div>
      <div class="p-circle">🔢<div class="p-badge">03</div></div>
    </div>

    <div class="p-step">
      <div class="p-circle">⚖️<div class="p-badge">04</div></div>
      <div class="p-vline"></div>
      <div class="p-text">
        <div class="num">STEP 04</div>
        <div class="title">항목별<br>적정성 판정</div>
        <div class="desc">점검 기준<br>자동 적용</div>
      </div>
    </div>

    <div class="p-step">
      <div class="p-text">
        <div class="num">STEP 05</div>
        <div class="title">의견 확인<br>및 수정</div>
        <div class="desc">항목별<br>수정 지시 반영</div>
      </div>
      <div class="p-vline"></div>
      <div class="p-circle">✏️<div class="p-badge">05</div></div>
    </div>

    <div class="p-step">
      <div class="p-circle">📥<div class="p-badge">06</div></div>
      <div class="p-vline"></div>
      <div class="p-text">
        <div class="num">STEP 06</div>
        <div class="title">엑셀<br>다운로드</div>
        <div class="desc">검토자문의견서<br>즉시 출력</div>
      </div>
    </div>

  </div>
</div>

---

<!-- _class: result -->

<div class="eyebrow">결과물 미리보기</div>
<h2 style="margin-bottom:12px">화면 구성 및 출력 내용</h2>

<div class="info-strip">
  <div class="info-field"><div class="info-label">건축물명</div><div class="info-val">○○빌딩</div></div>
  <div class="info-field"><div class="info-label">성능점검업체</div><div class="info-val">○○기계설비(주)</div></div>
  <div class="info-field"><div class="info-label">대표자</div><div class="info-val">홍길동</div></div>
  <div class="info-field"><div class="info-label">건물 주소</div><div class="info-val">서울특별시 강남구 ○○로 123</div></div>
  <div class="info-field"><div class="info-label">관리주체</div><div class="info-val">○○자산관리(주)</div></div>
</div>

<div class="result-panels">
  <div class="res-panel">
    <div class="res-head ok">🟢 적정 항목</div>
    <div class="res-item">
      <div class="ri-title">냉동기 성능 점검 — 냉수 출구 온도</div>
      <div class="ri-meta">기준: 설정치 ±1℃  |  판독값: 6.8℃  |  기재값: 6.8℃</div>
      <div class="ri-op ok">사진에서 확인된 냉수 출구 온도 6.8℃는 설정치 7℃ 대비 허용 범위 이내로, 업체의 [적합] 판정이 적정합니다.</div>
    </div>
    <div class="res-item">
      <div class="ri-title">공기조화기 — 필터 압력차</div>
      <div class="ri-meta">기준: 초기값 대비 2배 이내  |  판독값: 45Pa  |  기재값: 45Pa</div>
      <div class="ri-op ok">측정값 45Pa은 초기 설정값 대비 기준 이내이며, 업체의 [적합] 판정은 적정합니다.</div>
    </div>
  </div>
  <div class="res-panel">
    <div class="res-head bad">🔴 보완필요 항목</div>
    <div class="res-item">
      <div class="ri-title">급탕 펌프 — 토출 압력</div>
      <div class="ri-meta">기준: 설계값 ±5%  |  판독값: 판독 불가  |  기재값: 0.32 MPa</div>
      <div class="ri-op bad">사진에서 계측기 수치를 확인할 수 없어 기재값(0.32 MPa)의 정확성을 검증할 수 없습니다. 재첨부 필요.</div>
    </div>
    <div class="res-item">
      <div class="ri-title">보일러 — 배기가스 온도</div>
      <div class="ri-meta">기준: 230℃ 이하  |  판독값: 248℃  |  기재값: 220℃</div>
      <div class="ri-op bad">사진 판독값 248℃가 기준치(230℃)를 초과합니다. 기재값(220℃)과 불일치, [적합] 판정 부적정.</div>
    </div>
  </div>
</div>

---

<!-- _class: effect -->

<div class="effect-left">
  <h2>기대<br>효과</h2>
  <p>도입으로<br>달라지는 것</p>
</div>

<div class="effect-grid">
  <div class="eff-card">
    <div class="eff-bg-num">100</div>
    <div class="eff-num">100%</div>
    <div class="eff-title">점검 항목 완전 추출</div>
    <div class="eff-desc">AI가 문서 전체를 분석해<br>항목 누락 없이 모든 내용을 검토합니다.</div>
    <div class="eff-bar"></div>
  </div>
  <div class="eff-card">
    <div class="eff-bg-num">90</div>
    <div class="eff-num">↓90%</div>
    <div class="eff-title">검토 소요 시간 단축</div>
    <div class="eff-desc">수일이 걸리던 보고서 검토를<br>수분 내에 완료합니다.</div>
    <div class="eff-bar"></div>
  </div>
  <div class="eff-card">
    <div class="eff-bg-num">0</div>
    <div class="eff-num">0 편차</div>
    <div class="eff-title">일관된 기준 판정 적용</div>
    <div class="eff-desc">검토자 경험·전문성과 무관하게<br>동일한 점검 기준이 적용됩니다.</div>
    <div class="eff-bar"></div>
  </div>
  <div class="eff-card">
    <div class="eff-bg-num">→</div>
    <div class="eff-num">즉시</div>
    <div class="eff-title">엑셀 검토자문의견서</div>
    <div class="eff-desc">검토 결과를 즉시 엑셀로 내려받아<br>공문 작성에 바로 활용합니다.</div>
    <div class="eff-bar"></div>
  </div>
</div>
