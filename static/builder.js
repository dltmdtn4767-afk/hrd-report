/* =====================================================
   보고서 빌더 (builder.js)
   ===================================================== */

// ── 템플릿 업로드 ────────────────────────────────────
async function uploadTemplate(input) {
  const file = input.files[0];
  if (!file) return;
  const sid = window.sessionId || '';
  if (!sid) { alert('먼저 데이터 파일을 업로드하세요.'); return; }

  const statusEl = document.getElementById('tplStatus');
  statusEl.textContent = '⏳ 업로드 중...';
  statusEl.style.color = '#fbbf24';

  const fd = new FormData();
  fd.append('file', file);

  try {
    const r = await fetch(`/api/upload_template/${sid}`, { method: 'POST', body: fd });
    const json = await r.json();
    if (json.ok) {
      statusEl.textContent = `✅ ${file.name} (${json.slide_count}장)`;
      statusEl.style.color = '#34d399';
    } else {
      throw new Error(JSON.stringify(json));
    }
  } catch(e) {
    statusEl.textContent = '❌ 업로드 실패';
    statusEl.style.color = '#f87171';
    console.error(e);
  }
  input.value = '';
}

// ── 빌더 전역 상태 ──────────────────────────────────
const builderState = {
  activeIdx: 0,
  slides: [],       // 슬라이드 목록
  quantGroups: [],  // 정량 그룹 [{id, title, questionIds:[]}]
  qualEdits: {},    // {qId: [{text, merged}]}
  previewCharts: {} // 차트 인스턴스 캐시
};

let nextGroupId = 1;

// ── 슬라이드 타입 메타 ──────────────────────────────
const SLIDE_TYPES = {
  cover:        { label: '앞표지',           icon: '🏠', editable: true },
  toc:          { label: '목차',             icon: '📋', editable: false },
  section1:     { label: '섹션 I — 과정 개요', icon: '📌', editable: false },
  overview:     { label: '과정 개요 표',      icon: '📄', editable: true },
  schedule:     { label: '교육 일정표',       icon: '📅', editable: true },
  photo:        { label: '포토 섹션',         icon: '🖼', editable: false },
  section2:     { label: '섹션 II — 만족도'  , icon: '📌', editable: false },
  summary:      { label: 'Executive Summary', icon: '📊', editable: true },
  quant_chart:  { label: '정량 차트 그룹',    icon: '📈', editable: true },
  qual_text:    { label: '정성 평가 결과',    icon: '💬', editable: true },
  back_cover:   { label: '뒤표지',            icon: '🔚', editable: false },
};

// ── 초기화 (업로드 완료 후 호출) ─────────────────────
function initBuilder(dashboardData) {
  const s = dashboardData.summary || {};
  const multi = dashboardData.multi_result || {};
  const combined = multi.combined || {};
  const questions = combined.questions || [];
  const sessions = multi.sessions || [];

  // 슬라이드 초기 구성 (템플릿 순서대로)
  builderState.slides = [
    {
      type: 'cover',
      data: { company: s.company || '', course: s.course_name || '' }
    },
    { type: 'toc', data: {} },
    { type: 'section1', data: {} },
    {
      type: 'overview',
      data: {
        rows: [
          { key: '교육목적', val: '' },
          { key: '교육대상', val: '' },
          { key: '교육일시', val: '' },
          { key: '교육장소', val: '' },
          { key: '강사명',   val: '' },
          { key: '응답인원', val: `${s.response_count || 0}명` },
        ]
      }
    },
    {
      type: 'schedule',
      data: {
        rows: sessions.map(sess => ({
          label: sess.label || '',
          date: '', place: '', count: `${sess.response_count || 0}명`
        }))
      }
    },
    { type: 'section2', data: {} },
    {
      type: 'summary',
      data: {
        categories: combined.categories || [],
        overall: s.overall_average || 0,
        response_count: s.response_count || 0,
      }
    },
  ];

  // 정량 기본 그룹: 전체 문항 → 6개씩 분할
  builderState.quantGroups = [];
  nextGroupId = 1;
  const chunkSize = 6;
  for (let i = 0; i < questions.length; i += chunkSize) {
    const chunk = questions.slice(i, i + chunkSize);
    builderState.quantGroups.push({
      id: nextGroupId++,
      title: `정량 결과 ${Math.floor(i / chunkSize) + 1}`,
      questions: chunk,
      allQuestions: questions,
    });
  }
  // 각 정량 그룹 슬라이드 추가
  builderState.quantGroups.forEach(g => {
    builderState.slides.push({ type: 'quant_chart', groupId: g.id, data: {} });
  });

  // 정성 슬라이드
  builderState.slides.push({ type: 'qual_text', data: {} });
  builderState.slides.push({ type: 'back_cover', data: {} });

  renderSlideList();
  selectSlide(0);
  refreshBuilderMiniList();
}

// ── 정량 탭에서 빌더에 그룹 추가 ─────────────────────
function addQuantGroupToBuilder(title, questions) {
  const all = window._builderData?.multi_result?.combined?.questions || questions;
  const newG = {
    id: nextGroupId++,
    title,
    questions,
    allQuestions: all,
  };
  builderState.quantGroups.push(newG);

  // qual_text / back_cover 앞에 삽입
  const qualIdx = builderState.slides.findIndex(s => s.type === 'qual_text');
  const insertAt = qualIdx >= 0 ? qualIdx : builderState.slides.length - 1;
  builderState.slides.splice(insertAt, 0, { type: 'quant_chart', groupId: newG.id, data: {} });

  renderSlideList();
  refreshBuilderMiniList();

  // 알림
  showToast(`✅ "${title}" 슬라이드가 빌더에 추가됐습니다`);
}

// ── 정성 탭에서 빌더에 슬라이드 추가 ─────────────────
function addQualSlideToBuilder(oeId, label, groups) {
  // 이미 qual_text 슬라이드에 누적 저장됨 (qualEdits에 추가)
  if (!builderState.qualEdits[oeId]) {
    builderState.qualEdits[oeId] = {};
    groups.forEach((g, i) => {
      builderState.qualEdits[oeId][i] = g.label;
    });
  }
  refreshBuilderMiniList();
  showToast(`✅ "${label}" 정성 내용이 빌더에 추가됐습니다`);
}

// ── 퀀트 탭 하단 미니 슬라이드 목록 ─────────────────
function refreshBuilderMiniList() {
  const panel = document.getElementById('builderSlidesPanel');
  const list  = document.getElementById('builderSlidesList');
  if (!panel || !list) return;

  const count = builderState.slides.length;
  panel.style.display = count > 0 ? '' : 'none';

  list.innerHTML = builderState.slides.map((slide, i) => {
    const meta = SLIDE_TYPES[slide.type] || { label: slide.type, icon: '📄' };
    let label = meta.label;
    if (slide.type === 'quant_chart') {
      const g = builderState.quantGroups.find(g => g.id === slide.groupId);
      if (g) label = g.title;
    }
    return `<span class="mini-slide-chip">${i+1}. ${meta.icon} ${label}</span>`;
  }).join('');
}

// ── 토스트 알림 ───────────────────────────────────────
function showToast(msg) {
  let t = document.getElementById('builderToast');
  if (!t) {
    t = document.createElement('div');
    t.id = 'builderToast';
    t.style.cssText = `position:fixed;bottom:24px;right:24px;background:#1e293b;color:#fff;padding:10px 18px;border-radius:8px;font-size:13px;font-weight:600;z-index:9999;opacity:0;transition:opacity .3s,transform .3s;transform:translateY(10px)`;
    document.body.appendChild(t);
  }
  t.textContent = msg;
  t.style.opacity = '1';
  t.style.transform = 'translateY(0)';
  clearTimeout(t._timer);
  t._timer = setTimeout(() => {
    t.style.opacity = '0';
    t.style.transform = 'translateY(10px)';
  }, 2500);
}


// ── 슬라이드 목록 렌더 ──────────────────────────────
function renderSlideList() {
  const list = document.getElementById('slideList');
  if (!list) return;
  list.innerHTML = '';
  builderState.slides.forEach((slide, idx) => {
    const meta = SLIDE_TYPES[slide.type] || { label: slide.type, icon: '📄' };
    const div = document.createElement('div');
    div.className = 'slide-thumb' + (idx === builderState.activeIdx ? ' active' : '');
    div.dataset.idx = idx;

    // 썸네일 미니 프리뷰 (간략 HTML)
    div.innerHTML = `
      <div class="slide-thumb-inner" style="pointer-events:none;">
        ${renderPreviewHTML(slide, true)}
      </div>
      <div class="slide-thumb-label">
        <span class="slide-thumb-num">${idx + 1}</span>
        <span>${meta.icon} ${meta.label}</span>
      </div>`;
    div.addEventListener('click', () => selectSlide(idx));
    list.appendChild(div);
  });

  // + 정량 슬라이드 추가 버튼
  const addBtn = document.createElement('button');
  addBtn.className = 'slide-add-btn';
  addBtn.innerHTML = '＋ 정량 차트 슬라이드 추가';
  addBtn.onclick = addQuantSlide;
  list.appendChild(addBtn);
}

// ── 슬라이드 선택 ────────────────────────────────────
function selectSlide(idx) {
  builderState.activeIdx = idx;
  // 썸네일 active 업데이트
  document.querySelectorAll('.slide-thumb').forEach((el, i) => {
    el.classList.toggle('active', i === idx);
  });
  const slide = builderState.slides[idx];
  renderPreview(slide);
  renderEditor(slide, idx);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 미리보기 (오른쪽 상단 큰 뷰)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function renderPreview(slide) {
  const el = document.getElementById('slidePreview');
  if (!el) return;
  Object.values(builderState.previewCharts).forEach(c => { try { c.destroy(); } catch(e){} });
  builderState.previewCharts = {};
  el.innerHTML = renderPreviewHTML(slide, false);
  if (slide.type === 'summary') renderSummaryPreviewChart(slide.data);
  if (slide.type === 'quant_chart') renderQuantPreviewChart(slide.groupId);
  if (slide.type === 'custom_quant') renderCustomQuantPreviewChart(slide);
}

function renderPreviewHTML(slide, mini) {
  const d = slide.data || {};
  const scale = mini ? 'style="transform:scale(0.28);transform-origin:top left;width:357%;height:357%"' : '';
  const font = "'Nanum Gothic',sans-serif";
  switch (slide.type) {
    case 'cover':
      return `<div class="builder-slide-preview" ${scale} style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:#fff;display:flex;flex-direction:column;justify-content:center;align-items:center;text-align:center;min-height:300px">
        <div style="font-size:16px;font-weight:300;opacity:.7;margin-bottom:8px">${d.company || '고객사명'}</div>
        <div style="font-size:22px;font-weight:800;margin-bottom:6px">${d.course || '과정명'}</div>
        <div style="font-size:14px;font-weight:700;color:#36A86F">과정운영 결과보고서</div>
      </div>`;
    case 'toc':
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">목 차</div>
        <table class="tpl-table-c" style="width:60%">
          <tr><td class="key">Ⅰ</td><td>과정 개요</td></tr>
          <tr><td class="key">Ⅱ</td><td>만족도 결과</td></tr>
        </table>
      </div>`;
    case 'section1':
      return `<div class="builder-slide-preview" ${scale} style="display:flex;flex-direction:column;justify-content:center;align-items:center;min-height:300px">
        <div class="section-num">Ⅰ</div>
        <div class="section-name">과정 개요</div>
        <div class="section-sub">실시 개요</div>
      </div>`;
    case 'section2':
      return `<div class="builder-slide-preview" ${scale} style="display:flex;flex-direction:column;justify-content:center;align-items:center;min-height:300px">
        <div class="section-num">Ⅱ</div>
        <div class="section-name">만족도 결과</div>
        <div class="section-sub">만족도 결과</div>
      </div>`;
    case 'overview':
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">과정 개요</div>
        <table class="tpl-table-c">${(d.rows||[]).map(r=>`<tr><td class="key">${r.key}</td><td>${r.val||'—'}</td></tr>`).join('')}</table>
      </div>`;
    case 'schedule':
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">교육 일정표</div>
        <table class="tpl-table-b">
          <thead><tr><th>차수</th><th>일자</th><th>장소</th><th>인원</th></tr></thead>
          <tbody>${(d.rows||[]).map(r=>`<tr><td>${r.label}</td><td>${r.date||'—'}</td><td>${r.place||'—'}</td><td>${r.count||'—'}</td></tr>`).join('')}</tbody>
        </table>
      </div>`;
    case 'summary': {
      const cats = d.categories || [];
      let tableHtml = '<table class="tpl-table-a"><thead><tr>';
      cats.forEach(c => { tableHtml += `<th>${c.name.length>8?c.name.slice(0,8)+'…':c.name}</th>`; });
      tableHtml += '<th class="highlight">과정 전반</th></tr></thead><tbody><tr>';
      cats.forEach(c => { tableHtml += `<td>${(c.avg||0).toFixed(2)}</td>`; });
      tableHtml += `<td><strong>${(d.overall||0).toFixed(2)}</strong></td></tr>`;
      tableHtml += `<tr class="avg-row"><td colspan="${cats.length+1}" style="text-align:left">전체 평균: <strong>${(d.overall||0).toFixed(2)}</strong>점 / 응답: ${d.response_count||0}명</td></tr></tbody></table>`;
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">Executive Summary</div>
        <div style="height:55%" id="pvSummaryChart"><canvas id="pvSummaryCanvas" style="max-height:100%"></canvas></div>
        ${tableHtml}
      </div>`;
    }
    case 'quant_chart': {
      const g = builderState.quantGroups.find(g => g.id === slide.groupId);
      let tableHtml = '';
      if (g && g.questions) {
        tableHtml = '<table class="tpl-table-b"><thead><tr><th>항목</th><th>문항</th><th>평균</th><th>응답</th></tr></thead><tbody>';
        g.questions.forEach(q => {
          const cls = q.avg >= 4.5 ? 'score-high' : q.avg < 3.5 ? 'score-low' : '';
          tableHtml += `<tr><td></td><td style="text-align:left;font-size:10px">${q.label}</td><td class="${cls}">${q.avg?.toFixed(2)||'-'}</td><td>${q.count||'-'}</td></tr>`;
        });
        tableHtml += '</tbody></table>';
      }
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">${g ? g.title : '정량 평가'}</div>
        <div style="height:50%" id="pvQuantChart${slide.groupId}"><canvas id="pvQuantCanvas${slide.groupId}" style="max-height:100%"></canvas></div>
        ${tableHtml}
      </div>`;
    }
    case 'custom_quant': {
      const cs = d || {};
      const groups = cs.groups || [];
      let chartHtml = '';
      if (cs.chartType !== 'none') {
        chartHtml = `<div style="height:45%" id="pvCustomChart_${slide.customSlideId}"><canvas id="pvCustomCanvas_${slide.customSlideId}" style="max-height:100%"></canvas></div>`;
      }
      let tableHtml = '';
      if (cs.tableStyle === 'A') {
        tableHtml = '<table class="tpl-table-a"><thead><tr>';
        groups.forEach(g => { tableHtml += `<th>${g.name}</th>`; });
        const allAvg = groups.length ? (groups.reduce((s,g)=>s+g.avg,0)/groups.length) : 0;
        tableHtml += '<th class="highlight">과정 전반</th></tr></thead><tbody><tr>';
        groups.forEach(g => { tableHtml += `<td>${g.avg.toFixed(2)}</td>`; });
        tableHtml += `<td><strong>${allAvg.toFixed(2)}</strong></td></tr></tbody></table>`;
      } else if (cs.tableStyle === 'B') {
        tableHtml = '<table class="tpl-table-b"><thead><tr><th>항목</th><th>문항</th><th>평균</th><th>응답</th></tr></thead><tbody>';
        groups.forEach(g => {
          tableHtml += `<tr class="cat-row"><td colspan="2">${g.name}</td><td>${g.avg.toFixed(2)}</td><td></td></tr>`;
          (g.questions||[]).forEach(q => {
            const cls = q.avg >= 4.5 ? 'score-high' : q.avg < 3.5 ? 'score-low' : '';
            tableHtml += `<tr><td></td><td style="text-align:left;font-size:10px">${q.label}</td><td class="${cls}">${q.avg?.toFixed(2)||'-'}</td><td>${q.count||'-'}</td></tr>`;
          });
        });
        tableHtml += '</tbody></table>';
      } else if (cs.tableStyle === 'C') {
        tableHtml = '<table class="tpl-table-c"><tbody>';
        groups.forEach(g => { tableHtml += `<tr><td class="key">${g.name}</td><td>${g.avg.toFixed(2)}점 (${g.qIds?.length||0}문항)</td></tr>`; });
        tableHtml += '</tbody></table>';
      }
      const isAbove = cs.tablePos === 'above';
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">${cs.title || '커스텀 슬라이드'}</div>
        ${isAbove ? tableHtml : ''}
        ${chartHtml}
        ${!isAbove ? tableHtml : ''}
      </div>`;
    }
    case 'qual_text':
      return `<div class="builder-slide-preview" ${scale}>
        <div class="slide-title">정성 평가 결과</div>
        <div style="font-size:10px;color:#475569;line-height:1.8;padding:8px">
          공통응답 그룹핑 결과가 이 슬라이드에 표시됩니다.
        </div>
      </div>`;
    case 'back_cover':
      return `<div class="builder-slide-preview" ${scale} style="background:linear-gradient(135deg,#1a1a2e,#16213e);color:#fff;display:flex;justify-content:center;align-items:center;min-height:300px;font-size:18px;font-weight:300;letter-spacing:3px">
        EXPERT CONSULTING
      </div>`;
    default:
      return `<div class="builder-slide-preview" ${scale}><div class="slide-title">${slide.type}</div></div>`;
  }
}

function renderSummaryPreviewChart(data) {
  const canvas = document.getElementById('pvSummaryCanvas');
  if (!canvas || !data.categories) return;
  const cats = data.categories;
  const colors = ['#36A86F','#4A90D9','#E67E22','#9B59B6','#E74C3C','#1ABC9C'];
  builderState.previewCharts.summary = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: cats.map(c => c.name.length > 8 ? c.name.slice(0,8)+'…' : c.name),
      datasets: [{ data: cats.map(c => c.avg), backgroundColor: cats.map((_,i) => colors[i%colors.length]+'99'), borderColor: cats.map((_,i) => colors[i%colors.length]), borderWidth: 2, borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { min: 3, max: 5 }, x: { ticks: { font: { size: 9, family: "'Nanum Gothic'" } } } }
    }
  });
}

function renderQuantPreviewChart(groupId) {
  const g = builderState.quantGroups.find(g => g.id === groupId);
  if (!g) return;
  const canvas = document.getElementById(`pvQuantCanvas${groupId}`);
  if (!canvas) return;
  builderState.previewCharts[`quant_${groupId}`] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: g.questions.map(q => q.label.length > 12 ? q.label.slice(0,12)+'…' : q.label),
      datasets: [{ data: g.questions.map(q => q.avg), backgroundColor: '#36A86F99', borderColor: '#36A86F', borderWidth: 2, borderRadius: 4 }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { min: 3, max: 5 }, x: { ticks: { font: { size: 8, family: "'Nanum Gothic'" } } } }
    }
  });
}

function renderCustomQuantPreviewChart(slide) {
  const cs = slide.data || {};
  const groups = cs.groups || [];
  if (cs.chartType === 'none' || !groups.length) return;
  const canvasId = `pvCustomCanvas_${slide.customSlideId}`;
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;
  const isHoriz = cs.chartType === 'horizontalBar';
  const type = cs.chartType === 'line' ? 'line' : 'bar';
  const colors = groups.map(g => g.color || '#36A86F');
  builderState.previewCharts[`custom_${slide.customSlideId}`] = new Chart(canvas, {
    type,
    data: {
      labels: groups.map(g => g.name),
      datasets: [{ data: groups.map(g => g.avg || 0), backgroundColor: colors.map(c=>c+'99'), borderColor: colors, borderWidth: 2, borderRadius: 4 }]
    },
    options: {
      indexAxis: isHoriz ? 'y' : 'x',
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        [isHoriz?'x':'y']: { min: 0, max: 5 },
        [isHoriz?'y':'x']: { ticks: { font: { size: 9, family: "'Nanum Gothic'" } } }
      }
    }
  });
}

// 빌더 리프레시 (custom_slides.js에서 호출)
function refreshBuilderSlideList() { renderSlideList(); }
function refreshBuilderPreview() {
  const slide = builderState.slides[builderState.activeIdx];
  if (slide) renderPreview(slide);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 편집기 (오른쪽 하단)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function renderEditor(slide, idx) {
  const el = document.getElementById('slideEditor');
  if (!el) return;
  switch (slide.type) {
    case 'cover':      el.innerHTML = editorCover(slide.data); break;
    case 'overview':   el.innerHTML = editorOverview(slide.data); break;
    case 'schedule':   el.innerHTML = editorSchedule(slide.data); break;
    case 'summary':    el.innerHTML = editorSummary(slide.data); break;
    case 'quant_chart': el.innerHTML = editorQuantChart(slide.groupId); break;
    case 'qual_text':  el.innerHTML = editorQualText(); break;
    default:
      el.innerHTML = `<div class="editor-empty">이 슬라이드는 자동 생성됩니다 (편집 불필요)</div>`;
  }
}

// ── Cover 편집기
function editorCover(data) {
  return `
  <div class="editor-section-title">앞표지 편집</div>
  <div class="editor-row">
    <div class="editor-field">
      <label>고객사명</label>
      <input id="ed_company" value="${data.company || ''}" oninput="updateSlide('cover','company',this.value)">
    </div>
    <div class="editor-field">
      <label>과정명</label>
      <input id="ed_course" value="${data.course || ''}" oninput="updateSlide('cover','course',this.value)">
    </div>
  </div>`;
}

// ── Overview 편집기
function editorOverview(data) {
  const rows = (data.rows || []).map((r, i) => `
    <div class="editor-row">
      <div class="editor-field" style="max-width:100px">
        <label>항목명</label>
        <input value="${r.key}" oninput="updateOverviewRow(${i},'key',this.value)">
      </div>
      <div class="editor-field">
        <label>내용</label>
        <input value="${r.val || ''}" oninput="updateOverviewRow(${i},'val',this.value)">
      </div>
    </div>`).join('');
  return `<div class="editor-section-title">과정 개요 편집</div>${rows}`;
}

// ── Schedule 편집기
function editorSchedule(data) {
  const rows = (data.rows || []).map((r, i) => `
    <div class="editor-row">
      <div class="editor-field"><label>차수</label>
        <input value="${r.label}" oninput="updateScheduleRow(${i},'label',this.value)"></div>
      <div class="editor-field"><label>일자</label>
        <input value="${r.date || ''}" oninput="updateScheduleRow(${i},'date',this.value)"></div>
      <div class="editor-field"><label>장소</label>
        <input value="${r.place || ''}" oninput="updateScheduleRow(${i},'place',this.value)"></div>
      <div class="editor-field"><label>인원</label>
        <input value="${r.count || ''}" oninput="updateScheduleRow(${i},'count',this.value)"></div>
    </div>`).join('');
  return `<div class="editor-section-title">교육 일정표 편집</div>${rows}`;
}

// ── Summary 편집기
function editorSummary(data) {
  return `<div class="editor-section-title">Executive Summary (자동 생성)</div>
  <p style="font-size:12px;color:var(--gray-500)">
    업로드된 데이터에서 영역별 평균이 자동으로 반영됩니다.<br>
    전체 평균: <strong>${(data.overall||0).toFixed(2)}점</strong> | 응답 인원: <strong>${data.response_count||0}명</strong>
  </p>`;
}

// ── Quant Chart 편집기 (문항 그룹 선택)
function editorQuantChart(groupId) {
  const g = builderState.quantGroups.find(g => g.id === groupId);
  if (!g) return '<div class="editor-empty">그룹 없음</div>';
  const allQs = g.allQuestions || [];

  const checks = allQs.map(q => {
    const checked = g.questions.find(gq => gq.id === q.id) ? 'checked' : '';
    const lbl = q.label.length > 40 ? q.label.slice(0,40)+'…' : q.label;
    return `<label class="q-check-item">
      <input type="checkbox" ${checked} onchange="toggleGroupQuestion(${groupId},'${q.id}',this.checked,'${q.label.replace(/'/g,"\\'")}',${q.avg},${q.count||0})">
      <span class="q-label">${q.id} ${lbl} (${q.avg})</span>
    </label>`;
  }).join('');

  return `
  <div class="editor-section-title">정량 차트 문항 선택
    <button onclick="autoSplitGroup(${groupId})" style="margin-left:8px;font-size:10px;padding:2px 8px;background:var(--brand);color:#fff;border:none;border-radius:4px;cursor:pointer">6개씩 자동 분할</button>
    <button onclick="deleteQuantSlide(${groupId})" style="margin-left:4px;font-size:10px;padding:2px 8px;background:var(--danger);color:#fff;border:none;border-radius:4px;cursor:pointer">슬라이드 삭제</button>
  </div>
  <div class="editor-field"><label>슬라이드 제목</label>
    <input value="${g.title}" oninput="updateGroupTitle(${groupId},this.value)">
  </div>
  <div style="columns:2;column-gap:12px;font-size:11px">${checks}</div>`;
}

// ── Qual Text 편집기
function editorQualText() {
  const gd = window._qualBuilderData;
  if (!gd || !gd.open_ended_grouped) {
    return `<div class="editor-empty">정성 평가 탭에서 '정성 분석'을 먼저 실행하세요.</div>`;
  }
  const items = gd.open_ended_grouped;
  const html = items.map((oe, oeIdx) => {
    const common = (oe.groups || []).filter(g => g.is_common);
    if (!common.length) return '';
    const groups = common.map((g, gi) => `
      <div class="qual-edit-item">
        <textarea class="qual-edit-text" rows="2"
          oninput="updateQualEdit('${oe.id}',${gi},this.value)"
        >${builderState.qualEdits[oe.id]?.[gi] ?? g.label}</textarea>
      </div>`).join('');
    return `<div style="margin-bottom:10px">
      <div class="editor-section-title">${oe.id || ''} ${oe.label}</div>
      <div class="qual-edit-list">${groups}</div>
    </div>`;
  }).join('');
  return html || '<div class="editor-empty">공통응답 데이터가 없습니다.</div>';
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 상태 업데이트 핸들러
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function updateSlide(type, key, val) {
  const slide = builderState.slides.find(s => s.type === type);
  if (slide) { slide.data[key] = val; refreshPreviewDebounced(); }
}
function updateOverviewRow(i, key, val) {
  const slide = builderState.slides.find(s => s.type === 'overview');
  if (slide && slide.data.rows[i]) { slide.data.rows[i][key] = val; refreshPreviewDebounced(); }
}
function updateScheduleRow(i, key, val) {
  const slide = builderState.slides.find(s => s.type === 'schedule');
  if (slide && slide.data.rows[i]) { slide.data.rows[i][key] = val; refreshPreviewDebounced(); }
}
function updateGroupTitle(groupId, val) {
  const g = builderState.quantGroups.find(g => g.id === groupId);
  if (g) { g.title = val; renderSlideList(); refreshPreviewDebounced(); }
}
function toggleGroupQuestion(groupId, qId, checked, label, avg, count) {
  const g = builderState.quantGroups.find(g => g.id === groupId);
  if (!g) return;
  if (checked) {
    if (!g.questions.find(q => q.id === qId))
      g.questions.push({ id: qId, label, avg, count });
  } else {
    g.questions = g.questions.filter(q => q.id !== qId);
  }
  refreshPreviewDebounced();
}
function updateQualEdit(oeId, groupIdx, val) {
  if (!builderState.qualEdits[oeId]) builderState.qualEdits[oeId] = {};
  builderState.qualEdits[oeId][groupIdx] = val;
}

// ── 6개씩 자동 분할
function autoSplitGroup(groupId) {
  const g = builderState.quantGroups.find(g => g.id === groupId);
  if (!g) return;
  const all = g.allQuestions;
  const size = 6;
  const groupIdx = builderState.quantGroups.indexOf(g);
  // 현재 그룹과 관련 슬라이드 삭제
  builderState.quantGroups.splice(groupIdx, 1);
  builderState.slides = builderState.slides.filter(s => s.groupId !== groupId);
  // 새 그룹들 생성
  const qualIdx = builderState.slides.findIndex(s => s.type === 'qual_text');
  for (let i = 0; i < all.length; i += size) {
    const ng = {
      id: nextGroupId++,
      title: `정량 결과 ${Math.floor(i/size)+1}`,
      questions: all.slice(i, i + size),
      allQuestions: all,
    };
    builderState.quantGroups.splice(groupIdx + Math.floor(i/size), 0, ng);
    builderState.slides.splice(qualIdx + Math.floor(i/size), 0, { type: 'quant_chart', groupId: ng.id, data: {} });
  }
  renderSlideList();
  selectSlide(builderState.activeIdx);
}

// ── 슬라이드 추가/삭제
function addQuantSlide() {
  const all = builderState.quantGroups[0]?.allQuestions ||
              (window._builderData?.multi_result?.combined?.questions || []);
  const newG = {
    id: nextGroupId++,
    title: `정량 결과 추가 ${builderState.quantGroups.length + 1}`,
    questions: all.slice(0, 6),
    allQuestions: all,
  };
  builderState.quantGroups.push(newG);
  const qualIdx = builderState.slides.findIndex(s => s.type === 'qual_text');
  const insertAt = qualIdx >= 0 ? qualIdx : builderState.slides.length - 1;
  builderState.slides.splice(insertAt, 0, { type: 'quant_chart', groupId: newG.id, data: {} });
  renderSlideList();
  selectSlide(insertAt);
}

function deleteQuantSlide(groupId) {
  builderState.quantGroups = builderState.quantGroups.filter(g => g.id !== groupId);
  const slideIdx = builderState.slides.findIndex(s => s.groupId === groupId);
  if (slideIdx >= 0) builderState.slides.splice(slideIdx, 1);
  renderSlideList();
  selectSlide(Math.min(builderState.activeIdx, builderState.slides.length - 1));
}

// ── 디바운스 미리보기 새로고침
let _previewTimer = null;
function refreshPreviewDebounced() {
  clearTimeout(_previewTimer);
  _previewTimer = setTimeout(() => {
    renderPreview(builderState.slides[builderState.activeIdx]);
  }, 300);
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// PPT 빌드 & 다운로드
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
async function buildPPT() {
  const btn = document.getElementById('buildPptBtn');
  btn.textContent = '⏳ 생성 중...';
  btn.disabled = true;

  // qualEdits에서 정성 데이터 수집
  const qualData = [];
  const gd = window._qualBuilderData;
  if (gd && gd.open_ended_grouped) {
    gd.open_ended_grouped.forEach(oe => {
      const common = (oe.groups || []).filter(g => g.is_common);
      const edits = builderState.qualEdits[oe.id] || {};
      qualData.push({
        id: oe.id,
        label: oe.label,
        groups: common.map((g, i) => ({
          label: edits[i] !== undefined ? edits[i] : g.label,
          count: g.count,
          answers: g.answers || [],
        }))
      });
    });
  }

  const payload = {
    slides: builderState.slides,
    quant_groups: builderState.quantGroups.map(g => ({
      id: g.id,
      title: g.title,
      questions: g.questions,
    })),
    qual_data: qualData,
  };

  try {
    const sid = window.sessionId || '';
    const r = await fetch(`/api/build_ppt/${sid}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!r.ok) throw new Error(await r.text());
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = r.headers.get('Content-Disposition')?.split('filename=')[1]?.replace(/"/g,'') || '결과보고서.pptx';
    a.click();
    URL.revokeObjectURL(url);
  } catch(e) {
    alert('PPT 생성 실패: ' + e.message);
  }
  btn.textContent = '📥 PPT 생성';
  btn.disabled = false;
}
