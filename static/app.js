/* =============================================
   HRD 결과보고서 분석기 — 메인 앱 로직 v3
   ============================================= */

let sessionId = null;
let analysisData = null;   // { summary, sessions, combined }
let currentSession = null; // 현재 선택된 차수 데이터
let charts = {};

// ── 섹션 접기/펼치기 ──
function toggleSection(headerEl) {
  const card = headerEl.closest('.section-card');
  if (!card) return;
  card.classList.toggle('collapsed');
  const arrow = headerEl.querySelector('.collapse-arrow');
  if (arrow) arrow.textContent = card.classList.contains('collapsed') ? '▶' : '▼';
}

// 정량↔정성 이동 상태
let movedToQual = new Map();   // id → {id, label, avg, count} (정량→정성으로 이동)
let movedToQuant = new Set();  // id set (정성→정량으로 이동됐지만 정성탭에서 숨김)
let qualDataCache = null;      // 정성 원본 데이터 캐시
let qualSortOrder = 'desc';    // 공통응답 정렬: 'desc' | 'asc' | 'id'


// ── 초기화 ──────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  checkAIStatus();
  setupUpload();
  setupTabs();
  document.getElementById('resetBtn').addEventListener('click', resetApp);
  document.getElementById('pptBtn').addEventListener('click', generatePPT);
});

async function checkAIStatus() {
  try {
    const r = await fetch('/api/status');
    const d = await r.json();
    const el = document.getElementById('aiStatus');
    if (d.ai_enabled) {
      el.innerHTML = '<span class="dot dot-on"></span> AI 연결됨';
      el.style.color = '#166534';
    } else {
      el.innerHTML = '<span class="dot dot-off"></span> AI 미연결 (규칙 기반)';
    }
  } catch(e) {}
}

// ── 파일 업로드 ─────────────────────────────
function setupUpload() {
  const input = document.getElementById('fileInput');
  const card  = document.querySelector('.upload-card');

  input.addEventListener('change', e => {
    if (e.target.files[0]) uploadFile(e.target.files[0]);
  });

  // 드래그 앤 드롭
  card.addEventListener('dragover', e => { e.preventDefault(); card.style.borderColor = 'var(--brand)'; });
  card.addEventListener('dragleave', () => { card.style.borderColor = ''; });
  card.addEventListener('drop', e => {
    e.preventDefault();
    card.style.borderColor = '';
    const f = e.dataTransfer.files[0];
    if (f) uploadFile(f);
  });
}

async function uploadFile(file) {
  showProgress('파일 분석 중...', 20);

  const fd = new FormData();
  fd.append('file', file);
  // sheet 미지정 → 전체 시트 자동 분석

  try {
    showProgress('시트 파싱 중...', 40);
    const r = await fetch('/api/upload', { method: 'POST', body: fd });
    const d = await r.json();
    if (!r.ok) throw new Error(d.detail || '업로드 실패');

    showProgress('정성 문항 분석 중...', 70);
    sessionId = d.session_id;
    window.sessionId = sessionId;

    // 정성 데이터 그룹핑 요청
    const gr = await fetch(`/api/analyze_qual/${sessionId}`);
    const gd = gr.ok ? await gr.json() : {};

    showProgress('완료!', 100);
    analysisData = { ...d, qual: gd };
    setTimeout(() => renderDashboard(d, gd), 400);

  } catch(e) {
    alert('오류: ' + e.message);
    hideProgress();
  }
}

function showProgress(text, pct) {
  document.getElementById('uploadProgress').style.display = 'block';
  document.getElementById('progressFill').style.width = pct + '%';
  document.getElementById('progressText').textContent = text;
}
function hideProgress() {
  document.getElementById('uploadProgress').style.display = 'none';
}

// ── 대시보드 렌더링 ──────────────────────────
function renderDashboard(d, gd) {
  document.getElementById('uploadSection').style.display = 'none';
  document.getElementById('dashboard').style.display = 'block';

  const s = d.summary;
  const multi = d.multi_result;

  // 과정 헤더
  document.getElementById('courseTitle').textContent =
    `${s.company} ${s.course_name}`.trim() || '교육 과정';
  document.getElementById('courseMeta').textContent =
    `응답 인원 ${s.response_count}명  •  문항 ${s.total_questions}개  •  평균 ${s.overall_average}점`;

  // 차수 배지
  const badgesEl = document.getElementById('sessionBadges');
  badgesEl.innerHTML = '';
  if (s.multi_session && s.sessions_info) {
    s.sessions_info.forEach(si => {
      const b = document.createElement('span');
      b.className = 'session-badge';
      b.textContent = `${si.label} (${si.response_count}명)`;
      badgesEl.appendChild(b);
    });
    const total = document.createElement('span');
    total.className = 'session-badge';
    total.style.background = 'rgba(255,255,255,.35)';
    total.textContent = `종합 ${s.response_count}명`;
    badgesEl.appendChild(total);
  }

  renderOverviewTab(d);
  renderQuantTab(d);
  qualDataCache = gd;
  renderQualTab(mergeQualData(gd));

  // 보고서 빌더 초기화
  window._builderData = d;
  window._qualBuilderData = gd;
  // custom_slides.js용 데이터 노출
  const _mr = d.multi_result || {};
  window.analysisData = {
    summary: d.summary,
    sessions: _mr.sessions || [],
    combined: _mr.combined || _mr.sessions?.[0] || {},
  };
  if (typeof initBuilder === 'function') initBuilder(d);
  if (typeof initCustomSlides === 'function') initCustomSlides();
}


// ── 탭: 개요 ─────────────────────────────────
function renderOverviewTab(d) {
  const s = d.summary;
  const multi = d.multi_result || {};
  const sessions = multi.sessions || [];

  // 기본 정보 테이블
  const rows = [
    ['고객사', s.company || '-'],
    ['과정명', s.course_name || '-'],
    ['총 응답 인원', `${s.response_count}명`],
    ['전체 평균', `${s.overall_average}점 / 5점`],
    ['문항 수', `${s.total_questions}개`],
    ['영역 수', `${s.categories}개 (${s.category_names.join(', ')})`],
    ['주관식 문항', `${s.open_ended_count}개`],
    ['차수 구분', s.multi_session ? `${s.session_count}차수` : '단일'],
  ];

  if (s.multi_session && s.sessions_info) {
    s.sessions_info.forEach(si => {
      rows.push([si.label, `응답 ${si.response_count}명 / 평균 ${si.overall_average?.toFixed(2) || '-'}점`]);
    });
  }

  document.getElementById('overviewBody').innerHTML = rows.map(([k,v]) =>
    `<tr><td>${k}</td><td>${v}</td></tr>`
  ).join('');

  // 통계 카드
  const tier = s.overall_tier;
  const tierColor = tier === '상' ? '#10b981' : tier === '중' ? '#f59e0b' : '#ef4444';
  document.getElementById('statsGrid').innerHTML = `
    <div class="stat-card">
      <div class="stat-value">${s.overall_average}</div>
      <div class="stat-label">전체 평균 (/ 5점)</div>
    </div>
    <div class="stat-card">
      <div class="stat-value" style="color:${tierColor}">${tier || '-'}</div>
      <div class="stat-label">만족도 등급</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">${s.response_count}</div>
      <div class="stat-label">총 응답 인원</div>
    </div>
    <div class="stat-card">
      <div class="stat-value">${s.session_count || 1}</div>
      <div class="stat-label">차수</div>
    </div>
  `;
}

// ── 탭: 정량 평가 ────────────────────────────
function renderQuantTab(d) {
  const s = d.summary;
  const multi = d.multi_result || {};
  const sessions = multi.sessions || [];
  const combined = multi.combined || {};

  // 차수 서브탭 생성
  const tabsEl = document.getElementById('sessionTabs');
  tabsEl.innerHTML = '';

  const allSessions = [];
  if (sessions.length > 1) {
    sessions.forEach((sess, i) => {
      allSessions.push({ label: sess.session_label || sess.sheet_name, data: sess });
    });
    allSessions.push({ label: '종합', data: combined });
  }

  if (allSessions.length > 0) {
    allSessions.forEach((sess, i) => {
      const btn = document.createElement('button');
      btn.className = 'session-tab' + (i === allSessions.length - 1 ? ' active' : '');
      btn.textContent = sess.label;
      btn.addEventListener('click', () => {
        document.querySelectorAll('.session-tab').forEach(b => b.classList.remove('active'));
        btn.classList.add('active');
        renderQuantData(sess.data, sess.label, sessions);
      });
      tabsEl.appendChild(btn);
    });
    // 초기: 종합 표시
    renderQuantData(combined, '종합', sessions);
  } else {
    renderQuantData(combined.questions ? combined : d, s.course_name, []);
  }

  // 차수 비교 (멀티세션)
  if (sessions.length > 1) {
    renderSessionCompare(sessions, combined);
  }
}

function renderQuantData(data, label, sessions) {
  const cats = data.categories || [];
  const qs = data.questions || [];
  const resp = data.response_count || 0;

  // 차수 전환 시 선택 초기화
  selectedIndices.clear();
  updateSelectionUI();

  document.getElementById('quantSummaryTitle').textContent = `정량 평가 종합 — ${label}`;

  // 영역별 요약 표
  const summaryBody = document.getElementById('quantSummaryBody');
  summaryBody.innerHTML = cats.map(c => {
    const cls = scoreClass(c.avg);
    return `<tr>
      <td>${c.name}</td>
      <td><span class="score-badge ${cls}">${c.avg.toFixed(2)}</span></td>
      <td class="tier-${tierLabel(c.avg)}">${tierLabel(c.avg)}</td>
    </tr>`;
  }).join('');

  // 문항별 상세 표 (체크박스 포함)
  renderDetailBody(qs, resp);

  // 차트 렌더링
  renderSummaryChart(cats);
  renderDetailChart(qs);

  // PPT 내보내기 버튼 삽입
  _addExportButtons(cats, qs, resp);
}


function renderSummaryChart(cats) {
  if (charts.summary) charts.summary.destroy();
  const ctx = document.getElementById('summaryChart').getContext('2d');
  charts.summary = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: cats.map(c => c.name),
      datasets: [{
        label: '영역별 평균',
        data: cats.map(c => c.avg),
        backgroundColor: cats.map(c => scoreColor(c.avg)),
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        y: { min: 3, max: 5, grid: { color: '#f1f5f9' } },
        x: { grid: { display: false } }
      }
    }
  });
}

// ── 선택 모드 상태 ──────────────────────────
let selectMode = false;
let selectedIndices = new Set();
let currentQuestions = [];  // 현재 표시 중인 문항 목록

function toggleSelectMode() {
  selectMode = !selectMode;
  const btn = document.getElementById('selectModeBtn');
  const hint = document.getElementById('selectHint');
  btn.classList.toggle('active', selectMode);
  hint.style.display = selectMode ? 'block' : 'none';
  if (!selectMode) {
    selectedIndices.clear();
    updateSelectionUI();
  }
  // 테이블 체크박스 보이기/숨기기
  document.querySelectorAll('.row-check, #selectAll').forEach(el => {
    el.style.display = selectMode ? '' : 'none';
  });
  // 빌더·내보내기 버튼 보이기/숨기기
  ['exportExcelBtn','exportClipBtn','addToBuilderBtn'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.style.display = selectMode ? '' : 'none';
  });
  // 차트 재렌더
  renderDetailChart(currentQuestions);
}

// ── 선택 문항 → 빌더 슬라이드 추가 ─────────────
function addSelectedToBuilder() {
  const qs = getSelectedQuestions();
  if (!qs.length) { alert('문항을 먼저 선택하세요.'); return; }

  const title = prompt('슬라이드 제목을 입력하세요', `정량 결과 (${qs.length}문항)`);
  if (title === null) return;  // 취소

  if (typeof addQuantGroupToBuilder === 'function') {
    addQuantGroupToBuilder(title, qs);
  } else {
    // builder.js 아직 안 로드된 경우 대비
    console.warn('builder.js not loaded');
  }

  // 선택 해제
  selectedIndices.clear();
  updateSelectionUI();
  renderDetailChart(currentQuestions);
}


function toggleSelectAll(cb) {
  if (cb.checked) {
    currentQuestions.forEach((_, i) => selectedIndices.add(i));
  } else {
    selectedIndices.clear();
  }
  document.querySelectorAll('.row-check').forEach((el, i) => {
    el.checked = cb.checked;
    el.closest('tr').classList.toggle('selected-row', cb.checked);
  });
  updateSelectionUI();
  renderDetailChart(currentQuestions);
}

function toggleRowSelect(idx, cb) {
  if (cb.checked) selectedIndices.add(idx);
  else selectedIndices.delete(idx);
  cb.closest('tr').classList.toggle('selected-row', cb.checked);
  updateSelectionUI();
  renderDetailChart(currentQuestions);
}

function updateSelectionUI() {
  const cnt = selectedIndices.size;
  const countEl = document.getElementById('selectedCount');
  const excelBtn = document.getElementById('exportExcelBtn');
  const clipBtn = document.getElementById('exportClipBtn');
  if (cnt > 0) {
    countEl.style.display = 'block';
    countEl.textContent = `✅ ${cnt}개 문항 선택됨`;
    excelBtn.style.display = '';
    clipBtn.style.display = '';
  } else {
    countEl.style.display = 'none';
    excelBtn.style.display = 'none';
    clipBtn.style.display = 'none';
  }
}

// ── 상세 차트 (클릭 선택 지원) ──────────────
function renderDetailChart(qs) {
  currentQuestions = qs;
  if (charts.detail) charts.detail.destroy();
  const ctx = document.getElementById('detailChart').getContext('2d');

  const bgColors = qs.map((q, i) =>
    selectedIndices.has(i)
      ? '#2563eb'
      : scoreColor(q.avg) + 'cc'
  );
  const borderColors = qs.map((q, i) =>
    selectedIndices.has(i) ? '#1d4ed8' : 'transparent'
  );

  charts.detail = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: qs.map(q => q.label.length > 14 ? q.label.slice(0,14)+'…' : q.label),
      datasets: [{
        label: '문항별 평균',
        data: qs.map(q => q.avg),
        backgroundColor: bgColors,
        borderColor: borderColors,
        borderWidth: 2,
        borderRadius: 4,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: (items) => qs[items[0].dataIndex]?.label || '',
            label: (item) => ` 평균: ${item.raw}점`
              + (selectedIndices.has(item.dataIndex) ? ' ✅' : '')
          }
        }
      },
      scales: {
        y: { min: 3, max: 5, grid: { color: '#f1f5f9' } },
        x: { grid: { display: false }, ticks: { font: { size: 11 } } }
      },
      onClick: (evt, elements) => {
        if (!selectMode || !elements.length) return;
        const idx = elements[0].index;
        if (selectedIndices.has(idx)) selectedIndices.delete(idx);
        else selectedIndices.add(idx);
        // 테이블 체크박스 동기화
        const checkboxes = document.querySelectorAll('.row-check');
        if (checkboxes[idx]) {
          checkboxes[idx].checked = selectedIndices.has(idx);
          checkboxes[idx].closest('tr').classList.toggle('selected-row', selectedIndices.has(idx));
        }
        updateSelectionUI();
        renderDetailChart(qs);
      },
      onHover: (evt) => {
        evt.native.target.style.cursor = selectMode ? 'pointer' : 'default';
      }
    }
  });
}

// ── 문항별 상세 표 (체크박스 + 이동 버튼 포함) ──────────
function renderDetailBody(qs, resp) {
  const body = document.getElementById('quantDetailBody');
  body.innerHTML = qs.map((q, i) => {
    const cls = scoreClass(q.avg);
    const isMovedToQual = movedToQual.has(q.id);
    return `<tr data-idx="${i}" ${isMovedToQual ? 'style="opacity:.4;text-decoration:line-through"' : ''}>
      <td><input type="checkbox" class="row-check"
          style="display:${selectMode ? '' : 'none'}"
          onchange="toggleRowSelect(${i}, this)"></td>
      <td style="color:var(--gray-500)">${q.id || i+1}</td>
      <td>${q.label}</td>
      <td><span class="score-badge ${cls}">${q.avg.toFixed(2)}</span></td>
      <td>${q.count || resp}명</td>
      <td><button class="move-btn" onclick="moveToQual('${q.id}','${q.label.replace(/'/g,"\\'")}',${i})"
          title="정성 탭으로 이동">📤 정성으로</button></td>
    </tr>`;
  }).join('');
}



function renderSessionCompare(sessions, combined) {
  document.getElementById('sessionCompareCard').style.display = 'block';
  const labels = sessions.map(s => s.session_label || s.sheet_name);
  const cats = (combined.categories || []).filter(c => c.per_session);

  // 차수 비교 차트
  if (charts.session) charts.session.destroy();
  const ctx = document.getElementById('sessionChart').getContext('2d');
  const colors = ['#2563eb','#0ea5e9','#10b981','#f59e0b','#8b5cf6'];
  charts.session = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: labels.concat(['종합']),
      datasets: cats.slice(0, 5).map((c, i) => ({
        label: c.name,
        data: labels.map(l => c.per_session?.[l] || 0).concat([c.avg]),
        backgroundColor: colors[i % colors.length] + 'bb',
        borderRadius: 4,
      }))
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'top' } },
      scales: {
        y: { min: 3, max: 5 },
        x: { grid: { display: false } }
      }
    }
  });

  // 차수 비교 표
  const head = document.getElementById('sessionCompareHead');
  head.innerHTML = `<tr><th>영역</th>${labels.map(l=>`<th>${l}</th>`).join('')}<th>종합</th></tr>`;

  const body = document.getElementById('sessionCompareBody');
  body.innerHTML = cats.map(c => {
    const cells = labels.map(l => {
      const v = c.per_session?.[l];
      return v ? `<td><span class="score-badge ${scoreClass(v)}">${v.toFixed(2)}</span></td>` : '<td>-</td>';
    }).join('');
    return `<tr><td>${c.name}</td>${cells}<td><span class="score-badge ${scoreClass(c.avg)}">${c.avg.toFixed(2)}</span></td></tr>`;
  }).join('');
}

// ── 탭: 정성 평가 ────────────────────────────
// ── 정성 정렬 ─────────────────────────────────────
function setQualSort(order) {
  qualSortOrder = order;
  // 버튼 active 토글
  ['sortDesc','sortAsc','sortId'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.classList.remove('active');
  });
  const activeId = order === 'desc' ? 'sortDesc' : order === 'asc' ? 'sortAsc' : 'sortId';
  const activeEl = document.getElementById(activeId);
  if (activeEl) activeEl.classList.add('active');
  // 재렌더
  if (qualDataCache) renderQualTab(mergeQualData(qualDataCache));
}

function renderQualTab(gd) {
  const container = document.getElementById('qualContent');
  const emptyEl = document.getElementById('qualEmpty');
  container.innerHTML = '';

  let items = gd.open_ended_grouped || [];
  if (!items.length) {
    emptyEl.style.display = 'block';
    return;
  }
  emptyEl.style.display = 'none';

  // 정렬 적용
  items = [...items];
  if (qualSortOrder === 'id') {
    items.sort((a, b) => {
      const numA = parseInt((a.id || '').replace(/\D/g, '')) || 999;
      const numB = parseInt((b.id || '').replace(/\D/g, '')) || 999;
      return numA - numB;
    });
  } else if (qualSortOrder === 'desc') {
    items.sort((a, b) => {
      const cntA = (a.groups||[]).filter(g=>g.is_common).reduce((s,g)=>s+g.count,0);
      const cntB = (b.groups||[]).filter(g=>g.is_common).reduce((s,g)=>s+g.count,0);
      return cntB - cntA;
    });
  } else { // asc
    items.sort((a, b) => {
      const cntA = (a.groups||[]).filter(g=>g.is_common).reduce((s,g)=>s+g.count,0);
      const cntB = (b.groups||[]).filter(g=>g.is_common).reduce((s,g)=>s+g.count,0);
      return cntA - cntB;
    });
  }

  items.forEach(oe => {
    const common = (oe.groups || []).filter(g => g.is_common);
    const indiv  = (oe.groups || []).filter(g => !g.is_common);
    const totalCount = oe.answers?.length || 0;

    const card = document.createElement('div');
    card.className = 'qual-question-card';

    // 헤더 — Q번호 뱃지 + 문항명
    const qNum = oe.id || '';
    card.innerHTML = `
      <div class="qual-question-header">
        <div class="qual-header-left">
          ${qNum ? `<span class="q-badge">${qNum}</span>` : ''}
          <h4>${oe.label || '주관식 문항'}</h4>
        </div>
        <span class="qual-count">총 ${totalCount}건 / 공통 ${common.length}그룹</span>
      </div>
    `;

    const body = document.createElement('div');
    body.className = 'qual-body';

    // 공통응답 (2건 이상)
    if (common.length > 0) {
      common.forEach(g => {
        const div = document.createElement('div');
        div.className = 'common-group';
        div.innerHTML = `
          <div class="common-group-label">${g.common_id}</div>
          <div class="common-group-text">${g.label}</div>
          <div class="common-group-count">(${g.count}건)</div>
          <div class="common-group-answers">
            ${g.answers.map(a => `<div class="individual-answer">${a}</div>`).join('')}
          </div>
        `;
        body.appendChild(div);
      });
    }

    // 개별 응답 (1건)
    if (indiv.length > 0) {
      if (common.length > 0) {
        const sep = document.createElement('div');
        sep.style.cssText = 'margin: 12px 0 8px; font-size:12px; color:var(--gray-500); font-weight:600;';
        sep.textContent = '▸ 개별 응답';
        body.appendChild(sep);
      }
      indiv.forEach(g => {
        const div = document.createElement('div');
        div.className = 'individual-answer';
        div.textContent = g.label;
        body.appendChild(div);
      });
    }

    if (!common.length && !indiv.length) {
      body.innerHTML = '<div class="individual-answer" style="color:var(--gray-500)">응답 없음</div>';
    }

    card.appendChild(body);

    // 복사 및 이동 버튼
    const copyRow = document.createElement('div');
    copyRow.className = 'qual-copy-row';

    const moveBtn = document.createElement('button');
    moveBtn.className = 'move-btn alt';
    const isMoved = movedToQuant.has(oe.id);
    moveBtn.textContent = isMoved ? '↩ 정성으로 돌리기' : '📥 정량으로';
    moveBtn.onclick = () => moveToQuant(oe.id || oe.label, oe.label);

    const copyBtn = document.createElement('button');
    copyBtn.className = 'copy-btn';
    copyBtn.textContent = '📋 복사';
    copyBtn.addEventListener('click', () => copyQualText(oe, common, indiv, copyBtn));

    const addBuilderBtn = document.createElement('button');
    addBuilderBtn.className = 'add-to-builder-btn';
    addBuilderBtn.textContent = '📄 빌더에 추가';
    addBuilderBtn.addEventListener('click', () => {
      if (typeof addQualSlideToBuilder === 'function') {
        addQualSlideToBuilder(oe.id || oe.label, oe.label, common);
      }
    });

    copyRow.appendChild(moveBtn);
    copyRow.appendChild(copyBtn);
    copyRow.appendChild(addBuilderBtn);
    card.appendChild(copyRow);

    container.appendChild(card);

  });
}

function copyQualText(oe, common, indiv, btn) {
  let text = `[${oe.id || ''}] ${oe.label || ''}\n${'─'.repeat(40)}\n`;
  if (common.length) {
    common.forEach(g => {
      text += `\n${g.common_id}: ${g.label} (${g.count}건)\n`;
      g.answers.forEach(a => { text += `  • ${a}\n`; });
    });
  }
  if (indiv.length) {
    text += '\n▸ 개별 응답\n';
    indiv.forEach(g => { text += `  • ${g.label}\n`; });
  }
  copyToClipboard(text, btn);
}

// ── 복사 유틸 ────────────────────────────────
function setupCopyButtons() {
  document.querySelectorAll('.copy-btn[data-target]').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = document.getElementById(btn.dataset.target);
      if (!target) return;
      const text = tableToText(target);
      copyToClipboard(text, btn);
    });
  });
}

function tableToText(table) {
  const rows = [];
  table.querySelectorAll('tr').forEach(tr => {
    const cells = [...tr.querySelectorAll('th,td')].map(c => c.innerText.trim());
    rows.push(cells.join('\t'));
  });
  return rows.join('\n');
}

function copyToClipboard(text, btn) {
  navigator.clipboard.writeText(text).then(() => {
    const orig = btn.textContent;
    btn.textContent = '✅ 복사됨';
    btn.classList.add('copied');
    setTimeout(() => {
      btn.textContent = orig;
      btn.classList.remove('copied');
    }, 1500);
  });
}

// ── 탭 전환 ──────────────────────────────────
function setupTabs() {
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab).classList.add('active');
      // 탭 전환 후 copy 버튼 연결
      setTimeout(setupCopyButtons, 50);
    });
  });
}

// ── 리셋 ─────────────────────────────────────
function resetApp() {
  sessionId = null;
  analysisData = null;
  movedToQual.clear();
  movedToQuant.clear();
  qualDataCache = null;
  document.getElementById('dashboard').style.display = 'none';
  document.getElementById('uploadSection').style.display = 'flex';
  document.getElementById('fileInput').value = '';
  hideProgress();
  Object.values(charts).forEach(c => { try { c.destroy(); } catch(e){} });
  charts = {};
}

// ── 정량 ↔ 정성 이동 ────────────────────────
function moveToQual(qId, qLabel, idx) {
  if (movedToQual.has(qId)) {
    movedToQual.delete(qId);
  } else {
    movedToQual.set(qId, { id: qId, label: qLabel, answers: [], groups: [], is_manual: true });
  }
  // 정량 표 새로고침
  renderDetailBody(currentQuestions, 0);
  // 정성 탭 업데이트 (qualDataCache 있으면 반영)
  if (qualDataCache) {
    const merged = mergeQualData(qualDataCache);
    renderQualTab(merged);
  }
}

async function moveToQuant(qId, qLabel) {
  if (movedToQuant.has(qId)) {
    // ── 되돌리기: 정량에서 제거, 정성에서 복원
    movedToQuant.delete(qId);
    currentQuestions = currentQuestions.filter(q => q.id !== qId || !q._manualFromQual);
  } else {
    movedToQuant.add(qId);
    // rawdata에서 해당 문항 점수 계산
    try {
      const rd = await fetch(`/api/rawdata/${sessionId}`);
      if (rd.ok) {
        const rdJson = await rd.json();
        const rawBySheet = rdJson.raw_by_sheet || {};
        let scores = [];
        for (const qs of Object.values(rawBySheet)) {
          const match = qs.find(q => q.id === qId || q.label === qLabel);
          if (match) {
            scores = (match.responses || [])
              .map(r => parseFloat(r))
              .filter(v => !isNaN(v) && v >= 1 && v <= 5);
            break;
          }
        }
        if (scores.length > 0) {
          const avg = Math.round(scores.reduce((a, b) => a + b, 0) / scores.length * 100) / 100;
          // currentQuestions에 삽입 (중복 방지)
          if (!currentQuestions.find(q => q.id === qId)) {
            currentQuestions = [...currentQuestions, {
              id: qId, label: qLabel, avg, count: scores.length,
              category: '이동(정성→정량)', _manualFromQual: true
            }];
          }
        }
      }
    } catch(e) { console.warn('rawdata 로드 실패', e); }
  }

  // 정량 차트·표 갱신
  renderDetailBody(currentQuestions, 0);
  renderDetailChart(currentQuestions);
  // 정성 탭 갱신
  if (qualDataCache) renderQualTab(mergeQualData(qualDataCache));
}


function mergeQualData(gd) {
  // 수동이동 항목 추가 + 정량이동 항목 숨김
  const base = (gd.open_ended_grouped || []).filter(oe => !movedToQuant.has(oe.id));
  const extras = [...movedToQual.values()].filter(m => m.is_manual);
  return { ...gd, open_ended_grouped: [...base, ...extras] };
}

// ── 헬퍼 ─────────────────────────────────────
function scoreClass(v) {
  if (v >= 4.5) return 'score-high';
  if (v >= 4.0) return 'score-mid';
  return 'score-low';
}
function scoreColor(v) {
  if (v >= 4.5) return '#10b981';
  if (v >= 4.0) return '#2563eb';
  return '#f59e0b';
}
function tierLabel(v) {
  if (v >= 4.5) return '상';
  if (v >= 4.0) return '중';
  return '하';
}


// ── PPT 차트용 Excel 내보내기 ─────────────────
function getSelectedQuestions() {
  if (selectedIndices.size === 0) return currentQuestions;
  return [...selectedIndices].sort((a,b) => a-b).map(i => currentQuestions[i]);
}

async function exportToExcel() {
  const qs = getSelectedQuestions();
  if (!qs.length) return;

  const btn = document.getElementById('exportExcelBtn');
  btn.textContent = '⏳ 준비 중...';
  btn.disabled = true;

  const wb = XLSX.utils.book_new();
  const courseName = document.getElementById('courseTitle').textContent || '결과보고서';

  // ── 시트 1: PPT 차트 데이터 (A=문항, B=평균) ──
  const chartData = [
    ['문항', '평균점수'],
    ...qs.map(q => [q.label, q.avg])
  ];
  const ws1 = XLSX.utils.aoa_to_sheet(chartData);
  ws1['!cols'] = [{ wch: 40 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws1, 'PPT차트데이터');

  // ── 시트 2: 상세 요약 ──
  const detailData = [
    ['번호', '문항', '카테고리', '평균', '응답수'],
    ...qs.map(q => [q.id || '', q.label, q.category || '', q.avg, q.count || ''])
  ];
  const ws2 = XLSX.utils.aoa_to_sheet(detailData);
  ws2['!cols'] = [{ wch: 8 }, { wch: 40 }, { wch: 12 }, { wch: 8 }, { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, ws2, '상세데이터');

  // ── 시트 3: 로우데이터 (문항별 개별 응답) ──
  try {
    const rd = await fetch(`/api/rawdata/${sessionId}`);
    if (rd.ok) {
      const rdJson = await rd.json();
      const rawBySheet = rdJson.raw_by_sheet || {};

      // 선택된 문항 ID 집합
      const selectedIds = new Set(qs.map(q => q.id));
      const selectedLabels = new Set(qs.map(q => q.label));

      // 모든 시트의 로우데이터 합침
      const rawRows = [['시트', '번호', '문항', '응답1', '응답2', '응답3', '...']];
      for (const [sheetName, qList] of Object.entries(rawBySheet)) {
        for (const q of qList) {
          if (!selectedIds.has(q.id) && !selectedLabels.has(q.label)) continue;
          const responses = (q.responses || []).map(r => r === null ? '' : r);
          rawRows.push([sheetName, q.id, q.label, ...responses]);
        }
      }
      const ws3 = XLSX.utils.aoa_to_sheet(rawRows);
      ws3['!cols'] = [{ wch: 10 }, { wch: 8 }, { wch: 40 }];
      XLSX.utils.book_append_sheet(wb, ws3, '로우데이터');
    }
  } catch(e) {
    console.warn('로우데이터 로드 실패:', e);
  }

  XLSX.writeFile(wb, `[차트데이터] ${courseName}.xlsx`);
  btn.textContent = '📊 Excel 내보내기';
  btn.disabled = false;
}


function exportToClipboard() {
  const qs = getSelectedQuestions();
  if (!qs.length) return;

  // 탭 구분 텍스트 (Excel에 붙여넣기 가능)
  const header = '문항\t평균점수\t카테고리\t응답수';
  const rows = qs.map(q =>
    `${q.label}\t${q.avg}\t${q.category || ''}\t${q.count || ''}`
  );
  const text = [header, ...rows].join('\n');

  navigator.clipboard.writeText(text).then(() => {
    const btn = document.getElementById('exportClipBtn');
    const orig = btn.textContent;
    btn.textContent = '✅ 복사됨!';
    setTimeout(() => { btn.textContent = orig; }, 1500);
  });
}

async function generatePPT() {
  if (!sessionId) return;
  const btn = document.getElementById('pptBtn');
  btn.textContent = '⏳ 생성 중...';
  btn.classList.add('loading');
  btn.disabled = true;

  try {
    const r = await fetch(`/api/generate/${sessionId}`, { method: 'POST' });
    const d = await r.json();

    if (d.success) {
      // 자동 다운로드
      const a = document.createElement('a');
      a.href = `/api/download/${sessionId}`;
      a.download = d.output_file || 'report.pptx';
      a.click();
      btn.textContent = '✅ 다운로드 완료';
      setTimeout(() => {
        btn.textContent = '📥 PPT 보고서 생성';
        btn.classList.remove('loading');
        btn.disabled = false;
      }, 2500);
    } else {
      alert('PPT 생성 실패: ' + (d.error || '알 수 없는 오류'));
      btn.textContent = '📥 PPT 보고서 생성';
      btn.classList.remove('loading');
      btn.disabled = false;
    }
  } catch(e) {
    alert('오류: ' + e.message);
    btn.textContent = '📥 PPT 보고서 생성';
    btn.classList.remove('loading');
    btn.disabled = false;
  }
}


// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// PPT 개별 내보내기 (차트/표 → 복사용 PPTX)
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function _addExportButtons(cats, qs, resp) {
  // 영역별 요약: section-card (#quantSummaryCard) 맨 아래에 삽입
  const summaryCard = document.getElementById('quantSummaryCard');
  if (summaryCard) {
    let existing = summaryCard.querySelector('.ppt-export-row');
    if (existing) existing.remove();
    let div = document.createElement('div');
    div.className = 'ppt-export-row';
    div.style.cssText = 'display:flex;gap:8px;padding:8px 16px 12px;justify-content:flex-end;border-top:1px solid #eee;margin-top:8px';
    div.innerHTML = `
      <button class="ppt-export-btn" onclick="exportChartToPPT('summary','영역별 종합')">📊 차트 PPT 복사</button>
      <button class="ppt-export-btn" onclick="exportTableToPPT('summary','영역별 요약 표')">📋 표 PPT 복사</button>
    `;
    summaryCard.appendChild(div);
  }
  // 문항별 상세: chart-wrap 바로 뒤에 별도 div 삽입 (table 위)
  const detailChartWrap = document.getElementById('detailChart')?.closest('.chart-wrap');
  if (detailChartWrap) {
    let existing = detailChartWrap.parentElement.querySelector('.ppt-export-row');
    if (existing) existing.remove();
    let div = document.createElement('div');
    div.className = 'ppt-export-row';
    div.style.cssText = 'display:flex;gap:8px;padding:8px 0 4px;justify-content:flex-end';
    div.innerHTML = `
      <button class="ppt-export-btn" onclick="exportChartToPPT('detail','문항별 상세')">📊 차트 PPT 복사</button>
      <button class="ppt-export-btn" onclick="exportTableToPPT('detail','문항별 상세 표')">📋 표 PPT 복사</button>
    `;
    // chart-wrap 바로 뒤에 삽입 (table 앞)
    detailChartWrap.parentElement.insertBefore(div, detailChartWrap.nextSibling);
  }
}

window._exportCache = {};

async function exportChartToPPT(chartId, title) {
  const data = _getChartDataForExport(chartId);
  if (!data) return alert('차트 데이터 없음');

  const payload = {
    type: 'chart',
    title: title,
    data: data
  };

  try {
    const r = await fetch('/api/export_element', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(payload)
    });
    if (!r.ok) throw new Error(await r.text());
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `[복사용] ${title}.pptx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch(e) {
    alert('내보내기 실패: ' + e.message);
  }
}

async function exportTableToPPT(tableId, title) {
  const data = _getTableDataForExport(tableId);
  if (!data) return alert('표 데이터 없음');

  const payload = {
    type: 'table',
    title: title,
    data: data
  };

  try {
    const r = await fetch('/api/export_element', {
      method: 'POST',
      headers: {'Content-Type': 'application/json'},
      body: JSON.stringify(payload)
    });
    if (!r.ok) throw new Error(await r.text());
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `[복사용] ${title}.pptx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch(e) {
    alert('내보내기 실패: ' + e.message);
  }
}

function _getChartDataForExport(chartId) {
  const ad = window.analysisData;
  if (!ad) return null;
  const combined = ad.combined || {};
  const cats = combined.categories || [];
  const qs = combined.questions || [];

  if (chartId === 'summary') {
    return {
      labels: cats.map(c => c.name),
      values: cats.map(c => c.avg),
      chartType: 'bar',
      colors: cats.map(c => c.avg >= 4.5 ? '#36A86F' : c.avg < 3.5 ? '#E74C3C' : '#4A90D9')
    };
  } else if (chartId === 'detail') {
    return {
      labels: qs.map(q => (q.label || '').substring(0, 20)),
      values: qs.map(q => q.avg),
      chartType: 'bar',
      colors: qs.map(q => q.avg >= 4.5 ? '#36A86F' : q.avg < 3.5 ? '#E74C3C' : '#4A90D9')
    };
  }
  return null;
}

function _getTableDataForExport(tableId) {
  const ad = window.analysisData;
  if (!ad) return null;
  const combined = ad.combined || {};
  const cats = combined.categories || [];
  const qs = combined.questions || [];

  if (tableId === 'summary') {
    return {
      headers: ['영역', '평균', '등급'],
      rows: cats.map(c => [c.name, c.avg.toFixed(2), c.avg >= 4.5 ? '우수' : c.avg >= 3.5 ? '보통' : '미흡'])
    };
  } else if (tableId === 'detail') {
    return {
      headers: ['번호', '문항', '평균', '응답'],
      rows: qs.map(q => [q.id || '', q.label || '', (q.avg||0).toFixed(2), q.count || ''])
    };
  }
  return null;
}
