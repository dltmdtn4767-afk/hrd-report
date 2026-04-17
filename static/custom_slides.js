/* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
   custom_slides.js — 문항 그룹핑 + 커스텀 슬라이드 시스템
   ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */

// ── 전역 상태 ──
const customSlides = [];   // [{id, title, chartType, tableStyle, tablePos, groups:[{name,qIds}]}]
let activeCSIdx = -1;
let csIdSeq = 1;

function getAllQuestions() {
  // app.js의 analysisData에서 전체 문항 가져오기
  if (!window.analysisData) return [];
  const combined = window.analysisData.combined || window.analysisData.sessions?.[0];
  if (!combined) return [];
  return combined.questions || [];
}

function getCategories() {
  if (!window.analysisData) return [];
  const combined = window.analysisData.combined || window.analysisData.sessions?.[0];
  return combined?.categories || [];
}

// ── 초기화 ──
function initCustomSlides() {
  renderCSTabs();
  renderCSPanel();
}

// ── 탭 렌더링 ──
function renderCSTabs() {
  const cont = document.getElementById('customSlideTabs');
  if (!cont) return;
  let html = '';
  customSlides.forEach((s, i) => {
    html += `<button class="cst-tab ${i === activeCSIdx ? 'active' : ''}"
              onclick="selectCSTab(${i})">${s.title || '슬라이드 ' + (i + 1)}</button>`;
  });
  html += `<button class="cst-tab add" onclick="addCustomSlide()">＋</button>`;
  cont.innerHTML = html;
}

function selectCSTab(idx) {
  activeCSIdx = idx;
  renderCSTabs();
  renderCSPanel();
}

// ── 슬라이드 추가 ──
function addCustomSlide() {
  const cats = getCategories();
  const qs = getAllQuestions();

  // 기본 그룹: 카테고리별 자동 생성
  const defaultGroups = [];
  if (cats.length > 0) {
    cats.forEach(cat => {
      const catQIds = (cat.questions || []).map(q => q.id);
      defaultGroups.push({ name: cat.name, qIds: catQIds, color: '#36A86F' });
    });
  } else if (qs.length > 0) {
    // 카테고리 없으면 6개씩 자동 분할
    for (let i = 0; i < qs.length; i += 6) {
      const chunk = qs.slice(i, i + 6);
      defaultGroups.push({
        name: `그룹 ${Math.floor(i / 6) + 1}`,
        qIds: chunk.map(q => q.id),
        color: ['#36A86F', '#4A90D9', '#E67E22', '#9B59B6', '#E74C3C'][Math.floor(i / 6) % 5]
      });
    }
  }

  customSlides.push({
    id: csIdSeq++,
    title: `커스텀 슬라이드 ${customSlides.length + 1}`,
    chartType: 'bar',
    tableStyle: 'A',
    tablePos: 'below',
    groups: defaultGroups,
  });
  activeCSIdx = customSlides.length - 1;
  renderCSTabs();
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

// ── 슬라이드 삭제 ──
function deleteCustomSlide(idx) {
  if (!confirm('이 슬라이드를 삭제하시겠습니까?')) return;
  customSlides.splice(idx, 1);
  activeCSIdx = Math.min(activeCSIdx, customSlides.length - 1);
  renderCSTabs();
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

// ── 패널 렌더링 ──
function renderCSPanel() {
  const cont = document.getElementById('customSlidePanel');
  if (!cont) return;

  if (activeCSIdx < 0 || activeCSIdx >= customSlides.length) {
    cont.innerHTML = `<div class="csp-empty">
      <div>＋ 버튼을 눌러 커스텀 슬라이드를 추가하세요</div>
      <small>문항을 자유롭게 그룹핑하여 새로운 차트와 표를 생성합니다</small>
    </div>`;
    return;
  }

  const s = customSlides[activeCSIdx];
  const qs = getAllQuestions();

  let html = '<div class="csp-form">';

  // ── 상단 설정 ──
  html += `<div class="csp-row">
    <div class="csp-field" style="flex:2">
      <label>슬라이드 제목</label>
      <input value="${s.title}" onchange="updateCS('title',this.value)" placeholder="제목 입력">
    </div>
    <div class="csp-field">
      <label>차트 유형</label>
      <select onchange="updateCS('chartType',this.value)">
        <option value="bar" ${s.chartType === 'bar' ? 'selected' : ''}>세로 막대</option>
        <option value="horizontalBar" ${s.chartType === 'horizontalBar' ? 'selected' : ''}>가로 막대</option>
        <option value="line" ${s.chartType === 'line' ? 'selected' : ''}>꺾은선</option>
        <option value="none" ${s.chartType === 'none' ? 'selected' : ''}>차트 없음</option>
      </select>
    </div>
    <div class="csp-field">
      <label>표 스타일</label>
      <select onchange="updateCS('tableStyle',this.value)">
        <option value="A" ${s.tableStyle === 'A' ? 'selected' : ''}>Executive Summary (가로)</option>
        <option value="B" ${s.tableStyle === 'B' ? 'selected' : ''}>정량 평가 (세로)</option>
        <option value="C" ${s.tableStyle === 'C' ? 'selected' : ''}>과정 개요 (2열)</option>
        <option value="none" ${s.tableStyle === 'none' ? 'selected' : ''}>표 없음</option>
      </select>
    </div>
    <div class="csp-field">
      <label>표 위치</label>
      <select onchange="updateCS('tablePos',this.value)">
        <option value="below" ${s.tablePos === 'below' ? 'selected' : ''}>차트 아래</option>
        <option value="above" ${s.tablePos === 'above' ? 'selected' : ''}>차트 위</option>
        <option value="only" ${s.tablePos === 'only' ? 'selected' : ''}>표만</option>
      </select>
    </div>
    <button class="csp-del-btn" onclick="deleteCustomSlide(${activeCSIdx})">삭제</button>
  </div>`;

  // ── 그룹 정의 ──
  html += `<div class="csp-groups-section">
    <div class="csp-groups-header">
      <span>📊 그룹 정의 (문항을 묶어서 평균값을 산출합니다)</span>
      <button class="csp-add-group-btn" onclick="addGroup()">＋ 그룹 추가</button>
    </div>
    <div class="color-toggle-row">
      <button class="color-toggle-btn ${s._colorUnified ? 'active' : ''}" onclick="toggleColorUnify(true)">🎨 색 통일</button>
      <button class="color-toggle-btn ${!s._colorUnified ? 'active' : ''}" onclick="toggleColorUnify(false)">🌈 색 다르게</button>
      ${s._colorUnified ? '<input type="color" value="'+(s._unifiedColor||'#36A86F')+'" style="width:28px;height:24px;border:none;cursor:pointer" onchange="setUnifiedColor(this.value)">' : ''}
    </div>`;

  s.groups.forEach((g, gi) => {
    const groupAvg = calcGroupAvg(g.qIds, qs);
    html += `<div class="csp-group-card" style="border-left-color:${g.color || '#36A86F'}">
      <div class="csp-group-card-header">
        <input class="csp-group-name-input" value="${g.name}"
               onchange="updateGroup(${gi},'name',this.value)" placeholder="그룹명">
        <span class="csp-group-avg" style="background:${g.color || '#36A86F'}">${groupAvg.toFixed(2)}점</span>
        <input type="color" value="${g.color || '#36A86F'}" style="width:28px;height:28px;border:none;cursor:pointer"
               onchange="updateGroup(${gi},'color',this.value)">
        <button class="csp-group-del" onclick="removeGroup(${gi})">✕</button>
      </div>
      <div class="csp-group-q-grid">`;

    qs.forEach(q => {
      const checked = g.qIds.includes(q.id) ? 'checked' : '';
      const label = q.label.length > 35 ? q.label.slice(0, 35) + '…' : q.label;
      html += `<label class="csp-gq-item">
        <input type="checkbox" ${checked}
               onchange="toggleGroupQ(${gi},'${q.id}',this.checked)">
        <span>${q.id}. ${label}</span>
        <span class="csp-gq-avg">${q.avg}</span>
      </label>`;
    });

    html += `</div></div>`;
  });
  html += `</div>`;

  // ── 미리보기 ──
  html += `<div class="csp-preview" id="cspPreview_${s.id}">
    <div class="csp-preview-title">${s.title}</div>
    <div class="csp-preview-inner" id="cspPreviewInner_${s.id}"></div>
  </div>`;

  // ── PPT 내보내기 버튼 ──
  html += `<div class="ppt-export-row" style="margin-top:12px">`;
  if (s.chartType !== 'none') {
    html += `<button class="ppt-export-btn" onclick="exportCSChartToPPT(${activeCSIdx})">📊 차트 PPT 복사</button>`;
  }
  if (s.tableStyle !== 'none') {
    html += `<button class="ppt-export-btn" onclick="exportCSTableToPPT(${activeCSIdx})">📋 표 PPT 복사</button>`;
  }
  html += `</div>`;

  html += '</div>';
  cont.innerHTML = html;

  // 차트 + 표 렌더링
  setTimeout(() => renderCSPreview(s), 50);
}

// ── 그룹 평균 계산 ──
function calcGroupAvg(qIds, allQs) {
  if (!qIds || qIds.length === 0) return 0;
  const matched = allQs.filter(q => qIds.includes(q.id));
  if (matched.length === 0) return 0;
  return matched.reduce((sum, q) => sum + (q.avg || 0), 0) / matched.length;
}

// ── 업데이트 함수들 ──
function updateCS(key, val) {
  if (activeCSIdx < 0) return;
  customSlides[activeCSIdx][key] = val;
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

function addGroup() {
  if (activeCSIdx < 0) return;
  const s = customSlides[activeCSIdx];
  const AUTO_COLORS = ['#36A86F','#4A90D9','#E67E22','#9B59B6','#E74C3C','#1ABC9C','#34495E','#F39C12','#2ECC71','#3498DB'];
  const gi = s.groups.length;
  const color = s._colorUnified ? (s._unifiedColor || '#36A86F') : AUTO_COLORS[gi % AUTO_COLORS.length];
  s.groups.push({
    name: `그룹 ${gi + 1}`,
    qIds: [],
    color: color,
  });
  renderCSPanel();
}

function toggleColorUnify(unified) {
  if (activeCSIdx < 0) return;
  const s = customSlides[activeCSIdx];
  s._colorUnified = unified;
  if (unified) {
    const baseColor = s._unifiedColor || '#36A86F';
    s.groups.forEach(g => { g.color = baseColor; });
  } else {
    const AUTO_COLORS = ['#36A86F','#4A90D9','#E67E22','#9B59B6','#E74C3C','#1ABC9C','#34495E','#F39C12','#2ECC71','#3498DB'];
    s.groups.forEach((g, i) => { g.color = AUTO_COLORS[i % AUTO_COLORS.length]; });
  }
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

function setUnifiedColor(color) {
  if (activeCSIdx < 0) return;
  const s = customSlides[activeCSIdx];
  s._unifiedColor = color;
  s.groups.forEach(g => { g.color = color; });
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

function removeGroup(gi) {
  if (activeCSIdx < 0) return;
  customSlides[activeCSIdx].groups.splice(gi, 1);
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

function updateGroup(gi, key, val) {
  if (activeCSIdx < 0) return;
  customSlides[activeCSIdx].groups[gi][key] = val;
  if (key === 'name') {
    // 이름만 바꿀 때는 전체 다시 렌더 안 함
    syncBuilderFromCustomSlides();
    return;
  }
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

function toggleGroupQ(gi, qId, checked) {
  if (activeCSIdx < 0) return;
  const g = customSlides[activeCSIdx].groups[gi];
  if (checked) {
    if (!g.qIds.includes(qId)) g.qIds.push(qId);
  } else {
    g.qIds = g.qIds.filter(id => id !== qId);
  }
  renderCSPanel();
  syncBuilderFromCustomSlides();
}

// ── 미리보기 렌더링 ──
function renderCSPreview(s) {
  const inner = document.getElementById(`cspPreviewInner_${s.id}`);
  if (!inner) return;
  const qs = getAllQuestions();

  let html = '';

  // 표 위 (above)
  if (s.tableStyle !== 'none' && s.tablePos === 'above') {
    html += renderPreviewTable(s, qs);
  }

  // 차트
  if (s.chartType !== 'none' && s.groups.length > 0) {
    html += `<div class="csp-chart-wrap"><canvas id="csChart_${s.id}" height="150"></canvas></div>`;
  }

  // 표 아래 (below) 또는 표만 (only)
  if (s.tableStyle !== 'none' && (s.tablePos === 'below' || s.tablePos === 'only')) {
    html += renderPreviewTable(s, qs);
  }

  inner.innerHTML = html;

  // 차트 그리기
  if (s.chartType !== 'none' && s.groups.length > 0) {
    renderCSChart(s, qs);
  }
}

// ── 차트 미리보기 ──
const csChartInstances = {};
function renderCSChart(s, qs) {
  const canvasId = `csChart_${s.id}`;
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  if (csChartInstances[s.id]) csChartInstances[s.id].destroy();

  const labels = s.groups.map(g => g.name);
  const values = s.groups.map(g => calcGroupAvg(g.qIds, qs));
  const colors = s.groups.map(g => g.color || '#36A86F');

  const isHorizontal = s.chartType === 'horizontalBar';
  const chartType = s.chartType === 'line' ? 'line' : 'bar';

  csChartInstances[s.id] = new Chart(canvas.getContext('2d'), {
    type: chartType,
    data: {
      labels,
      datasets: [{
        data: values,
        backgroundColor: colors.map(c => c + '99'),
        borderColor: colors,
        borderWidth: 2,
        borderRadius: 4,
      }]
    },
    options: {
      indexAxis: isHorizontal ? 'y' : 'x',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: ctx => ctx.parsed[isHorizontal ? 'x' : 'y'].toFixed(2) + '점' }
        }
      },
      scales: {
        [isHorizontal ? 'x' : 'y']: {
          beginAtZero: true,
          max: 5,
          ticks: { font: { size: 10 } }
        },
        [isHorizontal ? 'y' : 'x']: {
          ticks: { font: { size: 10, family: "'Nanum Gothic'" } }
        }
      }
    }
  });
}

// ── 표 미리보기 ──
function renderPreviewTable(s, qs) {
  if (s.tableStyle === 'A') return renderTableA(s, qs);
  if (s.tableStyle === 'B') return renderTableB(s, qs);
  if (s.tableStyle === 'C') return renderTableC(s, qs);
  return '';
}

// 스타일A: Executive Summary (가로)
function renderTableA(s, qs) {
  if (s.groups.length === 0) return '';
  const overall = calcGroupAvg(qs.map(q => q.id), qs);
  let html = '<table class="tpl-table-a"><thead><tr>';
  s.groups.forEach((g, i) => {
    const isLast = i === s.groups.length - 1;
    html += `<th${isLast ? ' class="highlight"' : ''}>${g.name}</th>`;
  });
  html += `<th class="highlight">과정 전반</th></tr></thead><tbody><tr>`;
  s.groups.forEach(g => {
    const avg = calcGroupAvg(g.qIds, qs);
    html += `<td>${avg.toFixed(2)}</td>`;
  });
  html += `<td><strong>${overall.toFixed(2)}</strong></td></tr>`;
  html += `<tr class="avg-row"><td colspan="${s.groups.length + 1}" style="text-align:left">
    전체 평균: <strong>${overall.toFixed(2)}</strong>점</td></tr>`;
  html += '</tbody></table>';
  return html;
}

// 스타일B: 정량 평가 (세로)
function renderTableB(s, qs) {
  let html = '<table class="tpl-table-b"><thead><tr>';
  html += '<th>항목</th><th>문항</th><th>평균</th><th>응답</th></tr></thead><tbody>';
  s.groups.forEach(g => {
    const avg = calcGroupAvg(g.qIds, qs);
    html += `<tr class="cat-row"><td colspan="2">${g.name}</td><td>${avg.toFixed(2)}</td><td></td></tr>`;
    const matched = qs.filter(q => g.qIds.includes(q.id));
    matched.forEach(q => {
      const cls = q.avg >= 4.5 ? 'score-high' : q.avg < 3.5 ? 'score-low' : '';
      html += `<tr><td></td><td style="text-align:left">${q.id}. ${q.label}</td>
               <td class="${cls}">${q.avg?.toFixed(2) || '-'}</td>
               <td>${q.count || '-'}</td></tr>`;
    });
  });
  html += '</tbody></table>';
  return html;
}

// 스타일C: 2열 키-값
function renderTableC(s, qs) {
  let html = '<table class="tpl-table-c"><tbody>';
  s.groups.forEach(g => {
    const avg = calcGroupAvg(g.qIds, qs);
    html += `<tr><td class="key">${g.name}</td><td>${avg.toFixed(2)}점 (${g.qIds.length}문항)</td></tr>`;
  });
  const overall = calcGroupAvg(qs.map(q => q.id), qs);
  html += `<tr><td class="key" style="background:var(--tpl-highlight)">과정 전반</td>
           <td><strong>${overall.toFixed(2)}점</strong></td></tr>`;
  html += '</tbody></table>';
  return html;
}

// ── 빌더 동기화 ──
function syncBuilderFromCustomSlides() {
  if (typeof builderState === 'undefined') return;
  const qs = getAllQuestions();

  // 기존 커스텀 슬라이드를 빌더에서 제거
  builderState.slides = builderState.slides.filter(s => s.type !== 'custom_quant');
  builderState.quantGroups = [];

  customSlides.forEach(cs => {
    builderState.slides.push({
      type: 'custom_quant',
      customSlideId: cs.id,
      label: cs.title,
      data: {
        title: cs.title,
        chartType: cs.chartType,
        tableStyle: cs.tableStyle,
        tablePos: cs.tablePos,
        groups: cs.groups.map(g => ({
          name: g.name,
          color: g.color,
          qIds: g.qIds,
          avg: calcGroupAvg(g.qIds, qs),
          questions: qs.filter(q => g.qIds.includes(q.id)),
        })),
      },
    });

    // quantGroups도 업데이트
    cs.groups.forEach(g => {
      builderState.quantGroups.push({
        id: `cs${cs.id}_g${cs.groups.indexOf(g)}`,
        title: g.name,
        questions: qs.filter(q => g.qIds.includes(q.id)),
      });
    });
  });

  // 빌더 리프레시
  if (typeof refreshBuilderSlideList === 'function') refreshBuilderSlideList();
  if (typeof refreshBuilderPreview === 'function') refreshBuilderPreview();
}

// ── PPT 내보내기 (커스텀 슬라이드) ──
async function exportCSChartToPPT(idx) {
  const s = customSlides[idx];
  if (!s) return;
  const qs = getAllQuestions();
  const labels = s.groups.map(g => g.name);
  const values = s.groups.map(g => calcGroupAvg(g.qIds, qs));
  const colors = s.groups.map(g => g.color || '#36A86F');

  const payload = {
    type: 'chart',
    title: s.title,
    data: {
      labels,
      values,
      chartType: s.chartType === 'horizontalBar' ? 'horizontalBar' : s.chartType === 'line' ? 'line' : 'bar',
      colors
    }
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
    a.download = `[복사용] ${s.title} 차트.pptx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch(e) {
    alert('차트 내보내기 실패: ' + e.message);
  }
}

async function exportCSTableToPPT(idx) {
  const s = customSlides[idx];
  if (!s) return;
  const qs = getAllQuestions();

  let headers = [];
  let rows = [];

  if (s.tableStyle === 'A') {
    // 가로형: 그룹명 = 열
    headers = [...s.groups.map(g => g.name), '과정 전반'];
    const vals = s.groups.map(g => calcGroupAvg(g.qIds, qs).toFixed(2));
    const overall = calcGroupAvg(qs.map(q => q.id), qs).toFixed(2);
    rows = [[...vals, overall]];
  } else if (s.tableStyle === 'B') {
    // 세로형: 항목/문항/평균/응답
    headers = ['항목', '문항', '평균', '응답'];
    s.groups.forEach(g => {
      const avg = calcGroupAvg(g.qIds, qs);
      rows.push([g.name, '', avg.toFixed(2), '']);
      qs.filter(q => g.qIds.includes(q.id)).forEach(q => {
        rows.push(['', q.label || '', (q.avg||0).toFixed(2), q.count || '']);
      });
    });
  } else if (s.tableStyle === 'C') {
    // 2열형
    headers = ['항목', '결과'];
    s.groups.forEach(g => {
      const avg = calcGroupAvg(g.qIds, qs);
      rows.push([g.name, `${avg.toFixed(2)}점 (${g.qIds.length}문항)`]);
    });
    const overall = calcGroupAvg(qs.map(q => q.id), qs);
    rows.push(['과정 전반', `${overall.toFixed(2)}점`]);
  }

  const payload = {
    type: 'table',
    title: s.title,
    data: { headers, rows }
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
    a.download = `[복사용] ${s.title} 표.pptx`;
    a.click();
    URL.revokeObjectURL(url);
  } catch(e) {
    alert('표 내보내기 실패: ' + e.message);
  }
}

// ── window에 노출 ──
window.initCustomSlides = initCustomSlides;
window.customSlides = customSlides;
window.exportCSChartToPPT = exportCSChartToPPT;
window.exportCSTableToPPT = exportCSTableToPPT;
window.analysisData = window.analysisData || null;
