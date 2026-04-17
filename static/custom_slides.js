/* =========================================================
   커스텀 슬라이드 구성 시스템 — custom_slides.js
   정량/정성 탭에서 슬라이드 단위로 보고서를 구성
   ========================================================= */

// ── 전역 상태 ────────────────────────────────────────────
const customSlides = [];   // [{id, title, chartType, questions:[], tablePos, tableTitle}]
let customActiveId = null;
let nextCustomId = 1;
let _previewDebounce = null;

// ── 초기화 (데이터 로드 후) ─────────────────────────────
function initCustomSlides(dashboardData) {
  window._dashData = dashboardData;
  const questions = dashboardData?.multi_result?.combined?.questions || [];
  window._allQuestions = questions;
  renderCustomSlideTabs();
  syncBuilderFromCustomSlides();
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 커스텀 슬라이드 CRUD
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

function addCustomSlide() {
  const id = nextCustomId++;
  const questions = window._allQuestions || [];
  customSlides.push({
    id,
    title: `슬라이드 ${id}`,
    chartType: 'bar',        // 'bar' | 'none'
    questions: questions.slice(0, 6).map(q => q.id),  // 기본: 첫 6문항
    tablePos: 'below',       // 'above' | 'below' | 'none'
    tableTitle: '문항별 평균 점수',
  });
  customActiveId = id;
  renderCustomSlideTabs();
  renderCustomSlidePanel(id);
  syncBuilderFromCustomSlides();
  if (typeof showToast === 'function') showToast('✅ 슬라이드 추가됨 — 보고서 빌더에 실시간 반영');
}

function deleteCustomSlide(id) {
  const idx = customSlides.findIndex(s => s.id === id);
  if (idx >= 0) customSlides.splice(idx, 1);
  if (customActiveId === id) {
    customActiveId = customSlides.length ? customSlides[customSlides.length - 1].id : null;
  }
  renderCustomSlideTabs();
  if (customActiveId) renderCustomSlidePanel(customActiveId);
  else document.getElementById('customSlidePanel').innerHTML =
    '<div class="csp-empty">+ 슬라이드 추가 버튼을 누르세요</div>';
  syncBuilderFromCustomSlides();
}

function selectCustomSlide(id) {
  customActiveId = id;
  renderCustomSlideTabs();
  renderCustomSlidePanel(id);
}

// ── 탭 네비게이션 렌더 ──────────────────────────────────
function renderCustomSlideTabs() {
  const nav = document.getElementById('customSlideTabs');
  if (!nav) return;
  nav.innerHTML = customSlides.map(s => `
    <button class="cst-tab ${s.id === customActiveId ? 'active' : ''}"
      onclick="selectCustomSlide(${s.id})" id="cst_${s.id}">
      📊 ${s.title}
    </button>`).join('') +
    `<button class="cst-tab add" onclick="addCustomSlide()">＋</button>`;
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 커스텀 슬라이드 편집 패널
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function renderCustomSlidePanel(id) {
  const slide = customSlides.find(s => s.id === id);
  const panel = document.getElementById('customSlidePanel');
  if (!slide || !panel) return;

  const allQs = window._allQuestions || [];

  const qChecks = allQs.map(q => {
    const checked = slide.questions.includes(q.id) ? 'checked' : '';
    const short = q.label.length > 45 ? q.label.slice(0, 45) + '…' : q.label;
    return `<label class="csp-q-item">
      <input type="checkbox" ${checked}
        onchange="toggleCustomQ(${id},'${q.id}',this.checked)">
      <span class="csp-q-label">${q.id} ${short} <em>${q.avg}</em></span>
    </label>`;
  }).join('');

  panel.innerHTML = `
  <div class="csp-form">
    <div class="csp-row">
      <div class="csp-field">
        <label>슬라이드 제목</label>
        <input value="${slide.title}" oninput="updateCustomSlide(${id},'title',this.value)">
      </div>
      <div class="csp-field" style="max-width:140px">
        <label>차트 종류</label>
        <select onchange="updateCustomSlide(${id},'chartType',this.value)">
          <option value="bar" ${slide.chartType==='bar'?'selected':''}>막대 차트</option>
          <option value="none" ${slide.chartType==='none'?'selected':''}>차트 없음</option>
        </select>
      </div>
      <div class="csp-field" style="max-width:140px">
        <label>표 위치</label>
        <select onchange="updateCustomSlide(${id},'tablePos',this.value)">
          <option value="below" ${slide.tablePos==='below'?'selected':''}>차트 아래</option>
          <option value="above" ${slide.tablePos==='above'?'selected':''}>차트 위</option>
          <option value="none"  ${slide.tablePos==='none' ?'selected':''}>표 없음</option>
        </select>
      </div>
      <div class="csp-field">
        <label>표 제목</label>
        <input value="${slide.tableTitle}" oninput="updateCustomSlide(${id},'tableTitle',this.value)">
      </div>
      <button class="csp-del-btn" onclick="deleteCustomSlide(${id})">🗑 슬라이드 삭제</button>
    </div>

    <div class="csp-q-section">
      <div class="csp-q-header">
        <span>포함 문항 선택 <em>${slide.questions.length}개 선택됨</em></span>
        <div style="display:flex;gap:6px">
          <button class="csp-qbtn" onclick="selectAllCustomQ(${id},true)">전체 선택</button>
          <button class="csp-qbtn" onclick="selectAllCustomQ(${id},false)">전체 해제</button>
          <button class="csp-qbtn" onclick="selectChunkCustomQ(${id},6)">6개씩 선택</button>
        </div>
      </div>
      <div class="csp-q-grid">${qChecks}</div>
    </div>

    <!-- 미리보기 (차트) -->
    <div class="csp-preview">
      <div class="csp-preview-title">📐 미리보기</div>
      <div class="csp-preview-inner" id="csp_preview_${id}">
        <canvas id="csp_chart_${id}" style="max-height:200px"></canvas>
      </div>
      <div class="csp-preview-table" id="csp_table_${id}"></div>
    </div>
  </div>`;

  renderCustomPreview(id);
}

// ── 미리보기 렌더 ────────────────────────────────────────
function renderCustomPreview(id) {
  const slide = customSlides.find(s => s.id === id);
  if (!slide) return;
  const allQs = window._allQuestions || [];
  const qs = allQs.filter(q => slide.questions.includes(q.id));

  // 차트
  const canvasEl = document.getElementById(`csp_chart_${id}`);
  if (canvasEl) {
    // 기존 차트 파기
    if (window._cspCharts && window._cspCharts[id]) {
      try { window._cspCharts[id].destroy(); } catch(e) {}
    }
    if (!window._cspCharts) window._cspCharts = {};

    if (slide.chartType === 'bar' && qs.length > 0) {
      canvasEl.style.display = '';
      window._cspCharts[id] = new Chart(canvasEl, {
        type: 'bar',
        data: {
          labels: qs.map(q => q.label.length > 14 ? q.label.slice(0,14)+'…' : q.label),
          datasets: [{
            data: qs.map(q => q.avg),
            backgroundColor: qs.map(q =>
              q.avg >= 4.5 ? '#10b981' : q.avg >= 4.0 ? '#2563eb' : '#f59e0b'
            ),
            borderRadius: 4,
          }]
        },
        options: {
          responsive: true,
          plugins: { legend: { display: false },
            tooltip: { callbacks: { label: ctx => `${ctx.raw.toFixed(2)}점` } }
          },
          scales: {
            y: { min: 3, max: 5, ticks: { stepSize: 0.5 } },
            x: { ticks: { font: { size: 10 } } }
          }
        }
      });
    } else {
      canvasEl.style.display = 'none';
    }
  }

  // 표
  const tableEl = document.getElementById(`csp_table_${id}`);
  if (tableEl && slide.tablePos !== 'none' && qs.length > 0) {
    tableEl.innerHTML = `
      <div class="csp-tbl-title">${slide.tableTitle}</div>
      <table class="score-table" style="font-size:12px">
        <thead><tr><th>번호</th><th>문항</th><th>평균</th><th>평가</th></tr></thead>
        <tbody>${qs.map(q => `
          <tr>
            <td>${q.id}</td>
            <td>${q.label}</td>
            <td class="${q.avg>=4.5?'score-high':q.avg>=4?'score-mid':'score-low'}">${q.avg.toFixed(2)}</td>
            <td>${q.avg>=4.5?'상':q.avg>=4?'중':'하'}</td>
          </tr>`).join('')}
        </tbody>
      </table>`;
    tableEl.style.display = '';
  } else if (tableEl) {
    tableEl.style.display = 'none';
  }
}

// ── 상태 업데이트 ────────────────────────────────────────
function updateCustomSlide(id, key, val) {
  const slide = customSlides.find(s => s.id === id);
  if (!slide) return;
  slide[key] = val;
  // 탭 이름 업데이트
  if (key === 'title') {
    const tabEl = document.getElementById(`cst_${id}`);
    if (tabEl) tabEl.textContent = `📊 ${val}`;
  }
  // 디바운스로 미리보기 + 빌더 동기화
  clearTimeout(_previewDebounce);
  _previewDebounce = setTimeout(() => {
    renderCustomPreview(id);
    syncBuilderFromCustomSlides();
  }, 400);
}

function toggleCustomQ(id, qId, checked) {
  const slide = customSlides.find(s => s.id === id);
  if (!slide) return;
  if (checked && !slide.questions.includes(qId)) slide.questions.push(qId);
  if (!checked) slide.questions = slide.questions.filter(q => q !== qId);
  // 선택 수 업데이트
  const panel = document.getElementById(`csp_preview_${id}`)?.closest('.csp-form');
  if (panel) {
    const em = panel.querySelector('.csp-q-header em');
    if (em) em.textContent = `${slide.questions.length}개 선택됨`;
  }
  clearTimeout(_previewDebounce);
  _previewDebounce = setTimeout(() => {
    renderCustomPreview(id);
    syncBuilderFromCustomSlides();
  }, 400);
}

function selectAllCustomQ(id, all) {
  const slide = customSlides.find(s => s.id === id);
  if (!slide) return;
  const allQs = window._allQuestions || [];
  slide.questions = all ? allQs.map(q => q.id) : [];
  renderCustomSlidePanel(id);
  syncBuilderFromCustomSlides();
}

function selectChunkCustomQ(id, size) {
  const slide = customSlides.find(s => s.id === id);
  if (!slide) return;
  const allQs = window._allQuestions || [];
  // 현재 선택된 것들의 마지막 인덱스 이후 size개
  const lastIdx = allQs.findIndex(q => slide.questions.includes(q.id));
  const startIdx = lastIdx < 0 ? 0 : lastIdx;
  slide.questions = allQs.slice(startIdx, startIdx + size).map(q => q.id);
  renderCustomSlidePanel(id);
  syncBuilderFromCustomSlides();
}

// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
// 빌더 자동 동기화
// ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
function syncBuilderFromCustomSlides() {
  if (typeof builderState === 'undefined') return;
  const allQs = window._allQuestions || [];

  // 기존 quant_chart 슬라이드 전부 제거
  builderState.slides = builderState.slides.filter(s => s.type !== 'quant_chart');
  builderState.quantGroups = [];

  // customSlides 기반으로 재생성
  const qualIdx = builderState.slides.findIndex(s => s.type === 'qual_text');
  const insertAt = qualIdx >= 0 ? qualIdx : builderState.slides.length - 1;

  customSlides.forEach((cs, i) => {
    const questions = allQs.filter(q => cs.questions.includes(q.id));
    const g = {
      id: cs.id,
      title: cs.title,
      questions,
      allQuestions: allQs,
      tablePos: cs.tablePos,
      tableTitle: cs.tableTitle,
      chartType: cs.chartType,
    };
    builderState.quantGroups.push(g);
    builderState.slides.splice(insertAt + i, 0, {
      type: 'quant_chart', groupId: cs.id, data: {}
    });
  });

  if (typeof renderSlideList === 'function') renderSlideList();
  if (typeof refreshBuilderMiniList === 'function') refreshBuilderMiniList();
}
