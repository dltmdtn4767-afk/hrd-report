/* HRD 결과보고서 생성기 — Frontend Logic */

let state = {
    sessionId: null,
    file: null,
    preview: [],
    currentSlide: 0,
};

// AI 상태 확인
(async function checkAIStatus() {
    try {
        const res = await fetch('/api/status');
        const status = await res.json();
        const badge = document.getElementById('aiStatusBadge');
        if (badge) {
            if (status.api_key_set) {
                badge.className = 'ai-status connected';
                badge.innerHTML = `<span class="dot"></span> AI 연결됨 (${status.model}) | 샘플 ${status.sample_count}개`;
            } else {
                badge.className = 'ai-status disconnected';
                badge.innerHTML = '<span class="dot"></span> AI 미연결 — config.json에 API 키를 입력하세요';
            }
        }
    } catch (e) { console.log('status check failed'); }
})();

// ═══════════════════════════════════════════
// Step 1: 파일 업로드
// ═══════════════════════════════════════════

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0 && files[0].name.match(/\.xlsx?$/i)) {
        state.file = files[0];
        handleFile(files[0]);
    }
});
fileInput.addEventListener('change', e => {
    if (e.target.files.length > 0) {
        state.file = e.target.files[0];
        handleFile(e.target.files[0]);
    }
});

function handleFile(file) {
    dropZone.innerHTML = `
        <div class="drop-icon">✅</div>
        <p class="drop-text">${file.name}</p>
        <p class="drop-hint">${(file.size / 1024).toFixed(0)} KB</p>
    `;
    // 바로 분석 시작
    analyzeData();
}

async function analyzeData() {
    if (!state.file) return;
    showLoading('데이터 분석 중...');
    
    const formData = new FormData();
    formData.append('file', state.file);
    
    const sheetSelect = document.getElementById('sheetSelect');
    if (sheetSelect.value) {
        formData.append('sheet', sheetSelect.value);
    }
    
    try {
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await res.json();
        
        if (res.ok) {
            state.sessionId = data.session_id;
            showAnalysis(data.summary, data.ai_result);
        } else {
            alert('분석 실패: ' + (data.detail || '알 수 없는 오류'));
        }
    } catch (err) {
        alert('서버 연결 실패: ' + err.message);
    } finally {
        hideLoading();
    }
}

// ═══════════════════════════════════════════
// Step 2: AI 분석 결과
// ═══════════════════════════════════════════

function showAnalysis(summary, aiResult) {
    showStep(2);
    
    // 데이터 요약
    const summaryHtml = `
        <div class="summary-item"><span class="summary-label">고객사</span><span class="summary-value">${summary.company}</span></div>
        <div class="summary-item"><span class="summary-label">과정명</span><span class="summary-value">${summary.course_name}</span></div>
        <div class="summary-item"><span class="summary-label">문항 수</span><span class="summary-value">${summary.total_questions}개</span></div>
        <div class="summary-item"><span class="summary-label">카테고리</span><span class="summary-value">${summary.categories}개</span></div>
        <div class="summary-item"><span class="summary-label">모듈</span><span class="summary-value">${summary.has_modules ? '✅ 있음' : '❌ 없음'}</span></div>
        <div class="summary-item"><span class="summary-label">강사</span><span class="summary-value">${summary.num_instructors}명</span></div>
        <div class="summary-item"><span class="summary-label">주관식</span><span class="summary-value">${summary.open_ended_count}개</span></div>
        <div class="summary-item"><span class="summary-label">응답인원</span><span class="summary-value">${summary.response_count}명</span></div>
        <div class="summary-item"><span class="summary-label">전체 평균</span><span class="summary-value" style="color: ${summary.overall_average >= 4.5 ? 'var(--success)' : summary.overall_average >= 4.0 ? 'var(--warning)' : 'var(--danger)'}">${summary.overall_average}점</span></div>
    `;
    document.getElementById('dataSummary').innerHTML = summaryHtml;
    
    // AI 추론 결과
    const matched = aiResult.matched_pattern || '규칙 기반';
    const matchedSample = aiResult.matched_sample || '';
    const similarity = aiResult.similarity ? `(유사도 ${Math.round(aiResult.similarity * 100)}%)` : '';
    const reasoning = aiResult.reasoning || '';
    const suggestions = (aiResult.additional_suggestions || []).map(s => `<li>${s}</li>`).join('');
    
    let aiHtml = '';
    if (matchedSample) {
        aiHtml += `<div class="summary-item">
            <span class="summary-label">디자인 베이스</span>
            <span class="summary-value" style="font-size:12px;max-width:250px;overflow:hidden;text-overflow:ellipsis">📋 ${matchedSample}</span>
        </div>`;
    }
    aiHtml += `<div class="summary-item">
        <span class="summary-label">추론 방식</span>
        <span class="summary-value">${matched} ${similarity}</span>
    </div>`;
    if (reasoning) aiHtml += `<div class="ai-reasoning">${reasoning}</div>`;
    if (suggestions) aiHtml += `<ul style="margin-top:12px;padding-left:20px;font-size:13px;color:var(--text-secondary)">${suggestions}</ul>`;
    
    document.getElementById('aiResult').innerHTML = aiHtml;
    
    // 슬라이드 구성
    const slides = aiResult.recommended_slides || [];
    const chipsHtml = slides.map(s => {
        let cls = '';
        if (s.type?.includes('exec')) cls = 'exec';
        else if (s.type?.includes('quant')) cls = 'quant';
        else if (s.type?.includes('qual')) cls = 'qual';
        return `<div class="slide-chip ${cls}">📄 ${s.title || s.type}</div>`;
    }).join('');
    document.getElementById('slideStructure').innerHTML = `<div class="slide-chips">${chipsHtml}</div>`;
}

// ═══════════════════════════════════════════
// Step 3: 보고서 생성 + 미리보기
// ═══════════════════════════════════════════

async function generateReport() {
    showLoading('보고서 생성 중...');
    
    try {
        const res = await fetch(`/api/generate/${state.sessionId}`, { method: 'POST' });
        const data = await res.json();
        
        if (data.success) {
            state.preview = data.preview;
            state.currentSlide = 0;
            showStep(3);
            renderPreview();
            renderReview(data.review);
        } else {
            alert('생성 실패: ' + data.error);
        }
    } catch (err) {
        alert('서버 오류: ' + err.message);
    } finally {
        hideLoading();
    }
}

function renderPreview() {
    const slides = state.preview;
    if (!slides.length) return;
    
    // 썸네일
    const thumbs = slides.map((s, i) => `
        <div class="thumbnail ${i === state.currentSlide ? 'active' : ''}" onclick="goToSlide(${i})">
            <div class="thumbnail-num">S${s.index}</div>
            <div class="thumbnail-title">${(s.title || s.layout).substring(0, 10)}</div>
        </div>
    `).join('');
    document.getElementById('slideThumbnails').innerHTML = thumbs;
    
    // 카운터
    document.getElementById('slideCounter').textContent = `${state.currentSlide + 1} / ${slides.length}`;
    
    // 슬라이드 내용
    renderSlideContent(slides[state.currentSlide]);
}

function renderSlideContent(slide) {
    let html = '';
    html += `<div class="slide-layout-badge">${slide.layout}</div>`;
    
    if (slide.title) {
        html += `<div class="slide-title">${slide.title}</div>`;
    }
    
    // 텍스트
    const otherTexts = (slide.texts || []).filter(t => t !== slide.title).slice(0, 5);
    if (otherTexts.length) {
        html += otherTexts.map(t => `<p style="margin:6px 0;color:var(--text-secondary);font-size:13px">${escapeHtml(t)}</p>`).join('');
    }
    
    // 표
    for (const tbl of (slide.tables || [])) {
        html += `<p style="margin:12px 0 6px;font-size:12px;color:var(--text-muted)">${tbl.name} (${tbl.size})</p>`;
        html += '<table class="preview-table"><thead><tr>';
        if (tbl.rows.length > 0) {
            html += tbl.rows[0].map(c => `<th>${escapeHtml(c)}</th>`).join('');
            html += '</tr></thead><tbody>';
            for (let r = 1; r < Math.min(tbl.rows.length, 8); r++) {
                html += '<tr>' + tbl.rows[r].map(c => `<td>${escapeHtml(c)}</td>`).join('') + '</tr>';
            }
            if (tbl.rows.length > 8) {
                html += `<tr><td colspan="${tbl.rows[0].length}" style="text-align:center;color:var(--text-muted)">... ${tbl.rows.length - 8}행 더</td></tr>`;
            }
            html += '</tbody>';
        }
        html += '</table>';
    }
    
    // 차트
    for (const chart of (slide.charts || [])) {
        html += `<div class="chart-bar-container">`;
        const maxVal = Math.max(...(chart.values || [1]), 5);
        (chart.values || []).forEach((v, i) => {
            const pct = (v / 5) * 100;
            const color = v >= 4.5 ? '#10b981' : v >= 4.0 ? '#6366f1' : '#ef4444';
            html += `
                <div class="chart-bar-item">
                    <span class="chart-bar-label">Q${i+1}</span>
                    <div class="chart-bar" style="width:${pct}%;background:${color}"></div>
                    <span class="chart-bar-value">${v.toFixed(2)}</span>
                </div>`;
        });
        html += '</div>';
    }
    
    // 그룹
    for (const grp of (slide.groups || [])) {
        html += `<div class="preview-group">`;
        if (grp[0]) html += `<div class="preview-group-title">${escapeHtml(grp[0]).substring(0, 50)}</div>`;
        if (grp[1]) html += `<div class="preview-group-text">${escapeHtml(grp[1]).substring(0, 200)}</div>`;
        html += '</div>';
    }
    
    if (!html.trim() || html.length < 100) {
        html += `<p style="text-align:center;color:var(--text-muted);padding:40px">이미지/도형 슬라이드</p>`;
    }
    
    document.getElementById('slideContent').innerHTML = html;
}

function goToSlide(idx) {
    state.currentSlide = idx;
    renderPreview();
}

function prevSlide() {
    if (state.currentSlide > 0) { state.currentSlide--; renderPreview(); }
}

function nextSlide() {
    if (state.currentSlide < state.preview.length - 1) { state.currentSlide++; renderPreview(); }
}

// ═══════════════════════════════════════════
// 검증 결과
// ═══════════════════════════════════════════

function renderReview(review) {
    if (!review) return;
    
    const scoreColor = review.score >= 80 ? 'var(--success)' : review.score >= 50 ? 'var(--warning)' : 'var(--danger)';
    let html = `<div class="review-score" style="color:${scoreColor}">${review.score}점</div>`;
    
    for (const check of (review.checks || [])) {
        const icon = check.status === 'pass' ? '✅' : check.status === 'warn' ? '⚠️' : '❌';
        const cls = `check-${check.status}`;
        html += `<div class="check-item ${cls}">${icon} ${check.detail}</div>`;
    }
    
    document.getElementById('reviewResult').innerHTML = html;
}

// ═══════════════════════════════════════════
// 수정 요청
// ═══════════════════════════════════════════

async function sendModification() {
    const input = document.getElementById('modifyInput');
    const msg = input.value.trim();
    if (!msg) return;
    
    addChatMessage(msg, 'user');
    input.value = '';
    
    try {
        const res = await fetch(`/api/modify/${state.sessionId}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ message: msg }),
        });
        const data = await res.json();
        
        if (data.success) {
            const detail = data.modification?.detail || '수정 반영됨';
            addChatMessage(`✅ ${detail}`, 'ai');
            state.preview = data.preview;
            renderPreview();
        }
    } catch (err) {
        addChatMessage('오류: ' + err.message, 'ai');
    }
}

function addChatMessage(text, type) {
    const container = document.getElementById('chatMessages');
    container.innerHTML += `<div class="chat-msg ${type}">${escapeHtml(text)}</div>`;
    container.scrollTop = container.scrollHeight;
}

// ═══════════════════════════════════════════
// 다운로드 / 재생성
// ═══════════════════════════════════════════

function downloadReport() {
    if (!state.sessionId) return;
    window.location.href = `/api/download/${state.sessionId}`;
}

async function regenerate() {
    await generateReport();
}

// ═══════════════════════════════════════════
// 샘플 관리
// ═══════════════════════════════════════════

function showSamples() {
    document.getElementById('sampleModal').classList.remove('hidden');
    loadSamples();
}

function closeSamples() {
    document.getElementById('sampleModal').classList.add('hidden');
}

async function loadSamples() {
    try {
        const res = await fetch('/api/samples');
        const patterns = await res.json();
        
        const html = patterns.map(p => `
            <div class="sample-item">
                <span>${p.name}</span>
                <span style="color:var(--text-muted)">${p.slide_count}슬라이드 | Exec:${p.counts?.exec_slides || 0} Quant:${p.counts?.quant_slides || 0}</span>
            </div>
        `).join('');
        document.getElementById('sampleList').innerHTML = html || '<p style="color:var(--text-muted);text-align:center">등록된 샘플 없음</p>';
    } catch (err) {
        console.error(err);
    }
}

const sampleDrop = document.getElementById('sampleDropZone');
const sampleInput = document.getElementById('sampleInput');
sampleDrop.addEventListener('click', () => sampleInput.click());
sampleDrop.addEventListener('dragover', e => { e.preventDefault(); sampleDrop.classList.add('dragover'); });
sampleDrop.addEventListener('dragleave', () => sampleDrop.classList.remove('dragover'));
sampleDrop.addEventListener('drop', async e => {
    e.preventDefault();
    sampleDrop.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        const formData = new FormData();
        formData.append('file', files[0]);
        showLoading('샘플 분석 중...');
        try {
            await fetch('/api/samples/add', { method: 'POST', body: formData });
            await loadSamples();
        } catch (err) { alert(err.message); }
        hideLoading();
    }
});

// ═══════════════════════════════════════════
// 유틸
// ═══════════════════════════════════════════

function showStep(num) {
    document.querySelectorAll('.step').forEach(s => s.classList.add('hidden'));
    document.getElementById(`step-${['', 'upload', 'analysis', 'preview'][num]}`).classList.remove('hidden');
}

function goBack(step) { showStep(step); }
function showLoading(text) { document.getElementById('loadingText').textContent = text; document.getElementById('loading').classList.remove('hidden'); }
function hideLoading() { document.getElementById('loading').classList.add('hidden'); }
function escapeHtml(text) { const d = document.createElement('div'); d.textContent = text; return d.innerHTML; }

// 키보드 단축키
document.addEventListener('keydown', e => {
    if (e.key === 'ArrowLeft') prevSlide();
    if (e.key === 'ArrowRight') nextSlide();
});
