/* =====================================================================
   Auditoria de Distrito de Gás — Main JavaScript
   ===================================================================== */

// --- State ---
let currentSection = 'config';
let lightboxImages = [];
let lightboxIndex = 0;
let evtSource = null;
let pipelineTimer = null;
let pipelineStartTime = 0;

// --- Navigation ---
document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        const section = item.dataset.section;
        navigateTo(section);
    });
});

function navigateTo(section) {
    currentSection = section;
    document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
    document.querySelector(`[data-section="${section}"]`).classList.add('active');
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.getElementById(`sec-${section}`).classList.add('active');

    // Lazy load content
    if (section === 'dados') loadDataPreview();
    if (section === 'graficos') loadGraficos();
    if (section === 'diagramas') loadDiagramas();
    if (section === 'textos') loadTextos();
    if (section === 'downloads') loadDownloads();
}

// --- Theme Toggle ---
const themeToggle = document.getElementById('theme-toggle');
const savedTheme = localStorage.getItem('theme');
if (savedTheme) document.documentElement.setAttribute('data-theme', savedTheme);

themeToggle.addEventListener('click', () => {
    const current = document.documentElement.getAttribute('data-theme');
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('theme', next);
});

// --- File Upload ---
const uploadZone = document.getElementById('upload-zone');
const fileInput = document.getElementById('file-input');
const uploadStatus = document.getElementById('upload-status');
const uploadStatusText = document.getElementById('upload-status-text');

uploadZone.addEventListener('click', () => fileInput.click());
uploadZone.addEventListener('dragover', (e) => { e.preventDefault(); uploadZone.classList.add('drag-over'); });
uploadZone.addEventListener('dragleave', () => uploadZone.classList.remove('drag-over'));
uploadZone.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadZone.classList.remove('drag-over');
    if (e.dataTransfer.files.length) uploadFile(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', () => { if (fileInput.files.length) uploadFile(fileInput.files[0]); });

async function uploadFile(file) {
    const form = new FormData();
    form.append('file', file);
    try {
        const res = await fetch('/api/upload', { method: 'POST', body: form });
        const data = await res.json();
        if (res.ok) {
            uploadStatus.classList.remove('hidden', 'error');
            uploadStatusText.textContent = `${data.filename} enviado (${data.size_mb} MB)`;
            showToast('Arquivo enviado com sucesso!', 'success');
            checkStartBtn();
        } else {
            uploadStatus.classList.remove('hidden');
            uploadStatus.classList.add('error');
            uploadStatusText.textContent = data.detail || 'Erro no upload';
            showToast(data.detail || 'Erro no upload', 'error');
        }
    } catch (err) {
        showToast('Erro de conexão', 'error');
    }
}

// --- API Key ---
const apiKeyInput = document.getElementById('api-key-input');
const toggleKeyBtn = document.getElementById('toggle-key-btn');
const keyStatus = document.getElementById('key-status');

toggleKeyBtn.addEventListener('click', () => {
    apiKeyInput.type = apiKeyInput.type === 'password' ? 'text' : 'password';
});

apiKeyInput.addEventListener('input', () => {
    const val = apiKeyInput.value.trim();
    if (val.length > 10) {
        keyStatus.classList.remove('hidden');
        checkStartBtn();
    } else {
        keyStatus.classList.add('hidden');
        checkStartBtn();
    }
});

// --- Start Button ---
const btnStart = document.getElementById('btn-start');

function checkStartBtn() {
    const hasFile = !uploadStatus.classList.contains('hidden') && !uploadStatus.classList.contains('error');
    const hasKey = apiKeyInput.value.trim().length > 10;
    const mode = document.querySelector('input[name="mode"]:checked').value;
    // For "montar" and "step" modes, only file is needed (no API key initially)
    btnStart.disabled = !(hasFile && (hasKey || mode === 'montar' || mode === 'step'));
}

document.querySelectorAll('input[name="mode"]').forEach(r => r.addEventListener('change', checkStartBtn));

btnStart.addEventListener('click', startPipeline);

// --- Cancel Button ---
const btnCancel = document.getElementById('btn-cancel');
btnCancel.addEventListener('click', async () => {
    try {
        await fetch('/api/pipeline/cancel', { method: 'POST' });
        btnCancel.classList.add('hidden');
    } catch {}
});

// --- Pipeline Execution ---
async function startPipeline() {
    const apiKey = apiKeyInput.value.trim();
    const mode = document.querySelector('input[name="mode"]:checked').value;

    // Step-by-step mode: run extraction first, then navigate to Dados tab
    if (mode === 'step') {
        await runPhaseExtract();
        return;
    }

    btnStart.disabled = true;
    btnStart.textContent = 'Iniciando...';

    try {
        const res = await fetch('/api/pipeline/start', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ api_key: apiKey, mode: mode }),
        });
        const data = await res.json();
        if (res.ok) {
            navigateTo('pipeline');
            connectSSE();
            pipelineStartTime = Date.now();
            pipelineTimer = setInterval(updateTimer, 1000);
            document.getElementById('pipeline-subtitle').textContent = 'Pipeline em execução...';
            btnCancel.classList.remove('hidden');
            showToast('Pipeline iniciado!', 'info');
        } else {
            showToast(data.detail || 'Erro ao iniciar', 'error');
            btnStart.disabled = false;
            btnStart.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> Iniciar Pipeline';
        }
    } catch (err) {
        showToast('Erro de conexão', 'error');
        btnStart.disabled = false;
    }
}

function connectSSE() {
    if (evtSource) evtSource.close();
    evtSource = new EventSource('/api/pipeline/events');

    evtSource.addEventListener('phase_start', (e) => {
        const data = JSON.parse(e.data);
        updatePhase(data.phase, 'active');
        addLog(`Fase ${data.phase}: ${data.phase_name}`, 'step-start');
    });

    evtSource.addEventListener('step_start', (e) => {
        const data = JSON.parse(e.data);
        updateSubstep(data.step, 'active');
        updateChapterCard(data.chapter, 'active');
        updateProgress(data.progress, data.completed || 0, 36);
        addLog(`  Iniciando: ${data.step_label}`, 'step-start');
    });

    evtSource.addEventListener('step_complete', (e) => {
        const data = JSON.parse(e.data);
        updateSubstep(data.step, 'complete');
        updateProgress(data.progress, data.completed, data.total);
        updateChapterProgress(data.chapter, data.step);
        addLog(`  Concluído: ${data.step_label}`, 'step-complete');
    });

    evtSource.addEventListener('phase_complete', (e) => {
        const data = JSON.parse(e.data);
        updatePhase(data.phase, 'complete');
    });

    evtSource.addEventListener('done', (e) => {
        evtSource.close();
        clearInterval(pipelineTimer);
        updateProgress(1.0, 36, 36);
        document.getElementById('pipeline-subtitle').textContent = 'Pipeline concluído com sucesso!';
        document.getElementById('pipeline-subtitle').style.color = 'var(--verde)';
        addLog('Pipeline concluído!', 'step-complete');
        // Mark all phases complete
        [0,1,2,3,4].forEach(p => updatePhase(p, 'complete'));
        showToast('Pipeline concluído! Relatório disponível para download.', 'success');
        btnStart.disabled = false;
        btnCancel.classList.add('hidden');
        btnStart.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> Iniciar Pipeline';
    });

    evtSource.addEventListener('error', (e) => {
        let detail = 'Erro desconhecido';
        try { detail = JSON.parse(e.data).detail; } catch {}
        evtSource.close();
        clearInterval(pipelineTimer);
        document.getElementById('pipeline-subtitle').textContent = 'Erro na execução';
        document.getElementById('pipeline-subtitle').style.color = 'var(--vermelho)';
        addLog(`ERRO: ${detail}`, 'error');
        showToast(`Erro: ${detail}`, 'error');
        btnStart.disabled = false;
        btnCancel.classList.add('hidden');
        btnStart.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> Iniciar Pipeline';
    });

    evtSource.addEventListener('cancelled', (e) => {
        evtSource.close();
        clearInterval(pipelineTimer);
        document.getElementById('pipeline-subtitle').textContent = 'Pipeline cancelado';
        addLog('Pipeline cancelado pelo usuário', 'error');
        showToast('Pipeline cancelado', 'info');
        btnStart.disabled = false;
        btnCancel.classList.add('hidden');
        btnStart.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polygon points="5 3 19 12 5 21 5 3"/></svg> Iniciar Pipeline';
    });

    // Handle connection errors
    evtSource.onerror = () => {
        // EventSource auto-reconnects, but if pipeline is done, close cleanly
        setTimeout(async () => {
            try {
                const res = await fetch('/api/pipeline/status');
                const st = await res.json();
                if (!st.running) evtSource.close();
            } catch {}
        }, 2000);
    };
}

// --- Progress Updates ---
function updateProgress(progress, completed, total) {
    const pct = Math.round(progress * 100);
    document.getElementById('progress-fill').style.width = pct + '%';
    document.getElementById('progress-pct').textContent = pct + '%';
    document.getElementById('progress-steps').textContent = `${completed}/${total} etapas`;
}

function updateTimer() {
    const elapsed = Math.floor((Date.now() - pipelineStartTime) / 1000);
    const min = Math.floor(elapsed / 60);
    const sec = String(elapsed % 60).padStart(2, '0');
    document.getElementById('progress-time').textContent = `${min}:${sec}`;
}

function updatePhase(phase, status) {
    const el = document.querySelector(`.phase-step[data-phase="${phase}"]`);
    if (!el) return;
    el.classList.remove('active', 'complete');
    el.classList.add(status);
}

function updateSubstep(step, status) {
    const el = document.querySelector(`.substep[data-step="${step}"]`);
    if (!el) return;
    el.classList.remove('pending', 'active', 'complete', 'error');
    el.classList.add(status);
}

function updateChapterCard(chapter, status) {
    if (!chapter) return;
    const card = document.getElementById(`chapter-card-${chapter}`);
    if (card && status === 'active' && !card.classList.contains('complete')) {
        card.classList.add('active');
    }
}

function updateChapterProgress(chapter, step) {
    if (!chapter) return;
    const card = document.getElementById(`chapter-card-${chapter}`);
    if (!card) return;
    const substeps = card.querySelectorAll('.substep');
    const completed = card.querySelectorAll('.substep.complete').length;
    const pct = Math.round((completed / substeps.length) * 100);
    const fill = document.getElementById(`chapter-progress-${chapter}`);
    if (fill) fill.style.width = pct + '%';
    if (pct === 100) {
        card.classList.remove('active');
        card.classList.add('complete');
    }
}

function addLog(text, cls) {
    const entries = document.getElementById('log-entries');
    const entry = document.createElement('div');
    entry.className = `log-entry ${cls || ''}`;
    const time = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    entry.textContent = `[${time}] ${text}`;
    entries.appendChild(entry);
    entries.scrollTop = entries.scrollHeight;
}

// =====================================================================
// DADOS - Phase-by-phase extraction & preview
// =====================================================================
let dataPreviewLoaded = false;

// Button references
const btnExtract = document.getElementById('btn-extract');
const btnGenGraphs = document.getElementById('btn-gen-graphs');
const dataStatus = document.getElementById('data-status');
const dataStatusText = document.getElementById('data-status-text');

btnExtract.addEventListener('click', () => runPhaseExtract());
btnGenGraphs.addEventListener('click', () => runPhaseGraphs());

async function runPhaseExtract() {
    btnExtract.disabled = true;
    btnExtract.textContent = 'Extraindo...';
    dataStatus.classList.add('hidden');

    try {
        const res = await fetch('/api/phase/extract', { method: 'POST' });
        const data = await res.json();
        if (res.ok) {
            dataStatus.classList.remove('hidden', 'error');
            const r = data.resumo;
            dataStatusText.textContent =
                `Dados extraídos: ${r.vol_total_nm3.toLocaleString('pt-BR')} Nm³ total, ` +
                `PCS médio ${r.pcs_media_kcal} kcal, ${r.n_clientes} clientes, ` +
                `Balanço: ${r.balanco_resultado}`;
            showToast('Dados extraídos com sucesso!', 'success');
            btnGenGraphs.disabled = false;

            // Navigate to Dados tab and load preview
            navigateTo('dados');
            dataPreviewLoaded = false;
            loadDataPreview();
        } else {
            dataStatus.classList.remove('hidden');
            dataStatus.classList.add('error');
            dataStatusText.textContent = data.detail || 'Erro na extração';
            showToast(data.detail || 'Erro na extração', 'error');
        }
    } catch (err) {
        showToast('Erro de conexão', 'error');
    }
    btnExtract.disabled = false;
    btnExtract.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 3h18v18H3z"/><path d="M3 9h18"/><path d="M9 3v18"/></svg> Extrair Dados';
}

async function loadDataPreview() {
    if (dataPreviewLoaded) return;
    const grid = document.getElementById('data-preview-grid');
    try {
        const res = await fetch('/api/phase/extract/preview');
        if (!res.ok) return; // no data yet
        const data = await res.json();
        grid.innerHTML = '';

        // Config card
        if (data.config) {
            const card = document.createElement('div');
            card.className = 'data-card';
            card.innerHTML = `
                <div class="data-card-header">
                    <h4>Configuração</h4>
                </div>
                <div class="data-card-body">
                    <div class="data-item"><span class="data-label">Período</span><span class="data-value">${data.config.periodo}</span></div>
                    <div class="data-item"><span class="data-label">Dias</span><span class="data-value">${data.config.dias}</span></div>
                </div>
            `;
            grid.appendChild(card);
        }

        // Section cards
        for (const [name, text] of Object.entries(data.sections || {})) {
            const card = document.createElement('div');
            card.className = 'data-card';
            const lines = text.split('\n').filter(l => l.trim());
            const previewLines = lines.slice(0, 12);
            const hasMore = lines.length > 12;
            card.innerHTML = `
                <div class="data-card-header">
                    <h4>${name}</h4>
                    <span class="data-card-badge">${lines.length} linhas</span>
                </div>
                <div class="data-card-body">
                    <pre class="data-preview-text">${previewLines.join('\n')}</pre>
                    ${hasMore ? '<div class="data-more">... mais dados</div>' : ''}
                </div>
            `;
            // Toggle full view on click
            card.addEventListener('click', () => {
                const pre = card.querySelector('.data-preview-text');
                const more = card.querySelector('.data-more');
                if (card.classList.contains('expanded')) {
                    pre.textContent = previewLines.join('\n');
                    if (more) more.style.display = '';
                    card.classList.remove('expanded');
                } else {
                    pre.textContent = text;
                    if (more) more.style.display = 'none';
                    card.classList.add('expanded');
                }
            });
            grid.appendChild(card);
        }

        dataPreviewLoaded = true;
        if (!grid.children.length) {
            grid.innerHTML = '<div class="gallery-loading">Nenhum dado extraído ainda.</div>';
        }
    } catch {
        // No data extracted yet — keep placeholder
    }
}

async function runPhaseGraphs() {
    btnGenGraphs.disabled = true;
    btnGenGraphs.textContent = 'Gerando gráficos...';

    try {
        const res = await fetch('/api/phase/graphs', { method: 'POST' });
        const data = await res.json();
        if (res.ok) {
            showToast(`${data.count} gráficos gerados!`, 'success');
            // Reset gallery cache so it reloads fresh
            graficosData = null;
            // Navigate to Graficos tab
            navigateTo('graficos');
        } else {
            showToast(data.detail || 'Erro ao gerar gráficos', 'error');
        }
    } catch {
        showToast('Erro de conexão', 'error');
    }
    btnGenGraphs.disabled = false;
    btnGenGraphs.innerHTML = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="m9 17 3-6 3 6"/></svg> Gerar Gráficos';
}

// =====================================================================
// GALLERY - Gráficos
// =====================================================================
let graficosData = null;

async function loadGraficos() {
    if (graficosData) return;
    const gallery = document.getElementById('graficos-gallery');
    try {
        const res = await fetch('/api/outputs/graficos');
        graficosData = await res.json();
        renderGraficos('all');
        setupGraficoTabs();
    } catch {
        gallery.innerHTML = '<div class="gallery-loading">Erro ao carregar gráficos</div>';
    }
}

function renderGraficos(filter) {
    const gallery = document.getElementById('graficos-gallery');
    gallery.innerHTML = '';
    lightboxImages = [];

    for (const [capNum, capData] of Object.entries(graficosData)) {
        if (filter !== 'all' && String(capNum) !== String(filter)) continue;
        for (const g of capData.graphs) {
            if (!g.exists) continue;
            const idx = lightboxImages.length;
            lightboxImages.push({ url: g.url, caption: g.caption });

            const card = document.createElement('div');
            card.className = 'gallery-card';
            card.innerHTML = `<img src="${g.url}" alt="${g.caption}" loading="lazy"><div class="caption">${g.caption}</div>`;
            card.addEventListener('click', () => openLightbox(idx));
            gallery.appendChild(card);
        }
    }

    if (!gallery.children.length) {
        gallery.innerHTML = '<div class="gallery-loading">Nenhum gráfico encontrado</div>';
    }
}

function setupGraficoTabs() {
    document.querySelectorAll('#graficos-tabs .tab').forEach(tab => {
        tab.addEventListener('click', () => {
            document.querySelectorAll('#graficos-tabs .tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            renderGraficos(tab.dataset.cap);
        });
    });
}

// =====================================================================
// GALLERY - Diagramas
// =====================================================================
let diagramasLoaded = false;

async function loadDiagramas() {
    if (diagramasLoaded) return;
    const gallery = document.getElementById('diagramas-gallery');
    try {
        const res = await fetch('/api/outputs/diagramas');
        const data = await res.json();
        gallery.innerHTML = '';
        data.forEach((d, idx) => {
            const card = document.createElement('div');
            card.className = 'gallery-card';
            card.innerHTML = `<img src="${d.url}" alt="${d.caption}" loading="lazy"><div class="caption">${d.caption}</div>`;
            card.addEventListener('click', () => {
                lightboxImages = data.map(x => ({ url: x.url, caption: x.caption }));
                openLightbox(idx);
            });
            gallery.appendChild(card);
        });
        diagramasLoaded = true;
        if (!data.length) gallery.innerHTML = '<div class="gallery-loading">Nenhum diagrama encontrado</div>';
    } catch {
        gallery.innerHTML = '<div class="gallery-loading">Erro ao carregar diagramas</div>';
    }
}

// =====================================================================
// LIGHTBOX
// =====================================================================
const lightbox = document.getElementById('lightbox');
const lbImg = document.getElementById('lightbox-img');
const lbCaption = document.getElementById('lightbox-caption');

function openLightbox(idx) {
    lightboxIndex = idx;
    updateLightbox();
    lightbox.hidden = false;
}

function updateLightbox() {
    const item = lightboxImages[lightboxIndex];
    if (!item) return;
    lbImg.src = item.url;
    lbCaption.textContent = item.caption;
}

document.getElementById('lightbox-close').addEventListener('click', () => lightbox.hidden = true);
document.querySelector('.lightbox-overlay').addEventListener('click', () => lightbox.hidden = true);
document.getElementById('lightbox-prev').addEventListener('click', () => {
    lightboxIndex = (lightboxIndex - 1 + lightboxImages.length) % lightboxImages.length;
    updateLightbox();
});
document.getElementById('lightbox-next').addEventListener('click', () => {
    lightboxIndex = (lightboxIndex + 1) % lightboxImages.length;
    updateLightbox();
});
document.addEventListener('keydown', (e) => {
    if (lightbox.hidden) return;
    if (e.key === 'Escape') lightbox.hidden = true;
    if (e.key === 'ArrowLeft') { lightboxIndex = (lightboxIndex - 1 + lightboxImages.length) % lightboxImages.length; updateLightbox(); }
    if (e.key === 'ArrowRight') { lightboxIndex = (lightboxIndex + 1) % lightboxImages.length; updateLightbox(); }
});

// =====================================================================
// TEXTOS (Accordion)
// =====================================================================
let textosLoaded = false;
const md = typeof markdownit !== 'undefined' ? markdownit() : null;

const SECTION_LABELS = {
    a: { metodologia: 'Metodologia', conteudo: 'Conteúdo' },
    b: { dados: 'Dados', sintese: 'Síntese' },
    c: { graficos: 'Gráficos' },
    d: { sintese: 'Síntese' },
};

async function loadTextos() {
    if (textosLoaded) return;
    const accordion = document.getElementById('textos-accordion');
    try {
        const res = await fetch('/api/outputs/cache');
        const data = await res.json();
        accordion.innerHTML = '';

        // Group by chapter
        const groups = {};
        data.forEach(item => {
            const match = item.filename.match(/^(cap\d+|conclusoes|resumo_executivo)/);
            const group = match ? match[1] : item.filename;
            if (!groups[group]) groups[group] = [];
            groups[group].push(item);
        });

        const groupLabels = {
            cap1: 'Cap. 1 — Visão Geral',
            cap2: 'Cap. 2 — Volumes de Entrada',
            cap3: 'Cap. 3 — PCS',
            cap4: 'Cap. 4 — Energia',
            cap5: 'Cap. 5 — Perfis de Consumo',
            cap6: 'Cap. 6 — Incertezas',
            cap7: 'Cap. 7 — Balanço de Massa',
            conclusoes: 'Conclusões e Recomendações',
            resumo_executivo: 'Resumo Executivo',
        };

        for (const [group, items] of Object.entries(groups)) {
            const label = groupLabels[group] || group;
            const item = document.createElement('div');
            item.className = 'accordion-item';
            item.innerHTML = `
                <div class="accordion-header">
                    <h4>${label}</h4>
                    <span class="arrow">&#9662;</span>
                </div>
                <div class="accordion-buttons">
                    ${items.map(i => `<button class="sub-btn" data-file="${i.filename}">${getSubLabel(i.filename)}</button>`).join('')}
                </div>
                <div class="accordion-content"><div class="md-content" id="md-${group}"></div></div>
            `;
            accordion.appendChild(item);

            // Toggle accordion
            item.querySelector('.accordion-header').addEventListener('click', () => {
                item.classList.toggle('open');
            });

            // Sub-buttons load content
            item.querySelectorAll('.sub-btn').forEach(btn => {
                btn.addEventListener('click', async () => {
                    item.querySelectorAll('.sub-btn').forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                    item.classList.add('open');
                    const contentDiv = item.querySelector('.md-content');
                    contentDiv.innerHTML = '<em>Carregando...</em>';
                    try {
                        const r = await fetch(`/api/outputs/cache/${btn.dataset.file}`);
                        const d = await r.json();
                        contentDiv.innerHTML = md ? md.render(d.content) : `<pre>${d.content}</pre>`;
                    } catch {
                        contentDiv.innerHTML = '<em>Erro ao carregar</em>';
                    }
                });
            });
        }
        textosLoaded = true;
    } catch {
        accordion.innerHTML = '<div class="accordion-loading">Erro ao carregar seções</div>';
    }
}

function getSubLabel(filename) {
    const parts = filename.split('_');
    const letter = parts[1] || '';
    const type = parts.slice(2).join('_') || parts[1] || filename;
    const labels = { metodologia: 'Metodologia', dados: 'Dados', graficos: 'Gráficos', sintese: 'Síntese', conteudo: 'Conteúdo' };
    return labels[type] || labels[letter] || filename;
}

// =====================================================================
// DOWNLOADS
// =====================================================================
let downloadsLoaded = false;

async function loadDownloads() {
    if (downloadsLoaded) return;
    const grid = document.getElementById('download-grid');
    try {
        const res = await fetch('/api/outputs/downloads');
        const data = await res.json();
        grid.innerHTML = '';

        if (!data.length) {
            grid.innerHTML = '<div class="gallery-loading">Nenhum documento disponível. Execute o pipeline primeiro.</div>';
            return;
        }

        data.forEach(f => {
            const iconClass = f.type === 'docx' ? 'docx' : 'pptx';
            const iconText = f.type.toUpperCase();
            const card = document.createElement('div');
            card.className = 'download-card';
            card.innerHTML = `
                <div class="download-icon ${iconClass}">${iconText}</div>
                <h4>${f.filename}</h4>
                <div class="file-info">${f.size_mb} MB</div>
                <a href="${f.url}" download class="btn-primary">
                    <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                    Baixar
                </a>
            `;
            grid.appendChild(card);
        });

        // Add PPTX generation button if no PPTX exists
        const hasPptx = data.some(f => f.type === 'pptx');
        if (!hasPptx) {
            const genCard = document.createElement('div');
            genCard.className = 'download-card';
            genCard.innerHTML = `
                <div class="download-icon pptx">PPTX</div>
                <h4>Apresentação</h4>
                <div class="file-info">Ainda não gerada</div>
                <button class="btn-secondary" onclick="generatePresentation(this)">Gerar Apresentação</button>
            `;
            grid.appendChild(genCard);
        }

        downloadsLoaded = true;
    } catch {
        grid.innerHTML = '<div class="gallery-loading">Erro ao carregar downloads</div>';
    }
}

async function generatePresentation(btn) {
    btn.disabled = true;
    btn.textContent = 'Gerando...';
    try {
        const res = await fetch('/api/presentation/generate', { method: 'POST' });
        const data = await res.json();
        if (res.ok) {
            showToast(`Apresentação gerada! (${data.size_mb} MB)`, 'success');
            downloadsLoaded = false;
            loadDownloads();
        } else {
            showToast(data.detail || 'Erro ao gerar', 'error');
        }
    } catch {
        showToast('Erro de conexão', 'error');
    }
    btn.disabled = false;
    btn.textContent = 'Gerar Apresentação';
}

// =====================================================================
// TOAST
// =====================================================================
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    toast.textContent = message;
    container.appendChild(toast);
    setTimeout(() => toast.remove(), 3500);
}
