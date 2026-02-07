/**
 * FormatWord — Main Application Entry Point
 *
 * Wires up all UI events and orchestrates the pipeline:
 * 1. Upload source .docx
 * 2. Paste or upload translation .txt
 * 3. Parse → Match → Rebuild → Download
 */

import { parseDocx, getDocumentInfo } from './engine/docxParser.js';
import { parsePptx, getPptxInfo } from './engine/pptxParser.js';
import { matchTexts, getMatchSummary } from './engine/textMatcher.js';
import { rebuildDocx } from './engine/docxRebuilder.js';
import { rebuildPptx } from './engine/pptxRebuilder.js';
import { saveAs } from 'file-saver';

// ══════════════════════════════════════════════════
// State
// ══════════════════════════════════════════════════
const state = {
    sourceFile: null,
    sourceBuffer: null,
    translationText: '',
    generatedBlob: null,
    sourceFileName: '',
    fileType: '', // 'docx' or 'pptx'
};

// ══════════════════════════════════════════════════
// DOM References
// ══════════════════════════════════════════════════
const els = {
    // Source panel
    dropZoneSource: document.getElementById('dropZoneSource'),
    fileInputSource: document.getElementById('fileInputSource'),
    sourceFileInfo: document.getElementById('sourceFileInfo'),
    sourceFileName: document.getElementById('sourceFileName'),
    sourceFileSize: document.getElementById('sourceFileSize'),
    sourceRemove: document.getElementById('sourceRemove'),
    sourceDocInfo: document.getElementById('sourceDocInfo'),
    sourceDocInfoList: document.getElementById('sourceDocInfoList'),

    // Translation panel
    tabPaste: document.getElementById('tabPaste'),
    tabFile: document.getElementById('tabFile'),
    tabContentPaste: document.getElementById('tabContentPaste'),
    tabContentFile: document.getElementById('tabContentFile'),
    translationTextarea: document.getElementById('translationTextarea'),
    charCount: document.getElementById('charCount'),
    paraCount: document.getElementById('paraCount'),
    dropZoneTranslation: document.getElementById('dropZoneTranslation'),
    fileInputTranslation: document.getElementById('fileInputTranslation'),
    transFileInfo: document.getElementById('transFileInfo'),
    transFileName: document.getElementById('transFileName'),
    transFileSize: document.getElementById('transFileSize'),
    transRemove: document.getElementById('transRemove'),

    // Output panel
    outputPlaceholder: document.getElementById('outputPlaceholder'),
    progressContainer: document.getElementById('progressContainer'),
    progressSteps: document.getElementById('progressSteps'),
    progressBarFill: document.getElementById('progressBarFill'),
    progressMessage: document.getElementById('progressMessage'),
    outputResult: document.getElementById('outputResult'),
    resultStats: document.getElementById('resultStats'),
    outputErrors: document.getElementById('outputErrors'),
    errorBox: document.getElementById('errorBox'),

    // Actions
    btnGenerate: document.getElementById('btnGenerate'),
    btnDownload: document.getElementById('btnDownload'),

    // Modal
    matchReviewModal: document.getElementById('matchReviewModal'),
    matchReviewBody: document.getElementById('matchReviewBody'),
    modalClose: document.getElementById('modalClose'),
    modalSkip: document.getElementById('modalSkip'),
    modalConfirm: document.getElementById('modalConfirm'),
};

// ══════════════════════════════════════════════════
// Utility
// ══════════════════════════════════════════════════
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' o';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' Ko';
    return (bytes / (1024 * 1024)).toFixed(1) + ' Mo';
}

function updateGenerateButton() {
    const hasSource = state.sourceBuffer !== null;
    const hasTranslation = state.translationText.trim().length > 0;
    els.btnGenerate.disabled = !(hasSource && hasTranslation);
}

function showPanel(panel) {
    els.outputPlaceholder.style.display = 'none';
    els.progressContainer.style.display = 'none';
    els.outputResult.style.display = 'none';
    els.outputErrors.style.display = 'none';

    if (panel === 'placeholder') els.outputPlaceholder.style.display = '';
    if (panel === 'progress') els.progressContainer.style.display = '';
    if (panel === 'result') els.outputResult.style.display = '';
    if (panel === 'error') els.outputErrors.style.display = '';
}

// ══════════════════════════════════════════════════
// Source File Handling
// ══════════════════════════════════════════════════
function setupDropZone(dropZone, fileInput, onFileAccepted) {
    // Click to browse
    dropZone.addEventListener('click', (e) => {
        if (e.target.closest('.file-remove')) return;
        fileInput.click();
    });

    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            onFileAccepted(fileInput.files[0]);
        }
    });

    // Drag & Drop
    dropZone.addEventListener('dragenter', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('drag-over');
    });

    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault();
        if (!dropZone.contains(e.relatedTarget)) {
            dropZone.classList.remove('drag-over');
        }
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            onFileAccepted(e.dataTransfer.files[0]);
        }
    });
}

function handleSourceFile(file) {
    const ext = file.name.split('.').pop().toLowerCase();
    if (ext !== 'docx' && ext !== 'pptx') {
        showError('❌ Veuillez sélectionner un fichier .docx ou .pptx');
        return;
    }

    state.sourceFile = file;
    state.sourceFileName = file.name.replace(/\.(docx|pptx)$/i, '');
    state.fileType = ext;

    // Read as ArrayBuffer
    const reader = new FileReader();
    reader.onload = async (e) => {
        state.sourceBuffer = e.target.result;

        // Show file info
        els.dropZoneSource.querySelector('.drop-zone-content').style.display = 'none';
        els.sourceFileInfo.style.display = 'flex';
        els.sourceFileName.textContent = file.name;
        els.sourceFileSize.textContent = formatFileSize(file.size);
        els.dropZoneSource.classList.add('has-file');

        // Update file icon based on type
        const fileIcon = els.sourceFileInfo.querySelector('.file-info-icon');
        fileIcon.textContent = ext === 'pptx' ? '📊' : '📄';

        // Quick parse to show document info
        try {
            if (ext === 'pptx') {
                const { xmlFiles, segments, slideCount } = await parsePptx(state.sourceBuffer);
                const info = getPptxInfo(xmlFiles, slideCount);

                els.sourceDocInfoList.innerHTML = '';
                const items = [
                    `📊 Présentation PowerPoint`,
                    `${info.slideCount} slides`,
                    `${segments.length} segments de texte`,
                    `${info.shapeCount} zones de texte`,
                ].filter(Boolean);

                for (const item of items) {
                    const li = document.createElement('li');
                    li.textContent = item;
                    els.sourceDocInfoList.appendChild(li);
                }
            } else {
                const { xmlFiles, segments } = await parseDocx(state.sourceBuffer);
                const info = getDocumentInfo(xmlFiles);

                els.sourceDocInfoList.innerHTML = '';
                const items = [
                    `📄 Document Word`,
                    `${segments.length} segments de texte`,
                    `${info.paragraphCount} paragraphes dans le corps`,
                    info.hasTables ? '📊 Contient des tableaux' : null,
                    info.hasImages ? '🖼️ Contient des images' : null,
                    info.hasHeaders ? '📑 En-têtes détectés' : null,
                    info.hasFooters ? '📑 Pieds de page détectés' : null,
                    info.hasFootnotes ? '📝 Notes de bas de page' : null,
                ].filter(Boolean);

                for (const item of items) {
                    const li = document.createElement('li');
                    li.textContent = item;
                    els.sourceDocInfoList.appendChild(li);
                }
            }
            els.sourceDocInfo.style.display = '';
        } catch (err) {
            console.error('Quick parse error:', err);
        }

        updateGenerateButton();
    };
    reader.readAsArrayBuffer(file);
}

function removeSourceFile() {
    state.sourceFile = null;
    state.sourceBuffer = null;
    state.sourceFileName = '';
    state.fileType = '';

    els.dropZoneSource.querySelector('.drop-zone-content').style.display = '';
    els.sourceFileInfo.style.display = 'none';
    els.dropZoneSource.classList.remove('has-file');
    els.sourceDocInfo.style.display = 'none';
    els.fileInputSource.value = '';

    // Reset file icon
    const fileIcon = els.sourceFileInfo.querySelector('.file-info-icon');
    fileIcon.textContent = '📄';

    updateGenerateButton();
}

// ══════════════════════════════════════════════════
// Translation Text Handling
// ══════════════════════════════════════════════════
function handleTranslationFile(file) {
    if (!file.name.endsWith('.txt')) {
        showError('❌ Veuillez sélectionner un fichier .txt');
        return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
        state.translationText = e.target.result;
        els.translationTextarea.value = state.translationText;
        updateTextCounts();

        // Show file info in the file tab
        els.dropZoneTranslation.querySelector('.drop-zone-content').style.display = 'none';
        els.transFileInfo.style.display = 'flex';
        els.transFileName.textContent = file.name;
        els.transFileSize.textContent = formatFileSize(file.size);
        els.dropZoneTranslation.classList.add('has-file');

        // Switch to paste tab to show the loaded text
        switchTab('paste');
        updateGenerateButton();
    };
    reader.readAsText(file, 'utf-8');
}

function removeTranslationFile() {
    state.translationText = '';
    els.translationTextarea.value = '';

    els.dropZoneTranslation.querySelector('.drop-zone-content').style.display = '';
    els.transFileInfo.style.display = 'none';
    els.dropZoneTranslation.classList.remove('has-file');
    els.fileInputTranslation.value = '';

    updateTextCounts();
    updateGenerateButton();
}

function updateTextCounts() {
    const text = els.translationTextarea.value;
    state.translationText = text;

    els.charCount.textContent = `${text.length} caractères`;
    const paras = text
        .split(/\r?\n/)
        .filter((l) => l.trim().length > 0);
    els.paraCount.textContent = `${paras.length} paragraphes`;

    updateGenerateButton();
}

function switchTab(tab) {
    els.tabPaste.classList.toggle('active', tab === 'paste');
    els.tabFile.classList.toggle('active', tab === 'file');
    els.tabContentPaste.style.display = tab === 'paste' ? '' : 'none';
    els.tabContentFile.style.display = tab === 'file' ? '' : 'none';
}

// ══════════════════════════════════════════════════
// Progress & Error Display
// ══════════════════════════════════════════════════
const STEPS = [
    { key: 'analyze', icon: '⏳', label: 'Analyse du document source...' },
    { key: 'extract', icon: '🔍', label: 'Extraction de la structure...' },
    { key: 'match', icon: '🔗', label: 'Mapping texte source ↔ traduction...' },
    { key: 'rebuild', icon: '📝', label: 'Reconstruction du document...' },
    { key: 'finalize', icon: '✅', label: 'Validation et génération...' },
];

function initProgressSteps() {
    els.progressSteps.innerHTML = '';
    for (const step of STEPS) {
        const div = document.createElement('div');
        div.className = 'progress-step';
        div.id = `step-${step.key}`;
        div.innerHTML = `
      <span class="step-icon">${step.icon}</span>
      <span>${step.label}</span>
    `;
        els.progressSteps.appendChild(div);
    }
}

function setProgressStep(stepKey, percent, message) {
    // Update step states
    let found = false;
    for (const step of STEPS) {
        const el = document.getElementById(`step-${step.key}`);
        if (!el) continue;

        if (step.key === stepKey) {
            el.classList.add('active');
            el.classList.remove('done');
            found = true;
        } else if (!found) {
            el.classList.remove('active');
            el.classList.add('done');
        } else {
            el.classList.remove('active', 'done');
        }
    }

    els.progressBarFill.style.width = `${percent}%`;
    els.progressMessage.textContent = message || '';
}

function showError(message, type = 'error') {
    showPanel('error');
    els.errorBox.textContent = message;
    els.errorBox.className = `error-box ${type}`;
}

// ══════════════════════════════════════════════════
// Main Pipeline
// ══════════════════════════════════════════════════
async function runPipeline() {
    showPanel('progress');
    initProgressSteps();

    const btnText = els.btnGenerate.querySelector('.btn-text');
    const btnLoader = els.btnGenerate.querySelector('.btn-loader');
    const btnIcon = els.btnGenerate.querySelector('.btn-icon');
    btnText.textContent = 'Traitement en cours...';
    btnLoader.style.display = '';
    btnIcon.style.display = 'none';
    els.btnGenerate.disabled = true;

    const isPptx = state.fileType === 'pptx';
    const typeLabel = isPptx ? 'PPTX' : 'DOCX';

    try {
        // ── Step 1: Analyze ──
        setProgressStep('analyze', 10, `Décompression de l'archive ${typeLabel}...`);
        await sleep(200);

        let zip, xmlFiles, segments, slideCount;

        if (isPptx) {
            const result = await parsePptx(state.sourceBuffer, (p) => {
                setProgressStep('analyze', 15, p.message);
            });
            zip = result.zip;
            xmlFiles = result.xmlFiles;
            segments = result.segments;
            slideCount = result.slideCount;
        } else {
            const result = await parseDocx(state.sourceBuffer, (p) => {
                setProgressStep('analyze', 15, p.message);
            });
            zip = result.zip;
            xmlFiles = result.xmlFiles;
            segments = result.segments;
        }

        // ── Step 2: Extract ──
        if (isPptx) {
            setProgressStep('extract', 30, `${segments.length} segments extraits de ${slideCount} slides`);
        } else {
            setProgressStep('extract', 30, `${segments.length} segments de texte extraits`);
            const info = getDocumentInfo(xmlFiles);
            setProgressStep('extract', 40, `Structure analysée : ${info.paragraphCount} paragraphes`);
        }
        await sleep(300);

        // ── Step 3: Match ──
        setProgressStep('match', 50, 'Alignement du texte source et traduit...');
        await sleep(200);

        const { mapping, unmatched, stats } = matchTexts(
            segments,
            state.translationText,
            (p) => {
                setProgressStep('match', 60, p.message);
            }
        );

        setProgressStep('match', 65, getMatchSummary(stats));
        await sleep(300);

        // Handle unmatched segments
        if (unmatched.length > 0) {
            const shouldContinue = await showMatchReview(unmatched, mapping);
            if (!shouldContinue) {
                resetButton();
                showPanel('placeholder');
                return;
            }
        }

        // ── Step 4: Rebuild ──
        setProgressStep('rebuild', 70, `Remplacement du texte dans le ${typeLabel}...`);
        await sleep(200);

        let blob;
        if (isPptx) {
            blob = await rebuildPptx(zip, xmlFiles, segments, mapping, (p) => {
                const percent = p.step === 'rebuild-zip' ? 85 : 80;
                setProgressStep('rebuild', percent, p.message);
            });
        } else {
            blob = await rebuildDocx(zip, xmlFiles, segments, mapping, (p) => {
                const percent = p.step === 'rebuild-zip' ? 85 : 80;
                setProgressStep('rebuild', percent, p.message);
            });
        }

        state.generatedBlob = blob;

        // ── Step 5: Finalize ──
        setProgressStep('finalize', 100, 'Document généré avec succès !');
        await sleep(500);

        // Show result
        showPanel('result');
        els.resultStats.textContent = getMatchSummary(stats);

    } catch (err) {
        console.error('Pipeline error:', err);
        showError(`❌ Erreur : ${err.message}`);
    } finally {
        resetButton();
    }
}

function resetButton() {
    const btnText = els.btnGenerate.querySelector('.btn-text');
    const btnLoader = els.btnGenerate.querySelector('.btn-loader');
    const btnIcon = els.btnGenerate.querySelector('.btn-icon');
    btnText.textContent = 'Analyser et Générer';
    btnLoader.style.display = 'none';
    btnIcon.style.display = '';
    updateGenerateButton();
}

function sleep(ms) {
    return new Promise((r) => setTimeout(r, ms));
}

// ══════════════════════════════════════════════════
// Match Review Modal
// ══════════════════════════════════════════════════
function showMatchReview(unmatched, mapping) {
    return new Promise((resolve) => {
        els.matchReviewBody.innerHTML = '';

        for (const seg of unmatched) {
            const item = document.createElement('div');
            item.className = 'match-item';
            item.innerHTML = `
        <div class="match-item-header">Segment #${seg.id + 1}</div>
        <div class="match-source">${escapeHtml(seg.text)}</div>
        <textarea class="match-translation-input" data-seg-id="${seg.id}"
          placeholder="Entrez la traduction correspondante..."></textarea>
      `;
            els.matchReviewBody.appendChild(item);
        }

        els.matchReviewModal.style.display = '';

        const cleanup = () => {
            els.matchReviewModal.style.display = 'none';
            els.modalClose.removeEventListener('click', handleClose);
            els.modalSkip.removeEventListener('click', handleSkip);
            els.modalConfirm.removeEventListener('click', handleConfirm);
        };

        const handleClose = () => {
            cleanup();
            resolve(false);
        };

        const handleSkip = () => {
            // Skip unmatched — leave them as original text
            cleanup();
            resolve(true);
        };

        const handleConfirm = () => {
            // Read user-provided translations
            const inputs = els.matchReviewBody.querySelectorAll('.match-translation-input');
            for (const input of inputs) {
                const segId = parseInt(input.dataset.segId, 10);
                const text = input.value.trim();
                if (text) {
                    mapping.set(segId, text);
                }
            }
            cleanup();
            resolve(true);
        };

        els.modalClose.addEventListener('click', handleClose);
        els.modalSkip.addEventListener('click', handleSkip);
        els.modalConfirm.addEventListener('click', handleConfirm);
    });
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ══════════════════════════════════════════════════
// Download
// ══════════════════════════════════════════════════
function handleDownload() {
    if (!state.generatedBlob) return;
    const ext = state.fileType || 'docx';
    const filename = `${state.sourceFileName}_traduit.${ext}`;
    saveAs(state.generatedBlob, filename);
}

// ══════════════════════════════════════════════════
// Event Binding
// ══════════════════════════════════════════════════
function init() {
    // Source drop zone
    setupDropZone(els.dropZoneSource, els.fileInputSource, handleSourceFile);
    els.sourceRemove.addEventListener('click', (e) => {
        e.stopPropagation();
        removeSourceFile();
    });

    // Translation tabs
    els.tabPaste.addEventListener('click', () => switchTab('paste'));
    els.tabFile.addEventListener('click', () => switchTab('file'));

    // Translation textarea
    els.translationTextarea.addEventListener('input', updateTextCounts);

    // Translation drop zone
    setupDropZone(els.dropZoneTranslation, els.fileInputTranslation, handleTranslationFile);
    els.transRemove.addEventListener('click', (e) => {
        e.stopPropagation();
        removeTranslationFile();
    });

    // Generate button
    els.btnGenerate.addEventListener('click', runPipeline);

    // Download button
    els.btnDownload.addEventListener('click', handleDownload);

    // Initialize UI
    showPanel('placeholder');
    updateGenerateButton();
}

init();
