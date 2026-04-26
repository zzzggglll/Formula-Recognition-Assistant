const fileInput = document.getElementById('file-input');
const dropzone = document.getElementById('dropzone');
const statusPill = document.getElementById('status-pill');
const modePill = document.getElementById('mode-pill');
const noticeBox = document.getElementById('notice-box');
const uploadMeta = document.getElementById('upload-meta');
const imagePreview = document.getElementById('image-preview');
const previewFrame = document.getElementById('preview-frame');
const recognizeBtn = document.getElementById('recognize-btn');
const clearBtn = document.getElementById('clear-btn');
const latexOutput = document.getElementById('latex-output');
const saveCorrectionBtn = document.getElementById('save-correction-btn');
const restoreSavedBtn = document.getElementById('restore-saved-btn');
const restoreOriginalBtn = document.getElementById('restore-original-btn');
const editBadge = document.getElementById('edit-badge');
const editMetaText = document.getElementById('edit-meta-text');
const originalLatexOutput = document.getElementById('original-latex-output');
const copyLatexBtn = document.getElementById('copy-latex-btn');
const copyDisplayBtn = document.getElementById('copy-display-btn');
const downloadWordBtn = document.getElementById('download-word-btn');
const mathPreview = document.getElementById('math-preview');
const toast = document.getElementById('toast');

let defaultNotice = '';
const DEFAULT_EDIT_META = '识别完成后，你可以在这里手动修正公式。复制、Word 文档导出和后续操作都会使用当前修正版。';
const EMPTY_ORIGINAL_TEXT = '识别完成后，这里会显示原始识别结果。';

const runtimeConfig = {
  demoDeployment: false,
  recognizerMode: 'unknown',
  wordExportEnabled: true,
};

let selectedFile = null;
let toastTimer = null;
let cachedWordPayload = null;
let originalRecognizedLatex = '';
let savedCorrectedLatex = '';

function showToast(message, ok = false) {
  toast.textContent = message;
  toast.hidden = false;
  toast.classList.toggle('is-success', ok);
  if (toastTimer) {
    window.clearTimeout(toastTimer);
  }
  toastTimer = window.setTimeout(() => {
    toast.hidden = true;
  }, 2800);
}

function setStatus(label, state = 'ready') {
  statusPill.textContent = label;
  statusPill.classList.remove('is-ready', 'is-busy', 'is-done');
  if (state === 'ready') statusPill.classList.add('is-ready');
  if (state === 'busy') statusPill.classList.add('is-busy');
  if (state === 'done') statusPill.classList.add('is-done');
}

function setMode(label) {
  modePill.textContent = label;
}

function setNotice(text) {
  const value = (text || '').trim();
  noticeBox.textContent = value;
  noticeBox.hidden = !value;
}

function setLoading(loading) {
  recognizeBtn.disabled = loading;
  recognizeBtn.textContent = loading ? '识别中...' : '开始识别';
}

function setWordLoading(loading) {
  downloadWordBtn.disabled = loading || !runtimeConfig.wordExportEnabled;
  downloadWordBtn.textContent = loading ? '生成中...' : '下载 Word 原生公式 .docx';
}

function resetPreview() {
  imagePreview.removeAttribute('src');
  previewFrame.classList.remove('has-image');
}

function invalidateWordCache() {
  cachedWordPayload = null;
}

function normalizeLatexValue(value) {
  return (value || '').trim();
}

function setEditState(label, state = 'idle') {
  editBadge.textContent = label;
  editBadge.classList.remove('is-idle', 'is-dirty', 'is-saved', 'is-original');
  if (state === 'dirty') editBadge.classList.add('is-dirty');
  else if (state === 'saved') editBadge.classList.add('is-saved');
  else if (state === 'original') editBadge.classList.add('is-original');
  else editBadge.classList.add('is-idle');
}

function setEditMeta(text) {
  editMetaText.textContent = text;
}

function setOriginalLatex(text) {
  originalLatexOutput.textContent = normalizeLatexValue(text) ? text : EMPTY_ORIGINAL_TEXT;
}

function syncWordExportAvailability() {
  downloadWordBtn.disabled = !runtimeConfig.wordExportEnabled;
  downloadWordBtn.title = runtimeConfig.wordExportEnabled
    ? ''
    : '当前部署仅做界面演示，Word 导出已关闭。';
}

function updateCorrectionState() {
  const current = normalizeLatexValue(latexOutput.value);
  const original = normalizeLatexValue(originalRecognizedLatex);
  const saved = normalizeLatexValue(savedCorrectedLatex);

  saveCorrectionBtn.disabled = !current;
  restoreOriginalBtn.disabled = !original;
  restoreSavedBtn.disabled = !saved || saved === current;

  if (!original && !current) {
    setEditState('等待识别', 'idle');
    setEditMeta(DEFAULT_EDIT_META);
    return;
  }

  if (!current) {
    setEditState('内容为空', 'dirty');
    setEditMeta('当前编辑区已经被清空。你可以恢复原始识别结果，或重新输入一版修正版。');
    return;
  }

  if (!original) {
    if (saved && current === saved) {
      setEditState('已保存修正版', 'saved');
      setEditMeta('当前内容来自你手动输入并已保存。本页中的复制、Word 导出和后续操作都会使用这版修正版。');
      return;
    }

    setEditState('手动输入中', 'dirty');
    setEditMeta('当前内容来自手动输入，尚未保存为修正版。');
    return;
  }

  if (current === original) {
    if (saved && saved !== original) {
      setEditState('当前为原始识别', 'original');
      setEditMeta('当前显示的是模型原始识别结果。你之前保存过一版修正版，可点击“恢复已保存修正版”继续使用。');
      return;
    }

    setEditState('未修改', 'idle');
    setEditMeta('这是模型原始识别结果。你可以直接在这里修正，复制、Word 导出和后续操作都会使用当前内容。');
    return;
  }

  if (saved && current === saved) {
    setEditState('已保存修正版', 'saved');
    setEditMeta('当前正在使用你保存的修正版。复制、Word 导出和后续操作都会使用这版内容。');
    return;
  }

  setEditState('已修改，未保存', 'dirty');
  setEditMeta('你已经手动修改了公式，但还没有保存为修正版。');
}

function setEditorLatex(value) {
  latexOutput.value = value || '';
  invalidateWordCache();
  renderLatex();
  updateCorrectionState();
}

function loadRecognizedResult(latex) {
  originalRecognizedLatex = latex || '';
  savedCorrectedLatex = latex || '';
  setOriginalLatex(originalRecognizedLatex);
  setEditorLatex(originalRecognizedLatex);
}

function resetCorrectionState() {
  originalRecognizedLatex = '';
  savedCorrectedLatex = '';
  setOriginalLatex('');
  latexOutput.value = '';
  updateCorrectionState();
}

function updateImagePreview(file) {
  const reader = new FileReader();
  reader.onload = () => {
    imagePreview.src = reader.result;
    previewFrame.classList.add('has-image');
  };
  reader.readAsDataURL(file);
}

function renderLatex() {
  const latex = normalizeLatexValue(latexOutput.value);
  if (!latex) {
    mathPreview.innerHTML = '<div class="render-panel__empty">识别完成后，这里会显示渲染后的公式。</div>';
    return;
  }

  if (window.katex) {
    mathPreview.innerHTML = '';
    try {
      window.katex.render(latex, mathPreview, {
        displayMode: true,
        throwOnError: false,
        strict: 'ignore',
      });
    } catch (error) {
      mathPreview.innerHTML = `<div class="render-panel__fallback">渲染失败，请检查 LaTeX 语法。<br><br>${error.message}</div>`;
    }
    return;
  }

  mathPreview.innerHTML = '<div class="render-panel__fallback">KaTeX 资源未加载，当前先展示源码。你仍然可以复制 LaTeX 使用。</div>';
}

function handleIncomingFile(file) {
  if (!file) return;

  const allowedTypes = ['image/png', 'image/jpeg', 'image/webp'];
  if (!allowedTypes.includes(file.type)) {
    selectedFile = null;
    invalidateWordCache();
    resetPreview();
    uploadMeta.textContent = '仅支持 png / jpg / jpeg / webp 图片';
    setStatus('文件格式不支持');
    showToast('仅支持图片文件');
    return;
  }

  if (file.size > 5 * 1024 * 1024) {
    selectedFile = null;
    invalidateWordCache();
    resetPreview();
    uploadMeta.textContent = '图片超过 5MB，请压缩后重试';
    setStatus('图片过大');
    showToast('图片超过 5MB');
    return;
  }

  selectedFile = file;
  uploadMeta.textContent = `已选择：${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  updateImagePreview(file);
  invalidateWordCache();
  setStatus('图片已准备好', 'ready');
}

async function copyText(text, successMessage) {
  if (!normalizeLatexValue(text)) {
    showToast('当前没有可复制的内容');
    return;
  }

  try {
    await navigator.clipboard.writeText(text);
    showToast(successMessage, true);
  } catch {
    showToast('复制失败，请手动选择文本');
  }
}

async function readApiJson(response, fallbackMessage) {
  const contentType = response.headers.get('Content-Type') || '';
  if (contentType.includes('application/json')) {
    return await response.json();
  }

  const text = await response.text();
  if (text.trim().startsWith('<')) {
    throw new Error(`${fallbackMessage}，后端返回了 HTML 错误页，请查看 Flask 控制台日志。`);
  }

  throw new Error(text.trim() || fallbackMessage);
}

async function fetchWordPayload(force = false) {
  if (!runtimeConfig.wordExportEnabled) {
    throw new Error('当前部署仅做界面演示，Word 导出已关闭。');
  }

  const latex = normalizeLatexValue(latexOutput.value);
  if (!latex) {
    throw new Error('当前没有可转换的 LaTeX 公式');
  }

  if (!force && cachedWordPayload?.sourceLatex === latex) {
    return cachedWordPayload;
  }

  const response = await fetch('/api/word/prepare', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ latex }),
  });
  const payload = await readApiJson(response, 'Word 公式转换失败');

  if (!response.ok || !payload.ok) {
    throw new Error(payload.error || 'Word 公式转换失败');
  }

  cachedWordPayload = {
    ...payload,
    sourceLatex: latex,
  };
  return cachedWordPayload;
}

async function downloadWordDocx() {
  if (!runtimeConfig.wordExportEnabled) {
    showToast('当前部署仅做界面演示，Word 导出已关闭。');
    return;
  }

  const latex = normalizeLatexValue(latexOutput.value);
  if (!latex) {
    showToast('当前没有可导出的 LaTeX 公式');
    return;
  }

  try {
    setWordLoading(true);
    const preparedPayload = await fetchWordPayload();
    const response = await fetch('/api/word/export', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        latex,
        filename: selectedFile?.name || 'formula',
        omml_xml: preparedPayload.omml_xml,
        word_ooxml: preparedPayload.word_ooxml,
      }),
    });

    if (!response.ok) {
      let errorMessage = 'Word 文档生成失败';
      try {
        const payload = await readApiJson(response, '识别服务异常');
        errorMessage = payload.error || errorMessage;
      } catch {
        // ignore json parsing failure
      }
      throw new Error(errorMessage);
    }

    const blob = await response.blob();
    const disposition = response.headers.get('Content-Disposition') || '';
    const fileNameMatch = disposition.match(/filename\*=(?:UTF-8'')?([^;]+)|filename="?([^";]+)"?/i);
    const downloadName = decodeURIComponent(fileNameMatch?.[1] || fileNameMatch?.[2] || 'formula-word-equation.docx');
    const objectUrl = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = downloadName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    window.URL.revokeObjectURL(objectUrl);

    showToast('Word 原生公式文档已生成', true);
  } catch (error) {
    showToast(error.message || 'Word 文档生成失败');
  } finally {
    setWordLoading(false);
  }
}

async function loadRuntimeConfig() {
  try {
    const response = await fetch('/api/runtime-info');
    const payload = await readApiJson(response, '读取运行时配置失败');

    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || '读取运行时配置失败');
    }

    runtimeConfig.demoDeployment = Boolean(payload.demo_deployment);
    runtimeConfig.recognizerMode = payload.recognizer_mode || 'unknown';
    runtimeConfig.wordExportEnabled = Boolean(payload.word_export_enabled);

    if (runtimeConfig.demoDeployment) {
      defaultNotice = '当前是 Vercel Demo 版：可演示上传、结果展示、LaTeX 编辑、复制和预览；Word 导出默认关闭。';
      setNotice(defaultNotice);
    }
  } catch {
    // ignore runtime-info failure and keep safe defaults
  } finally {
    syncWordExportAvailability();
  }
}

fileInput.addEventListener('change', (event) => {
  handleIncomingFile(event.target.files?.[0]);
});

dropzone.addEventListener('dragover', (event) => {
  event.preventDefault();
  dropzone.classList.add('is-dragover');
});

dropzone.addEventListener('dragleave', () => {
  dropzone.classList.remove('is-dragover');
});

dropzone.addEventListener('drop', (event) => {
  event.preventDefault();
  dropzone.classList.remove('is-dragover');
  handleIncomingFile(event.dataTransfer?.files?.[0]);
});

document.addEventListener('paste', (event) => {
  const items = Array.from(event.clipboardData?.items || []);
  const imageItem = items.find((item) => item.type.startsWith('image/'));
  if (!imageItem) return;
  const file = imageItem.getAsFile();
  if (!file) return;
  showToast('已接收剪贴板截图', true);
  handleIncomingFile(file);
});

recognizeBtn.addEventListener('click', async () => {
  if (!selectedFile) {
    showToast('请先上传一张公式图片');
    return;
  }

  const formData = new FormData();
  formData.append('file', selectedFile);

  setLoading(true);
  setStatus('识别进行中', 'busy');
  setMode('请求中');
  setNotice('正在上传并请求识别服务，请稍等几秒。');

  try {
    const response = await fetch('/api/recognize', {
      method: 'POST',
      body: formData,
    });
    const payload = await readApiJson(response, '识别服务异常');

    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || '识别失败');
    }

    loadRecognizedResult(payload.latex || '');
    setMode(`${payload.provider} · ${payload.mode}`);
    setNotice(payload.notice || '识别完成');
    setStatus('识别完成', 'done');
    showToast('识别完成，已进入可修正状态', true);
  } catch (error) {
    setStatus('请重试');
    setMode('识别失败');
    setNotice(error.message || '识别失败，请稍后再试。');
    showToast(error.message || '识别失败');
  } finally {
    setLoading(false);
  }
});

clearBtn.addEventListener('click', () => {
  selectedFile = null;
  fileInput.value = '';
  uploadMeta.textContent = '还没有选择图片';
  setStatus('准备就绪', 'ready');
  setMode('等待识别');
  setNotice(defaultNotice);
  invalidateWordCache();
  resetCorrectionState();
  resetPreview();
  renderLatex();
  syncWordExportAvailability();
});

saveCorrectionBtn.addEventListener('click', () => {
  const current = normalizeLatexValue(latexOutput.value);
  if (!current) {
    showToast('当前没有可保存的修正版');
    return;
  }

  savedCorrectedLatex = latexOutput.value;
  updateCorrectionState();
  setNotice('已保存修正版。后续复制与可用的导出操作都会使用当前版本。');
  showToast('修正版已保存', true);
});

restoreSavedBtn.addEventListener('click', () => {
  const saved = normalizeLatexValue(savedCorrectedLatex);
  if (!saved) {
    showToast('当前还没有保存过修正版');
    return;
  }

  setEditorLatex(savedCorrectedLatex);
  setNotice('已恢复你保存的修正版。');
  showToast('已恢复保存的修正版', true);
});

restoreOriginalBtn.addEventListener('click', () => {
  const original = normalizeLatexValue(originalRecognizedLatex);
  if (!original) {
    showToast('当前还没有原始识别结果');
    return;
  }

  setEditorLatex(originalRecognizedLatex);
  setNotice('已恢复模型原始识别结果，你可以继续修正。');
  showToast('已恢复原始识别结果', true);
});

copyLatexBtn.addEventListener('click', () => {
  copyText(latexOutput.value, 'LaTeX 源码已复制');
});

copyDisplayBtn.addEventListener('click', () => {
  const latex = normalizeLatexValue(latexOutput.value);
  copyText(`$$\n${latex}\n$$`, '带 $$ 的公式已复制');
});

downloadWordBtn.addEventListener('click', downloadWordDocx);

latexOutput.addEventListener('input', () => {
  invalidateWordCache();
  renderLatex();
  updateCorrectionState();
});

setStatus('准备就绪', 'ready');
updateCorrectionState();
renderLatex();
syncWordExportAvailability();
loadRuntimeConfig();
