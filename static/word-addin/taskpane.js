const fileInput = document.getElementById('file-input');
const dropzone = document.getElementById('dropzone');
const uploadMeta = document.getElementById('upload-meta');
const imagePreview = document.getElementById('image-preview');
const imagePreviewFrame = document.getElementById('image-preview-frame');
const imagePreviewEmpty = document.getElementById('image-preview-empty');
const recognizeBtn = document.getElementById('recognize-btn');
const resetBtn = document.getElementById('reset-btn');
const hostPill = document.getElementById('host-pill');
const resultPill = document.getElementById('result-pill');
const noticeBox = document.getElementById('notice-box');
const providerChip = document.getElementById('provider-chip');
const modeChip = document.getElementById('mode-chip');
const latexOutput = document.getElementById('latex-output');
const mathPreview = document.getElementById('math-preview');
const insertBtn = document.getElementById('insert-btn');
const downloadBtn = document.getElementById('download-btn');
const wordHint = document.getElementById('word-hint');
const toast = document.getElementById('toast');

const DEFAULT_NOTICE = '请先在 Word 中打开这个任务窗格，再上传公式截图。当前 MVP 不提供 LaTeX 手改功能。';
const DEFAULT_HINT = '插入按钮会使用后端生成的 Word OOXML 结果，失败时可以退回下载 .docx 的方式。';

let selectedFile = null;
let recognitionPayload = null;
let toastTimer = null;
let isWordHostReady = false;

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

function setNotice(text) {
  noticeBox.textContent = text;
}

function setResultState(label) {
  resultPill.textContent = label;
}

function updateActionState() {
  insertBtn.disabled = !(isWordHostReady && recognitionPayload?.word_ooxml);
  downloadBtn.disabled = !recognitionPayload?.latex;
}

function setHostState(label, ready = false) {
  isWordHostReady = ready;
  hostPill.textContent = label;
  hostPill.classList.toggle('is-ready', ready);
  updateActionState();
}

function resetImagePreview() {
  imagePreview.removeAttribute('src');
  imagePreviewFrame.classList.remove('has-image');
  imagePreviewEmpty.hidden = false;
}

function updateImagePreview(file) {
  const reader = new FileReader();
  reader.onload = () => {
    imagePreview.src = reader.result;
    imagePreviewFrame.classList.add('has-image');
    imagePreviewEmpty.hidden = true;
  };
  reader.readAsDataURL(file);
}

function renderLatex(latex) {
  const value = (latex || '').trim();
  if (!value) {
    mathPreview.innerHTML = '<div class="render-panel__empty">识别完成后，这里会显示渲染后的公式。</div>';
    return;
  }

  if (!window.katex) {
    mathPreview.innerHTML = '<div class="render-panel__empty">KaTeX 尚未加载，当前只展示 LaTeX 文本。</div>';
    return;
  }

  mathPreview.innerHTML = '';
  try {
    window.katex.render(value, mathPreview, {
      displayMode: true,
      throwOnError: false,
      strict: 'ignore',
    });
  } catch (error) {
    mathPreview.innerHTML = `<div class="render-panel__empty">预览渲染失败：${error.message}</div>`;
  }
}

function resetResult() {
  recognitionPayload = null;
  providerChip.textContent = '识别引擎：未开始';
  modeChip.textContent = '插入状态：等待识别';
  latexOutput.textContent = '识别完成后，这里会显示结果。';
  wordHint.textContent = DEFAULT_HINT;
  setResultState('尚未识别');
  renderLatex('');
  updateActionState();
}

function handleIncomingFile(file) {
  if (!file) return;

  const allowedTypes = ['image/png', 'image/jpeg', 'image/webp'];
  if (!allowedTypes.includes(file.type)) {
    selectedFile = null;
    resetImagePreview();
    uploadMeta.textContent = '仅支持 png / jpg / jpeg / webp 图片';
    showToast('仅支持图片文件');
    return;
  }

  if (file.size > 5 * 1024 * 1024) {
    selectedFile = null;
    resetImagePreview();
    uploadMeta.textContent = '图片超过 5MB，请压缩后重试';
    showToast('图片超过 5MB');
    return;
  }

  selectedFile = file;
  uploadMeta.textContent = `已选择：${file.name} · ${(file.size / 1024 / 1024).toFixed(2)} MB`;
  updateImagePreview(file);
  resetResult();
  setNotice('图片已准备好，点击“识别并准备插入”后即可生成可插入 Word 的公式。');
}

async function readApiJson(response, fallbackMessage) {
  const contentType = response.headers.get('Content-Type') || '';
  if (contentType.includes('application/json')) {
    return await response.json();
  }

  const text = await response.text();
  if (text.trim().startsWith('<')) {
    throw new Error(`${fallbackMessage}，后端返回了 HTML 页面，请检查 Flask 是否正常运行。`);
  }

  throw new Error(text.trim() || fallbackMessage);
}

async function recognizeFormula() {
  if (!selectedFile) {
    showToast('请先选择一张公式图片');
    return;
  }

  const formData = new FormData();
  formData.append('file', selectedFile);

  recognizeBtn.disabled = true;
  recognizeBtn.textContent = '识别中...';
  setResultState('识别中');
  setNotice('正在请求后端识别并生成 Word 插件插入数据，请稍候。');

  try {
    const response = await fetch('/api/word-addin/recognize', {
      method: 'POST',
      body: formData,
    });
    const payload = await readApiJson(response, 'Word 插件识别失败');

    if (!response.ok || !payload.ok) {
      throw new Error(payload.error || 'Word 插件识别失败');
    }

    recognitionPayload = payload;
    providerChip.textContent = `识别引擎：${payload.provider || '未知'}`;
    modeChip.textContent = `插入状态：${payload.mode || '已生成插入数据'}`;
    latexOutput.textContent = payload.latex || '';
    wordHint.textContent = payload.word_hint || DEFAULT_HINT;
    renderLatex(payload.latex || '');
    setResultState('可插入');
    setNotice(payload.notice || '识别完成，已经生成可插入 Word 的公式。');
    updateActionState();
    showToast('识别完成，可以插入到当前 Word 文档了', true);
  } catch (error) {
    recognitionPayload = null;
    setResultState('识别失败');
    setNotice(error.message || '识别失败，请稍后再试。');
    updateActionState();
    showToast(error.message || '识别失败');
  } finally {
    recognizeBtn.disabled = false;
    recognizeBtn.textContent = '识别并准备插入';
  }
}

async function insertIntoWord() {
  if (!recognitionPayload?.word_ooxml) {
    showToast('当前没有可插入的识别结果');
    return;
  }

  if (!isWordHostReady || !window.Word) {
    showToast('请在 Word 客户端任务窗格中使用插入功能');
    return;
  }

  insertBtn.disabled = true;
  insertBtn.textContent = '插入中...';

  try {
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertOoxml(recognitionPayload.word_ooxml, Word.InsertLocation.replace);
      await context.sync();
    });
    modeChip.textContent = '插入状态：已插入当前文档';
    showToast('公式已插入当前 Word 文档', true);
  } catch (error) {
    showToast(`插入失败：${error.message || error}`);
    setNotice('Word 插入失败。你可以先试试下载 Word 原生公式 .docx 作为兜底方案。');
  } finally {
    insertBtn.textContent = '插入到当前 Word 文档';
    updateActionState();
  }
}

async function downloadWordDocx() {
  if (!recognitionPayload?.latex) {
    showToast('当前没有可下载的识别结果');
    return;
  }

  downloadBtn.disabled = true;
  downloadBtn.textContent = '生成中...';

  try {
    const response = await fetch('/api/word/export', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        latex: recognitionPayload.latex,
        filename: recognitionPayload.filename || selectedFile?.name || 'formula',
        omml_xml: recognitionPayload.omml_xml,
        word_ooxml: recognitionPayload.word_ooxml,
      }),
    });

    if (!response.ok) {
      let errorMessage = 'Word 文档生成失败';
      try {
        const payload = await readApiJson(response, 'Word 文档生成失败');
        errorMessage = payload.error || errorMessage;
      } catch {
        // ignore parsing failure
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
    downloadBtn.textContent = '下载 Word 原生公式 .docx';
    updateActionState();
  }
}

function resetAll() {
  selectedFile = null;
  fileInput.value = '';
  uploadMeta.textContent = '还没有选择图片';
  setNotice(DEFAULT_NOTICE);
  resetImagePreview();
  resetResult();
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

recognizeBtn.addEventListener('click', recognizeFormula);
insertBtn.addEventListener('click', insertIntoWord);
downloadBtn.addEventListener('click', downloadWordDocx);
resetBtn.addEventListener('click', resetAll);

if (window.Office?.onReady) {
  Office.onReady((info) => {
    const isWord = info.host === Office.HostType.Word || info.host === 'Word';
    if (isWord) {
      setHostState('已连接 Word 宿主', true);
      setNotice('Word 宿主已连接。上传截图后，识别结果会直接准备成可插入当前文档的公式。');
    } else {
      setHostState('当前不在 Word 宿主内', false);
      setNotice('这个页面更适合在 Word 任务窗格中打开。现在你也可以先在浏览器里预览识别效果。');
    }
  });
} else {
  setHostState('未检测到 Office.js 宿主', false);
  setNotice('当前是浏览器预览模式。真正插入 Word 时，请通过 Word 插件任务窗格打开本页。');
}

resetAll();
