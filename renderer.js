const editor = document.getElementById('editor');
const preview = document.getElementById('preview');
const exportBtn = document.getElementById('exportBtn');
const tableStyleSelect = document.getElementById('tableStyleSelect');
let md;

function normalizeAiMathInput(markdown) {
  if (!markdown) {
    return '';
  }

  let text = markdown
    .replace(/\r\n/g, '\n')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ');

  text = text.replace(/```(?:latex|tex|math)\s*([\s\S]*?)```/gi, (_m, formulaBody) => {
    const content = String(formulaBody || '').trim();
    if (!content) {
      return '';
    }
    return `\n$$\n${content}\n$$\n`;
  });

  text = text.replace(/\\+\[([\s\S]*?)\\+\]/g, (_m, inner) => `\n$$\n${String(inner).trim()}\n$$\n`);
  text = text.replace(/\\+\(([^\n]*?)\\+\)/g, (_m, inner) => `$${String(inner).trim()}$`);

  text = text.replace(/(^|\n)\[\s*\n([\s\S]*?)\n\](?=\n|$)/g, (m, prefix, body) => {
    const content = String(body || '').trim();
    const looksLikeMath = /\\[a-zA-Z]+|\^|_|=|Γ|α|β/.test(content);
    if (!content || !looksLikeMath) {
      return m;
    }
    return `${prefix}$$\n${content}\n$$`;
  });

  return text;
}

function loadSettings() {
  const savedTableStyle = localStorage.getItem('tableStyle') || 'normal';
  tableStyleSelect.value = savedTableStyle;
  applyPreviewTableStyle();
}

function bindSettings() {
  tableStyleSelect.addEventListener('change', () => {
    localStorage.setItem('tableStyle', tableStyleSelect.value);
    applyPreviewTableStyle();
    render();
  });
}

function applyPreviewTableStyle() {
  preview.classList.remove('preview-table-normal', 'preview-table-minimal');
  const styleClass = tableStyleSelect.value === 'minimal' ? 'preview-table-minimal' : 'preview-table-normal';
  preview.classList.add(styleClass);
}

async function initMarkdownEngine() {
  if (window.mdDepsReady) {
    await window.mdDepsReady;
  }

  if (!window.markdownit || !window.katex || !window.texmath) {
    throw new Error('Markdown/KaTeX 依赖加载失败，请检查本地依赖是否存在。');
  }

  md = window.markdownit({
    html: true,
    breaks: true,
    linkify: true
  });

  md.use(window.texmath, {
    engine: window.katex,
    delimiters: 'dollars',
    katexOptions: {
      throwOnError: false
    }
  });
}

function render() {
  if (!md) {
    return;
  }
  preview.innerHTML = md.render(normalizeAiMathInput(editor.value));
}

async function exportDocx() {
  const content = editor.value;
  exportBtn.disabled = true;
  exportBtn.textContent = '导出中...';

  try {
    const result = await window.api.exportDocx(content, {
      tableStyle: tableStyleSelect.value
    });
    if (result.canceled) {
      return;
    }
    alert(`导出成功: ${result.outputPath}`);
  } catch (err) {
    console.error(err);
    alert(`导出失败: ${err.message || '未知错误'}`);
  } finally {
    exportBtn.disabled = false;
    exportBtn.textContent = '导出 Word';
  }
}

editor.addEventListener('input', render);
exportBtn.addEventListener('click', exportDocx);
loadSettings();
bindSettings();

initMarkdownEngine()
  .then(() => {
    render();
  })
  .catch((err) => {
    console.error(err);
    preview.innerHTML = '<p style="color:#b91c1c;">预览依赖加载失败，请检查网络或本地依赖文件。</p>';
  });
