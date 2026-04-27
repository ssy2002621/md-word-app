const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { execFile } = require('child_process');
const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');

function resolvePandocExecutable() {
  const bundledCandidates = [];

  if (app.isPackaged) {
    bundledCandidates.push(path.join(process.resourcesPath, 'vendor', 'pandoc', 'pandoc.exe'));
  }

  bundledCandidates.push(path.join(__dirname, 'vendor', 'pandoc', 'pandoc.exe'));

  for (const candidate of bundledCandidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return 'pandoc';
}

function getDefaultExportPath() {
  const baseDir = app.isPackaged ? path.dirname(process.execPath) : __dirname;
  return path.join(baseDir, 'output.docx');
}

function normalizeAiMathInput(markdown) {
  if (!markdown) {
    return '';
  }

  let text = markdown
    .replace(/\r\n/g, '\n')
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\u00A0/g, ' ');

  // Convert common AI fenced math snippets into block math.
  text = text.replace(/```(?:latex|tex|math)\s*([\s\S]*?)```/gi, (_m, formulaBody) => {
    const content = String(formulaBody || '').trim();
    if (!content) {
      return '';
    }
    return `\n$$\n${content}\n$$\n`;
  });

  // Support copied forms like \(...\), \\(...\\), \[...\], \\[...\\].
  text = text.replace(/\\+\[([\s\S]*?)\\+\]/g, (_m, inner) => `\n$$\n${String(inner).trim()}\n$$\n`);
  text = text.replace(/\\+\(([^\n]*?)\\+\)/g, (_m, inner) => `$${String(inner).trim()}$`);

  // Heuristic fallback: some copy buttons strip backslashes and leave standalone [ ... ] blocks.
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

function forceHeadingStylesBlack(docxPath) {
  const zip = new AdmZip(docxPath);
  const stylesEntry = zip.getEntry('word/styles.xml');
  if (!stylesEntry) {
    return;
  }

  let stylesXml = stylesEntry.getData().toString('utf8');
  const targetStyleIds = ['Title', 'Heading1', 'Heading2', 'Heading3', 'Heading4', 'Heading5', 'Heading6'];

  for (const styleId of targetStyleIds) {
    const styleRegex = new RegExp(
      `(<w:style[^>]*w:styleId=\"${styleId}\"[\\s\\S]*?<w:rPr>)([\\s\\S]*?)(</w:rPr>)`,
      'g'
    );

    stylesXml = stylesXml.replace(styleRegex, (_match, start, rprContent, end) => {
      // Remove existing color definitions (including themeColor) and enforce pure black.
      const cleaned = rprContent.replace(/<w:color\b[^>]*\/>/g, '');
      return `${start}<w:color w:val=\"000000\"/>${cleaned}${end}`;
    });
  }

  zip.updateFile('word/styles.xml', Buffer.from(stylesXml, 'utf8'));
  zip.writeZip(docxPath);
}

function upsertPropertyInPr(xml, prTag, propertyTag, propertyXml, containerTagRegex) {
  const propertyBlockRegex = new RegExp(`<w:${propertyTag}\\b[^>]*(?:/>|>[\\s\\S]*?</w:${propertyTag}>)`, 'g');

  if (containerTagRegex.test(xml)) {
    return xml.replace(containerTagRegex, (prXml) => {
      const cleaned = prXml.replace(propertyBlockRegex, '');
      return cleaned.replace(`</w:${prTag}>`, `${propertyXml}</w:${prTag}>`);
    });
  }

  return xml.replace(/<w:tbl\b[^>]*>/, (openTag) => `${openTag}<w:${prTag}>${propertyXml}</w:${prTag}>`);
}

function clearCellBorders(tcXml) {
  return setCellBorders(tcXml, {
    top: 'nil',
    left: 'nil',
    right: 'nil',
    bottom: 'nil'
  });
}

function styleHeaderRowCell(tcXml) {
  return setCellBorders(tcXml, {
    top: 'single',
    left: 'nil',
    right: 'nil',
    bottom: 'single'
  });
}

function styleLastRowCell(tcXml) {
  return setCellBorders(tcXml, {
    top: 'nil',
    left: 'nil',
    right: 'nil',
    bottom: 'single'
  });
}

function setCellBorders(tcXml, borders) {
  const tcPrRegex = /<w:tcPr\b[^>]*>[\s\S]*?<\/w:tcPr>/;
  const line = (val) => {
    if (val === 'single') {
      return `w:val=\"single\" w:sz=\"8\" w:space=\"0\" w:color=\"000000\"`;
    }
    return 'w:val=\"nil\"';
  };

  const tcBordersXml =
    '<w:tcBorders>' +
    `<w:top ${line(borders.top)}/>` +
    `<w:left ${line(borders.left)}/>` +
    `<w:right ${line(borders.right)}/>` +
    `<w:bottom ${line(borders.bottom)}/>` +
    '</w:tcBorders>';

  if (tcPrRegex.test(tcXml)) {
    return tcXml.replace(tcPrRegex, (tcPrXml) => {
      const cleaned = tcPrXml.replace(/<w:tcBorders\b[^>]*>[\s\S]*?<\/w:tcBorders>/g, '');
      return cleaned.replace('</w:tcPr>', `${tcBordersXml}</w:tcPr>`);
    });
  }

  return tcXml.replace(/<w:tc\b[^>]*>/, (openTag) => `${openTag}<w:tcPr>${tcBordersXml}</w:tcPr>`);
}

function applyTableStyle(docxPath, tableStyle) {
  const zip = new AdmZip(docxPath);
  const docEntry = zip.getEntry('word/document.xml');
  if (!docEntry) {
    return;
  }

  const normalBorders =
    '<w:tblBorders>' +
    '<w:top w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:left w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:right w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:insideH w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '<w:insideV w:val="single" w:sz="8" w:space="0" w:color="000000"/>' +
    '</w:tblBorders>';

  const minimalBorders =
    '<w:tblBorders>' +
    '<w:top w:val="nil"/>' +
    '<w:left w:val="nil"/>' +
    '<w:bottom w:val="nil"/>' +
    '<w:right w:val="nil"/>' +
    '<w:insideH w:val="nil"/>' +
    '<w:insideV w:val="nil"/>' +
    '</w:tblBorders>';

  let documentXml = docEntry.getData().toString('utf8');
  const tableRegex = /<w:tbl\b[\s\S]*?<\/w:tbl>/g;
  const bordersXml = tableStyle === 'minimal' ? minimalBorders : normalBorders;

  documentXml = documentXml.replace(tableRegex, (tableXml) => {
    let updatedTable = upsertPropertyInPr(
      tableXml,
      'tblPr',
      'tblBorders',
      bordersXml,
      /<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>/
    );

    updatedTable = upsertPropertyInPr(
      updatedTable,
      'tblPr',
      'jc',
      '<w:jc w:val="center"/>',
      /<w:tblPr\b[^>]*>[\s\S]*?<\/w:tblPr>/
    );

    if (tableStyle === 'minimal') {
      updatedTable = updatedTable.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cellXml) => clearCellBorders(cellXml));
      const rowRegex = /<w:tr\b[\s\S]*?<\/w:tr>/g;
      const rows = updatedTable.match(rowRegex) || [];
      let rowIndex = 0;
      const totalRows = rows.length;

      updatedTable = updatedTable.replace(rowRegex, (rowXml) => {
        const currentIndex = rowIndex;
        rowIndex += 1;

        if (currentIndex === 0) {
          return rowXml.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cellXml) => styleHeaderRowCell(cellXml));
        }

        if (totalRows > 1 && currentIndex === totalRows - 1) {
          return rowXml.replace(/<w:tc\b[\s\S]*?<\/w:tc>/g, (cellXml) => styleLastRowCell(cellXml));
        }

        return rowXml;
      });
    }

    return updatedTable;
  });

  zip.updateFile('word/document.xml', Buffer.from(documentXml, 'utf8'));
  zip.writeZip(docxPath);
}

function createWindow() {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false
    }
  });

  win.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

ipcMain.handle('export-docx', async (_event, payload) => {
  const rawMarkdown = typeof payload === 'string' ? payload : payload?.markdown || '';
  const tableStyle = typeof payload === 'string' ? 'normal' : payload?.tableStyle || 'normal';
  const inputPath = path.join(app.getPath('temp'), 'md-word-export-temp.md');

  const { canceled, filePath } = await dialog.showSaveDialog({
    title: '导出 Word',
    defaultPath: getDefaultExportPath(),
    filters: [{ name: 'Word 文档', extensions: ['docx'] }]
  });

  if (canceled || !filePath) {
    return { canceled: true };
  }

  const outputPath = filePath.toLowerCase().endsWith('.docx')
    ? filePath
    : `${filePath}.docx`;

  const markdown = normalizeAiMathInput(rawMarkdown);
  fs.writeFileSync(inputPath, markdown, 'utf8');

  return new Promise((resolve, reject) => {
    const args = [
      inputPath,
      '-f',
      'markdown+tex_math_dollars+tex_math_single_backslash+tex_math_double_backslash',
      '-t',
      'docx',
      '-o',
      outputPath
    ];

    execFile(resolvePandocExecutable(), args, (err, _stdout, stderr) => {
      if (err) {
        reject(new Error(stderr || `${err.message}\n请确认已安装 Pandoc，或使用内置 Pandoc 的完整版安装包。`));
        return;
      }

      try {
        forceHeadingStylesBlack(outputPath);
        applyTableStyle(outputPath, tableStyle);
      } catch (styleErr) {
        reject(new Error(`导出成功，但应用文档样式失败: ${styleErr.message}`));
        return;
      }

      resolve({ canceled: false, outputPath });
    });
  });
});
