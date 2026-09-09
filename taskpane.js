// taskpane.js — WordAI Pro core logic

function initApp() {
  loadSettings();
  bindTabs();
  bindQuickActions();
  bindGenerateButtons();
  bindEditButtons();
  bindToneButtons();
  bindChatButtons();
  bindPromptBox();
  bindOutputActions();
  bindSettings();
  bindResearchButtons();
  bindDocumentButtons();
  bindFileUpload();
  bindDarkMode();
  bindWordCountTracker();
}

if (typeof Office !== 'undefined') {
  Office.onReady(() => initApp());
} else {
  document.addEventListener('DOMContentLoaded', initApp);
}

// ─── Word Count Tracker ──────────────────────────────────────────────────

let wcInterval = null;

function bindWordCountTracker() {
  const goal = parseInt(localStorage.getItem('wordai_wc_goal') || '0');
  if (goal > 0) showWcTracker(goal);

  document.getElementById('btn-wc-settings').addEventListener('click', () => {
    const setup = document.getElementById('wc-setup');
    setup.classList.toggle('hidden');
    if (!setup.classList.contains('hidden'))
      document.getElementById('wc-goal-input').focus();
  });

  document.getElementById('btn-wc-save').addEventListener('click', () => {
    const val = parseInt(document.getElementById('wc-goal-input').value);
    if (!val || val < 1) return;
    localStorage.setItem('wordai_wc_goal', val);
    document.getElementById('wc-setup').classList.add('hidden');
    showWcTracker(val);
  });

  document.getElementById('btn-wc-clear').addEventListener('click', () => {
    localStorage.removeItem('wordai_wc_goal');
    clearInterval(wcInterval);
    document.getElementById('wc-tracker').classList.add('hidden');
    document.getElementById('wc-setup').classList.add('hidden');
    document.getElementById('wc-goal-input').value = '';
  });

  // Show setup if no goal set yet — show a subtle hint button in header
  if (!goal) {
    document.getElementById('wc-setup').classList.remove('hidden');
  }
}

function showWcTracker(goal) {
  document.getElementById('wc-tracker').classList.remove('hidden');
  document.getElementById('wc-goal-display').textContent = goal.toLocaleString();
  updateWcBar(goal);
  clearInterval(wcInterval);
  wcInterval = setInterval(() => updateWcBar(goal), 3000);
}

async function updateWcBar(goal) {
  if (typeof Word === 'undefined') {
    document.getElementById('wc-live').textContent = '— words';
    return;
  }
  try {
    const text = await Word.run(async ctx => {
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      return body.text.trim();
    });
    const count = text ? text.split(/\s+/).filter(Boolean).length : 0;
    const pct = Math.min(100, Math.round((count / goal) * 100));
    const remaining = Math.max(0, goal - count);

    document.getElementById('wc-live').textContent = count.toLocaleString() + ' words';
    const bar = document.getElementById('wc-bar');
    bar.style.width = pct + '%';
    bar.classList.toggle('done', count >= goal);
    document.getElementById('wc-bar-label').textContent =
      count >= goal ? '🎉 Goal reached!' : `${remaining.toLocaleString()} words to go (${pct}%)`;
  } catch (_) {}
}

// ─── Dark Mode ─────────────────────────────────────────────────────────────

function bindDarkMode() {
  const btn = document.getElementById('btn-darkmode');
  const isDark = localStorage.getItem('wordai_dark') === '1';
  if (isDark) document.body.classList.add('dark');
  btn.textContent = isDark ? '☀️' : '🌙';
  btn.addEventListener('click', () => {
    const dark = document.body.classList.toggle('dark');
    localStorage.setItem('wordai_dark', dark ? '1' : '0');
    btn.textContent = dark ? '☀️' : '🌙';
  });
}

// ─── Helpers ────────────────────────────────────────────────────────────────

function showLoading(show) {
  document.getElementById('loading').classList.toggle('hidden', !show);
}

function markdownToHtml(text) {
  return text
    // Tables
    .replace(/^\|(.+)\|$/gm, (row) => '<tr>' + row.slice(1,-1).split('|').map(c => `<td>${c.trim()}</td>`).join('') + '</tr>')
    .replace(/(<tr>.*<\/tr>\n?)+/g, (t) => {
      const rows = t.trim().split('\n');
      const header = rows[0].replace(/<td>/g,'<th>').replace(/<\/td>/g,'</th>');
      const body = rows.slice(2).join('\n');
      return `<table>${header}${body}</table>`;
    })
    // Headings
    .replace(/^### (.+)$/gm, '<h3>$1</h3>')
    .replace(/^## (.+)$/gm, '<h2>$1</h2>')
    .replace(/^# (.+)$/gm, '<h1>$1</h1>')
    // Bold + Italic
    .replace(/\*\*\*(.+?)\*\*\*/g, '<strong><em>$1</em></strong>')
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\*(.+?)\*/g, '<em>$1</em>')
    // Horizontal rule
    .replace(/^---+$/gm, '<hr>')
    // Bullet lists
    .replace(/^[*-] (.+)$/gm, '<li>$1</li>')
    .replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>')
    // Numbered lists
    .replace(/^\d+\. (.+)$/gm, '<li>$1</li>')
    // Line breaks
    .replace(/\n{2,}/g, '</p><p>')
    .replace(/\n/g, '<br>');
}

function showOutput(text) {
  const section = document.getElementById('output-section');
  const el = document.getElementById('output-text');
  el.innerHTML = '<p>' + markdownToHtml(text) + '</p>';
  section.classList.remove('hidden');
}

function hideOutput() {
  document.getElementById('output-section').classList.add('hidden');
}

async function getContext(mode) {
  if (typeof Word === 'undefined') return '';
  return Word.run(async (ctx) => {
    if (mode === 'selection') {
      const sel = ctx.document.getSelection();
      sel.load('text');
      await ctx.sync();
      return sel.text.trim();
    }
    if (mode === 'document') {
      const body = ctx.document.body;
      body.load('text');
      await ctx.sync();
      return body.text.trim();
    }
    return '';
  });
}

async function insertText(rawText, replace = false) {
  if (typeof Word === 'undefined') { showError('Word API not available.'); return; }
  return Word.run(async (ctx) => {
    const sel = ctx.document.getSelection();
    if (replace) sel.insertText('', Word.InsertLocation.replace);
    const insertRange = replace ? sel : sel.getRange(Word.RangeLocation.after);

    // Parse markdown into blocks
    const lines = rawText.split('\n');
    const blocks = [];
    let i = 0;
    while (i < lines.length) {
      const line = lines[i];
      // Table block
      if (/^\|/.test(line)) {
        const tableLines = [];
        while (i < lines.length && /^\|/.test(lines[i])) { tableLines.push(lines[i]); i++; }
        blocks.push({ type: 'table', lines: tableLines });
        continue;
      }
      blocks.push({ type: 'line', text: line });
      i++;
    }

    let cursor = insertRange;

    for (const block of blocks) {
      if (block.type === 'table') {
        const rows = block.lines.filter(l => !/^\|[-:\s|]+$/.test(l));
        const parsed = rows.map(r => r.replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim()));
        const colCount = Math.max(...parsed.map(r => r.length));
        const table = cursor.insertTable(parsed.length, colCount, Word.InsertLocation.after, parsed);
        table.styleBuiltIn = Word.Style.tableGrid;
        // Bold header row
        const headerRow = table.rows.getFirst();
        headerRow.load('cells');
        await ctx.sync();
        headerRow.cells.items.forEach(cell => {
          cell.body.paragraphs.getFirst().font.bold = true;
        });
        cursor = table.getRange(Word.RangeLocation.after);
        await ctx.sync();
        continue;
      }

      const text = block.text;

      // Skip separator lines
      if (/^[-*]{3,}$/.test(text.trim())) {
        const p = cursor.insertParagraph('', Word.InsertLocation.after);
        p.styleBuiltIn = Word.Style.normal;
        p.font.size = 4;
        cursor = p.getRange(Word.RangeLocation.after);
        continue;
      }

      // Empty line
      if (text.trim() === '') {
        cursor = cursor.insertParagraph('', Word.InsertLocation.after).getRange(Word.RangeLocation.after);
        continue;
      }

      // Headings
      const h1 = text.match(/^# (.+)/);
      const h2 = text.match(/^## (.+)/);
      const h3 = text.match(/^### (.+)/);
      if (h1 || h2 || h3) {
        const content = (h1 || h2 || h3)[1].replace(/\*\*/g, '');
        const p = cursor.insertParagraph(content, Word.InsertLocation.after);
        p.styleBuiltIn = h1 ? Word.Style.heading1 : h2 ? Word.Style.heading2 : Word.Style.heading3;
        cursor = p.getRange(Word.RangeLocation.after);
        continue;
      }

      // Bullet list
      const bullet = text.match(/^[*-] (.+)/);
      if (bullet) {
        const p = cursor.insertParagraph(stripInline(bullet[1]), Word.InsertLocation.after);
        p.styleBuiltIn = Word.Style.listParagraph;
        p.listItem.listLevelType = Word.ListLevelType.bullet;
        applyInlineFormats(p, bullet[1]);
        cursor = p.getRange(Word.RangeLocation.after);
        continue;
      }

      // Numbered list
      const numbered = text.match(/^\d+\. (.+)/);
      if (numbered) {
        const p = cursor.insertParagraph(stripInline(numbered[1]), Word.InsertLocation.after);
        p.styleBuiltIn = Word.Style.listParagraph;
        p.listItem.listLevelType = Word.ListLevelType.number;
        applyInlineFormats(p, numbered[1]);
        cursor = p.getRange(Word.RangeLocation.after);
        continue;
      }

      // Normal paragraph
      const p = cursor.insertParagraph(stripInline(text), Word.InsertLocation.after);
      p.styleBuiltIn = Word.Style.normal;
      applyInlineFormats(p, text);
      cursor = p.getRange(Word.RangeLocation.after);
    }

    await ctx.sync();
  });
}

// Strip inline markdown markers for plain text insertion
function stripInline(text) {
  return text
    .replace(/\*\*\*(.+?)\*\*\*/g, '$1')
    .replace(/\*\*(.+?)\*\*/g, '$1')
    .replace(/\*(.+?)\*/g, '$1')
    .replace(/_(.+?)_/g, '$1');
}

// Apply bold/italic to a paragraph using Word search
function applyInlineFormats(paragraph, rawText) {
  // Bold+Italic
  [...rawText.matchAll(/\*\*\*(.+?)\*\*\*/g)].forEach(m => {
    try {
      const r = paragraph.search(m[1], { matchCase: false });
      r.load('items');
      r.items.forEach(item => { item.font.bold = true; item.font.italic = true; });
    } catch (_) {}
  });
  // Bold
  [...rawText.matchAll(/\*\*(.+?)\*\*/g)].forEach(m => {
    try {
      const r = paragraph.search(m[1], { matchCase: false });
      r.load('items');
      r.items.forEach(item => { item.font.bold = true; });
    } catch (_) {}
  });
  // Italic
  [...rawText.matchAll(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g)].forEach(m => {
    try {
      const r = paragraph.search(m[1], { matchCase: false });
      r.load('items');
      r.items.forEach(item => { item.font.italic = true; });
    } catch (_) {}
  });
}

async function runAI(systemPrompt, userContent) {
  const lang = localStorage.getItem('wordai_language') || 'English';
  const messages = [
    { role: 'system', content: `${systemPrompt} Respond in ${lang}.` },
    { role: 'user', content: userContent || '(no content provided)' },
  ];
  const { text } = await AIProvider.call(messages);
  updateTokenDisplay();
  return text;
}

function updateTokenDisplay() {
  const el = document.getElementById('token-count');
  if (el) el.textContent = localStorage.getItem('wordai_tokens') || '0';
  const working = localStorage.getItem('wordai_working_model');
  const model = localStorage.getItem('wordai_model');
  const status = document.getElementById('fetch-status');
  if (status && model === '__auto__' && working) {
    status.textContent = `⚡ Using: ${working}`;
  }
}

function showError(msg) {
  showOutput(`⚠️ ${msg}`);
}

// ─── Tabs ────────────────────────────────────────────────────────────────────

function bindTabs() {
  document.querySelectorAll('.tab').forEach(tab => {
    tab.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
      document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
      tab.classList.add('active');
      document.getElementById(`tab-${tab.dataset.tab}`).classList.add('active');
    });
  });
}

// ─── Document Tab ───────────────────────────────────────────────────────────

function docStatus(msg, isError) {
  const el = document.getElementById('doc-status');
  el.textContent = msg;
  el.className = 'doc-status ' + (isError ? 'doc-status-error' : 'doc-status-ok');
  el.classList.remove('hidden');
  setTimeout(() => el.classList.add('hidden'), 4000);
}

function bindDocumentButtons() {
  // Font family + size apply on change
  document.getElementById('font-family-select').addEventListener('change', () => applyFontFormat());
  document.getElementById('font-size-select').addEventListener('change', () => applyFontFormat());

  // Inline format buttons
  document.querySelectorAll('.fmt-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const fmt = btn.dataset.fmt;
      if (typeof Word === 'undefined') return docStatus('Word API unavailable', true);
      await Word.run(async ctx => {
        const sel = ctx.document.getSelection();
        sel.load('font');
        await ctx.sync();
        if (fmt === 'bold')      sel.font.bold      = !sel.font.bold;
        if (fmt === 'italic')    sel.font.italic    = !sel.font.italic;
        if (fmt === 'underline') sel.font.underline = sel.font.underline === 'None' ? 'Single' : 'None';
        if (fmt === 'strike')    sel.font.strikeThrough = !sel.font.strikeThrough;
        await ctx.sync();
      });
    });
  });

  // All data-action doc buttons
  document.querySelectorAll('.doc-btn[data-action]').forEach(btn => {
    btn.addEventListener('click', () => handleDocAction(btn.dataset.action));
  });
}

async function applyFontFormat() {
  if (typeof Word === 'undefined') return;
  const family = document.getElementById('font-family-select').value;
  const size   = parseInt(document.getElementById('font-size-select').value);
  await Word.run(async ctx => {
    const sel = ctx.document.getSelection();
    sel.font.name = family;
    sel.font.size = size;
    await ctx.sync();
  });
}

async function handleDocAction(action) {
  if (typeof Word === 'undefined') { docStatus('Word API unavailable', true); return; }
  try {
    await Word.run(async ctx => {

      // ─ Word Count
      if (action === 'wordcount') {
        const body = ctx.document.body;
        body.load('text');
        await ctx.sync();
        const text = body.text.trim();
        const words = text ? text.split(/\s+/).length : 0;
        const chars = text.length;
        const paras = text.split(/\n+/).filter(Boolean).length;
        docStatus(`🔢 Words: ${words} | Chars: ${chars} | Paragraphs: ${paras}`);
        return;
      }

      // ─ Properties
      if (action === 'properties') {
        const props = ctx.document.properties;
        props.load('title,author,subject,keywords,lastModifiedBy');
        await ctx.sync();
        docStatus(`📋 Title: ${props.title||'—'} | Author: ${props.author||'—'} | Modified by: ${props.lastModifiedBy||'—'}`);
        return;
      }

      // ─ Outline (headings list)
      if (action === 'outline') {
        const paras = ctx.document.body.paragraphs;
        paras.load('text,styleBuiltIn');
        await ctx.sync();
        const headings = paras.items
          .filter(p => ['Heading1','Heading2','Heading3'].includes(p.styleBuiltIn))
          .map(p => `${p.styleBuiltIn.replace('Heading','H')}: ${p.text.trim()}`)
          .join('\n');
        showOutput(headings || 'No headings found in document.');
        return;
      }

      // ─ Bookmarks
      if (action === 'bookmarks') {
        const bms = ctx.document.bookmarks;
        bms.load('items/name');
        await ctx.sync();
        const names = bms.items.map(b => b.name).join(', ');
        docStatus(`🔖 Bookmarks: ${names || 'None found'}`);
        return;
      }

      // ─ Find
      if (action === 'find') {
        const term = document.getElementById('find-input').value.trim();
        if (!term) { docStatus('Enter text to find', true); return; }
        const results = ctx.document.body.search(term, { matchCase: false, matchWholeWord: false });
        results.load('items/text');
        await ctx.sync();
        results.items.forEach(r => { r.font.highlightColor = 'Yellow'; });
        await ctx.sync();
        docStatus(`🔍 Found ${results.items.length} match(es) — highlighted in yellow`);
        return;
      }

      // ─ Replace All
      if (action === 'replace') {
        const find    = document.getElementById('find-input').value.trim();
        const replace = document.getElementById('replace-input').value;
        if (!find) { docStatus('Enter text to find', true); return; }
        const results = ctx.document.body.search(find, { matchCase: false });
        results.load('items/text');
        await ctx.sync();
        results.items.forEach(r => r.insertText(replace, Word.InsertLocation.replace));
        await ctx.sync();
        docStatus(`🔄 Replaced ${results.items.length} instance(s)`);
        return;
      }

      // ─ Add Comment
      if (action === 'add-comment') {
        const sel = ctx.document.getSelection();
        sel.load('text');
        await ctx.sync();
        if (!sel.text.trim()) { docStatus('Select text to comment on', true); return; }
        const commentText = prompt('Enter comment:');
        if (!commentText) return;
        sel.insertComment(commentText);
        await ctx.sync();
        docStatus('✅ Comment added');
        return;
      }

      // ─ List Comments
      if (action === 'list-comments') {
        const comments = ctx.document.body.getComments();
        comments.load('items/authorName,items/content/text');
        await ctx.sync();
        const list = comments.items.map((c,i) => `${i+1}. [${c.authorName}]: ${c.content.text}`).join('\n');
        showOutput(list || 'No comments found.');
        return;
      }

      // ─ Delete All Comments
      if (action === 'delete-comments') {
        const comments = ctx.document.body.getComments();
        comments.load('items');
        await ctx.sync();
        comments.items.forEach(c => c.delete());
        await ctx.sync();
        docStatus(`🗑️ Deleted ${comments.items.length} comment(s)`);
        return;
      }

      // ─ Resolve All Comments
      if (action === 'resolve-comments') {
        const comments = ctx.document.body.getComments();
        comments.load('items');
        await ctx.sync();
        comments.items.forEach(c => { c.resolved = true; });
        await ctx.sync();
        docStatus(`✅ Resolved ${comments.items.length} comment(s)`);
        return;
      }

      // ─ Apply Styles
      const styleMap = {
        'apply-heading1': Word.Style.heading1,
        'apply-heading2': Word.Style.heading2,
        'apply-heading3': Word.Style.heading3,
        'apply-normal':   Word.Style.normal,
        'apply-quote':    Word.Style.quote,
        'apply-code':     'Code',
      };
      if (styleMap[action]) {
        const sel = ctx.document.getSelection();
        const paras = sel.paragraphs;
        paras.load('items');
        await ctx.sync();
        paras.items.forEach(p => {
          try { p.styleBuiltIn = styleMap[action]; } catch(_) { p.style = styleMap[action]; }
        });
        await ctx.sync();
        docStatus(`🎨 Style applied: ${action.replace('apply-','')}`);
        return;
      }

      // ─ Alignment
      const alignMap = { 'align-left': 'Left', 'align-center': 'Centered', 'align-right': 'Right' };
      if (alignMap[action]) {
        const sel = ctx.document.getSelection();
        const paras = sel.paragraphs;
        paras.load('items');
        await ctx.sync();
        paras.items.forEach(p => { p.alignment = alignMap[action]; });
        await ctx.sync();
        docStatus(`↔ Aligned: ${alignMap[action]}`);
        return;
      }

      // ─ Page Orientation
      if (action === 'page-portrait' || action === 'page-landscape') {
        const sections = ctx.document.sections;
        sections.load('items/body/parentSection');
        await ctx.sync();
        sections.items.forEach(s => {
          s.pageSetup.orientation = action === 'page-landscape'
            ? Word.PageOrientation.landscape
            : Word.PageOrientation.portrait;
        });
        await ctx.sync();
        docStatus(`📄 Page set to ${action === 'page-landscape' ? 'Landscape' : 'Portrait'}`);
        return;
      }

      // ─ Page Break
      if (action === 'insert-pagebreak') {
        const sel = ctx.document.getSelection();
        sel.insertBreak(Word.BreakType.page, Word.InsertLocation.after);
        await ctx.sync();
        docStatus('⏎ Page break inserted');
        return;
      }

      // ─ Insert TOC placeholder
      if (action === 'insert-toc') {
        const sel = ctx.document.getSelection();
        const p = sel.insertParagraph('TABLE OF CONTENTS', Word.InsertLocation.before);
        p.styleBuiltIn = Word.Style.heading1;
        const body = ctx.document.body;
        const paras = body.paragraphs;
        paras.load('items/text,items/styleBuiltIn');
        await ctx.sync();
        const headings = paras.items.filter(p =>
          ['Heading1','Heading2','Heading3'].includes(p.styleBuiltIn)
        );
        let tocText = '';
        headings.forEach(h => {
          const indent = h.styleBuiltIn === 'Heading1' ? '' : h.styleBuiltIn === 'Heading2' ? '  ' : '    ';
          tocText += `${indent}${h.text.trim()}\n`;
        });
        const tocPara = sel.insertParagraph(tocText || '(No headings found)', Word.InsertLocation.before);
        tocPara.styleBuiltIn = Word.Style.normal;
        await ctx.sync();
        docStatus('📚 TOC inserted');
        return;
      }

      // ─ Track Changes
      if (action === 'accept-all') {
        ctx.document.body.getTrackedChanges().load('items');
        await ctx.sync();
        ctx.document.body.getTrackedChanges().items.forEach(c => c.accept());
        await ctx.sync();
        docStatus('✅ All changes accepted');
        return;
      }
      if (action === 'reject-all') {
        ctx.document.body.getTrackedChanges().load('items');
        await ctx.sync();
        ctx.document.body.getTrackedChanges().items.forEach(c => c.reject());
        await ctx.sync();
        docStatus('❌ All changes rejected');
        return;
      }
    });
  } catch(err) {
    docStatus(`⚠️ ${err.message}`, true);
  }
}

// ─── File Upload ─────────────────────────────────────────────────────────────

let uploadedFilesContent = []; // { name, text, isImage, base64, mimeType }

function bindFileUpload() {
  const input = document.getElementById('research-file-input');
  const area  = document.getElementById('upload-area');
  const clearBtn = document.getElementById('btn-clear-uploads');

  input.addEventListener('change', () => processFiles(input.files));

  area.addEventListener('dragover', e => { e.preventDefault(); area.classList.add('drag-over'); });
  area.addEventListener('dragleave', () => area.classList.remove('drag-over'));
  area.addEventListener('drop', e => {
    e.preventDefault();
    area.classList.remove('drag-over');
    processFiles(e.dataTransfer.files);
  });

  clearBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    uploadedFilesContent = [];
    input.value = '';
    document.getElementById('upload-file-list').innerHTML = '';
    document.getElementById('upload-file-list').classList.add('hidden');
    clearBtn.classList.add('hidden');
    document.getElementById('upload-label').querySelector('span').textContent = 'PDF, Word, Image, TXT, CSV, JSON\u2026';
  });
}

async function processFiles(files) {
  if (!files || !files.length) return;
  const list = document.getElementById('upload-file-list');
  const clearBtn = document.getElementById('btn-clear-uploads');
  list.classList.remove('hidden');
  clearBtn.classList.remove('hidden');

  for (const file of files) {
    const tag = document.createElement('span');
    tag.className = 'upload-file-tag';
    tag.textContent = '⏳ ' + file.name;
    list.appendChild(tag);

    try {
      const result = await readFile(file);
      uploadedFilesContent.push({ name: file.name, ...result });
      tag.textContent = '✅ ' + file.name;
    } catch (err) {
      tag.className = 'upload-file-tag error';
      tag.textContent = '❌ ' + file.name;
    }
  }

  const label = document.getElementById('upload-label').querySelector('span');
  label.textContent = `${uploadedFilesContent.length} file(s) loaded`;
}

function readFile(file) {
  return new Promise((resolve, reject) => {
    const name = file.name.toLowerCase();
    const isImage = /\.(png|jpe?g|gif|webp|bmp)$/.test(name);

    if (isImage) {
      const reader = new FileReader();
      reader.onload = e => resolve({ isImage: true, base64: e.target.result, mimeType: file.type, text: '' });
      reader.onerror = reject;
      reader.readAsDataURL(file);
      return;
    }

    if (name.endsWith('.pdf')) {
      const reader = new FileReader();
      reader.onload = async e => {
        try {
          const pdfjsLib = window['pdfjs-dist/build/pdf'];
          pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
          const pdf = await pdfjsLib.getDocument({ data: new Uint8Array(e.target.result) }).promise;
          let text = '';
          for (let i = 1; i <= Math.min(pdf.numPages, 30); i++) {
            const page = await pdf.getPage(i);
            const content = await page.getTextContent();
            text += content.items.map(s => s.str).join(' ') + '\n';
          }
          resolve({ isImage: false, text: text.trim() });
        } catch (err) { reject(err); }
      };
      reader.onerror = reject;
      reader.readAsArrayBuffer(file);
      return;
    }

    // All text-based files (txt, csv, json, xml, md, rtf, docx raw, etc.)
    const reader = new FileReader();
    reader.onload = e => resolve({ isImage: false, text: e.target.result });
    reader.onerror = reject;
    reader.readAsText(file);
  });
}

function getUploadedContext() {
  if (!uploadedFilesContent.length) return '';
  return uploadedFilesContent.map(f => {
    if (f.isImage) return `[Image file: ${f.name} — attached as vision input]`;
    return `--- File: ${f.name} ---\n${f.text.slice(0, 8000)}`;
  }).join('\n\n');
}

function getUploadedImages() {
  return uploadedFilesContent.filter(f => f.isImage);
}

// ─── Research Tab ────────────────────────────────────────────────────────────

function bindResearchButtons() {
  document.querySelectorAll('.res-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const action = btn.dataset.action;
      const topic = document.getElementById('research-topic').value.trim();
      const citStyle = document.getElementById('research-citation-style').value;
      const contextMode = document.getElementById('context-select').value;
      const docText = await getContext(contextMode === 'none' ? 'document' : contextMode);
      const topicLine = topic ? `Research topic: "${topic}"` : '';

      const prompts = {
        planner: `You are an expert academic research advisor. Given the research topic, generate:
1. 5 focused research questions
2. A clear hypothesis
3. Suggested methodology (qualitative/quantitative/mixed)
4. Chapter structure (Introduction, Literature Review, Methodology, Results, Discussion, Conclusion)
5. Key variables or concepts to investigate
Be specific and PhD-level rigorous.`,

        litreview: `You are an expert academic researcher. Analyze the provided content and generate a structured literature review that includes:
1. Major themes and sub-themes
2. Key agreements among scholars
3. Contradictions and debates in the literature
4. Research gaps this work could address
5. Theoretical frameworks relevant to the topic
Use formal academic language. Flag any claim that needs a real citation with [CITATION NEEDED].`,

        evidence: `You are an academic research assistant. From the provided content, build an Evidence Matrix in this exact format for each key claim:
| Claim | Evidence Type | Source Needed | Suggested Search Terms |
Identify at least 5 major claims. Mark unsupported claims with ⚠️. Suggest where to find supporting evidence (Google Scholar, JSTOR, PubMed, etc.).`,

        argument: `You are a critical thinking expert and academic editor. Analyze the provided content and identify:
1. ⚠️ Unsupported claims (assertions without evidence)
2. 🔄 Logical gaps or non-sequiturs
3. ❌ Contradictions within the text
4. 💡 Suggestions to strengthen each weak argument
Be specific, cite the exact sentence or claim that has the issue.`,

        defense: `You are a PhD dissertation committee examiner. Based on the provided content, generate:
1. 10 tough examiner questions the student must be able to answer
2. For each question, explain WHY an examiner would ask it
3. Suggest what a strong answer should cover
4. Identify the 3 most vulnerable areas of the research
This is to help the student prepare, not to intimidate.`,

        abstract: `You are an academic writing expert. Write a structured abstract (250-300 words) following this format:
- Background/Context (1-2 sentences)
- Problem Statement (1-2 sentences)
- Objectives (1-2 sentences)
- Methodology (1-2 sentences)
- Key Findings or Expected Contributions (2-3 sentences)
- Conclusion/Implications (1-2 sentences)
Use formal academic language. Do not fabricate data.`,

        methodology: `You are a research methodology expert. Based on the topic and content, write a detailed Methodology section including:
1. Research Design (justify the choice)
2. Data Collection Methods
3. Sampling Strategy and Sample Size justification
4. Data Analysis Approach
5. Validity and Reliability measures
6. Ethical Considerations
Use formal academic language appropriate for a PhD thesis.`,

        bibliography: `You are an academic citation expert. From the content provided, identify all sources mentioned or implied. For each:
1. Format a proper ${citStyle} citation (use placeholder details if exact info is missing, marked with [VERIFY])
2. Note what claim it supports
3. Flag any source that appears to be missing key details
Also list 5 highly relevant real journals/databases where the student should search for sources on this topic.`,
      };

      const uploadedCtx = getUploadedContext();
      const uploadedImgs = getUploadedImages();
      const docLine = docText ? `\n\nDocument content:\n${docText.slice(0, 4000)}` : '';
      const uploadLine = uploadedCtx ? `\n\nUploaded files:\n${uploadedCtx}` : '';
      const userContent = [topicLine, docLine, uploadLine].filter(Boolean).join('') || 'No topic or document provided.';

      if (uploadedImgs.length) {
        await handleAIActionWithImages(prompts[action], userContent, uploadedImgs);
      } else {
        await handleAIAction(prompts[action], userContent);
      }
    });
  });
}

// ─── Quick Actions ───────────────────────────────────────────────────────────

function bindQuickActions() {
  const prompts = {
    rewrite:   'Rewrite the following text to improve clarity and flow. Return only the rewritten text.',
    summarize: 'Summarize the following text concisely. Return only the summary.',
    improve:   'Improve the writing quality of the following text. Return only the improved text.',
    translate: `Translate the following text to ${localStorage.getItem('wordai_language') || 'Urdu'}. Return only the translation.`,
    expand:    'Expand the following text with more detail and depth. Return only the expanded text.',
    shorten:   'Shorten the following text while keeping the key points. Return only the shortened text.',
  };

  document.querySelectorAll('.quick-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const action = btn.dataset.action;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode === 'none' ? 'selection' : contextMode);
      if (!context) return showError('Please select some text first.');
      await handleAIAction(prompts[action], context);
    });
  });
}

// ─── Generate Buttons ────────────────────────────────────────────────────────

function bindGenerateButtons() {
  document.querySelectorAll('.gen-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode);
      const prompt = btn.dataset.prompt;
      const userContent = context ? `${prompt}:\n\n${context}` : prompt;
      await handleAIAction(prompt, userContent);
    });
  });
}

// ─── Edit Buttons ────────────────────────────────────────────────────────────

function bindEditButtons() {
  const editPrompts = {
    grammar:  'Check the following text for grammar, spelling, and punctuation errors. List each issue with a suggested fix.',
    'fix-all':'Fix all grammar, spelling, punctuation, and clarity issues in the following text. Return only the corrected text.',
    clarity:  'Improve the clarity and readability of the following text. Return only the improved text.',
    passive:  'Rewrite the following text to eliminate passive voice. Return only the rewritten text.',
    simplify: 'Simplify the language in the following text so it is easy to understand. Return only the simplified text.',
    score:    'Analyze the following text and give a writing quality score out of 100. Break down: grammar, clarity, readability, word choice, and structure.',
  };

  document.querySelectorAll('.edit-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const action = btn.dataset.action;
      const contextMode = document.getElementById('context-select').value;
      const context = await getContext(contextMode === 'none' ? 'document' : contextMode);
      if (!context) return showError('No text found. Select text or choose "Entire Document".');
      await handleAIAction(editPrompts[action], context);
    });
  });
}

// ─── Tone Buttons ────────────────────────────────────────────────────────────

function bindToneButtons() {
  document.querySelectorAll('.tone-btn').forEach(btn => {
    btn.addEventListener('click', async () => {
      const tone = btn.dataset.tone;
      const context = await getContext('selection');
      if (!context) return showError('Please select some text first.');
      await handleAIAction(
        `Rewrite the following text in a ${tone} tone. Return only the rewritten text.`,
        context
      );
    });
  });
}

// ─── Chat ────────────────────────────────────────────────────────────────────

const chatHistory = [];

function bindChatButtons() {
  document.querySelectorAll('.suggest-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.getElementById('chat-input').value = btn.dataset.msg;
      sendChat();
    });
  });

  document.getElementById('btn-chat-send').addEventListener('click', sendChat);

  document.getElementById('chat-input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      sendChat();
    }
  });
}

async function sendChat() {
  const input = document.getElementById('chat-input');
  const userMsg = input.value.trim();
  if (!userMsg) return;
  input.value = '';

  appendChatMessage(userMsg, 'user');

  showLoading(true);
  try {
    const docText = await getContext('document');
    const systemPrompt = docText
      ? `You are an AI assistant helping with a Word document. Here is the document content:\n\n${docText.slice(0, 8000)}\n\nAnswer questions about this document helpfully and concisely.`
      : 'You are a helpful AI writing assistant inside Microsoft Word.';

    chatHistory.push({ role: 'user', content: userMsg });

    const lang = localStorage.getItem('wordai_language') || 'English';
    const messages = [
      { role: 'system', content: `${systemPrompt} Respond in ${lang}.` },
      ...chatHistory.slice(-10),
    ];

    const { text } = await AIProvider.call(messages);
    chatHistory.push({ role: 'assistant', content: text });
    appendChatMessage(text, 'ai');
    updateTokenDisplay();
  } catch (err) {
    appendChatMessage(`⚠️ ${err.message}`, 'ai');
  } finally {
    showLoading(false);
  }
}

function appendChatMessage(text, role) {
  const container = document.getElementById('chat-messages');
  const div = document.createElement('div');
  div.className = role === 'user' ? 'msg-user' : 'msg-ai';
  div.textContent = text;
  container.appendChild(div);
  container.scrollTop = container.scrollHeight;
}

// ─── Prompt Box ──────────────────────────────────────────────────────────────

function bindPromptBox() {
  document.getElementById('btn-generate').addEventListener('click', async () => {
    const prompt = document.getElementById('main-prompt').value.trim();
    if (!prompt) return;
    const contextMode = document.getElementById('context-select').value;
    const context = await getContext(contextMode);
    const userContent = context ? `${prompt}\n\n${context}` : prompt;
    await handleAIAction('You are a helpful AI writing assistant inside Microsoft Word.', userContent);
  });
}

// ─── Shared AI Action Handler ────────────────────────────────────────────────

async function handleAIAction(systemPrompt, userContent) {
  showLoading(true);
  hideOutput();
  try {
    const result = await runAI(systemPrompt, userContent);
    showOutput(result);
  } catch (err) {
    showError(err.message);
  } finally {
    showLoading(false);
  }
}

async function handleAIActionWithImages(systemPrompt, userContent, images) {
  showLoading(true);
  hideOutput();
  try {
    const lang = localStorage.getItem('wordai_language') || 'English';
    const messages = [
      { role: 'system', content: `${systemPrompt} Respond in ${lang}.` },
      {
        role: 'user',
        content: [
          { type: 'text', text: userContent || '(no content provided)' },
          ...images.map(img => ({ type: 'image_url', image_url: { url: img.base64 } }))
        ]
      }
    ];
    const { text } = await AIProvider.call(messages);
    updateTokenDisplay();
    showOutput(text);
  } catch (_) {
    // Fallback: vision not supported, send as text-only
    try {
      const result = await runAI(systemPrompt, userContent + '\n[Images attached but vision not supported by current model.]');
      showOutput(result);
    } catch (err) { showError(err.message); }
  } finally {
    showLoading(false);
  }
}

// ─── Output Actions ──────────────────────────────────────────────────────────

function bindOutputActions() {
  document.getElementById('btn-insert').addEventListener('click', async () => {
    const text = document.getElementById('output-text').innerText;
    await insertText(text, false);
  });

  document.getElementById('btn-replace').addEventListener('click', async () => {
    const text = document.getElementById('output-text').innerText;
    await insertText(text, true);
  });

  document.getElementById('btn-copy').addEventListener('click', () => {
    const text = document.getElementById('output-text').innerText;
    navigator.clipboard.writeText(text);
  });

  document.getElementById('btn-discard').addEventListener('click', hideOutput);
}

// ─── Settings ────────────────────────────────────────────────────────────────

function bindSettings() {
  document.getElementById('btn-settings').addEventListener('click', () => {
    document.getElementById('settings-panel').classList.remove('hidden');
  });

  document.getElementById('btn-close-settings').addEventListener('click', () => {
    document.getElementById('settings-panel').classList.add('hidden');
  });

  document.getElementById('btn-fetch-models').addEventListener('click', async () => {
    const provider = document.getElementById('provider-select').value;
    const key = document.getElementById('api-key-input').value.trim();
    const status = document.getElementById('fetch-status');
    if (!key) { status.textContent = '⚠️ Enter your API key first.'; return; }
    const btn = document.getElementById('btn-fetch-models');
    btn.textContent = '⏳ Fetching...';
    btn.disabled = true;
    status.textContent = '';
    const models = await AIProvider.fetchModels(provider, key);
    const select = document.getElementById('model-select');
    select.innerHTML = '';
    if (models.length === 0) {
      select.innerHTML = '<option value="__auto__">🤖 Auto (fallback)</option><option value="">No models found — check your API key</option>';
      status.textContent = '❌ Could not load models.';
    } else {
      select.innerHTML = '<option value="__auto__">🤖 Auto (try best available)</option>';
      models.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m; opt.textContent = m;
        select.appendChild(opt);
      });
      const saved = localStorage.getItem('wordai_model');
      if (saved && [...select.options].some(o => o.value === saved)) {
        select.value = saved;
      }
      status.textContent = `✅ ${models.length} models loaded.`;
    }
    btn.textContent = '🔄 Fetch Models';
    btn.disabled = false;
  });

  document.getElementById('btn-save-settings').addEventListener('click', () => {
    const model = document.getElementById('model-select').value;
    localStorage.setItem('wordai_provider', document.getElementById('provider-select').value);
    localStorage.setItem('wordai_api_key', document.getElementById('api-key-input').value);
    localStorage.setItem('wordai_model', model);
    localStorage.setItem('wordai_language', document.getElementById('language-select').value);
    document.getElementById('settings-panel').classList.add('hidden');
  });

  document.getElementById('btn-reset-tokens').addEventListener('click', () => {
    localStorage.setItem('wordai_tokens', '0');
    updateTokenDisplay();
  });
}

function loadSettings() {
  const provider = localStorage.getItem('wordai_provider') || 'openai';
  const key = localStorage.getItem('wordai_api_key') || '';
  const model = localStorage.getItem('wordai_model') || '';
  const lang = localStorage.getItem('wordai_language') || 'English';

  document.getElementById('provider-select').value = provider;
  document.getElementById('api-key-input').value = key;
  document.getElementById('language-select').value = lang;

  const select = document.getElementById('model-select');
  if (model && model !== '__auto__') {
    select.innerHTML = `<option value="__auto__">🤖 Auto (try best available)</option><option value="${model}">${model}</option>`;
    select.value = model;
  } else {
    select.innerHTML = '<option value="__auto__">🤖 Auto (try best available)</option>';
    select.value = '__auto__';
  }
  updateTokenDisplay();
}
